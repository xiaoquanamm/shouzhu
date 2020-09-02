using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using System.Data;
using System.Windows.Forms;
using System.ComponentModel;
using System.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections;
using System.Management;
namespace Interop.Office.Core
{

    internal class AllData
    {




        #region//SolidWorks对象－方法

        /// <summary>
        /// 带事件的
        /// </summary>
        internal static SldWorks SwEventPtr = null;
        //是否是已经提示用户打开SolidWorks了
        private static bool balerted = false;
        internal static ISldWorks iSwApp
        {
            get
            {
            Top:
                if (iswapp == null)
                {
                    //提醒启动SolidWorks2次
                    System.Diagnostics.Process[] processArr = null;
                    for (int i = 0; i < 2; i++)
                    {
                        processArr = System.Diagnostics.Process.GetProcessesByName("SLDWORKS");
                        if (processArr.Length == 0)
                        {
                            ShowMyDialogOK sOK = new ShowMyDialogOK("继续操作之前需要手工启动SolidWorks", "请手工打开SolidWorks，待完全打开后点击【确定】按钮继续。");
                            if (sOK.ShowDialog() == DialogResult.Cancel) return null;

                            System.Threading.Thread.Sleep(2000);
                        }
                        else
                        {
                            //System.Diagnostics.Process P = processArr[0];
                            break;
                        }
                    }

                    if (processArr.Length > 0)
                    {
                        try
                        {
                            iswapp = (ISldWorks)System.Runtime.InteropServices.Marshal.GetActiveObject("SldWorks.Application");
                        }
                        catch//(Exception ea)
                        {
                            //StringOperate.Alert("Marshal Error:" + ea.Message);
                        }

                        if (iswapp == null)
                        {
                            for (int i = 18; i < 30; i++)
                            {
                                try
                                {
                                    iswapp = (ISldWorks)System.Runtime.InteropServices.Marshal.GetActiveObject("SldWorks.Application." + i.ToString());
                                    if (iswapp != null)
                                    {
                                        break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    //MessageBox.Show(ex.Message);
                                }
                            }
                        }

                        if (iswapp == null)
                        {
                            try
                            {
                                iswapp = (ISldWorks)System.Runtime.InteropServices.Marshal.GetActiveObject("SldWorks.ISldWorks");
                            }
                            catch //(Exception ea)
                            {
                               // StringOperate.Alert("Marshal Error:" + ea.Message);
                            }
                        }
                    }
                    if(iswapp ==null)//如果没有打开SolidWorks
                    {
                        try
                        {
                            iswapp = new SldWorksClass();
                        }
                        catch { }
                        if (iswapp == null)
                        {
                            try
                            {
                                iswapp = new SldWorks();
                            }
                            catch { }
                        }
                    }

                    //最后如果还是没有成功,并且还没有提示,提示用户打开SolidWorks
                    if (iswapp == null && !balerted && processArr.Length == 0)
                    {
                        StringOperate.Alert("打开SolidWorks失败！请首先打开SolidWorks！"); balerted = true;
                    }
                }

                if (iswapp != null)
                {
                    try
                    {
                        iswapp.Visible = true;
                    }
                    catch// (Exception ea)
                    {
                        iswapp = null;
                        goto Top;
                    }

                }

                return iswapp;
            }
            set
            {
                iswapp = value;
                iswapp.Visible = true;
            }
        }
        internal static ISldWorks iswapp = null;


        /// <summary>
        /// 版本号：15－2007，16－2008，17－2009，18－2010，19－2011
        /// </summary>
        internal static int iSWVersion
        {
            get
            {
                if (zzziversion == -1)
                {
                    string strVersion = iSwApp.RevisionNumber();
                    strVersion = strVersion.Remove(strVersion.IndexOf('.'));
                    zzziversion = Convert.ToInt16(strVersion);
                }
                return zzziversion;
            }
        }
        private static int zzziversion = -1;



        /// <summary>
        /// SolidWorks语言（简体中文1--chinese-simplified）， （英文2-english）， （繁体中文3--chinese）
        /// </summary>
        internal static int iSWLanguage
        {
            get
            {
                if (iswlanguagename == -1)
                {
                    string s = AllData.iSwApp.GetCurrentLanguage();
                    bool b1 = AllData.iSwApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swUseEnglishLanguage);
                    bool b2 = AllData.iSwApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swUseEnglishLanguageFeatureNames);

                    if (s == "english")
                    {
                        iswlanguagename = 2;
                    }
                    else if (s == "chinese")
                    {
                        iswlanguagename = 3;
                        if (b1 || b2) iswlanguagename = 2;
                    }
                    else if (s == "chinese-simplified")
                    {
                        iswlanguagename = 1;
                        if (b1 || b2) iswlanguagename = 2;
                    }
                    else
                    {
                        if (b1 || b2) iswlanguagename = 2;
                        else
                        {
                            iswlanguagename = 1;  //简体中文
                        }
                    }
                }
                return iswlanguagename;
            }
        }
        private static int iswlanguagename = -1;//简体中文1chinese-simplified， 英文2-english， 繁体中文3--chinese



        /// <summary>
        /// 在SolidWorks中新建一个零件，并返回
        /// </summary>
        /// <param name="bCheckRefPlaneChineseName">是否判断特征的名称是中文 如:前视基准面</param>
        /// <param name="bAlertForm">模板不是中文时是否提醒</param>
        /// <param name="templatePath">指定模板路径或空</param>
        internal static IModelDoc2 NewDoc(bool bCheckRefPlaneChineseName, bool bAlertForm,string templatePath)
        {
            string strDefault = "";
            if (templatePath.Trim().Length > 0)
            {
                strDefault = templatePath;
            }
            else
            {
                strDefault = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            }
            
            IModelDoc2 modDoc = null;
            try
            {
                if(File.Exists(strDefault))
                {
                    //modDoc = (IModelDoc2)AllData.iSwApp.NewDocument(strDefault, 0, 0.0, 0.0);
                    modDoc = (IModelDoc2)AllData.iSwApp.INewDocument2(strDefault, 0, 0.0d, 0.0d);
                }
                else
                {
                    modDoc = (IModelDoc2)AllData.iSwApp.NewPart();
                }
            }
            catch (Exception ea)
            {
                StringOperate.Alert("创建空零件出错：" + ea.Message);
            }
            string defaultTempPath = AllData.StartUpPath + "\\迈迪模板\\迈迪标准模版\\迈迪零件模板.PRTDOT";

        

            if (modDoc == null)
            {
                if (File.Exists(defaultTempPath))
                {
                    modDoc = (IModelDoc2)AllData.iSwApp.INewDocument2(defaultTempPath, 0, 0.0d, 0.0d);
                }
                else
                {
                    modDoc = (IModelDoc2)AllData.iSwApp.NewPart();
                }
            }
            if (modDoc == null)
            {
                StringOperate.Alert("新建SolidWorks零件失败！请检测模板文件是否存在！");
            }
            //检查是否是中文
            if (bCheckRefPlaneChineseName)
            {
                if (modDoc != null)
                {
                    SolidWorksMethod swm = new SolidWorksMethod();
                    Interop.Office.Core.SolidWorksMethod.swRefPlaneNames swRef = swm.GetRefPlaneNames(modDoc);
                    if(swRef.FrontPlaneName!= "前视基准面")
                    {
                        if (File.Exists(defaultTempPath))
                        {
                            //在这里有可能创建一个空的零件。
                            modDoc = (IModelDoc2)AllData.iSwApp.INewDocument2(defaultTempPath, 0, 0.0d, 0.0d);
                        }
                        else if (bAlertForm)
                        {
                            StringOperate.Alert("检测到您的SolidWorks默认模板是英文特征名称,有可能造成生成零件错误!建议更换中文默认模板!");
                        }
                    }
                }
            }

            //2014-12-27添加，检查用户是不是在草图设置中选中了“打开零件时直接打开草图”
            bool b = AllData.iSwApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchCreateSketchOnNewPart);
            if (b)
            {
                string strA="您在SolidWorks工具》选项》草图 中选中了【打开新零件时直接打开草图】";
                strA += System.Environment.NewLine;
                strA += "这有可能导致生成模型出错！要关闭这个选项吗？";
                if (StringOperate.Alert(strA, "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    AllData.iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchCreateSketchOnNewPart, false);
                }
            }

            return modDoc;
        }


 

        #endregion//---



        #region//路径属性


        /// <summary>
        /// 程序的运行目录最后没有"\"
        /// </summary>
        internal static string StartUpPath
        {
            get
            {
                //StringOperate.Alert("Assembly.GetExecutingAssembly().Location=="+Assembly.GetExecutingAssembly().Location + "  system.windows.forms.startuppath=" + System.Windows.Forms.Application.StartupPath );
                //return System.Windows.Forms.Application.StartupPath.ToString();
                return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            }
        }
        /// <summary>
        /// 王逸心程序专用路径,StartUpPath + "\\WYX\\";
        /// </summary>
        internal static string WYXPath
        {
            get
            {
                return StartUpPath + "\\WYX\\";
            }
        }
        /// <summary>
        /// 帮助文件的路径
        /// </summary>
        internal static string strHelpFile
        {
            get
            {
                return StartUpPath + "\\MD3DtoolsHelp.chm";
            }
        }
        /// <summary>
        /// 用于拖放过程中的图标
        /// </summary>
        internal static Cursor CurDrog
        {
            get
            {
                if (zzcurdrog == null)
                {
                    zzcurdrog = new Cursor(AllData.StartUpPath + "\\IMG\\move.cur");
                }
                return zzcurdrog;
            }
        }
        internal static Cursor zzcurdrog;



        /// <summary>
        /// 设置零件的材质,成功返回True（swModel和swPart二者指定其一）
        /// </summary>
        /// <param name="DBFullName">材料库路径，可以不指定</param>
        /// <param name="cfgName">零件配置名称</param>
        /// <param name="strValue">材质名称</param>
        internal static bool setPartMaterial(IModelDoc2 swModel, PartDoc swPart, string DBFullName, string cfgName, string strValue)
        {
            if (swPart == null)
            {
                swPart = (PartDoc)swModel;
            }

            //先设定默认材料库
            if (DBFullName.Length == 0)
            {
                swPart.SetMaterialPropertyName2(cfgName, null, strValue);
            }
            else
            {
                swPart.SetMaterialPropertyName2(cfgName, DBFullName, strValue);
            }

            string strdb = "";
            string strMt = swPart.GetMaterialPropertyName2(cfgName, out strdb);
            if (strMt == strValue) return true;

            //再重新指定
            foreach (object o in AllMaterialDBs)
            {
                try
                {
                    swPart.SetMaterialPropertyName2(cfgName, o.ToString(), strValue);
                    strMt = swPart.GetMaterialPropertyName2(cfgName, out strdb);
                    if (strMt == strValue) return true;
                }
                catch { }
            }
            return false;
        }
        /// <summary>
        /// 得到所有的材料文件全名称
        /// </summary>
        internal static ArrayList AllMaterialDBs
        {
            get
            {
                if (zzallmatel == null)
                {
                    zzallmatel = new ArrayList();

                    //先获取迈迪的材料库
                    DirectoryInfo dir = new DirectoryInfo(AllData.StartUpPath + "\\迈迪材质");
                    if (dir.Exists)
                    {
                        FileInfo[] fiArr = dir.GetFiles("*.sldmat");
                        foreach (FileInfo fi in fiArr)
                        {
                            if (zzallmatel.Contains(fi.FullName)) continue;
                            if (fi.Name.IndexOf("赛迪") != -1)//赛迪的放在最前面，第一个引用
                            {
                                zzallmatel.Insert(0, fi.FullName);
                            }
                            else //07的不再做判断
                            {
                                zzallmatel.Add(fi.FullName);
                            }
                        }
                    }
                    //然后再获取SolidWorks的材料库
                    string txt = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsMaterialDatabases);
                    string[] strArr = txt.Split(new char[] { ';' });
                    foreach (string s in strArr)
                    {
                        //迈迪的材质因为上面加入了，这里就不再加入了。
                        if (s.Length == 0 || !Directory.Exists(s) || s.IndexOf("迈迪") != -1) continue;
                        DirectoryInfo info = new DirectoryInfo(s);
                        foreach (FileInfo fi in info.GetFiles("*.sldmat"))
                        {
                            if (!zzallmatel.Contains(fi.FullName)) zzallmatel.Add(fi.FullName);
                        }
                    }

                }
                return zzallmatel;
            }
        }
        private static ArrayList zzallmatel = null;



        /// <summary>
        /// 得到新的数据库链接对象DBMethod
        /// </summary>
        internal static DBMethod getDBMethod(dbnames i)
        {
            //如果是Access数据库，用Dbtool，从6.0开始已经不用Sql数据库了。
            DBMethod db;

            string Connstr = "";
            if (i == dbnames.gb_user)
            {
                Connstr = GetConnStr("\\prtdot_user\\gb_user.mdb", "");
            }
            else if (i == dbnames.gb)
            {
                Connstr = GetConnStr("\\Database\\gb.mdb", "jnmd#tangkai#2008");
            }
            else if (i == dbnames.property)
            {
                Connstr = GetConnStr("\\Database\\property.mdb", "");
            }
            else if (i == dbnames.belt)
            {
                Connstr = GetConnStr("\\WYX\\Belt\\belt.mdb", "123&abc");
            }
            else if (i == dbnames.chain)
            {
                Connstr = GetConnStr("\\WYX\\chain\\chain.mdb", "abc123321");
            }
            else if (i == dbnames.WYX)
            {
                Connstr = GetConnStr("\\Database\\WYX.mdb", "");
            }
            else if (i == dbnames.gear)
            {
                Connstr = GetConnStr("\\Database\\gear.mdb", "");
            }
            else if (i == dbnames.rijiyuelei)
            {
                Connstr = GetConnStr("\\日积月累DB\\rijiyuelei.mdb", "tangkai&rijiyuelei");
            }
            else if (i == dbnames.symbol)
            {
                Connstr = GetConnStr("\\symbol\\symbol.mdb", "abc@123");
            }
            else if (i == dbnames.bend)
            {
                Connstr = GetConnStr("\\WYX\\Bend\\bend.mdb", "1876405");
            }
            else if (i == dbnames.DKG)
            {
                Connstr = GetConnStr("\\WYX\\DianKongGui\\cabinet.mdb", "1876405");
            }

            if ( Sys.Is64BitOPSystem == false)
               
            {
                db = new Dbtool(Connstr);
            }
            else
            {
                db = new DbClient(Connstr);
            }
            return db;

        }

   

        private static string GetConnStr(string dbname, string password)
        {
            //如果是Access版本，把上面的注释去掉
            string ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="+ StartUpPath;
            if (password == "")
            {
                return ConnStr + dbname ;
            }
            else
            {
                return ConnStr + dbname + ";Jet OLEDB:database Password=" + password + ";Persist Security Info=True";
            }
        }
        private static string connstrsql = "";
        private static string connstrsql2 = "";


        #endregion//path info

    }

    //所有数据库的名称
    internal enum dbnames
    {
        gb = 1,
        property = 2,
        belt = 3,
        chain = 4,

        gb_user = 5,
        gear = 6,
        WYX = 7,

        rijiyuelei=8,
        symbol=9,

        bend = 10,
        DKG=11
    }
}