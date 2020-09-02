using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using System.IO;
using System.Collections;
using SolidWorks.Interop.swconst;


namespace Interop.Office.Core
{
    internal partial class PSetUp : Form
    {
        internal PSetUp()
        {
            InitializeComponent(); 
            this.helpProvider1.HelpNamespace = AllData.strHelpFile;
        }
        internal PSetUp(int iPageIndex)
        {
            InitializeComponent(); 
            this.helpProvider1.HelpNamespace = AllData.strHelpFile;
            this.tabControl1.SelectedIndex = iPageIndex;
        }
        private void PSetUp_Load(object sender, EventArgs e)
        {
            //面板0
            this.chkOnePartManyConf.Checked = Interop.Office.Core.Properties.Settings.Default.bOnePartManyConf;//同一零件不同配置
            this.txtWorkPath.Text = Interop.Office.Core.Properties.Settings.Default.strWorkPath;//工作目录
            this.chkUseProxy.Checked = (PSetUp.WebProxy != null);//是否使用代理上网

            this.rdoBig5.Checked = PSetUp.bIsBig5Encoding;
            this.rdoGB2312.Checked = !PSetUp.bIsBig5Encoding;

            //面板1
            this.chkThreadMaps.Checked = Interop.Office.Core.Properties.Settings.Default.frmAll_bThreadMaps;//是否显示螺纹贴图
            this.chkShowStartForm.Checked = Interop.Office.Core.Properties.Settings.Default.bShowStartForm;//加载插件时显示启动画面
            this.chkFrmAllUseCache.Checked = Interop.Office.Core.Properties.Settings.Default.frmAll_bUseCache;//缓存标准件树视图
            this.chkDelFormula.Checked = Interop.Office.Core.Properties.Settings.Default.frmAll_bDelFormula;//删除方程式


            //网络版设置--------------------------
            this.cmbServerIP.Text = fntWeb_ServerIP;
            this.cmbServerPort.Text = fntWeb_ServerPort.ToString();
            this.txtFntWebPath.Text = MD3DParts;//配件网的跟目录


            //管理员设置
            if (bIsManager)
            {
                this.txtMgrRegCode.Text = "您是管理员";
                this.groupRegMgr.Enabled = false;
                this.groupMgr.Enabled = true;
            }
            this.chkCanPart1.Checked = bCanParts1;
            this.chkCanPart2.Checked = bCanParts2;
            this.chkCanPart3.Checked = bCanParts3;

            this.cmbNameStyle.Text = strNameStyle;//5.1新的命名规则
            this.cmbNameStyleWeb.Text = strNameStyleWeb;

            this.chkSetUniteIsKG.Checked = bSetUniteIsKG;
            this.nmrcDecimals.Value = iDecimals;

            this.txtGBBasePath.Text = MDGBParts;
            this.chkUseGBParts.Checked = bUseGBPath;
        }
        private void btnok_Click(object sender, EventArgs e)
        {
            //0
            Interop.Office.Core.Properties.Settings.Default.bOnePartManyConf = this.chkOnePartManyConf.Checked;

            //1
            Interop.Office.Core.Properties.Settings.Default.frmAll_bThreadMaps = this.chkThreadMaps.Checked;
            Interop.Office.Core.Properties.Settings.Default.bShowStartForm = this.chkShowStartForm.Checked;
            Interop.Office.Core.Properties.Settings.Default.frmAll_bUseCache=this.chkFrmAllUseCache.Checked ;//缓存标准件树视图
            Interop.Office.Core.Properties.Settings.Default.frmAll_bDelFormula = this.chkDelFormula.Checked;


            strNameStyle = this.cmbNameStyle.Text.TrimEnd();
            strNameStyleWeb = this.cmbNameStyleWeb.Text.TrimEnd();


            //是否使用代理上网
            PSetUp.WebProxy = (this.chkUseProxy.Checked ? System.Net.WebRequest.GetSystemWebProxy() : null);

            //是否使用繁体中文上网
            PSetUp.bIsBig5Encoding = (this.rdoBig5.Checked);


            //最后保存
            Interop.Office.Core.Properties.Settings.Default.Save();

            PSetUp.bCanParts1 = this.chkCanPart1.Checked;
            PSetUp.bCanParts2 = this.chkCanPart2.Checked;
            PSetUp.bCanParts3 = this.chkCanPart3.Checked;

            bSetUniteIsKG = this.chkSetUniteIsKG.Checked;
            iDecimals = Convert.ToInt16(this.nmrcDecimals.Value);

            iusergbpath = -1;//如果不设置这一项，下次重启程序时设置才生效
            bUseGBPath = this.chkUseGBParts.Checked;

            mdgbparts = MDGBParts = this.txtGBBasePath.Text;
            

            this.Close();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private static cfgInfo cinfo = new cfgInfo();


        /// <summary>
        /// 启用配置模式时, 指定工作路径，每次打开SolidWorks都要重新指定
        /// </summary>
        internal static string strWorkPath
        {
            get
            {
                while(Properties.Settings.Default.strWorkPath.Length ==0)
                {
                    FolderBrowserDialog fd = new FolderBrowserDialog();
                    fd.ShowNewFolderButton = true;
                    fd.Description = "设置工作目录！";
                    if (fd.ShowDialog() == DialogResult.OK)
                    {
                        Interop.Office.Core.Properties.Settings.Default.strWorkPath = fd.SelectedPath;
                        Interop.Office.Core.Properties.Settings.Default.Save();
                    }
                }
                return Interop.Office.Core.Properties.Settings.Default.strWorkPath;
            }
        }
        /// <summary>
        /// 指定模型库的路径,全局性的,后面有\\ ,如果数据库中没有,就返回"....\\MD3DParts所在的目录"
        /// 在单独的EXE程序MDPKGS.exe中也用到这个
        /// </summary>
        internal static string MD3DParts
        {
            get
            {
                string str = cinfo.getValue("strMD3DPartsPath");
                if (str.Length == 0 || !Directory.Exists(str))
                {
                    str = AllData.StartUpPath + "\\MD3DParts\\";
                }
                else
                {
                    if (!str.EndsWith("\\")) str = str + "\\";
                }
                return str;
            }
            set
            {
                cinfo.setValue("strMD3DPartsPath", value);
            }
        }


        /// <summary>
        /// 作为单机版使用StandAlone，作为客户端使用Client，作为服务器使用Server
        /// </summary>
        internal static string fntWeb_UsingType
        {
            get
            {
                return cinfo.getValue(SetupNames.UsingType.ToString());
            }
            set
            {
                cinfo.setValue(SetupNames.UsingType.ToString(), value);
            }
        }
        /// <summary>
        /// 是否作为内网客户端使用
        /// </summary>
        internal static bool bUsingType_Client
        {
            get
            {
                return (fntWeb_UsingType == "Client");
            }
        }
        /// <summary>
        /// 与服务器通信的Ip地址,这里必需写入注册表,和UpdateLAN.exe共享
        /// </summary>
        internal static string fntWeb_ServerIP
        {
            get
            {
                return cinfo.getValue(SetupNames.strServerIP.ToString());
            }
            set
            {
                cinfo.setValue(SetupNames.strServerIP.ToString(), value);
            }
        }
        /// <summary>
        /// 与服务器通信的端口号,默认是3548
        /// </summary>
        internal static int fntWeb_ServerPort
        {
            get
            {
                int a = 3548;
                try
                {
                    a = Convert.ToInt16(cinfo.getValue(SetupNames.strServerPort.ToString()));
                }
                catch (Exception e)
                {
                    StringOperate.Alert("读取端口错误：" + e.Message + ";使用默认端口3548。");
                }
                return a;
            }
            set
            {
                try
                {
                    cinfo.setValue(SetupNames.strServerPort.ToString(), value.ToString());
                }
                catch (Exception e)
                {
                    StringOperate.Alert("设置端口错误：" + e.Message);
                }
            }
        }


        /// <summary>
        /// 是否是迈迪
        /// </summary>
        internal static bool bIsMD
        {
            get
            {
                FrmMgrWindow fmgr = new FrmMgrWindow();
                string mdcode =fmgr.getMDCode(GetMgrCode());

                //mdcode = fmgr.getMDCode("636305154856");


                //首先查找新的存放位置
                string mVaule = cinfo.getValue("mdcode", "regcode.cfg");
                if (mVaule.IndexOf(mdcode) != -1) return true;

                //再查找原先的
                string scode = cinfo.getValue("mdcode");
                if (scode.IndexOf(mdcode) != -1) return true;

                //最后返回否
                return false;
            }
        }
        /// <summary>
        /// 是否是管理员,可以存储多个管理员注册码，中间用“，”隔开
        /// </summary>
        internal static bool bIsManager
        {
            get
            {
                string mgrcode = GetMgrCode();

                //首先查找新的存放位置
                string mVaule = cinfo.getValue(Dbtool2.strError.ToString(), "regcode.cfg");
                if (mVaule.IndexOf(mgrcode) != -1) return true;

                //再查找原先的
                string scode = cinfo.getValue("mgrcode");
                if (scode.IndexOf(mgrcode) != -1) return true;

                //最后返回否
                return false;
            }
            set
            {
                string mCode = Dbtool2.strError.ToString();
                string mVaule = cinfo.getValue(mCode, "regcode.cfg");
                string txt = GetMgrCode();
                if (value)
                {
                    if (mVaule.IndexOf(txt) == -1)
                    {
                        mVaule = mVaule + "," + txt;
                    }
                }
                else
                {
                    mVaule = mVaule.Replace(txt, "").Replace(",,", "").Replace("，，", "");
                }
                cinfo.setValue(mCode, mVaule, "regcode.cfg");
            }
        }
        private static string GetMgrCode()
        {
            char[] cArr = Dbtool2.strError.ToString().ToCharArray();
            byte[] bArr = new byte[cArr.Length];    //本机码 byte[]数组
            for (int i = 0; i < cArr.Length; i++)
            {
                bArr[i] = (byte)cArr[i];
            }

            for (int i = 0; i < bArr.Length; i++)
            {
                byte bi = bArr[i];
                if (bi > 47 && bi < 53)//0-4
                {
                    bi = (byte)(bi + 5);
                }
                else if (bi > 52)//5--9
                {
                    bi = (byte)(bi - 5);
                }
                if (bi > 64 && bi < 96)//大写
                {
                    bi = (byte)(bi - 35);
                }
                else if (bi > 96)//小写
                {
                    bi = (byte)(bi - 32);
                }
                bArr[i] = bi;//付给它新值，
            }

            //把ByteToStr(byte[] btArr, int len)移动进来了。
            byte[] btArr=bArr ;
            int len = 16;
            string strResult = "";
            for (int i = 0; i < btArr.Length; i++)
            {
                strResult += btArr[i].ToString();
            }
            if (strResult.Length > len + 4)//长度大于16位
            {
                strResult = strResult.Substring(3, len);
            }
            else if (strResult.Length < len)//长度小于16位
            {
                bool first = true;
                Random r = new Random();
                while (strResult.Length < len)
                {
                    if (first) strResult = strResult + r.Next(0, 9).ToString();
                    else strResult = r.Next(0, 9).ToString() + strResult;
                    first = !first;
                }
            }
            else //16--20
            {
                strResult = strResult.Substring(0, len);
            }

            //返回计算的结果
            string strRet = strResult;
            int n = Convert.ToInt16(strRet.Substring(0, 1)) + 1;
            strRet = strRet.Substring(0, n) + "1" + strRet.Substring(n + 1);
            strRet = strRet.Substring(1, 4) + strRet.Substring(6, 4) + strRet.Substring(11, 4);
            return strRet;
        }


        /// <summary>
        /// 是否可以使用本机库,默认是
        /// </summary>
        internal static bool bCanParts1
        {
            get
            {
                string s1 = cinfo.getValue("bEnable1");
                return (s1.ToString().ToLower() != "false");
            }
            set
            {
                cinfo.setValue("bEnable1", value.ToString());
            }
        }
        /// <summary>
        /// 是否可以使用内网库
        /// </summary>
        internal static bool bCanParts2
        {
            get
            {
                string s1 = cinfo.getValue("bEnable2");
                return (s1.ToString().ToLower() != "false");
            }
            set
            {
                cinfo.setValue("bEnable2", value.ToString());
            }
        }
        /// <summary>
        /// 是否可以使用在线库
        /// </summary>
        internal static bool bCanParts3
        {
            get
            {
                string s1 = cinfo.getValue("bEnalbe3");
                return (s1.ToString().ToLower() != "false");
            }
            set
            {
                cinfo.setValue("bEnalbe3", value.ToString());
            }
        }
        /// <summary>
        /// 配件库中的表,第一列是代号还是规格，默认是代号
        /// </summary>
        internal static string FirstColName
        {
            get
            {
                if (firstcolname.Length == 0)
                {
                    firstcolname = cinfo.getValue("FirstColName");
                }
                if (firstcolname.Length == 0)
                {
                    firstcolname = "代号";
                }
                return firstcolname;
            }
            set
            {
                cinfo.setValue("FirstColName", value);
            }
        }
        private static string firstcolname = "";



        /// <summary>
        /// 是否使用浏览器代理上网，建议不使用
        /// </summary>
        internal static System.Net.IWebProxy WebProxy
        {
            get
            {
                if (useWebProxy == "-1")
                {
                    useWebProxy = cinfo.getValue("useWebProxy");
                    if (useWebProxy == "1")
                    {
                        PSetUp.useWebProxyAddress = cinfo.getValue("useWebProxyAddress");
                    }
                }

                if (useWebProxy=="1")
                {
                    //如果用户指定了代理服务器IP地址和端口号，这种情况很少见，需要手动打开记事本中去写
                    if (PSetUp.useWebProxyAddress.Length > 10)
                    {
                        System.Net.WebProxy wp = new System.Net.WebProxy(PSetUp.useWebProxyAddress, true);
                        wp.Credentials = System.Net.CredentialCache.DefaultCredentials;
                        return wp;
                    }
                    else
                    {
                        return System.Net.WebRequest.GetSystemWebProxy();
                    }
                }
                else
                {
                    //默认代理是开启的,故只有等待超时后才会绕过代理,这就阻塞了.
                    //如果不加上这一句，会很慢，甚至造成链接超时。
                    return null;
                }
            }
            set
            {
                if (value == null)
                {
                    //不使用代理
                    useWebProxy = "0";
                    cinfo.setValue("useWebProxy", "0");
                }
                else
                {
                    //使用代理
                    useWebProxy = "1";
                    cinfo.setValue("useWebProxy", "1");
                }
            }
        }
        private static string useWebProxy ="-1";//-1：还没读，1:使用代理上网
        private static string useWebProxyAddress = "";//如：“192.168.0.1:8001”



        /// <summary>
        /// 是否使用繁体中文版本
        /// </summary>
        internal static bool bIsBig5Encoding
        {
            get
            {
                if (sCharacter == "-1")
                {
                    sCharacter = cinfo.getValue("Character");
                }

                return (sCharacter == "2");
            }
            set
            {
                if (value == true)
                {
                    cinfo.setValue("Character", "2");
                }
                else
                {
                    cinfo.setValue("Character", "1");
                }
                sCharacter = cinfo.getValue("Character");
            }
        }
        private static string sCharacter = "-1";//-1：还没读，1:简体，2：繁体;3:英文



        /// <summary>
        /// 标准件重量单位显示几位小数,默认3
        /// </summary>
        internal static int iDecimals
        {
            get
            {
                if (idecimals == -1)
                {
                    try
                    {
                        string str = cinfo.getValue("Decimals");
                        idecimals = Convert.ToInt16(str);
                    }
                    catch
                    {
                        idecimals = 3;
                    }
                }
                return idecimals;
            }
            set
            {
                value = Math.Max(0, value);
                value = Math.Min(value, 4);
                cinfo.setValue("Decimals", value.ToString());
            }
        }
        private static int idecimals = -1;
        /// <summary>
        /// 重量单位默认是千克,默认true
        /// </summary>
        internal static bool bIsKg
        {
            get
            {
                if (biskg == -1)
                {
                    try
                    {
                        bool b = Convert.ToBoolean(cinfo.getValue("bIsKg"));
                        if (b) biskg = 1;
                        else biskg = 0;
                    }
                    catch
                    {
                        biskg = 1;
                    }
                }
                return (biskg == 1);
            }
        }
        private static int biskg = -1;
        /// <summary>
        /// 生成零件后设置文件的重量单位为KG,默认True
        /// </summary>
        internal static bool bSetUniteIsKG
        {
            get
            {
                if (bsetuniteiskg == "")
                {
                    try
                    {
                        bsetuniteiskg =cinfo.getValue("bSetUniteIsKG").ToLower();
                    }
                    catch
                    {
                        bsetuniteiskg = "true";
                    }
                }
                return (bsetuniteiskg !="false");
            }
            set
            {
                cinfo.setValue("bSetUniteIsKG", value.ToString());
            }
        }
        private static string  bsetuniteiskg="";
        /// <summary>
        /// 是否重写属性,默认全部重写，1:全部重写，2只重写非空属性 ，3重写非空且非默认值的属性 4全部不重写
        /// </summary>
        internal static int ReWriteAttType
        {
            get
            {
                if (rewriteattthype == -1)
                {
                    try
                    {
                        rewriteattthype = Convert.ToInt16(cinfo.getValue("ReWriteAttType"));
                    }
                    catch
                    {
                        rewriteattthype = 1;
                    }
                }
                return rewriteattthype;
            }
        }
        private static int rewriteattthype = -1;

        /// <summary>
        /// 是否使用标准件统一存放目录,默认为False
        /// </summary>
        internal static bool bUseGBPath
        {
            get
            {
                if (iusergbpath ==-1)
                {
                    try
                    {
                        bool b = Convert.ToBoolean(cinfo.getValue("bUseGBPath"));
                        if (b) iusergbpath = 1;
                        else iusergbpath = 0;
                    }
                    catch
                    {
                        iusergbpath =0;
                    }
                }
                return (iusergbpath == 1);
            }
            set
            {
                cinfo.setValue("bUseGBPath", value.ToString());
            }
        }
        private static int iusergbpath = -1;//-1:还没读，0：false 1:true
        /// <summary>
        /// 生成标准件后存入指定目录（已有标准件不再生成），若MDGBParts.Lengh>0,就是
        /// </summary>
        internal static string MDGBParts
        {
            get
            {
                if (mdgbparts == "null")
                {
                    mdgbparts = cinfo.getValue("MDGBParts");
                }
                return mdgbparts;
            }
            set
            {

                cinfo.setValue("MDGBParts", value.ToString());
            }
        }
        private static string mdgbparts = "null";



        /// <summary>
        /// 重量属性始终为读取系统重量,材料属性始终为获取系统材质,默认是
        /// </summary>
        internal static bool bMassPPTIsAuto
        {
            get
            {
                if (bmasspptisauto == "")
                {
                    try
                    {
                        bmasspptisauto = cinfo.getValue("bMassPPTIsAuto").ToLower();
                    }
                    catch
                    {
                    }
                }
                return (bmasspptisauto != "false");
            }
        }
        private static string bmasspptisauto = "";

        /// <summary>
        /// 重量信息:auto:自动计算,text:文本,sw:和SolidWorks的统一起来
        /// </summary>
        internal static string WeightStyle
        {
            get
            {
                if (weightstyle == "")
                {
                    try
                    {
                        weightstyle = cinfo.getValue("WeightStyle");
                    }
                    catch
                    {
                        weightstyle = "auto";
                    }
                }
                return weightstyle;
            }
        }
        private static string weightstyle = "";
        /// <summary>
        /// 材质特性,auto:自动计算,text:文本,sw:和SolidWorks的统一起来
        /// </summary>
        internal static string MaterialStyle
        {
            get
            {
                if (materialstyle == "")
                {
                    try
                    {
                        materialstyle = cinfo.getValue("MaterialStyle");
                    }
                    catch
                    {
                        materialstyle = "auto";
                    }
                }
                return materialstyle;
            }
        }
        private static string materialstyle = "";



        /// <summary>
        /// 生成企业配件后不隐藏所有线段，尺寸等等。
        /// </summary>
        internal static bool bHiddenDims
        {
            get
            {
                if (bhiddendims == "")
                {
                    try
                    {
                        bhiddendims = cinfo.getValue("hiddenDimensions").ToLower();
                    }
                    catch
                    {
                    }
                }
                return (bhiddendims != "false");
            }
        }
        private static string bhiddendims = "";



        /// <summary>
        /// 标准件保存时要替换掉的字符
        /// 1:去掉空格,2去掉反斜杠,3去掉点,4去掉下划线,5去掉中划线,6替换反斜杠为中文,7替换点为中文
        /// </summary>
        private static string strNameReplace = "-1";
        /// <summary>
        /// 命名规则,或名称,或标准，这是标准件的
        /// </summary>
        internal static string strNameStyle
        {
            get
            {
                string str=cinfo.getValue("strNameStyle");
                if (str.Length == 0)
                {
                    str = "标准名称规格(材料)";
                }
                return str;
            }
            set
            {
                cinfo.setValue("strNameStyle", value);
            }
        }
        /// <summary>
        /// 命名规则,或名称,或标准,这是配件库的
        /// </summary>
        internal static string strNameStyleWeb
        {
            get
            {
                string str=cinfo.getValue("strNameStyleWeb");
                if(str.Length ==0)
                {
                    str = "标准名称规格";
                }
                return str;
            }
            set
            {
                cinfo.setValue("strNameStyleWeb", value);
            }
        }
        /// <summary>
        /// 根据命名规则,输入[名称],[标准],[规格],[材料],[级别],返回零件的保存名称
        /// </summary>
        internal static string getName(string name, string stand, string style, string cailiao, string level, bool bStandard)
        {
            string str = "";
            if (bStandard)
            {
                str = strNameStyle;
            }
            else
            {
                str = strNameStyleWeb;
            }
            str = str.Replace("名称", "*名称*");
            str = str.Replace("规格", "*规格*");
            str = str.Replace("标准", "*标准*");
            str = str.Replace("代号", "*代号*");
            str = str.Replace("材料", "*材料*");
            str = str.Replace("级别", "*级别*");
            str = str.Replace("*名称*", name);
            str = str.Replace("*规格*", style);
            str = str.Replace("*标准*", stand);
            str = str.Replace("*代号*", stand);
            str = str.Replace("*材料*", cailiao);
            str = str.Replace("*级别*", level);

            str = getNameReplaceCode(str);

            return str;
        }
        internal static string getNameReplaceCode(string str)
        {
            str = str.Replace("[]", "").Replace("<未指定>", "");

            if (strNameReplace == "-1")
            {
                strNameReplace = cinfo.getValue("strNameReplace");
                if (strNameReplace.Length == 0)
                {
                    strNameReplace = "67";
                }
            }
            if (strNameReplace.IndexOf("1") != -1)//去掉空格
            {
                str = str.Replace(" ", "");
            }
            if (strNameReplace.IndexOf("2") != -1)//去掉反斜杠
            {
                str = str.Replace("/", "");
            }
            if (strNameReplace.IndexOf("3") != -1)//去掉点
            {
                str = str.Replace(".", "");
            }
            if (strNameReplace.IndexOf("4") != -1)//去掉下划线
            {
                str = str.Replace("_", "");
            }
            if (strNameReplace.IndexOf("5") != -1)//去掉中划线
            {
                str = str.Replace("-", "");
            }
            if (strNameReplace.IndexOf("6") != -1)//替换反斜杠为中文
            {
                str = str.Replace('/', '／');
            }
            if (strNameReplace.IndexOf("7") != -1)//替换点为中文
            {
                str = str.Replace('.', '．');
            }

            return str;
        }


        //启用迈迪文件模版,把发恩特的图纸模版家进去
        private void chkUserMDTemplates_CheckedChanged(object sender, EventArgs e)
        {
            if (AllData.iSwApp == null)
            {
                StringOperate.Alert("没有打开SolidWorks,无法设置,请按如下步骤手工设置:\n\n【工具】》【选项】》【文件位置】》【添加】》选择本软件安装目录下的文件夹【迈迪】完成设置。");
                return;
            }

            //首先设置使用迈迪模板
            this.SetUseMDTemplate();

            //选中状态下打开设置窗口
            if (chkUserMDTemplates.Checked)
            {
                PSetUp_Ref sr = new PSetUp_Ref();
                sr.ShowDialog();
            }
        }
        /// <summary>
        /// 在SolidWorks中设置使用迈迪模板库。
        /// </summary>
        internal void SetUseMDTemplate()
        {
            //防止SW默认模板路径丢失
            string sPart = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            string sAsm = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
            string sDwg = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);

            //不能删除原来的路径
            string dwgTempPath = AllData.StartUpPath + "\\迈迪模板";
            string txt = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates);

            if (txt.IndexOf("\\迈迪模板") == -1)
            {
                dwgTempPath = txt + ";" + dwgTempPath;
                AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates, dwgTempPath);
            }
            //这里是V3.01以后的版本了
            //已经是第二次安装了,并且换了安装位置
            else
            {
                string strPath = "";
                string[] strArr = txt.Split(new char[] { ';' });
                foreach (string s in strArr)
                {
                    if (s.IndexOf("迈迪模板") != -1)
                    {
                        strPath += dwgTempPath + ";";
                    }
                    else
                    {
                        strPath += s + ";";
                    }
                }
                strPath = strPath.TrimEnd(new char[] { ';' });
                AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates, strPath);
            }

            //设置默认模板路径
            if(sPart.Length>0)AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart, sPart);
            if (sAsm.Length > 0) AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly, sAsm);
            if (sDwg.Length > 0) AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing, sDwg);
        }
        //使用迈迪材质库
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (AllData.iSwApp == null)
            {
                StringOperate.Alert("无法设置,请按如下步骤手工设置:\n\n【工具】》【选项】》【文件位置】》【添加】》选择本软件安装目录下的文件夹【迈迪】完成设置。");
            }
            else
            {
                this.SetUseMDMaterial();
            }
        }
        /// <summary>
        /// 设置使用迈迪材料库
        /// </summary>
        internal void SetUseMDMaterial()
        {
            //防止SW默认模板路径丢失
            string sPart = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            string sAsm = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
            string sDwg = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);

            //不能删除原来的路径
            string dwgTempPath = AllData.StartUpPath + "\\迈迪材质";
            string txt = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsMaterialDatabases);

            if (txt.IndexOf("\\迈迪材质") == -1)
            {
                dwgTempPath = txt + ";" + dwgTempPath;
                AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsMaterialDatabases, dwgTempPath);
            }
            else//已经是第二次安装了,并且换了安装位置
            {
                string strPath = "";
                string[] strArr = txt.Split(new char[] { ';' });
                foreach (string s in strArr)
                {
                    if (s.IndexOf("迈迪材质") != -1)
                    {
                        strPath += dwgTempPath + ";";
                    }
                    else
                    {
                        strPath += s + ";";
                    }
                }
                strPath = strPath.TrimEnd(new char[] { ';' });
                AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsMaterialDatabases, strPath);
            }

            //设置默认模板路径
            if (sPart.Length > 0) AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart, sPart);
            if (sAsm.Length > 0) AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly, sAsm);
            if (sDwg.Length > 0) AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing, sDwg);
        }

        //使用焊件轮廓库
        private void chkWeldment_CheckedChanged(object sender, EventArgs e)
        {
            if (AllData.iSwApp == null)
            {
                StringOperate.Alert("无法设置,请按如下步骤手工设置:\n\n【工具】》【选项】》【文件位置】》【焊件轮廓】》选择文件夹【" + AllData.StartUpPath + "\\WeldProfiles】完成设置。");
            }
            else
            {
                this.SetUseMDWeldmentProfile();
            }
        }
        /// <summary>
        /// 设置使用迈迪焊件轮廓库
        /// </summary>
        internal void SetUseMDWeldmentProfile()
        {
            //防止SW默认模板路径丢失
            string sPart = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            string sAsm = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
            string sDwg = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);

            //不能删除原来的路径
            string dwgTempPath = AllData.StartUpPath + "\\WeldProfiles";
            string txt = AllData.iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsWeldmentProfiles);

            if (txt.ToLower().IndexOf("\\WeldProfiles") == -1)
            {
                dwgTempPath = txt + ";" + dwgTempPath;
                AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsWeldmentProfiles, dwgTempPath);
            }
            else
            {
                string strPath = "";
                string[] strArr = txt.Split(new char[] { ';' });
                foreach (string s in strArr)
                {
                    if (s.IndexOf("\\WeldProfiles") != -1)
                    {
                        strPath += dwgTempPath + ";";
                    }
                    else
                    {
                        strPath += s + ";";
                    }
                }
                strPath = strPath.TrimEnd(new char[] { ';' });
                AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsWeldmentProfiles, strPath);
            }

            //设置默认模板路径
            if (sPart.Length > 0) AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart, sPart);
            if (sAsm.Length > 0) AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly, sAsm);
            if (sDwg.Length > 0) AllData.iSwApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing, sDwg);
        }



        //标准件统一路径
        private void chkUseGBParts_CheckedChanged(object sender, EventArgs e)
        {
            if (chkUseGBParts.Checked)
            {
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.ShowNewFolderButton = true;
                fd.Description = "选择一个本机或网络目录，生成零件后自动保存到目录中，如果已存在同名零件就直接打开现有零件，不再重新生成。";
                if (Directory.Exists(this.txtGBBasePath.Text))
                {
                    fd.SelectedPath = txtGBBasePath.Text;
                }
                if (fd.ShowDialog() == DialogResult.OK)
                {
                    MDGBParts = txtGBBasePath.Text = fd.SelectedPath;
                }
            }
        }

        //同一零件，多个配置//选中此项后弹出选择目录
        private void chkOnePartManyConf_CheckedChanged(object sender, EventArgs e)
        {
            if (chkOnePartManyConf.Checked)
            {
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.ShowNewFolderButton = true;
                fd.Description = "请选择一个工作路径，生成零件时如果路径中已存在同名零件就在零件中添加一个配置并打开！";
                if (Directory.Exists(txtWorkPath.Text))
                {
                    fd.SelectedPath = txtWorkPath.Text;
                }
                if (fd.ShowDialog() == DialogResult.OK)
                {
                    txtWorkPath.Text = fd.SelectedPath;
                    Interop.Office.Core.Properties.Settings.Default.strWorkPath = txtWorkPath.Text;
                    Interop.Office.Core.Properties.Settings.Default.Save();
                }
            }
        }
        



        //查找IP
        private void btnSelServerIP_Click(object sender, EventArgs e)
        {
            MDClient.SelIPAddress selIP = new MDClient.SelIPAddress();
            if (selIP.ShowDialog() == DialogResult.OK)
            {
                this.cmbServerIP.Text = selIP.txtSelectIP.Text;
            }
        }
        //扫描哪一台是主机
        private void btnScanIP_Click(object sender, EventArgs e)
        {
            Transport tsp = new Transport();
            string serverIP = tsp.GetServerIPByScreen();
            if (serverIP.Length > 0)
            {
                this.cmbServerIP.Text = serverIP;
            }
        }
        //测试连接
        private void btnServerBack_Click(object sender, EventArgs e)
        {
            fntWeb_ServerIP = this.cmbServerIP.Text;
            fntWeb_ServerPort = Convert.ToInt32(this.cmbServerPort.Text);
            Transport tsp = new Transport();
            string str = tsp.CanLinkServer;
            if (str.Length >0)
            {
                StringOperate.Alert("测试失败:" + str );
            }
            else
            {
                StringOperate.Alert("测试成功!");
            }
        }
        //配件网模型库路径
        private void btnSetFntWebPath_Click(object sender, EventArgs e)
        {
            //这个地方还有一个作用就是控制是否是MD
            try
            {
                string txt = this.txtFntWebPath.Text.Trim().Replace(" ", "");
                if (txt.Length == 8)//76543210
                {
                    long a = Convert.ToInt64(txt);
                    Interop.Office.Core.Properties.Settings.Default.exe1ength = a;
                    Interop.Office.Core.Properties.Settings.Default.Save();
                    return;
                }
                else if (txt.Length == 24)//输入盗版SW序列号
                {
                    if (bIsManager)
                    {
                        backMethod bm = new backMethod();
                        string newVal = bm.mAdd(txt);
                        cinfo.addValue("SNumber",newVal);
                        StringOperate.Alert(txt + "=" + newVal + " 添加成功！");
                    }
                }
            }
            catch { }


            //作为服务器运行的话就无法设置
            if (cinfo.getValue(SetupNames.UsingType.ToString()) == SetupNames.Server.ToString())
            {
                StringOperate.Alert("您在初始化中选择了作为服务器运行，无法设置此路径！只能使用默认路径。"); return;
            }

            //选择目录
            FolderBrowserDialog fb = new FolderBrowserDialog();
            if(this.txtFntWebPath.Text.Length>0 && Directory.Exists(this.txtFntWebPath.Text))
            {
                fb.SelectedPath = this.txtFntWebPath.Text;
            }
            if (fb.ShowDialog() == DialogResult.OK)
            {
                MD3DParts = this.txtFntWebPath.Text = fb.SelectedPath;
            }
        }
        //配件库路径设为默认值
        private void btnSetFntWebPathDefault_Click(object sender, EventArgs e)
        {
            MD3DParts = this.txtFntWebPath.Text = AllData.StartUpPath + "\\MD3DParts\\";
            //这里要指定默认值“”
            cinfo.setValue("strMD3DPartsPath", "");
        }



        //联网申请管理员注册码
        private void btnAskFor_Click(object sender, EventArgs e)
        {
            bool bhasOK = Dbtool2.strError.AutoSize ;
            if (!bhasOK)
            {
                StringOperate.Alert("您当前还没有注册，只有正式注册用户才能申请管理员注册码！"); return;
            }

            string mCode = Dbtool2.strError.ToString();
            string rCode = Dbtool2.strError.Text;

            string url = "http://www.my3dparts.com/Opinion_Soft/RegisterMgr.aspx?MachineCode=" + Dbtool2.strError.ToString() + "&RegCode=" + Dbtool2.strError.Text;

            System.Diagnostics.Process.Start(url);
        }
        //将管理员注册码写入到注册表
        private void btnRegToMgr_Click(object sender, EventArgs e)
        {
            string txt = this.txtMgrRegCode.Text.Trim();
            if (txt !=  GetMgrCode() )
            {
                StringOperate.Alert("您输入的管理员注册码不正确，请输入正确的12位管理员注册码！"); return;
            }
            else
            {
                PSetUp.bIsManager = true;
                StringOperate.Alert("注册管理员成功！");
            }
        }
        //申请说明
        private void btnMgrDesc_Click(object sender, EventArgs e)
        {
            string str= "如需开放管理员工具，请联系迈迪公司获取管理员注册码" + System.Environment.NewLine + System.Environment.NewLine;
            str += "注册管理员需提供机器码和注册码。" + System.Environment.NewLine + System.Environment.NewLine;
            str += "只有企业用户可申请注册为管理员。个人用户无需使用管理员工具。" + System.Environment.NewLine + System.Environment.NewLine;
            str += "注册为管理员后可以在标准件库，配件库中使用管理工具。" + System.Environment.NewLine + System.Environment.NewLine;

            StringOperate.Alert(str);
        }


        /// <summary>
        /// Excel文件的扩展名是多少,返回小写,没有设置之前返回“.xls”
        /// </summary>
        internal static string ExcelExtension
        {
            get
            {
                string exten = Properties.Settings.Default.ExcelExtension.ToLower();
                if (exten.Length == 0)
                {
                    Setting_Excel sel = new Setting_Excel();
                    sel.ShowDialog();
                }
                exten = Properties.Settings.Default.ExcelExtension.ToLower();
                if (exten.Length == 0)
                {
                    exten = ".xls";
                }
                return exten;
            }
        }
        //设置当前使用的Excel版本
        private void btnSettingExcelExten_Click(object sender, EventArgs e)
        {
            Setting_Excel sel = new Setting_Excel();
            sel.ShowDialog();
        }












    }
}