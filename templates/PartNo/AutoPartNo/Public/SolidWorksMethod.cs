using System;
using System.Collections.Generic;
using System.Text;
using SolidWorks.Interop.sldworks;
using System.Collections;
using System.Windows.Forms;
using System.Data;
using System.IO;
using SolidWorks.Interop.swconst;

namespace Interop.Office.Core
{

    ////获取当前创建的草图,如果不用这句，下面也获取不到skgear
    //swPart.SketchManager.InsertSketch(true);
    //Sketch skgear = (Sketch)swPart.GetActiveSketch2();

    ////获取当前的草图//获取当前创建的草图
    //SelectionMgr selMgr = (SelectionMgr)swPart.SelectionManager;
    //Feature ftSketch = (Feature)selMgr.GetSelectedObject5(1);//注意这里从1开始

    //对当前的特征重命名,这样就不用判断是什么语言了。
    //swPart.SelectedFeatureProperties(0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, true, false, "LuoXuanXianA");

    internal class SolidWorksMethod
    {
        internal SolidWorksMethod()
        {
        }

        /// <summary>
        /// 根据Comp和属性名称得到零件的属性,并保存(返回是否成功）
        /// </summary>
        internal bool SetAtrByCompAndAtrName(Component2 comp, string AtrName, object objValue)
        {
            if (comp == null) return false;

            //是否需要自动解压
            bool bSuppressed = comp.IsSuppressed();//如果是轻化状态，先还原，完成后再转到轻化
            int iOldState = comp.GetSuppression(); //得到还原前是什么状态
            if (bSuppressed)
            {
                comp.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved);
            }
            ModelDoc2 swModel = (ModelDoc2)comp.GetModelDoc();

            bool bOK = false;
            int a = 0;
            if (swModel != null)
            {
                try
                {
                    CustomPropertyManager cpmMdl = swModel.Extension.get_CustomPropertyManager(comp.ReferencedConfiguration);
                    a = cpmMdl.GetType2(AtrName);
                    if (a != 0 && a != -1)
                    {
                        a = cpmMdl.Set(AtrName, objValue.ToString());//0=ok,1 = err
                        bOK = (a == 0);
                    }
                    else
                    {
                        Type T = objValue.GetType();
                        if (T == typeof(System.String)) a = (int)swCustomInfoType_e.swCustomInfoText;
                        else if (T == typeof(System.Boolean)) a = (int)swCustomInfoType_e.swCustomInfoYesOrNo;
                        else if (T == typeof(System.Single)) a = (int)swCustomInfoType_e.swCustomInfoDouble;
                        else if (T == typeof(System.Int32)) a = (int)swCustomInfoType_e.swCustomInfoNumber;
                        else if (T == typeof(System.DateTime)) a = (int)swCustomInfoType_e.swCustomInfoDate;

                        //1--ok ,0--err,-1--alerady exists
                        a = cpmMdl.Add2(AtrName, a, objValue.ToString());//添加到自定义属性中
                        bOK = (a == 1);
                    }
                    swModel.Save();
                }
                catch (Exception ea) { StringOperate.Alert(ea.Message + ".......写属性......."); }
            }

            if (bSuppressed)//再回到轻化状态
            {
                comp.SetSuppression2(iOldState);
            }

            return bOK ;
        }
        /// <summary>
        /// 设置切割清单特征的用户定义属性(不保存）(返回是否成功）
        /// </summary>
        internal bool SetAtrByWeldmentFeature(Feature docFeat, string AtrName, object objValue)
        {
            if (docFeat == null) return false;

            CustomPropertyManager CPMgr = null;
            try
            {
                 CPMgr = docFeat.CustomPropertyManager;
            }
            catch
            {
                return false;
            }
            if (CPMgr != null)
            {
                //判断是否已经存在这个属性,如果有，先删除
                string[] strArr = (string[])CPMgr.GetNames();
                foreach (string s in strArr)
                {
                    if (s == AtrName)
                    {
                        int iok = CPMgr.Delete(AtrName); break;
                    }
                }

                int a = CPMgr.Add2(AtrName, (int)swCustomInfoType_e.swCustomInfoText  , objValue.ToString());//添加到自定义属性中
                if (a < 1)
                {
                    a = CPMgr.Set(AtrName, objValue.ToString());
                    if (a < 1)
                    {
                        StringOperate.AlertDebug("添加焊件属性失败！");
                    }
                }
                return (a > 0);
            }
            return false;
        }
        /// <summary>
        /// 设置零件的属性(当前配置)
        /// </summary>
        internal bool SetAtrByModel(string AttName, string Value, ModelDoc2 swModel)
        {
            string cfgName = "";
            object obj = swModel.ConfigurationManager.ActiveConfiguration;
            if (obj != null)
            {
                cfgName = (obj as Configuration).Name;
            }

            CustomPropertyManager cpmMdl = swModel.Extension.get_CustomPropertyManager(cfgName);//自定义属性
            cpmMdl.Delete(AttName);
            int a = cpmMdl.Add2(AttName, (int)swCustomInfoType_e.swCustomInfoText, Value);
            return (a > 0);
        }
        /// <summary>
        /// 设置零件的属性(自定义属性,齿轮专用)
        /// </summary>
        internal void SetAtrByModel(string AttName, string Value, IModelDoc2 swModel)
        {
            string cfgName = "";
            object obj = swModel.ConfigurationManager.ActiveConfiguration;
            if (obj != null)
            {
                cfgName = (obj as Configuration).Name;
            }

            swModel.AddCustomInfo3(cfgName, AttName, (int)swCustomInfoType_e.swCustomInfoText, Value);
            swModel.set_CustomInfo2(cfgName, AttName, Value);
        }
        /// <summary>
        /// 设置零件的属性(配置为空时写入到自定义属性中)
        /// </summary>
        internal bool SetAtrByModel(string AttName, string Value, ModelDoc2 swModel,string cfgName)
        {
            CustomPropertyManager cpmMdl = swModel.Extension.get_CustomPropertyManager(cfgName);//自定义属性
            cpmMdl.Delete(AttName);
            int a= cpmMdl.Add2(AttName, (int)swCustomInfoType_e.swCustomInfoText, Value);
            return (a > 0);
        }





        /// <summary>
        /// 根据Component得到零件的属性,一次可以得到多个属性
        /// </summary>
        internal string[] GetAtrByCompAndAtrName(string[] AttNameArr, Component2 comp)
        {
            //是否需要自动解压
            bool bSuppressed = comp.IsSuppressed();//如果是轻化状态，先还原，完成后再转到轻化
            int iOldState = comp.GetSuppression(); //得到还原前是什么状态
            if (bSuppressed)
            {
                comp.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved);
            }
            ModelDoc2 swModel = (ModelDoc2)comp.GetModelDoc();
            if (swModel != null)
            {
                CustomPropertyManager cpmMdl = swModel.Extension.get_CustomPropertyManager(comp.ReferencedConfiguration);

                for (int i = 0; i < AttNameArr.Length; i++)
                {
                    string sVal = "";
                    string strResult = "";
                    cpmMdl.Get2(AttNameArr[i], out sVal, out strResult);
                    AttNameArr[i] = strResult;//把结果保存到数组中
                }
            }

            if (bSuppressed)//再回到轻化状态
            {
                comp.SetSuppression2(iOldState);
            }

            return AttNameArr;
        }
        /// <summary>
        /// 根据ModelDoc得到零件的属性
        /// </summary>
        internal string GetAttValueByName(string AtrName, ModelDoc2 swModel)
        {
            string cfgname2 = swModel.ConfigurationManager.ActiveConfiguration.Name;

            CustomPropertyManager cpmMdl = swModel.Extension.get_CustomPropertyManager(cfgname2);//自定义属性
            string outval = "";
            string sovVal = "";
            cpmMdl.Get2(AtrName, out outval, out sovVal);

            return sovVal;
        }
        /// <summary>
        /// 根据ModelDoc得到零件的属性
        /// </summary>
        internal string GetAttValueByName(string AtrName, ModelDoc2 swModel,string cfgName)
        {
            CustomPropertyManager cpmMdl = swModel.Extension.get_CustomPropertyManager(cfgName);//自定义属性
            string outval = "";
            string sovVal = "";
            cpmMdl.Get2(AtrName, out outval, out sovVal);

            return sovVal;
        }
        /// <summary>
        /// 根据weldmentFeature得到焊件的属性
        /// </summary>
        internal OneSldPPt GetAttValueByName(string AtrName, Feature docFeat)
        {
            CustomPropertyManager cpmgr = docFeat.CustomPropertyManager;

            OneSldPPt op = new OneSldPPt(AtrName);
            cpmgr.Get2(op.name, out op.value, out op.solValue);
            op.iType = cpmgr.GetType2(op.name);
            if (op.solValue == "<未指定>")
            {
                op.solValue = "";
                if (op.value == "<未指定>") op.value = "";
            }
            return op;
        }



        /// <summary>
        /// 得到当前配置下所有自定义属性的名称
        /// </summary>
        internal string[] GetAllCustomPropertyName(ModelDoc2 swModel)
        {
            Configuration curcfg2 = (Configuration)swModel.ConfigurationManager.ActiveConfiguration;
            if (curcfg2 == null) return null;//工程图下可能有错

            string cfgname2 = curcfg2.Name;

            ModelDocExtension MDExtension = swModel.Extension;
            CustomPropertyManager cpm = MDExtension.get_CustomPropertyManager(cfgname2);
            string[] strArr = (string[])cpm.GetNames();
            return strArr;
        }
        
        /// <summary>
        /// 得到一个Component的所有属性，未指定的值为"",是否包括配置属性默认是，是否包括全局的属性，如果包括全局属性在最后
        /// </summary>
        internal List<OneSldPPt> GetAllCustomPropertyInfo(ModelDoc2 swModel, Component2 comp, bool bIncludeCfg, bool bIncludeGlobal, Feature weldmentFeature)
        {
            //是否需要自动解压
            bool bSuppressed = false;
            int iOldState = 0;
            string cfgName = "";
            if (weldmentFeature != null)
            {
            }
            else if (swModel == null)
            {
                cfgName = comp.ReferencedConfiguration;
                bSuppressed = comp.IsSuppressed();//如果是轻化状态，先还原，完成后再转到轻化
                iOldState = comp.GetSuppression(); //得到还原前是什么状态
                if (bSuppressed)
                {
                    comp.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved);
                }
                swModel = (ModelDoc2)comp.GetModelDoc();
                if (swModel == null) return null;
            }
            else
            {
                object obj=swModel.GetActiveConfiguration();
                if (obj != null)
                {
                    Configuration swConf = (Configuration)obj;
                    cfgName = swConf.Name;
                }
            }

            List<OneSldPPt> listPPt = new List<OneSldPPt>();

            //默认包括配置的属性
            if (bIncludeCfg)
            {
                OneSldPPt[] pptArr = this.GetAllCustomPropertyInfo(swModel, cfgName, null, weldmentFeature);
                if (pptArr != null && pptArr.Length > 0)
                {
                    listPPt.AddRange(pptArr);
                }
            }

            //如果包括全局属性
            if (bIncludeGlobal)
            {
                OneSldPPt[] pptArr = this.GetAllCustomPropertyInfo(swModel, "", null, weldmentFeature);
                if (pptArr != null && pptArr.Length > 0)
                {
                    listPPt.AddRange(pptArr);
                }
            }

            if (bSuppressed && comp != null)//再回到轻化状态
            {
                comp.SetSuppression2(iOldState);
            }

            return listPPt;
        }


         /// <summary>
         /// 得到一个配置下所有的属性名和值，如果是全局属性cfgname="",当前配置的属性cfgname="ActiveCfg" ，未指定的值为"" 
         /// </summary>
         /// <param name="swModel">swModel和weldmentFeature二者要指定其一</param>
        /// <param name="cfgname">配置名称，读取焊件清单属性时不用指定</param>
         /// <param name="arAttNames"></param>
         /// <param name="weldmentFeature">焊件清单特征</param>
         /// <returns></returns>
        internal OneSldPPt[] GetAllCustomPropertyInfo(ModelDoc2 swModel, string cfgname, ArrayList arAttNames, Feature weldmentFeature)
        {
            CustomPropertyManager cpmgr = null;
            if (weldmentFeature != null)
            {
                cpmgr = weldmentFeature.CustomPropertyManager;//焊件结构专用
            }
            else
            {
                if (cfgname.ToLower() == "activecfg")
                {
                    cfgname = swModel.ConfigurationManager.ActiveConfiguration.Name;
                }
                if (cfgname != "")
                {
                    cpmgr = swModel.Extension.get_CustomPropertyManager(cfgname); //curcfg2.CustomPropertyManager; 
                }
                else
                {
                    cpmgr = swModel.Extension.get_CustomPropertyManager(string.Empty);//自定义属性
                }
            }

                string[] strArr = (string[])cpmgr.GetNames();
            if (arAttNames == null)
            {
                arAttNames = new ArrayList();
                if (strArr != null && strArr.Length > 0) arAttNames.AddRange(strArr);//读取现有的所有属性
            }
            else
            {
                //2015-5-26只读取特定的属性, 其它的属性过滤掉
                ArrayList arHas2 = new ArrayList();
                foreach (object o in arAttNames)
                {
                    foreach (string s in strArr)
                    {
                        if (s == o.ToString())
                        {
                            arHas2.Add(s); break;
                        }
                    }
                }
                arAttNames = arHas2;
            }
            int iCount = arAttNames.Count;
            if (iCount == 0) return null;

            OneSldPPt[] pptArr = new OneSldPPt[iCount];
            for(int i=0;i<iCount ;i++)
            {
                OneSldPPt op = new OneSldPPt(arAttNames[i].ToString());
                cpmgr.Get2(op.name, out op.value, out op.solValue);
                op.iType = cpmgr.GetType2(op.name);
                if (op.solValue == "<未指定>")
                {
                    op.solValue = "";
                    if (op.value == "<未指定>") op.value = "";
                }
                pptArr[i] = op;
            }

            return pptArr;
        }
        /// <summary>
        /// 得到所有的配置，每个配置所有的属性信息
        /// </summary>
        internal Hashtable GetAllCfgCustomPropertyInfo(ModelDoc2 swModel)
        {
            Hashtable ht = new Hashtable();
            string[] strArr = (string[])swModel.GetConfigurationNames();
            foreach (string s in strArr)
            {
                ht.Add(s, this.GetAllCustomPropertyInfo(swModel, s, null,null));
            }
            return ht;
        }




        /// <summary>
        /// 自动根据每个配置的属性生成一个表，包括所有component和配置，所有的属性，但是不包括尺寸。
        /// </summary>
        internal DataTable CreateDesignTableFromModelDoc(ModelDoc2 swModel)
        {
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                return this.CreateDesignTableFromPart(swModel);
            }

            AssemblyDoc swasm = (AssemblyDoc)swModel;

            //生成一个表
            DataTable dt = new DataTable();
            dt.Columns.Add("配置名", typeof(System.String));
            string[] strArr = (string[])swModel.GetConfigurationNames();
            foreach (string s in strArr)
            {
                DataRow dr = dt.NewRow();
                dr[0] = s;
                dt.Rows.Add(dr);
            }

            //生成每个组建的列
            ArrayList arComp = this.GetAllAssemblyComponents(swasm);
            foreach (object o in arComp)
            {
                Component2 swcomp = (Component2)o;
                
                //string sName = "$配置@" + swcomp.Name2.Remove(swcomp.Name2.LastIndexOf('-'));
                //sName = sName + "<" + swcomp.Name2.Substring(swcomp.Name2.LastIndexOf('-') + 1);
                //sName = sName + ">";

                dt.Columns.Add(swcomp.Name2 , typeof(System.String));
            }

            //获取每个组建的配置信息
            foreach (DataRow dr in dt.Rows)
            {
                string cfgname = dr[0].ToString();
                swModel.ShowConfiguration2(cfgname);

                //获取一个配置的所有属性并列表
                OneSldPPt[] pptArr = this.GetAllCustomPropertyInfo(swModel, cfgname, null, null);
                for (int i = 0; i < pptArr.Length; i++)
                {
                    OneSldPPt ppt = pptArr[i];
                    if (!dt.Columns.Contains(ppt.name))
                    {
                        dt.Columns.Add(ppt.name, typeof(System.String));
                    }
                    dr[ppt.name] = ppt.value;
                }

                //获取一个配置的所有子零件对应的配置
                object[] objArr = (object[])swasm.GetComponents(false);
                foreach (object obj in objArr)
                {
                    Component2 comp = (Component2)obj;
                    dr[comp.Name2] = comp.ReferencedConfiguration;
                }
            }

            return dt;
        }
        /// <summary>
        /// 自动根据每个配置的属性生成一个表，包括所有的属性，但是不包括尺寸
        /// </summary>
        internal DataTable CreateDesignTableFromPart(ModelDoc2 swModel)
        {
            //生成一个表
            DataTable dt = new DataTable();
            dt.Columns.Add("配置名", typeof(System.String));
            string[] strArr = (string[])swModel.GetConfigurationNames();
            foreach (string s in strArr)
            {
                DataRow dr = dt.NewRow();
                dr[0] = s;
                dt.Rows.Add(dr);
            }

            //获取每个组建的配置信息
            foreach (DataRow dr in dt.Rows)
            {
                string cfgname = dr[0].ToString();
                //获取一个配置的所有属性并列表
                bool b= swModel.ShowConfiguration2(cfgname);
                if (b == false)
                {
                    StringOperate.Alert(swModel.GetPathName() + " 的配置[" + cfgname + "]设为当前失败！");
                }

                OneSldPPt[] pptArr = this.GetAllCustomPropertyInfo(swModel, cfgname, null, null);
                for (int i = 0; i < pptArr.Length; i++)
                {
                    OneSldPPt ppt = pptArr[i];
                    if (!dt.Columns.Contains(ppt.name))
                    {
                        dt.Columns.Add(ppt.name, typeof(System.String));
                    }
                    dr[ppt.name] = ppt.value;
                }
            }

            return dt;
        }




        /// <summary>
        /// 获取方程式中全局变量的值，Golbal Variables
        /// </summary>
        /// <param name="VariableName">全局变量的名称</param>
        /// <param name="Part">模型</param>
        /// <returns>如果没有找到这个全局变量，返回-1</returns>
        internal double getGolbalVariablesValue(string VariableName, ModelDoc2 Part)
        {
            /*
             * 没有好的办法，用EquationMgr遍历所有的方程式，
             * 判断一下方程式的名称，
             * 如果名称是你定义的全局变量名如\"X\"= \"Y\"，
             * 然后获取其值
             */

            EquationMgr emgr = (EquationMgr)Part.GetEquationMgr();

            int ic = emgr.GetCount();
            for (int i = 0; i < ic; i++)
            {
                string s = emgr.get_Equation(i);//\"Y\"= \"X\"
                s = s.Replace("\"", "");//Y=X
                s = s.Remove(s.IndexOf('='));
                s = s.Trim();

                if (s.ToLower() == VariableName.ToLower())
                {
                    return emgr.get_Value(i);
                }
            }
            return -1;
        }
        /// <summary>
        /// 得到一个零件下的所有尺寸信息。，并返回ArrayList组
        /// </summary>
        internal ArrayList GetAllDimensions(ModelDoc2 swModel)
        {
            ArrayList arDims = new ArrayList();
            Feature f = (Feature)swModel.FirstFeature();
            while (f != null)
            {
                DisplayDimension disDim = (DisplayDimension)f.GetFirstDisplayDimension();
                while (disDim != null)
                {
                    Dimension dim = (Dimension)disDim.GetDimension();
                    arDims.Add(dim);

                    disDim = disDim.GetNext5();
                }

                f = (Feature)f.GetNextFeature();
            }
            return arDims;
        }

        /// <summary>
        /// 得到一个特征下的所有尺寸信息,可以是特征，也可以是草图等等，如果Feature为空，自动获取当前选中的Feature
        /// </summary>
        internal List<DisplayDimension> GetAllDimensions(ModelDoc2 Part, Feature ft)
        {
            if (Part == null && ft == null) return null;
            if (ft == null)
            {
                SelectionMgr selMgr = (SelectionMgr)Part.SelectionManager;

                if (selMgr.GetSelectedObjectCount() == 0) return null;

                //获取草图，获取当前选中的草图
                ft = (Feature)selMgr.GetSelectedObject5(1);
            }


            //先全部显示才行
            bool b = Part.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayFeatureDimensions);
            if (!b) Part.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayFeatureDimensions, true);

            List<DisplayDimension> arDims = new List<DisplayDimension>();

            //还是用Enum不容易出错
            EnumDisplayDimensions enumdis = ft.EnumDisplayDimensions();
            if (enumdis != null)
            {

                DisplayDimension disDim = null;
                int PceltFetched = 0;
                enumdis.Next(1, out disDim, ref PceltFetched);
                while (disDim != null)
                {
                    arDims.Add(disDim);
                    //disDim.MarkedForDrawing = true;//带到工程图中去

                    enumdis.Next(1, out disDim, ref PceltFetched);
                }
            }

            if (!b) Part.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayFeatureDimensions, b);


            return arDims;
        }


        /// <summary>
        /// 得到三个基准面的名称:1:前视基准面,2上视基准面,3右视基准面
        /// </summary>
        internal swRefPlaneNames GetRefPlaneNames(IModelDoc2 swModel)
        {
            swRefPlaneNames swRef = new swRefPlaneNames();
            string[] strArr = new string[3] { "", "", "" };
            Feature swFeat = (Feature)swModel.FirstFeature();
            int iCount = 0;
            while (swFeat != null)
            {
                if (swFeat.GetTypeName() == "RefPlane")
                {
                    iCount++;
                    if (iCount == 1)
                    {
                        swRef.FrontPlaneName = swFeat.Name;
                    }
                    else if (iCount == 2)
                    {
                        swRef.TopPlaneName = swFeat.Name;
                    }
                    else if (iCount == 3)
                    {
                        swRef.RightPlaneName = swFeat.Name;
                    }
                    else
                    {
                        break;
                    }
                }
                swFeat = (Feature)swFeat.GetNextFeature();
            }
            return swRef;
        }
        /// <summary>
        /// 前视，上视，右视，三个基准面的名称。
        /// </summary>
        internal struct swRefPlaneNames
        {
            /// <summary>
            /// 前视基准面
            /// </summary>
            internal string FrontPlaneName ;
            /// <summary>
            /// 上视基准面
            /// </summary>
            internal string TopPlaneName  ;
            /// <summary>
            /// 右视基准面
            /// </summary>
            internal string RightPlaneName  ;
        }

        /// <summary>
        /// 得到第一个基准轴的名称,如果没有基准轴，返回“”
        /// </summary>
        internal string getAxisName(IModelDoc2 swModel)
        {
            Feature swFeat = (Feature)swModel.FirstFeature();
            while (swFeat != null)
            {
                //string s = swFeat.GetTypeName2();
                if (swFeat.GetTypeName2().ToUpper() == "REFAXIS")
                {
                    return swFeat.Name;
                }
                swFeat = (Feature)swFeat.GetNextFeature();
            }
            return "";
        }


        /// <summary>
        /// 自动计算装配体的重量, 出错就返回0 WeightAttName=重量，单重等等。
        /// </summary>
        internal double GetAssemblyTotalWeight(ModelDoc2 swModel, bool bAutoDeSuppress, string WeightAttName)
        {
            if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return 0;

            //计算它的所有零件的重量总和，子装配如果有重量就不向下读取了，否则继续读取子装配的重量
            Configuration swConf = (Configuration)swModel.GetActiveConfiguration();
            Component2 swRootComp = (Component2)swConf.GetRootComponent();

            object[] objArr = (object[])swRootComp.GetChildren();//装配的顶级零件


            double dtotalWeight = 0;       //总重量
            Hashtable ht = new Hashtable();//全路径~配置名-----单件重量

            //开始遍历
            this.GetAssemblyTotalWeight_Ref(objArr, bAutoDeSuppress, WeightAttName, ref dtotalWeight, ref ht);

            return dtotalWeight;
        }
        private void GetAssemblyTotalWeight_Ref(object[] objArr, bool bAutoDeSuppress, string WdightAttName, ref double dtotalWeight, ref Hashtable ht)
        {
            foreach (object obj in objArr)
            {
                Component2 comp = (Component2)obj;
                string key = comp.GetPathName() + "ぺ" + comp.ReferencedConfiguration;
                if (ht.ContainsKey(key))
                {
                    dtotalWeight += Convert.ToDouble(ht[key]); continue;//已经有了，不用再计算了
                }


                bool bSuppreed = false;//完成后是否需要压缩
                int iOldState = comp.GetSuppression(); //得到还原前是什么状态
                if (comp.IsSuppressed())
                {
                    if (bAutoDeSuppress)
                    {
                        comp.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved);
                        bSuppreed = true;
                    }
                    else
                    {
                        continue;
                    }
                }

                //“重量”这个属性有可能是方程式，也有可能是输入的值，这就看用户的了，。

                //不管是零件还是子装配，如果有单重，直接读取单重
                string[] SWeight = this.GetAtrByCompAndAtrName(new string[] { WdightAttName }, comp);
                if (SWeight[0].Length > 0 && SWeight[0] != "未知" && SWeight[0] != "null")//
                {
                    double thisWight = 0;
                    try
                    {
                        thisWight = Convert.ToDouble(SWeight[0]);
                    }
                    catch
                    {
                        StringOperate.Alert("零件[" + comp.Name2 + "]缺少单重信息");
                    }

                    //如果没有这个属性，读取材质重量
                    if (thisWight <= 0)
                    {
                        ModelDoc2 swModel = (ModelDoc2)comp.GetModelDoc2();
                        thisWight=GetPartWeight(swModel,5);
                    }

                    dtotalWeight += thisWight;
                    ht.Add(key, thisWight);
                }
                else//如果没有读取单重，且如果是装配，读取它的下一级
                {
                    //如果这是个装配，读取他的所有零件
                    object[] compArr = (object[])comp.GetChildren();
                    if (compArr != null && compArr.Length > 0)
                    {
                        this.GetAssemblyTotalWeight_Ref(compArr, bAutoDeSuppress, WdightAttName, ref dtotalWeight, ref ht);
                    }
                    else
                    {
                        ModelDoc2 swModel=(ModelDoc2)comp.GetModelDoc2();
                        double Weig= this.GetPartVolumn(swModel, 5);
                        dtotalWeight += Weig;
                        ht.Add(key, Weig);
                    }
                }

                //最后是否需要再次压缩
                if (bSuppreed)
                {
                    comp.SetSuppression2(iOldState);
                }
            }
        }


        /// <summary>
        /// 得到图纸的大小
        /// </summary>
        internal double[] GetDwgSize(ModelDoc2 swModel)
        {
            //得到图纸的大小
            DrawingDoc swDraw = (DrawingDoc)swModel;

            Sheet sht = (Sheet)swDraw.GetCurrentSheet();
            double SheetWidth = 0;
            double SheetHeight = 0;

            //SolidWorks2007
            int a = sht.GetSize(ref SheetWidth, ref SheetHeight);

            return new double[] { SheetWidth, SheetHeight };
        }
        /// <summary>
        /// 由swDwgPaperSizes_e得到图纸名称
        /// </summary>
        internal string GetDwgSizeNameByDwgSizeID(int iSize)
        {
            string sName = "";
            switch (iSize)
            {
                case 11:
                    {
                        sName = "A0"; break;
                    }
                case 10:
                    {
                        sName = "A1"; break;
                    }
                case 9:
                    {
                        sName = "A2"; break;
                    }
                case 8:
                    {
                        sName = "A3"; break;
                    }
                case 6:
                    {
                        sName = "A4"; break;
                    }
                case 7:
                    {
                        sName = "A4V"; break;
                    }
                case 0:
                    {
                        sName = "A"; break;
                    }
                case 1:
                    {
                        sName = "AV"; break;
                    }
                case 2:
                    {
                        sName = "B"; break;
                    }
                case 3:
                    {
                        sName = "C"; break;
                    }
                case 4:
                    {
                        sName = "D"; break;
                    }
                case 5:
                    {
                        sName = "E"; break;
                    }
                default:
                    {
                        sName = "X";
                        break;
                    }
            }
            return sName;
            //swDwgPaperA0size 11 
            //swDwgPaperA1size 10 
            //swDwgPaperA2size 9 
            //swDwgPaperA3size 8 
            //swDwgPaperA4size 6 
            //swDwgPaperA4sizeVertical 7 
            //swDwgPaperAsize 0 
            //swDwgPaperAsizeVertical 1 
            //swDwgPaperBsize 2 
            //swDwgPaperCsize 3 
            //swDwgPaperDsize 4 
            //swDwgPaperEsize 5 
            //swDwgPapersUserDefined 12 
        }


        /// <summary>
        /// 由两点坐标得到距离
        /// </summary>
        internal double GetTwoPointDistinct(double dx, double dy, double dx2, double dy2)
        {
            double x = dx2 - dx;
            x = x * x;
            double y = dy2 - dy;
            y = y * y;
            return Math.Sqrt(x + y);
        }


        /// <summary>
        /// 知道两条线段的四个点(线一点一,线一点二,线二点一,线二点二)判断两条线段是否相交,
        /// </summary>
        internal bool CheckTwoLineCross(double xa1, double ya1, double xa2, double ya2, double xb1, double yb1, double xb2, double yb2)
        {
            //线一的直线方程
            //y=kx + b 
            //ya1= ka* xal + ba 
            //ya2= kb* xa2 + bb
            //得出k 和 b

            double ka = (ya2 - ya1) / (xa2 - xa1);
            double ba = ya1 - ((ya2 - ya1) * xa1) / (xa2 - xa1);

            double kb = (yb2 - yb1) / (xb2 - xb1);
            double bb = yb1 - ((yb2 - yb1) * xb1) / (xb2 - xb1);

            if (ka == kb) return false;//这两条直线是平行的

            //两条直线相交
            //交电X,Y相等
            // ka* xal + ba == kb* xa2 + bb
            double x = (bb - ba) / (ka - kb);//焦点的X坐标

            if (x > Math.Max(xa1, xa2) || x < Math.Min(xa1, xa2)) return false;//坐标不在这两条线段的范围内,而是在延长线上
            if (x > Math.Max(xb1, xb2) || x < Math.Min(xb1, xb2)) return false;

            return true;
        }
        /// <summary>
        /// 知道一条线段上的两点，和第三点的Y坐标，求第三点的X坐标
        /// </summary>
        internal double getPointInLineX(double x1, double y1, double x2, double y2, double y3)
        {
            //线一的直线方程
            //y=kx + b 
            //x = (y - b) /k

            //y1= k* xl + b 
            //得出k 和 b

            double k = (y2 - y1) / (x2 - x1);
            double b = y1 - ((y2 - y1) * x1) / (x2 - x1);

            return (y3 - b) / k;
        }
        /// <summary>
        /// 知道一条线段上的两点，和第三点的X坐标，求第三点的Y坐标
        /// </summary>
        internal double getPointInLineY(double x1, double y1, double x2, double y2, double x3)
        {
            //线一的直线方程
            //y=kx + b 

            //ya1= ka* xal + ba 
            //得出k 和 b

            double k = (y2 - y1) / (x2 - x1);
            double b = y1 - ((y2 - y1) * x1) / (x2 - x1);

            return k * x3 + b;
        }



        /// <summary>
        /// 在SW打开的文件中，根据提供的路径返回，
        /// </summary>
        internal ModelDoc2 GetModelDoc2FromOpenDocs(string fullPathName)
        {
            if (AllData.iSwApp == null) return null;
            EnumDocuments2 enumDoc = AllData.iSwApp.EnumDocuments2();
            if (enumDoc == null) return null;

            ModelDoc2 swModel = null;
            int PceltFetched = 0;
            enumDoc.Next(1, out swModel, ref  PceltFetched);

            while (swModel != null)
            {
                if (fullPathName == swModel.GetPathName())
                {
                    return swModel;
                }

                //转向下一个
                enumDoc.Next(1, out swModel, ref PceltFetched);
            }

            return null;
        }
        /// <summary>
        /// 得到所有打开的文件，可以设置 参数（是否只显示装配）
        /// </summary>
        internal ArrayList GetAllOpenDocuments(bool onlyAssembly)
        {
            if (AllData.iSwApp == null) return null;
            EnumDocuments2 enumDoc = AllData.iSwApp.EnumDocuments2();
            if (enumDoc == null) return null;

            ModelDoc2 swModel = null;
            int PceltFetched = 0;
            enumDoc.Next(1, out swModel, ref  PceltFetched);
            ArrayList ar = new ArrayList();

            while (swModel != null)
            {
                if (onlyAssembly)//只显示装配 
                {
                    if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY) ar.Add(swModel);
                }
                else
                {
                    ar.Add(swModel);
                }

                //转向下一个
                enumDoc.Next(1, out swModel, ref PceltFetched);
            }
            return ar;
        }
        /// <summary>
        /// 得到所有打开的文件，可以设置 参数（是否只显示装配）
        /// </summary>
        internal List<string> GetAllOpenDocsPath(bool onlyAssembly)
        {
            if (AllData.iSwApp == null) return null;
            EnumDocuments2 enumDoc = AllData.iSwApp.EnumDocuments2();
            if (enumDoc == null) return null;

            ModelDoc2 swModel = null;
            int PceltFetched = 0;
            enumDoc.Next(1, out swModel, ref  PceltFetched);

            List<string> listArr = new List<string>();

            while (swModel != null)
            {
                if (onlyAssembly)//只显示装配 
                {
                    if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY) listArr.Add(swModel.GetPathName());
                }
                else
                {
                    listArr.Add(swModel.GetPathName());
                }

                //转向下一个
                enumDoc.Next(1, out swModel, ref PceltFetched);
            }
            return listArr;
        }


        /// <summary>
        /// 得到一个装配体下所有的文件路径
        /// </summary>
        internal ArrayList GetAllAssemblyDocs(AssemblyDoc swAssem)
        {
            ArrayList arFiles = new ArrayList();

            if (swAssem == null) return arFiles;

            object[] objArr = (object[])swAssem.GetComponents(false);//得到所有的
            foreach (object obj in objArr)
            {
                Component2 comp = (Component2)obj;
                if (!arFiles.Contains(comp.GetPathName()))
                {
                    arFiles.Add(comp.GetPathName());
                }
            }

            return arFiles;
        }
        /// <summary>
        /// 得到一个装配体下所有的零件Component
        /// </summary>
        internal ArrayList GetAllAssemblyComponents(AssemblyDoc swAssem)
        {
            ArrayList arFiles = new ArrayList();

            if (swAssem == null) return arFiles;

            object[] objArr = (object[])swAssem.GetComponents(false);//得到所有的
            foreach (object obj in objArr)
            {
                Component2 comp = (Component2)obj;
                if (!arFiles.Contains(comp))
                {
                    arFiles.Add(comp);
                }
            }

            return arFiles;
        }

        /// <summary>
        /// 得到零件的体积
        /// </summary>
        internal double GetPartVolumn(ModelDoc2 Part, int idecimal)
        {
            MassProperty mass = (MassProperty)Part.Extension.CreateMassProperty();
            //mass.UseSystemUnits = false;
            double Volu = mass.Volume;//立方毫米
            Volu = Math.Round(Volu, idecimal);
            return Volu;
        }
        /// <summary>
        /// 得到质量， 小数位数
        /// </summary>
        internal double GetPartWeight(ModelDoc2 Part, int idecimal)
        {
            MassProperty mass = (MassProperty)Part.Extension.CreateMassProperty();
            //mass.UseSystemUnits = false;
            double Volu = mass.Mass;
            Volu = Math.Round(Volu, idecimal);
            return Volu;



        }
        /// <summary>
        /// 得到零件的表面积
        /// </summary>
        internal double GetPartSurfaceArea(ModelDoc2 Part, int idecimal)
        {
            return 0;


        }
        internal string GetPartBoxXYZ(ModelDoc2 Part, Component2 comp, Body2 swBody, Feature swfeat)
        {

            return "";

        }
        /// <summary>
        /// 得到弯管，工字钢的表示方式，没有返回""
        /// </summary>
        internal string GetPartBoxXYZ_ref(ModelDoc2 Part, Body2 swBody, bool bCheckSweep, ref double dbVolume)
        {
            //是否要检查焊件轮廓
            if (bCheckSweep == false) return "";

            //进一步检查是不是弯管（扫描）
            Object[] ftArr = (Object[])swBody.GetFeatures();
            if (ftArr == null || ftArr.Length == 0 || ftArr[0] == null) return "";

            Feature ft = (Feature)ftArr[0];
            string sType = ft.GetTypeName2();
            if (sType.ToLower() == "sweep")//扫描特征，如圆管，弯管 
            {
                #region
                double LTotal = 0; //弯管长度
                double dmin = 0;//内径
                double dmax = 0;//外径
                //得到扫描特征数据
                SweepFeatureData swSweep = (SweepFeatureData)ft.GetDefinition();

                //计算弯管内外直径
                Feature swProfFeat = (Feature)swSweep.Profile;//扫描的轮廓，外形草图特征
                Sketch swProfSketch = (Sketch)swProfFeat.GetSpecificFeature(); //草图
                if (swProfFeat != null)
                {
                    object[] objArr = (object[])swProfSketch.GetSketchSegments();
                    foreach (object oarr in objArr)
                    {
                        SketchSegment sssg = (SketchSegment)oarr;
                        Curve swCV = (Curve)sssg.GetCurve();
                        if (swCV.IsCircle())
                        {
                            double[] dbArr = (double[])swCV.CircleParams;
                            double dbR = dbArr[6] * 2;//得出直径来

                            double cc1 = swSweep.GetWallThickness(true); //壁厚-前(外）
                            double cc2 = swSweep.GetWallThickness(false); //壁厚-后（内)

                            dmax = dbR + cc1 * 2;
                            dmin = dbR - cc2 * 2;
                            dmax = Math.Round(dmax * 1000d, PSetUp.iDecimals);
                            dmin = Math.Round(dmin * 1000d, PSetUp.iDecimals);

                            break;
                        }
                    }
                }

                //计算弯管长度
                if (Part == null)
                {
                    Part = (ModelDoc2)AllData.iSwApp.ActiveDoc;
                }
                bool bRet = swSweep.AccessSelections(Part, null);//必须加上这一句下面的草图才能找到，否则找不到
                Feature swPathFeat = (Feature)swSweep.Path;
                Sketch swPathSketch = (Sketch)swPathFeat.GetSpecificFeature();
                if (swPathSketch != null)
                {
                    object[] objArr = (object[])swPathSketch.GetSketchSegments();
                    foreach (object ooog in objArr)
                    {
                        SketchSegment sssg = (SketchSegment)ooog;
                        LTotal += sssg.GetLength();
                    }
                    LTotal = Math.Round(LTotal * 1000d, PSetUp.iDecimals );
                }

                //释放选择
                swSweep.ReleaseSelectionAccess();

                //返回字符串
                if (LTotal > 0 && dmin > 0 && dmax > 0)
                {
                    return "-Φ" + dmin.ToString() + "/Φ" + dmax.ToString() + "-" + LTotal.ToString();
                }
                #endregion
            }
            else if (sType.ToLower() == "weldmemberfeat")//工字钢，槽钢等焊接件
            {
                #region
                StructuralMemberFeatureData swWeldFeatData = (StructuralMemberFeatureData)ft.GetDefinition();

                string spath = swWeldFeatData.WeldmentProfilePath; //获取焊件轮廓路径，如：D:\\槽钢GB／T 707-1988\\280X84X9.5 型号：28b#.sldlfp
                if (spath.IndexOf('.') != -1)
                {
                    spath = spath.Remove(spath.LastIndexOf('.'));
                }
                if (spath.IndexOf('\\') != -1)
                {
                    spath = spath.Substring(spath.LastIndexOf('\\') + 1);
                }
                if (spath.IndexOf("型号：") != -1)
                {
                    spath = spath.Substring(spath.IndexOf("型号：") + 3);
                }
                spath = spath.Replace("#", "");

                SketchSegment swSegg = swWeldFeatData.IGetPathSegments(0);//
                double LTotal = swSegg.GetLength(); //获取焊件长度
                LTotal = Math.Round(LTotal * 1000d, PSetUp.iDecimals);

                //返回字符串
                if (LTotal > 0)
                {
                    return "ㄈ" + spath + "-" + LTotal.ToString();
                }

                ////再进一步判断是不是圆管，这里临时没必要了，只要是钢结构，就好了
                ////得到轮廓草图，然后得到草图的线条
                //Feature ftsub = ft.IGetFirstSubFeature();//得到子特征
                //while (ftsub != null)
                //{
                //    string sName = ftsub.Name;
                //    string tName = ftsub.GetTypeName2().ToLower();
                //    if (tName == "profilefeature")
                //    {
                //        break;
                //    }
                //    else
                //    {
                //        ftsub = (Feature)ftsub.GetNextSubFeature();
                //    }
                //}
                //if (ftsub != null)
                //{
                //    Sketch skch = (Sketch)ftsub.GetSpecificFeature2();
                //    object[] objArr = (object[])skch.GetSketchSegments();
                //    if (objArr.Length < 4)//两条曲线肯定是圆
                //    {
                //        bool bAllCircle = true; //判断这个焊件轮廓是不是圆
                //        foreach (object ooog in objArr)
                //        {
                //            SketchSegment sssg = (SketchSegment)ooog;
                //            Curve swCV = (Curve)sssg.GetCurve();
                //            if (!swCV.IsCircle())
                //            {
                //                bAllCircle = false;
                //            }
                //        }
                //        if (bAllCircle)
                //        {
                //            //这个焊件轮廓是圆
                //        }
                //        else
                //        {
                //            //这个焊件轮廓不是圆
                //        }
                //    }
                //}
                #endregion
            }
            else
            {
                //如果特征有四个面，而且两个平面，两个圆柱面，很可能是圆管
                #region
                //string s = ft.Name;
                int iFace = swBody.GetFaceCount();
                if (iFace == 4)
                {
                    Face2 swPlane1 = null;
                    Face2 swPlane2 = null;
                    Surface swCylinder1 = null;
                    Surface swCylinder2 = null;
                    object[] faceArr = (object[])swBody.GetFaces();
                    foreach (object oof in faceArr)
                    {
                        Face2 swface = (Face2)oof;
                        Surface surFc = (Surface)swface.GetSurface();
                        if (surFc.IsCylinder())
                        {
                            if (swCylinder1 == null) swCylinder1 = surFc;
                            else if (swCylinder2 == null) swCylinder2 = surFc;
                        }
                        else if (surFc.IsPlane())
                        {
                            if (swPlane1 == null) swPlane1 = swface;
                            else if (swPlane2 == null) swPlane2 = swface;
                        }
                    }

                    if (swPlane1 != null && swPlane2 != null && swCylinder1 != null && swCylinder2 != null)
                    {
                        //得到圆柱面的半径 
                        object objArr1 = swCylinder1.CylinderParams;//圆柱面的所有参数
                        double[] dbArr1 = (double[])objArr1;
                        double dbRadius1 = Math.Round(dbArr1[6] * 1000d, PSetUp.iDecimals);
                        dbRadius1 = dbRadius1 * 2.0;//半径到直径


                        //得到圆柱面的半径 
                        object objArr2 = swCylinder2.CylinderParams;//圆柱面的所有参数
                        double[] dbArr2 = (double[])objArr2;
                        double dbRadius2 = Math.Round(dbArr2[6] * 1000d, PSetUp.iDecimals);
                        dbRadius2 = dbRadius2 * 2.0;//半径到直径


                        //得到圆管的长度
                        object ptArr1;
                        object ptArr2;
                        double dist = Part.ClosestDistance(swPlane1, swPlane2, out ptArr1, out ptArr2);
                        dist = Math.Round(dist * 1000d, PSetUp.iDecimals);
                        //返回结果
                        return dist.ToString() + "-Φ" + Math.Min(dbRadius1, dbRadius2).ToString() + "/Φ" + Math.Max(dbRadius1, dbRadius2).ToString();
                    }
                }
                #endregion
            }

            return "";
        }



        /// <summary>
        /// 得到模型对应的工程图的纸张大小, 返回如A0，A1
        /// </summary>
        internal string GetPartDwgSize(ModelDoc2 Part, Component2 comp)
        {
            List<string> strOpenArr= this.GetAllOpenDocsPath(false);

            DrawingDoc swDraw = null;
            bool bClosePart = true;
            
            //首先得到模型对应的图纸
            string strDwgFullPath = "";
            if (Part != null)
            {
                strDwgFullPath = Part.GetPathName();
            }
            else if (comp != null)
            {
                strDwgFullPath = comp.GetPathName();
            }

            if (strDwgFullPath.ToLower().EndsWith(".slddrw"))
            {
                bClosePart = false;
                if (Part != null)
                {
                    swDraw = (DrawingDoc)Part;
                }
                else if (comp != null)
                {
                    swDraw = (DrawingDoc)comp.GetModelDoc2();
                }
            }
            else
            {
                strDwgFullPath = strDwgFullPath.Remove(strDwgFullPath.LastIndexOf('.'));
                strDwgFullPath = strDwgFullPath + ".slddrw";
                if (File.Exists(strDwgFullPath))
                {
                    //看看在打开的文件中是否有这个图纸
                    Part = this.GetModelDoc2FromOpenDocs(strDwgFullPath);
                    if (Part == null)
                    {
                        int iErrors = 0;
                        int iWarning = 0;
                        swDraw = (DrawingDoc)AllData.iSwApp.OpenDoc6(strDwgFullPath, 3, 4, "", ref iErrors, ref iWarning);
                    }
                    else
                    {
                        swDraw = (DrawingDoc)Part;
                        bClosePart = false;
                    }
                }
            }

            if (swDraw != null)
            {
                Sheet sht = (Sheet)swDraw.GetCurrentSheet();
                if (sht == null) return "";

                double SheetWidth = 0;
                double SheetHeight = 0;

                //SolidWorks2007
                int a = sht.GetSize(ref SheetWidth, ref SheetHeight);
                if (bClosePart && !strOpenArr.Contains((swDraw as ModelDoc2).GetPathName()))
                {
                    AllData.iSwApp.CloseDoc((swDraw as ModelDoc2).GetTitle());
                }

                string sname = this.GetDwgSizeNameByDwgSizeID(a);

                return sname;
            }

            return "";
        }



        /// <summary>
        /// 判断一下是不是焊件零件
        /// </summary>
        internal bool  bIsWeldment(IModelDoc2 swModel, Component2 comp )
        {
            if (comp != null)
            {
                swModel = (IModelDoc2)comp.GetModelDoc2();
            }
            
            if (swModel != null)
            {
                int iType = swModel.GetType();
                if (iType == (int)swDocumentTypes_e.swDocPART)
                {
                    PartDoc swPart = (PartDoc)swModel;

                    return swPart.IsWeldment();
                }
            }

            return false;
        }
        /// <summary>
        /// 得到所有的切割清单特征
        /// </summary>
        internal List<Feature> getAllWeldmentFeature(ModelDoc2 swModel)
        {
            List<Feature> LFT = new List<Feature>();

            //Get Weldment Cut-list Feature and Annotations
            Feature thisFeat = (Feature)swModel.FirstFeature();
            while (thisFeat != null)
            {
                string s = thisFeat.GetTypeName2();

                //子特征
                Feature subFt = (Feature)thisFeat.GetFirstSubFeature();
                while (subFt != null)
                {
                    string SN2 = subFt.Name;
                    string ST2 = subFt.GetTypeName2();
                    if (ST2 == "CutListFolder" || ST2 == "SubWeldFolder")
                    {
                        BodyFolder bFolder = (BodyFolder)subFt.GetSpecificFeature2();
                        if (bFolder.GetBodyCount() > 0)
                        {
                            bool b = bFolder.UpdateCutList();//先更新一下
                            LFT.Add(subFt);
                            //int a = bFolder.GetBodyCount();
                        }
                    }

                    subFt = (Feature)subFt.GetNextSubFeature();
                }

                thisFeat = (Feature)thisFeat.GetNextFeature();
            }

            return LFT;
        }



        /// <summary>
        /// 由两个点得到两点之间的距离,注意这里是三坐标（x,y,z)
        /// </summary>
        internal string GetDistanceByTwoVertex(double[] dbA, double[] dbB)
        {
            try
            {
                double ax = dbA[0];
                double ay = dbA[1];
                double az = dbA[2];
                double bx = dbB[0];
                double by = dbB[1];
                double bz = dbB[2];
                double DX = Math.Abs(ax - bx);
                double DY = Math.Abs(ay - by);
                double DZ = Math.Abs(az - bz);
                double distance = Math.Sqrt(DX * DX + DY * DY);
                distance = Math.Sqrt(distance * distance + DZ * DZ);
                distance = distance * 1000;//毫米
                distance = distance + 0.4999999;

                string Ret = distance.ToString();
                if (Ret.IndexOf('.') != -1) Ret = Ret.Remove(Ret.IndexOf('.'));
                return Ret;
            }
            catch
            {
                return "0";
            }
        }

    }




    /// <summary>
    /// 修改了材料,单重也变化
    /// </summary>
    internal class SolidWorksMethod2
    {
        /// <summary>
        /// 验证配置名称是否正确
        /// 如果不正确，返回一个正确的,即在名称后加-2,-3,-4....
        /// </summary>
        internal string CheckConfigName(ModelDoc2 swModel, string configName)
        {
            string[] strArr = (string[])swModel.GetConfigurationNames();

            ArrayList ar = new ArrayList();
            foreach (string s in strArr)
            {
                ar.Add(s.ToLower());
            }

            if (ar.Contains(configName.ToLower()))
            {
                for (int i = 2; i < 10000; i++)
                {
                    string s = configName + "-" + i.ToString();
                    if (!ar.Contains(s.ToLower()))
                    {
                        configName = s; break;
                    }
                }
            }
            return configName;
        }


        /// <summary>
        /// 根据特征类别名称，在特征树中得到一个特征( 特征类别名称，是否查找子特征）
        /// </summary>
        internal Feature getFeatureByFeatureTypeName(ModelDoc2 swModel, string strFTypeName, bool bIncludeSubFeature)
        {
            if (swModel == null) return null;

            strFTypeName = strFTypeName.ToLower();

            Feature thisFeat = (Feature)swModel.FirstFeature();
            while (thisFeat != null)
            {
                string sType = thisFeat.GetTypeName2().ToLower();
                if (sType == strFTypeName)
                {
                    return thisFeat;
                }

                //子特征
                if (bIncludeSubFeature)
                {
                    Feature subsubFt = getFeatureByFeatureTypeName_ref(thisFeat, strFTypeName, bIncludeSubFeature);
                    if (subsubFt != null) return subsubFt;
                }

                thisFeat = (Feature)thisFeat.GetNextFeature();
            }

            return null;
        }
        /// <summary>
        /// 根据特征类别名查找子特征
        /// </summary>
        internal Feature getFeatureByFeatureTypeName_ref(Feature thisFeat, string strFTypeName, bool bIncludeSubFeature)
        {
            if (thisFeat == null) return null;

            Feature subFt = (Feature)thisFeat.GetFirstSubFeature();
            while (subFt != null)
            {
                string ST2 = subFt.GetTypeName2().ToLower();
                if (ST2 == strFTypeName)
                {
                    return subFt;
                }
                else
                {
                    Feature subsubFt = getFeatureByFeatureTypeName_ref(subFt, strFTypeName, bIncludeSubFeature);
                    if (subsubFt != null) return subsubFt;
                }
                subFt = (Feature)subFt.GetNextSubFeature();
            }

            return null;
        }


    }













}
