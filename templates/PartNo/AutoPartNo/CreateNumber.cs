using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using System.Collections;
using SolidWorks.Interop.swconst;

namespace Interop.Office.Core
{
    internal partial class CreateNumber : Form
    {
        internal CreateNumber()
        {
            InitializeComponent();
        }
        private SolidWorksMethod swm = new SolidWorksMethod();

        private void CreateNumber_Load(object sender, EventArgs e)
        {
            //新方法,检查看看还能不能用
            if (Dbtool2.hasconn(true, this.FindForm()) == false)
            { }

            int iType = Interop.Office.Core.Properties.Settings.Default.swBomBomTemplate;
            if (iType == 1)
            {
                this.rdoTemp1.Checked = true;
            }
            else if (iType == 2)
            {
                this.rdoTemp2.Checked = true;
            }
            else if (iType == 3)
            {
                this.rdoTemp3.Checked = true;
            }
            else if (iType == 4)
            {
                this.rdoTemp4.Checked = true;
            }
            this.cmbPPtName.Text = Interop.Office.Core.Properties.Settings.Default.CreateNbr_PPtName;
            this.cmbStartIndex.Text = Interop.Office.Core.Properties.Settings.Default.CreateNbr_StartIndex;
        }
        private void CreateNumber_FormClosing(object sender, FormClosingEventArgs e)
        {
            //保存
            int i = 1;
            if (this.rdoTemp2.Checked) i = 2;
            else if (this.rdoTemp3.Checked) i = 3;
            else if (this.rdoTemp4.Checked) i = 4;

            Interop.Office.Core.Properties.Settings.Default.swBomBomTemplate = i;
            Interop.Office.Core.Properties.Settings.Default.CreateNbr_PPtName = this.cmbPPtName.Text;
            Interop.Office.Core.Properties.Settings.Default.CreateNbr_StartIndex = this.cmbStartIndex.Text;
            Interop.Office.Core.Properties.Settings.Default.Save();

            //System.Environment.Exit(System.Environment.ExitCode);

        }



        //生成零件号
        private void btnOK_Click(object sender, EventArgs e)
        {
            this.CreatePartNumber();
        }
        //重排序零件号
        private void btnReSort_Click(object sender, EventArgs e)
        {
            this.ReSortAllView();
        }





        /// <summary>
        /// 生成零件号
        /// </summary>
        private void CreatePartNumber()
        {
            ModelDoc2 swModel = (ModelDoc2)AllData.iSwApp.ActiveDoc;
            if (swModel == null) return;

            DrawingDoc swDraw = (DrawingDoc)swModel;
            if (swDraw == null) return;


            //序号为零件属性，还是项目号, 还是文字
            int iTextContent = -1;
            string strText = "";
            if (this.rdoTemp1.Checked)
            {
                iTextContent = (int)swBalloonTextContent_e.swBalloonTextItemNumber;
                strText = "";
            }
            else if (this.rdoTemp2.Checked)
            {
                iTextContent = (int)swBalloonTextContent_e.swBalloonTextCustomProperties;
                strText = Interop.Office.Core.Properties.Settings.Default.CreateNbr_PPtName; //"序号";
            }
            else if (this.rdoTemp3.Checked)
            {
                iTextContent = (int)swBalloonTextContent_e.swBalloonTextCustom;
                strText = "0";
            }
            else if (this.rdoTemp4.Checked)
            {
                iTextContent = (int)swBalloonTextContent_e.swBalloonTextQuantity;
                strText = "0";
            }


            //以正方形分布
            int iLayout = (int)swBalloonLayoutType_e.swDetailingBalloonLayout_Square;
            //下划线
            int iStyle = (int)swBalloonStyle_e.swBS_Underline;
            //下划线长度
            int iSize = (int)swBalloonFit_e.swBF_Tightest;


            //如果用户没有选中一个视图，就选中所有的视图
            SolidWorks.Interop.sldworks.View swView = (SolidWorks.Interop.sldworks.View)swDraw.ActiveDrawingView;
            if (swView == null)//对所有的视图进行重排序
            {
                swModel.ClearSelection2(true);
                swView = (SolidWorks.Interop.sldworks.View)swDraw.GetFirstView();
                while (swView != null)
                {
                    swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, true, 0, null, 0);
                    swView = (SolidWorks.Interop.sldworks.View)swView.GetNextView();
                }
            }

            //自动添加序号
            swDraw.AutoBalloon3(iLayout, true, iStyle, iSize, iTextContent, strText, iTextContent, strText, "");


            //Notes = Part.AutoBalloon3(1, True, 10, 2, 1, "", 2, "", "-无-")

            //swDraw.AutoBalloon3(iLayout, true, iStyle, -1, iTextContent, "xh", iTextContent, "xh", "");
            //swDraw.AutoBalloon3(1, true, 10, 2, 1, "", 2, "", "-无-");

            //.AutoBalloon3(1, True, 10, 2, 1, "", 2, "", "-无-")
            ////SolidWorks2006
            //swDraw.AutoBalloon2((int)swBalloonLayoutType_e.swDetailingBalloonLayout_Top, true);
        }






        #region//重排序零件号///////////////////////////////////////////////////////////


        /// <summary>
        /// 如果有一个视图为选中状态,则对这个视图进行重排序,
        /// 否则,对一个图纸下的所有视图进行重排序
        /// </summary>
        private void ReSortAllView()
        {
            this.arCheckRepect.Clear();
            if (this.rdoTemp1.Checked)
            {
                if (StringOperate.Alert("序号为项目号时，无法重排序，要继续吗？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            }

            ModelDoc2 swModel = (ModelDoc2)AllData.iSwApp.ActiveDoc;
            if (swModel == null) return;

            DrawingDoc swDraw = (DrawingDoc)swModel;
            SelectionMgr SelMgr = (SelectionMgr)swModel.SelectionManager;
            swModel.ClearSelection2(true);

            //101，$数量×"序号"，$数量×"属性名"
            string strTxt = this.cmbStartIndex.Text;
            int abc = 1;
            bool bText = false;
            try
            {
                abc = Convert.ToInt16(this.cmbStartIndex.Text);//序号从一开始
            }
            catch
            {
                bText = true;
            }

            SolidWorks.Interop.sldworks.View swView = (SolidWorks.Interop.sldworks.View)swDraw.ActiveDrawingView;
            if (swView == null)//对所有的视图进行重排序
            {
                swView = (SolidWorks.Interop.sldworks.View)swDraw.GetFirstView();
                //遍历视图
                while (swView != null)
                {
                    swDraw.ActivateView(swView.Name);

                    if (bText)
                    {
                        SaidiResortOneView(swView, swModel);//对这个视图进行排序
                    }
                    else
                    {
                        abc = ReSortOneView(abc, swView, swModel);//对这个视图进行排序
                    }

                    //转到下一个视图
                    swView = (SolidWorks.Interop.sldworks.View)swView.GetNextView();
                }
            }
            else//只对当前视图进行重排序
            {
                if (bText)
                {
                    SaidiResortOneView(swView, swModel);//对这个视图进行排序
                }
                else
                {
                    ReSortOneView(abc, swView, swModel);
                }
            }

            //重建,以更新明细表
            swModel.EditRebuild3();

            this.arCheckRepect.Clear();
        }
        /// <summary>
        /// 对于一个视图进行重排序,返回序号已经排到哪了
        /// </summary>
        private int ReSortOneView(int abc, SolidWorks.Interop.sldworks.View swView, ModelDoc2 swModel)
        {
            //先得到视图的对角线的三点坐标和比例//初始化视图
            this.GetViewThreePointAndScale(swView);

            //排序用
            ArrayList arSort = new ArrayList();
            //保存数据
            Hashtable ht = new Hashtable();
            //保存数据
            Hashtable ht2 = new Hashtable();
            //是否有在右上方的//如果没有则为0,否则保在最上方的最大值,如果最大值<0.15,说明虽然有在上方的,但可以忽略,因为离左上角很进.
            double dAtTopMaxValue = 0;


            //遍历注释
            Note not = (Note)swView.GetFirstNote();
            while (not != null)
            {
                if (not.IsBomBalloon())//只有符号这个条件才可以
                {
                    Annotation ann = (Annotation)not.GetAnnotation();//必须要有Annotation
                    if (ann == null) goto ABC;
                    object[] obArr = (object[])ann.GetAttachedEntities2();//必须要有Entities
                    if (obArr == null || obArr.Length == 0) goto ABC;

                    Entity ent = (Entity)obArr[0];
                    Component2 comp = (Component2)ent.GetComponent();

                    //看看是否有重复的，一个配置只允许标注一次，如果需要多次标注，只能用不同的配置。
                    string FullPath = comp.GetPathName() + comp.ReferencedConfiguration;
                    if (arCheckRepect.Contains(FullPath))
                    {
                        goto ABC;
                    }
                    else
                    {
                        arCheckRepect.Add(FullPath);
                    }

                    //string txt = not.GetText();
                    double[] dbArr = (double[])not.GetTextPoint2();

                    //计算角度
                    double dAngle = GetAngelByNotePoint(dbArr[0], dbArr[1]);

                    if (CheckUnderOrUpTheLine(dbArr[0], dbArr[1]) == true)//在上方//升序
                    {
                        if (dAtTopMaxValue < dAngle) dAtTopMaxValue = dAngle;//保存这个角度//这个角度是上方角度的最大值
                    }
                    else//下方的//降序
                    {
                        dAngle = 15 - dAngle;//dAngle一般不会大于二
                    }

                    arSort.Add(dAngle);//排序
                    ht.Add(dAngle, not);//保存数据
                    ht2.Add(dAngle, comp);//防止再算一遍
                }
            ABC:
                not = (Note)not.GetNext();
            }

            ///////////////////////////////////////////////////////////////////////////////////////////////
            arSort.Sort();//排序
            if (dAtTopMaxValue < 0.15) arSort.Reverse();//如果上方没有,或靠近左上角的地方有一两个(Angle<0.15),下方的也不会按DESC排序
            /////////////////////////////////////////////////////////////////////////////////////////////////


            //序号为项目号，还是零件属性，还是文本
            int iTextContent = -1;
            string strText = "";
            if (this.rdoTemp1.Checked)
            {
                iTextContent = (int)swDetailingNoteTextContent_e.swDetailingNoteTextItemNumber;
                strText = "";
            }
            else if (this.rdoTemp2.Checked)
            {
                iTextContent = (int)swDetailingNoteTextContent_e.swDetailingNoteTextCustomProperty;
                strText = "$PRPMODEL:\"序号\"";
            }
            else if (this.rdoTemp3.Checked)
            {
                iTextContent = (int)swDetailingNoteTextContent_e.swDetailingNoteTextCustom;
                strText = "";
            }
            else if (this.rdoTemp4.Checked)
            {
                iTextContent = (int)swBalloonTextContent_e.swBalloonTextQuantity;
                strText = "0";
            }


            //开始修改
            for (int i = 0; i < arSort.Count; i++)
            {
                Note notsub = (Note)ht[arSort[i]];
                Component2 comp = (Component2)ht2[arSort[i]];

                //选中这个注释
                swModel.Extension.SelectByID2(notsub.GetName() + "@" + swView.Name, "NOTE", 0, 0, 0, true, 0, null, 0);

                if (this.rdoTemp2.Checked)
                {
                    this.swm.SetAtrByCompAndAtrName(comp, "序号", abc);//设置属性,并保存到零件
                }
                else
                {
                    strText = abc.ToString();//序号
                }

                //notsub.PropertyLinkedText = "$PRPMODEL:\"序号\"";//这样可能也可以
                notsub.SetBomBalloonText(iTextContent, strText, iTextContent, strText);

                //////////////////////////////////////////////////////////////////////
                //检查看看是否正确,如果不正确,再设置，这里没起到任何作用
                string txt = notsub.GetText();
                if (txt != abc.ToString())
                {
                    notsub.SetText(abc.ToString());
                }
                ///////////////////////////////////////////////////////////////////////

                abc++;
            }
            return abc;
        }
        /// <summary>
        /// 如果不是数字冲排序，而是重新指定文字，特别是一些复杂的组合，如【项目数X代号】
        /// </summary>
        private void SaidiResortOneView(SolidWorks.Interop.sldworks.View swView, ModelDoc2 swModel)
        {
            //101，$数量×$图号  ，$数量（$属性名）等等各种格式
            string strText = this.cmbStartIndex.Text;

            //是否需要进行赛迪处理-页码-页数的时候，如果页数=1不显示，否则显示页码
            bool bSaidiCheck = false;


            char[] carr = new char[] { '(', ')', '×', 'x', 'X', '-', '_', '=', '（', '）', '[', ']', ' ','/','"' };
            string[] strArr = strText.Split(carr);

            //执行完切割后，要去掉引号，因为下面还要加引号
            strText = strText.Replace("\"", "");

            foreach (string s in strArr)
            {
                if (s != "$项目数" && s !="$项目号" && s.Length>1)
                {
                    strText = strText.Replace(s, "$PRPMODEL:\"" + s.Trim(new char[]{'$'}) + "\"");//如"$项目数×$PRPMODEL:\"名称\"";
                }
            }
            bool bXMS = (strText.IndexOf("$项目数") != -1);
            bool bXMH = (strText.IndexOf("$项目号") != -1);

            int iTextQuantity = (int)swDetailingNoteTextContent_e.swDetailingNoteTextQuantity;//数量
            int iTextNumber = (int)swDetailingNoteTextContent_e.swDetailingNoteTextItemNumber;//项目号
            int iTextProperty = (int)swDetailingNoteTextContent_e.swDetailingNoteTextCustomProperty;//属性
            
            //遍历注释
            Note not = (Note)swView.GetFirstNote();
            while (not != null)
            {
                if (not.IsBomBalloon())//只有符号这个条件才可以
                {
                    Annotation ann = (Annotation)not.GetAnnotation();//必须要有Annotation
                    if (ann == null) goto ABC;
                    object[] obArr = (object[])ann.GetAttachedEntities2();//必须要有Entities
                    if (obArr == null || obArr.Length == 0) goto ABC;

                    Entity ent = (Entity)obArr[0];
                    Component2 comp = (Component2)ent.GetComponent();

                    //看看是否有重复的，一个配置只允许标注一次，如果需要多次标注，只能用不同的配置。
                    string FullPath = comp.GetPathName() + comp.ReferencedConfiguration;

                    //先复制一份值
                    string thisTxt = strText;
                    //如果有项目数,读取项目数
                    if (bXMS)
                    {
                        //选中这个注释，然后设为项目数
                        swModel.Extension.SelectByID2(not.GetName() + "@" + swView.Name, "NOTE", 0, 0, 0, true, 0, null, 0);
                        not.SetBomBalloonText(iTextQuantity, "", iTextQuantity, "");
                        string txt = not.GetText();
                        if (txt.Length > 0)
                        {
                            thisTxt = thisTxt.Replace("$项目数", txt);
                        }
                    }
                    //如果有项目号，读取项目号
                    if (bXMH)
                    {
                        //选中这个注释，然后设为项目号
                        swModel.Extension.SelectByID2(not.GetName() + "@" + swView.Name, "NOTE", 0, 0, 0, true, 0, null, 0);
                        not.SetBomBalloonText(iTextNumber, "", iTextNumber, "");
                        string txt = not.GetText();
                        if (txt.Length > 0)
                        {
                            thisTxt = thisTxt.Replace("$项目号", txt);
                        }
                    }

                    //如果是赛迪，当页数>1时，加上-页码，否侧不加-页码
                    //先列出-页数-页码来，然后再去掉-1-1
                    if (bSaidiCheck)
                    {
                        //先设置页数属性链接
                        not.SetBomBalloonText(iTextProperty, "$PRPMODEL:\"页数\"", iTextProperty, "$PRPMODEL:\"页数\"");
                        string s = not.GetText();
                        if (s.Length > 0 && s.Length <4 && s != "1" && s != "0")
                        {
                            thisTxt = thisTxt.Replace("-$PRPMODEL:\"页数\"", "");
                        }
                        else
                        {
                            thisTxt = thisTxt.Replace("-$PRPMODEL:\"页码\"-$PRPMODEL:\"页数\"", "");
                        }
                    }

                    //最后设置新的文本
                    not.SetBomBalloonText(iTextProperty, thisTxt, iTextProperty, thisTxt);

                }
            ABC:
                not = (Note)not.GetNext();
            }
        }


        //有可能在下一个视图中有重复的
        private ArrayList arCheckRepect = new ArrayList();
        //默认的比例,一个视图一个样//即(X2-X1)/(Y2-Y1)
        private double scale = 0;
        //还有左上角到中心点的距离[6]
        private double distAC = 0;
        //线的三点(左上角[0][1],右下角[2][3],中心点[4][5])
        private double[] dbArrLine = new double[6];

        //得到线的三点,和比例//对一个视图的数据初始化
        private void GetViewThreePointAndScale(SolidWorks.Interop.sldworks.View swView)
        {
            double[] dbViewArr = (double[])swView.GetOutline();//得到两个点坐标:左下角,右上角

            this.dbArrLine[0] = dbViewArr[0];//左上角
            this.dbArrLine[1] = dbViewArr[3];
            this.dbArrLine[2] = dbViewArr[2];//右下角
            this.dbArrLine[3] = dbViewArr[1];
            this.dbArrLine[4] = (dbViewArr[0] + dbViewArr[2]) / 2;//中心点
            this.dbArrLine[5] = (dbViewArr[3] + dbViewArr[1]) / 2;

            //X长度//中心点到左上角
            double x0 = this.dbArrLine[4] - this.dbArrLine[0];
            double y0 = this.dbArrLine[5] - this.dbArrLine[1];

            //左上角到中心点的距离,用来计算角度//三边计算角度
            this.distAC = Math.Sqrt(x0 * x0 + y0 * y0);

            this.scale = x0 / y0;//比例关系
        }
        //判断一个点是在右上方吗?
        private bool CheckUnderOrUpTheLine(double x, double y)
        {
            //得到在线上任意一点x的Y坐标y1
            //公式:(x-this.dbArrLine[0])/(y-this.dbArrLine[1]) = (dbArrLine[2] - dbArrLine[0])/(dbArrLine[3] - dbArrLine[1])
            //得到:(x-this.dbArrLine[0])/(y-this.dbArrLine[1]) = this.scale
            double y1 = (x - this.dbArrLine[0]) / this.scale + this.dbArrLine[1];

            if (y1 > y)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        //由一个NOte的坐标得到它的Cos值
        private double GetAngelByNotePoint(double x, double y)
        {
            double a = this.distAC;//左上角到中心点的距离
            double b = this.swm.GetTwoPointDistinct(x, y, this.dbArrLine[4], this.dbArrLine[5]);//NOte到中心点
            double c = this.swm.GetTwoPointDistinct(x, y, this.dbArrLine[0], this.dbArrLine[1]);//Note到左上角

            x = (a * a + b * b - c * c) / (2 * a * b);
            //StringOperate.Alert(a.ToString() + "\n" + b.ToString() + "\n" + c.ToString() + "\n" + x.ToString() + "\n" + Math.Acos(x).ToString());

            return Math.Acos(x);
        }



        #endregion//end-----------------------------------------------------------












        #region//重排序零件号----原有的方法//////////////////////////////////////////////////////////////


        /// <summary>
        /// 得到所有的Note(与Component关联的),将Note的序号写入到Part的序号属性中
        /// </summary>
        private void GetAllPartSortNumber(ModelDoc2 swModel)
        {
            //DrawingDoc swDraw = (DrawingDoc)swModel;
            ////SelectionMgr SelMgr = (SelectionMgr)swModel.SelectionManager;
            //SolidWorks.Interop.sldworks.View swView = (SolidWorks.Interop.sldworks.View)swDraw.GetFirstView();
            ////遍历视图
            //while (swView != null)
            //{/////////////////////////////////////////////////
            //    //遍历注释
            //    Note not = (Note)swView.GetFirstNote();
            //    while (not != null)
            //    {////////////////////////////////////////////
            //        string txt = not.GetText();
            //        try
            //        {
            //            int a = Convert.ToInt32(txt);
            //            //如果不是数字,这一步会结束
            //            Annotation ann = (Annotation)not.GetAnnotation();
            //            object[] obArr = (object[])ann.GetAttachedEntities2();
            //            object oc = obArr[0];
            //            Entity ent = (Entity)oc;
            //            Component2 comp = (Component2)ent.GetComponent();

            //            this.SetAtrByCompAndAtrName("序号", comp, txt, (int)swCustomInfoType_e.swCustomInfoNumber);
            //        }
            //        catch
            //        { }

            //        not = (Note)not.GetNext();
            //    }///////////////////////////////////////////

            //    //转到下一个视图
            //    swView = (SolidWorks.Interop.sldworks.View)swView.GetNextView();
            //}/////////////////////////////////////////////////
        }


        /// <summary>
        /// 对零件的编号重新排序,但是要想显示更新(连接到属性),要手工再设置一次
        /// </summary>
        private void ReSortNoteNumber()
        {
            //ModelDoc2 swModel = (ModelDoc2)AllData.iSwApp.ActiveDoc;
            //DrawingDoc swDraw = (DrawingDoc)swModel;
            //SelectionMgr SelMgr = (SelectionMgr)swModel.SelectionManager;

            //bool balert = false;//最后出个提示
            //int abc = 1;//序号从一开始
            //swModel.ClearSelection2(true);

            //SolidWorks.Interop.sldworks.View swView = (SolidWorks.Interop.sldworks.View)swDraw.GetFirstView();
            ////遍历视图
            //while (swView != null)
            //{
            //    swDraw.ActivateView(swView.Name);

            //    Hashtable htAllNote = GetAllNote(swModel, swView, false);
            //    ArrayList arNote = SortNote(htAllNote, 0);
            //    //显示这个视图下有多少个零件号
            //    //StringOperate.Alert(htAllNote.Count.ToString());
            //    //StringOperate.Alert(arNote.Count.ToString());
            //    foreach (object o in arNote)
            //    {
            //        Note not = (Note)o;
            //        Annotation ann = (Annotation)not.GetAnnotation();

            //        object[] obArr = (object[])ann.GetAttachedEntities2();
            //        object oc = obArr[0];
            //        Entity ent = (Entity)oc;
            //        Component2 comp = (Component2)ent.GetComponent();

            //        //设置属性,并保存到零件
            //        SetAtrByCompAndAtrName("序号", comp, abc.ToString());

            //        //选中这个注释
            //        swModel.Extension.SelectByID2(not.GetName() + "@" + swView.Name, "NOTE", 0, 0, 0, true, 0, null, 0);


            //        not.SetText(abc.ToString());
            //        string txt = not.GetText();//看看是否出提示
            //        if (txt != abc.ToString())
            //        {
            //            balert = true;
            //        }

            //        abc++;
            //    }

            //    //转到下一个视图
            //    swView = (SolidWorks.Interop.sldworks.View)swView.GetNextView();
            //}
            //if (balert)
            //{
            //    StringOperate.Alert("已成功将最新的零件序号写入零件的属性:序号 中! 请手动设置序号的相关属性为:序号!");
            //}
        }
        //第一步,先得到一个视图下所有要重新排序的Note,这里一个视图一组,
        //返回HashTable,Note---Point
        private Hashtable GetAllNote(ModelDoc2 swModel, SolidWorks.Interop.sldworks.View swView, bool OnlySelected)
        {
            Hashtable htRet = new Hashtable();
            //if (OnlySelected)//只选择选中的
            //{
            //    SelectionMgr SelMgr = (SelectionMgr)swModel.SelectionManager;
            //    int Count = SelMgr.GetSelectedObjectCount2(0);//选中的个数
            //    for (int i = 1; i <= Count; i++)
            //    {
            //        try
            //        {
            //            object o = SelMgr.GetSelectedObject6(i, 0);//0:忽略当前的,-1:忽略全部
            //            Note not = (Note)o;
            //            double[] dbArr = (double[])not.GetTextPoint2();

            //            htRet.Add(not, dbArr);
            //        }
            //        catch
            //        { }
            //    }
            //}
            //else//所有的////////////////////////////////////////////////////////////////////////////
            //{
            //    //遍历注释
            //    Note not = (Note)swView.GetFirstNote();
            //    while (not != null)
            //    {
            //        string txt = not.GetText();
            //        try
            //        {
            //            int a = Convert.ToInt32(txt);
            //            //如果不是数字,这一步会结束
            //            Annotation ann = (Annotation)not.GetAnnotation();
            //            object[] obArr = (object[])ann.GetAttachedEntities2();
            //            object oc = obArr[0];
            //            Entity ent = (Entity)oc;
            //            Component2 comp = (Component2)ent.GetComponent();

            //            //添加到哈希表
            //            htRet.Add(not, (double[])not.GetTextPoint2());
            //        }
            //        catch
            //        { }
            //        not = (Note)not.GetNext();
            //    }
            //}
            return htRet;
        }
        //排序Note,先排横向,或竖向的
        //SortNumber:0横向,1竖向,2方形,3圆形,5不规则
        private ArrayList SortNote(Hashtable htNote, int SortNumber)
        {
            ArrayList arRet = new ArrayList();

            //if (SortNumber == 0 || SortNumber == 1)//横向,或纵向
            //{
            //    ArrayList arSort = new ArrayList();
            //    Hashtable ht2 = new Hashtable();

            //    double[] SheetSize = GetDwgSize((ModelDoc2)AllData.iSwApp.ActiveDoc);
            //    double Sx = SheetSize[0] * (-0.1);
            //    double Sy = SheetSize[1] * 1.2;
            //    foreach (object o in htNote.Keys)
            //    {
            //        double[] dbArr = (double[])htNote[o];
            //        double x = dbArr[0];
            //        double y = dbArr[1];

            //        double Dist = Math.Sqrt((x - Sx) * (x - Sx) + (y - Sy) * (y - Sy));
            //        arSort.Add(Dist);
            //        ht2.Add(Dist, o);
            //    }

            //    //排序
            //    arSort.Sort();
            //    foreach (object oa in arSort)
            //    {
            //        arRet.Add(ht2[oa]);
            //    }
            //}
            return arRet;
        }


        #endregion////////////////////////////////////////////////////////////////////////////////////




    }
}