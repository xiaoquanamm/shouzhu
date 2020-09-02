using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using System.IO;
using SolidWorks.Interop.swconst;

namespace Interop.Office.Core
{
    internal partial class PSetUp_Ref : Form
    {
        internal PSetUp_Ref()
        {
            InitializeComponent();
        }
        //取消修改
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void PSetUp_Ref_Load(object sender, EventArgs e)
        {
            this.txtOldName.Text = Interop.Office.Core.Properties.Settings.Default.dwgtemplateCompanyName;
        }

        private StringOperate pm = new StringOperate();

        //开始修改
        private void button1_Click(object sender, EventArgs e)
        {
            this.lblalert.Text = "";
            string strOldName = this.txtOldName.Text.Trim();
            string strNewName = this.txtNewName.Text.Trim() ;

            if(strNewName ==strOldName || strNewName.Trim().Length <4 )
            {
                StringOperate.Alert("请输入您的公司名称，注意不要和原有名称相同。");return ;
            }

            if (StringOperate.Alert("确实要修改所有的迈迪模版吗？这可能需要几分钟的时间。", "修改模版提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string dwgTempPath = AllData.StartUpPath + "\\迈迪模板";
                this.TraverseFile(new DirectoryInfo(dwgTempPath), strOldName, strNewName);


                //最后提示
                this.lblalert.Text = "成功修改了" + iOK.ToString() + " 张图纸模版";
                
                if (iOK > 0)
                {
                    Interop.Office.Core.Properties.Settings.Default.dwgtemplateCompanyName = strNewName;
                    Interop.Office.Core.Properties.Settings.Default.Save();
                }
            }
        }

        //修改的结果
        int iOK = 0;
        int iErr = 0;
        private void TraverseFile(DirectoryInfo dir,string OldName,string NewName)
        {
            foreach (FileInfo fi in dir.GetFiles())
            {
                if (fi.Extension.ToLower() != ".drwdot" && fi.Extension.ToLower() != ".slddrw") continue;//slddrt.

                string str = fi.FullName;
                ModelDoc2 swModel = (ModelDoc2)AllData.iSwApp.OpenDoc(str, (int)swDocumentTypes_e.swDocDRAWING);
                if (swModel == null)
                {
                    iErr++; continue;
                }
                DrawingDoc swDraw = (DrawingDoc)swModel;
                if (swDraw == null)
                {
                    iErr++; continue;
                }

                swDraw.EditTemplate();
                
                SolidWorks.Interop.sldworks.View swView = (SolidWorks.Interop.sldworks.View)swDraw.GetFirstView();
                Note swnote = (Note)swView.GetFirstNote();
                while (swnote != null)
                {
                    string s = swnote.GetText().ToString();
                    if (s == OldName)
                    {
                        swnote.SetText(NewName);
                        iOK++;
                        //这里需要保存
                        swDraw.EditSheet();
                        swModel.Save();
                    }
                    swnote = (Note)swnote.GetNext();
                }
                AllData.iSwApp.CloseDoc(swModel.GetTitle());
            }

            foreach (DirectoryInfo dirsub in dir.GetDirectories())
            {
                TraverseFile(dirsub,OldName ,NewName);
            }
        }
    }
}