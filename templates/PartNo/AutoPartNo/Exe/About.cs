using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.IO;

namespace Interop.Office.Core
{
    internal partial class About : Form
    {

        internal About()
        {
            InitializeComponent();
        }
        internal About(int SelectedPageIndex)//这里要先写入注册表,然后在加载时从注册表中读出来,以确定显示哪个窗口.
        {
            InitializeComponent();

            if (tabControl1.TabCount > SelectedPageIndex)
            {
                Interop.Office.Core.Properties.Settings.Default.AboutTabSelIndex = SelectedPageIndex;
            }
        }


        #region 程序集属性访问器

        /// <summary>
        /// 程序的版本
        /// </summary>
        internal static string AssemblyVersion
        {
            get
            {
                return "2015.4.28";
            }
        }

        internal string AssemblyTitle
        {
            get
            {
                return "MD3DTools";

                //// 获取此程序集上的所有 Title 属性
                //object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                //// 如果至少有一个 Title 属性
                //if (attributes.Length > 0)
                //{
                //    // 请选择第一个属性
                //    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                //    // 如果该属性为非空字符串，则将其返回
                //    if (titleAttribute.Title != "")
                //        return titleAttribute.Title;
                //}
                //// 如果没有 Title 属性，或者 Title 属性为一个空字符串，则返回 .exe 的名称
                //return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        internal string AssemblyDescription
        {
            get
            {
                return "迈迪设计宝";


                //// 获取此程序集的所有 Description 属性
                //object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                //// 如果 Description 属性不存在，则返回一个空字符串
                //if (attributes.Length == 0)
                //    return "";
                //// 如果有 Description 属性，则返回该属性的值
                //return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        internal string AssemblyProduct
        {
            get
            {
                return "迈迪软件";

                //// 获取此程序集上的所有 Product 属性
                //object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                //// 如果 Product 属性不存在，则返回一个空字符串
                //if (attributes.Length == 0)
                //    return "";
                //// 如果有 Product 属性，则返回该属性的值
                //return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        internal string AssemblyCopyright
        {
            get
            {
                return "迈迪公司";

                //// 获取此程序集上的所有 Copyright 属性
                //object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                //// 如果 Copyright 属性不存在，则返回一个空字符串
                //if (attributes.Length == 0)
                //    return "";
                //// 如果有 Copyright 属性，则返回该属性的值
                //return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        internal string AssemblyCompany
        {
            get
            {
                return "迈迪公司";

                //// 获取此程序集上的所有 Company 属性
                //object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                //// 如果 Company 属性不存在，则返回一个空字符串
                //if (attributes.Length == 0)
                //    return "";
                //// 如果有 Company 属性，则返回该属性的值
                //return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        
        #endregion



        private void About_Load(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = Interop.Office.Core.Properties.Settings.Default.AboutTabSelIndex;

            this.picBox1.ImageLocation = AllData.StartUpPath + "\\IMG\\about1.jpg";
            this.picBox2.ImageLocation = AllData.StartUpPath + "\\IMG\\about2.jpg";

            //这个页面是公司简介
            Uri uri = new Uri(AllData.StartUpPath + "\\IMG\\AboutPage2.htm");//也可以是本机上的一个.htm文件
            webBrowser1.Url = uri;

            //显示版本
            this.lblVersionName.Text = "企业版 " + About.AssemblyVersion;
         }
        //关闭窗体前,保存当前是那个面板
        //并且，看看是否是从网站上传回了东西
        private void About_FormClosing(object sender, FormClosingEventArgs e)
        {
            //写入注册表,打开了那个面板
            Interop.Office.Core.Properties.Settings.Default.AboutTabSelIndex = tabControl1.SelectedIndex;
            Interop.Office.Core.Properties.Settings.Default.Save();

        }


        

        internal bool panel2first = true;//在线交流,第一次打开这个面板
        internal bool panel3first = true;//软件注册,是否为第一次打开

        internal string MachineCode = "";//本机码
        internal string EndUrl = "";//Url的参数
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int a = tabControl1.SelectedIndex;
            if (this.EndUrl == "" || this.MachineCode =="")//第一次，计算出本机码
            {
                this.MachineCode = Dbtool2.strError.ToString();//得到本机码
                this.EndUrl = String.Format("?MachineCode={0}&HasRegister={1}", this.MachineCode, Dbtool2.strError.AutoSize);
            }
            if (a < 2)//0,1,2面板的大小固定
            {
                this.WindowState = System.Windows.Forms.FormWindowState.Normal;
                this.Width = 708;
                this.Height = 488;
                this.FormBorderStyle = FormBorderStyle.FixedSingle;
                this.MaximizeBox = false;
                //if (a == 2)//软件注册
                //{
                //    if (panel2first)//第一次打开这个面板
                //    {
                //        panel2first = false;

                //        Uri uri = new Uri("http://www.jnmedia.cn/payment/payment.asp" + this.EndUrl);
                //        webRegister.Url = uri;
                //    }
                //}
            }
            else//面板:问题反馈,自动更新
            {
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.MaximizeBox = true;

                if (a == 2)//意见反馈
                {
                    if (panel3first)//第一次打开这个面板
                    {
                        panel3first = false;

                        Uri uri = new Uri("http://www.my3dparts.com/Opinion_Soft/Default.aspx" + this.EndUrl);
                        //Uri uri = new Uri("http://www.jnmedia.cn/MD3Dtools/index.asp" + this.EndUrl);
                        //这个地方将问题反馈转移到三维配件网上,但注意,以前的连接也要有效

                        webBrowser2.Url = uri;
                    }
                }
            }
            if (a == 3)
            {
                if (this.webBrowser3.Url == null)
                {
                    string target = "http://www.my3dparts.com/Contact/WebParts.aspx" + "?a=" + Dbtool2.strError.ToString() + "&b=" + Dbtool2.strError.Text.ToString() + "&code=1458";
                    Uri uri = new Uri(target);
                    this.webBrowser3.Url = uri;
                }
            }

            if (tabControl1.TabPages[a].Controls.Count  > 0)
            {
                tabControl1.TabPages[a].Controls[0].Focus();//第一个控件得到焦点
            }
        }





        //关闭窗口
        private void okButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        //单击【注册信息】打开注册窗口
        private void button1_Click(object sender, EventArgs e)
        {
            Dbtool2.strError.ShowDialog();
        }

        //隐含的特殊操作，所有的重要设定
        private void logoPictureBox_DoubleClick(object sender, EventArgs e)
        {
            
        }
        //点迈－是管理员
        //点迪－不是管理员
        //点软－是迈迪
        //点件－不是迈迪
        private void logoPictureBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (e.X < 20 && e.Y < 20)
                {
                    FrmMgrWindow fmgr = new FrmMgrWindow();
                    fmgr.Show();
                }
            }
        }



    }
}