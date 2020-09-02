using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Interop.Office.Core
{
    internal  partial class FrmMgrWindow : Form
    {
        internal FrmMgrWindow()
        {
            InitializeComponent();
        }
        private void FrmMgrWindow_Load(object sender, EventArgs e)
        {
            cfgInfo cinfo = new cfgInfo();
            this.textBox1.Text = cinfo.getValue("mdcode",   "regcode.cfg");
        }

        //开始注册
        private void button1_Click(object sender, EventArgs e)
        {
            string txtcode = this.textBox1.Text.Trim().Replace(" ", "");
 
            cfgInfo cinfo = new cfgInfo();
            cinfo.setValue("mdcode", txtcode, "regcode.cfg");

            this.Close();
        }

        //取消注册
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        //获取迈迪注册码
        internal string getMDCode(string mgrcode)
        {
            string mdcode = "";
            string smcode = mgrcode.Replace(" ", "");
            for (int i=0;i<smcode.Length;i++)
            {
                string ss = smcode.Substring(i, 1);
                int aa = Convert.ToInt16(ss);
                aa *= aa;

                ss =  aa.ToString().Substring(aa.ToString().Length - 1);

                if (i % 2 == 0)
                {
                    mdcode += ss;
                }
                else
                {
                    mdcode = ss + mdcode;
                }
            }
            return mdcode;
        }



    }
}
