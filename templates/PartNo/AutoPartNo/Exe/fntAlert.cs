using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace Interop.Office.Core
{
    internal partial class fntAlert : Form
    {

        /// <summary>
        /// 如果网站配置有在线图片日期，就读取，然后保存，保存时保存成在线日期，下次检测如果本机图片是在线日期的，就不下载，如果本机图片老了，就重新下载。
        /// </summary>
        internal fntAlert()
        {
            InitializeComponent();
            //this.TopMost = false ;
            //this.WindowState = System.Windows.Forms.FormWindowState.Normal ;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
            About ab = new About();
            ab.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            Dbtool2.strError.ShowDialog();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.LoadRef();
        }
        private void LoadRef()
        {
            //名称是随机产生的,防止出现破解的情况
            Random rdm = new Random();
            int a = rdm.Next(1, 999999);
            string s = a.ToString();
            this.Name = s;
            this.Text = s;

            timer1_Tick(null, null);
            button1.Focus();
        }

        private int iMaxCount
        {
            get
            {
                if (izzmcount == -1)
                {
                    DirectoryInfo dir = new DirectoryInfo(AllData.StartUpPath + "\\IMG\\MDFile");
                    FileInfo[] fArr = dir.GetFiles("MDFile*");
                    izzmcount = fArr.Length;
                }
                return izzmcount;
            }
        }
        private int izzmcount = -1;
        private int imgIdx = 1;
        private Random radom = new Random();
        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                //5.5以后的版本，直接读取文件
                this.imgIdx = radom.Next(1, this.iMaxCount);
                this.setPic();
            }
            catch//(Exception ea)
            {
            }
        }
        private void picLeft_Click(object sender, EventArgs e)
        {
            timer1.Stop();

            this.imgIdx--;
            if (this.imgIdx <= 0) this.imgIdx = this.iMaxCount;

            this.setPic();
        }
        private void picRight_Click(object sender, EventArgs e)
        {
            timer1.Stop();

            this.imgIdx++;
            if (this.imgIdx > this.iMaxCount) this.imgIdx = 1;

            this.setPic();
        }

        private void setPic()
        {
            string filePath = AllData.StartUpPath + "\\IMG\\MDFile\\MDFile" + this.imgIdx.ToString();
            //如果在线
            if (TransWeb.bOnline && TransWeb.htWebCfg.ContainsKey("InitIMGData"))
            {
                try
                {
                    DateTime dtUpdate = Convert.ToDateTime(TransWeb.htWebCfg["InitIMGData"]);
                    bool bNeed = true;
                    if (File.Exists(filePath + ".jpg"))
                    {
                        DateTime dt = File.GetLastWriteTime(filePath + ".jpg");
                        if (dt >= dtUpdate)
                        {
                            bNeed = false;
                        }
                    }

                    //下载
                    if (bNeed)
                    {
                        this.timer1.Stop();
                        this.pictureBox1.ImageLocation = "http://www.my3dparts.com/MD/MDFile/MDFile" + this.imgIdx.ToString() + ".jpg";
                        return;
                    }
                }
                catch
                { }
            }

            if (System.IO.File.Exists(filePath + ".jpg"))
            {
                this.pictureBox1.ImageLocation = filePath + ".jpg";
            }
            else if (System.IO.File.Exists(filePath))
            {
                this.pictureBox1.ImageLocation = filePath;
            }
        }
        //图片加载完成
        private void pictureBox1_LoadCompleted(object sender, AsyncCompletedEventArgs e)
        {
            try
            {
                if (pictureBox1.ImageLocation.StartsWith("http://www"))
                {
                    this.timer1.Start();
                    if (pictureBox1.Image.Width < 200 && pictureBox1.Image.Height < 150) return;//说明没有成功加载

                    string imgname = pictureBox1.ImageLocation.Substring(pictureBox1.ImageLocation.LastIndexOf('/') + 1);
                    string cache = AllData.StartUpPath + "\\IMG\\MDFile\\" + imgname;

                    if (File.Exists(cache))
                    {
                        File.Delete(cache);
                    }
                    pictureBox1.Image.Save(cache);

                    DateTime dtUpdate = Convert.ToDateTime(TransWeb.htWebCfg["InitIMGData"]);
                    File.SetLastWriteTime(cache, dtUpdate);
                }
            }
            catch
            {
            }
        }



        private void fntAlert_FormClosing(object sender, FormClosingEventArgs e)
        {
            timer1.Stop();
            timer1.Dispose();
        }

        //单击转到网站并关闭当前窗口
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            string key = this.pictureBox1.ImageLocation;
            key = key.Substring(key.LastIndexOf('\\') + 1);

            //V6.0到标准网站查询
            TransWeb.openOnlineHelpPage("init", key,"");

            //关闭这个窗口
            this.button1.PerformClick();
        }

        
        
        //窗口的重画事件,在这里要加上调用绘图的方法
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            this.drawInPic(p1);
        }
        private Pen p1 = new Pen(Color.Green, 7);
        private Pen p2 = new Pen(Color.GreenYellow, 8);
        private void drawInPic(Pen p)
        {
            //开始绘图
            Graphics gleft = this.picLeft.CreateGraphics();
            gleft.DrawLine(p, 50, 0, 0, 50);
            gleft.DrawLine(p, 0, 50, 50, 100);

            Graphics gright = this.picRight.CreateGraphics();
            gright.DrawLine(p, 0, 0, 50, 50);
            gright.DrawLine(p, 50, 50, 0, 100);
        }

        private void picLeft_MouseEnter(object sender, EventArgs e)
        {
            this.drawInPic(this.p2);
        }

        private void picLeft_MouseLeave(object sender, EventArgs e)
        {
            this.drawInPic(this.p1);
        }

        private void pictureBox1_MouseEnter(object sender, EventArgs e)
        {
            pictureBox1.BorderStyle = BorderStyle.FixedSingle;
        }

        private void pictureBox1_MouseLeave(object sender, EventArgs e)
        {
            pictureBox1.BorderStyle = BorderStyle.None;
        }




    }
}