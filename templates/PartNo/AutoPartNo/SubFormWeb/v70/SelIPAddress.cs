using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.DirectoryServices;

namespace MDClient
{
    internal partial class SelIPAddress : Form
    {
        internal SelIPAddress()
        {
            InitializeComponent();
        }
        private void SelIPAddress_Load(object sender, EventArgs e)
        {
            GroupInfo();
        }
        private void GroupInfo()//工作组
        {
            DirectoryEntry MainGroup = new DirectoryEntry("WinNT:");
            foreach (DirectoryEntry domain in MainGroup.Children)
            {
                listBox1.Text = "";
                listBox1.Items.Add(domain.Name);
            }
        }
        private void ComputerInfo(string strname)//计算机
        {
            DirectoryEntry MainGroup = new DirectoryEntry("WinNT:");
            foreach (DirectoryEntry domain in MainGroup.Children)
            {
                if (domain.Name == strname)
                {
                    foreach (DirectoryEntry pc in domain.Children)
                    {
                        if (pc.Name != "Schema")//Schema是结束标记   
                            this.listBox2.Items.Add(pc.Name);
                    }
                }
            }
        }
       

        //选择一台机器
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox2.SelectedItem == null) return;
            txtSelectIP.Text = "";//先清空
            string txt = listBox2.SelectedItem.ToString();
            IPAddress[] ip = null;
            try
            {
                ip = Dns.GetHostAddresses(txt);
                foreach (System.Net.IPAddress ipaddr in ip)
                {
                    if (ipaddr.ToString().IndexOf('.') != -1)
                    {
                        txtSelectIP.Text = ipaddr.ToString();//再设置
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Interop.Office.Core.StringOperate.Alert(ex.Message);
                return;
            }
        }

        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            listBox2.Items.Add("正再搜索,请稍后...");
            listBox2.Refresh();
            ComputerInfo(this.listBox1.Text);
            listBox2.Items.RemoveAt(0);
        }



    }
}