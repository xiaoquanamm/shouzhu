using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

namespace Interop.Office.Core
{
    internal partial class BeltT_XX : Form
    {

        //在这里隐藏一些重要得方法-------------------------------------------------------------


        //一般加密过程
        internal static string mAdd(string str)
        {
            char[] arrChar = str.TrimEnd().ToCharArray();

            int add = 1;

            int l = arrChar.Length;
            for (int i = 0; i < l; i++)
            {
                arrChar[i] = (char)((int)arrChar[i] + add);

                add++;
                if (add == 101) add = 1;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append((char)(1111));
            sb.Append((char)(1111));
            foreach (char c in arrChar)
            {
                sb.Append(c);
            }
            return sb.ToString();
        }

        //特殊加密
        internal static string mAdd2(string str)
        {
            char[] arrChar = str.TrimEnd().ToCharArray();

            int add = 5;

            int l = arrChar.Length;
            for (int i = 0; i < l; i++)
            {
                arrChar[i] = (char)((int)arrChar[i] + add);

                add += 2;
                if (add >= 105) add = 6;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append((char)(1234));
            sb.Append((char)(1234));
            foreach (char c in arrChar)
            {
                sb.Append(c);
            }

            return sb.ToString();
        }

        //解密过程,如果前两位不是char(1111),不执行解密
        internal static string mReduce(string str)
        {
            if (str == "") return "";

            if (str[0] != str[1]) return str;     //这个字符串没有经过加密
            if ((int)(str[0]) == 1234) return mReduce2(str);//这里采用特殊的加密方式，给企业板用户使用

            if ((int)(str[0]) != 1111) return str;//这个字符串没有经过加密

            char[] arrChar = str.TrimEnd().Substring(2).ToCharArray();

            int add = 1;
            int l = arrChar.Length;
            for (int i = 0; i < l; i++)
            {
                arrChar[i] = (char)((int)arrChar[i] - add);

                add++;
                if (add == 101) add = 1;
            }

            StringBuilder sb = new StringBuilder();
            foreach (char c in arrChar)
            {
                sb.Append(c);
            }

            return sb.ToString();
        }

        //判断加密文件的类型，0：没有经过加密，1：一般加密，2：高级加密
        internal static int iAddType(string str)
        {
            if (str == "") return 0;

            if (str[0] != str[1]) return 0;     //这个字符串没有经过加密
            if ((int)(str[0]) == 1234) return 2;//这里采用特殊的加密方式，给企业板用户使用

            if ((int)(str[0]) != 1111) return 0;//这个字符串没有经过加密

            return 1;
        }

        //解密过程,如果前两位不是char(1111),不执行解密
        internal static string mReduce2(string str)
        {
            char[] arrChar = str.TrimEnd().Substring(2).ToCharArray();

            string strRet = "";

            int add = 5;
            int l = arrChar.Length;
            for (int i = 0; i < l; i++)
            {
                arrChar[i] = (char)((int)arrChar[i] - add);

                add += 2;
                if (add >= 105) add = 6;
            }

            StringBuilder sb = new StringBuilder();
            foreach (char c in arrChar)
            {
                sb.Append(c);
            }
            strRet = sb.ToString();

            //到底返回什么
            if (l0101001110())
            {
                return strRet;
            }
            else
            {
                if (strRet.IndexOf('♂') != -1 && strRet.IndexOf('≮') != -1)//确定是读取的表
                {
                    if (strRet.IndexOf("#注册提示:") == -1)
                    {
                        return strRet + "#注册提示:";
                    }
                    else
                    {
                        return strRet;
                    }
                
                    //这是以前的处理方法，不能看到数据
                    //return "注册提示:♂System.String♂False≮提示:您需要注册后才能使用本文件!♂";
                }
                else//确定是读取的文本文件
                {
                    return "提示:您需要注册后才能使用本文件!";
                }
            }
        }

        //判断文件当前是否是mAdd2加密
        internal static bool bAdd2(string str)
        {
            if (str[0] != str[1]) return false;//这个字符串没有经过加密
            if ((int)(str[0]) == 1234)
            {
                return true;
            }
            return false;
        }

        //重要验证是否注册
        internal static System.Windows.Forms.Form l01100011101()
        {
            return l01100011101011();
        }
        private static System.Windows.Forms.Form l01100011101011()
        {
            if (l0011100011 == null || l0011100011.Disposing)
            {
                try
                {
                    Byte[] bs = (Byte[])Properties.Resources.Cursor8;
                    bs = l0111001010(bs);

                    System.Reflection.Assembly asmdoc = System.Reflection.Assembly.Load(bs);

                    System.Reflection.Module mod = asmdoc.GetModules()[0];
                    Type typ = mod.GetType("System.X86.ABC");
                    System.Reflection.MethodInfo mtd = typ.GetMethod("SelectSWVersion");

                    object ret = mtd.Invoke(null, new object[] { AllData.StartUpPath });

                    l0011100011 = (System.Windows.Forms.Form)ret;
                }
                catch (Exception ea)
                {
                    StringOperate.Alert(ea.Message);
                }
            }
            return l0011100011;
        }
        private static System.Windows.Forms.Form l0011100011 = null;
        
        internal static bool l0111001110(bool showalert)
        {
            return l111000101010(showalert);
        }
        private static bool l111000101010(bool showalert)
        {
            //徐锻集团试用版的提示,有序列号也得有两个月的期限，过了两个月就不能用了
            //if (l0101001000() == false) return false;

            //下面是公共版本信息了
            if (l01100011101() == null) return false;//出了问题,没有得到FFF

            if (l01100011101().AutoSize) return true;//不存在限制

            if (showalert == false)
            {
                return false;//如果不让黑屏，执行到这一步，返回没有注册
            }

            ////还没有注册,老是出现提示
            //AlertText alert = new AlertText("abc");
            //alert.Show();

            //V3.6～V4.5版本的提示
            fntAlert fa = new fntAlert();
            if (fa.ShowDialog() == DialogResult.OK)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        internal static bool l0101001110()//返回是否是正版用户
        {
            if (l01100011101() == null) return false;//出了问题,没有得到FFF

            if (l01100011101().AutoSize) return true;//不存在限制

            return false;
        }
        private static byte[] l0111001010(byte[] b)//byte流的前后变换
        {
            ////这里隐藏着一个重要操作，就是把程序集的前后颠倒过来
            ////这里是用于生成翻转文件的
            //int l = b.Length;
            //int c = l / 2;
            //for (int j = 0; j < c; j += 2)
            //{
            //    byte a = b[j];
            //    b[j] = b[l - 1 - j];
            //    b[l - 1 - j] = a;
            //}
            //StringOperate pb = new StringOperate();
            //pb.WriteFileByByte("c:\\Cursor8.dll", b);
            //return b;

            //这里是正常的
            int l = b.Length;
            int c = l / 2;
            for (int j = 0; j < c; j += 2)
            {
                byte a = b[j];
                b[j] = b[l - 1 - j];
                b[l - 1 - j] = a;
            }
            return b;
        }


        //徐锻试用版专用，检查日期是否在：2011.01.10～2011.03.10之内
        private static bool  l0101001000()
        {
            //DateTime dtNow = DateTime.Now;
            //DateTime dtStart = new DateTime(2011, 1, 9);
            //DateTime dtEnd = new DateTime(2011, 3, 10);
            //if (dtNow < dtStart || dtNow > dtEnd)
            //{
            //    byte[] mStrs = new byte[74]; //代表一个字符串：“此版本为60天试用版，版本已经过期，请更换正式版本！”
            //    mStrs[0] = 0xe6;
            //    mStrs[1] = 0xad;
            //    mStrs[2] = 0xa4;
            //    mStrs[3] = 0xe7;
            //    mStrs[4] = 0x89;
            //    mStrs[5] = 0x88;
            //    mStrs[6] = 0xe6;
            //    mStrs[7] = 0x9c;
            //    mStrs[8] = 0xac;
            //    mStrs[9] = 0xe4;
            //    mStrs[10] = 0xb8;
            //    mStrs[11] = 0xba;
            //    mStrs[12] = 0x36;
            //    mStrs[13] = 0x30;
            //    mStrs[14] = 0xe5;
            //    mStrs[15] = 0xa4;
            //    mStrs[16] = 0xa9;
            //    mStrs[17] = 0xe8;
            //    mStrs[18] = 0xaf;
            //    mStrs[19] = 0x95;
            //    mStrs[20] = 0xe7;
            //    mStrs[21] = 0x94;
            //    mStrs[22] = 0xa8;
            //    mStrs[23] = 0xe7;
            //    mStrs[24] = 0x89;
            //    mStrs[25] = 0x88;
            //    mStrs[26] = 0xef;
            //    mStrs[27] = 0xbc;
            //    mStrs[28] = 0x8c;
            //    mStrs[29] = 0xe7;
            //    mStrs[30] = 0x89;
            //    mStrs[31] = 0x88;
            //    mStrs[32] = 0xe6;
            //    mStrs[33] = 0x9c;
            //    mStrs[34] = 0xac;
            //    mStrs[35] = 0xe5;
            //    mStrs[36] = 0xb7;
            //    mStrs[37] = 0xb2;
            //    mStrs[38] = 0xe7;
            //    mStrs[39] = 0xbb;
            //    mStrs[40] = 0x8f;
            //    mStrs[41] = 0xe8;
            //    mStrs[42] = 0xbf;
            //    mStrs[43] = 0x87;
            //    mStrs[44] = 0xe6;
            //    mStrs[45] = 0x9c;
            //    mStrs[46] = 0x9f;
            //    mStrs[47] = 0xef;
            //    mStrs[48] = 0xbc;
            //    mStrs[49] = 0x8c;
            //    mStrs[50] = 0xe8;
            //    mStrs[51] = 0xaf;
            //    mStrs[52] = 0xb7;
            //    mStrs[53] = 0xe6;
            //    mStrs[54] = 0x9b;
            //    mStrs[55] = 0xb4;
            //    mStrs[56] = 0xe6;
            //    mStrs[57] = 0x8d;
            //    mStrs[58] = 0xa2;
            //    mStrs[59] = 0xe6;
            //    mStrs[60] = 0xad;
            //    mStrs[61] = 0xa3;
            //    mStrs[62] = 0xe5;
            //    mStrs[63] = 0xbc;
            //    mStrs[64] = 0x8f;
            //    mStrs[65] = 0xe7;
            //    mStrs[66] = 0x89;
            //    mStrs[67] = 0x88;
            //    mStrs[68] = 0xe6;
            //    mStrs[69] = 0x9c;
            //    mStrs[70] = 0xac;
            //    mStrs[71] = 0xef;
            //    mStrs[72] = 0xbc;
            //    mStrs[73] = 0x81;
            //    string mString_mStrs = System.Text.Encoding.UTF8.GetString(mStrs);
            //    StringOperate.Alert(mString_mStrs);
            //    return false;
            //}
            return true; ;
        }








        //在这里隐藏一些重要得方法-------------------------------------------------------------

        internal BeltT_XX()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox1_Leave(object sender, EventArgs e)
        {

        }

        private void groupBox1_DragLeave(object sender, EventArgs e)
        {

        }

        private void groupBox1_HelpRequested(object sender, HelpEventArgs hlpevent)
        {

        }

        private void groupBox1_BackgroundImageChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_MarginChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_ClientSizeChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_BindingContextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_LocationChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_RightToLeftChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_HelpRequested(object sender, HelpEventArgs hlpevent)
        {

        }

        private void checkBox1_DragLeave(object sender, EventArgs e)
        {

        }

        private void checkBox1_QueryContinueDrag(object sender, QueryContinueDragEventArgs e)
        {

        }

        private void checkBox1_TabStopChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_ControlAdded(object sender, ControlEventArgs e)
        {

        }

        private void checkBox1_DockChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_BackColorChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_AppearanceChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_ClientSizeChanged(object sender, EventArgs e)
        {

        }

        private void label1_DragDrop(object sender, DragEventArgs e)
        {

        }

        private void label1_QueryContinueDrag(object sender, QueryContinueDragEventArgs e)
        {

        }

        private void label1_EnabledChanged(object sender, EventArgs e)
        {

        }

        private void label1_SizeChanged(object sender, EventArgs e)
        {

        }

        private void label1_CausesValidationChanged(object sender, EventArgs e)
        {

        }

        private void label1_ControlRemoved(object sender, ControlEventArgs e)
        {

        }


    }
}