using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Collections;
using System.Net;
using System.Net.Sockets;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Interop.Office.Core
{

    /// <summary>
    /// 字符串处理函数，文本文件读写，Bype[]转换
    /// </summary>
    internal class StringOperate
    {

        internal StringOperate() { }


        #region //文件到byte[]的转换

        /// <summary>
        /// 根据byte[]创建文件并写入指定的目录WritePath(包含文件名)下//false表示出错
        /// </summary>
        /// <param name="WritePath">要写入文件到什么地方//包含路径全名//文件名</param>
        /// <param name="byt">byte[];一个序列化的文件//要满的,即没有空的内容</param>
        /// <returns>是否成功</returns>
        internal bool WriteFileByByte(string WritePath, byte[] byt)
        {
            try
            {
                Stream s = File.Open(WritePath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                s.Write(byt, 0, byt.Length);
                s.Flush();
                s.Dispose();
                return true;
            }
            catch//(Exception e)
            {
                //Transport.Succeed = e.Message;
                return false;
            }
        }

        /// <summary>
        /// 根据byte[]创建文件并写入指定的目录WritePath(包含文件名)下//false表示出错
        /// 这个方法可以指定byte[]数组的写入长度,如byt.length=1000,但是我要写入前800
        /// </summary>
        internal bool WriteFileByByte(string WritePath, byte[] byt, int WriteLength)
        {
            try
            {
                Stream s = File.Open(WritePath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                s.Write(byt, 0, WriteLength);
                s.Flush();
                s.Dispose();
                return true;
            }
            catch// (Exception e)
            {
                //Transport.Succeed = e.Message;
                return false;
            }
        }

        /// <summary>
        /// 由文件全名序列化后得到byte[]//出错时返回1长度的byte,//文件不存在返回null,文件正在使用时无法读取
        /// </summary>
        /// <param name="fileFullName">文件的路径和名称</param>
        /// <returns>序列化后得到的byte</returns>
        internal byte[] GetByteByFilePath(string fileFullName)
        {
            try
            {
                ////FileInfo fi = new FileInfo(fileFullName);
                ////if (!fi.Exists) return null;
                ////int length = Convert.ToInt32(fi.Length);
                ////Stream s = File.Open(fileFullName, FileMode.Open);
                ////byte[] bytFile = new byte[length];
                ////length = s.Read(bytFile, 0, bytFile.Length);
                ////s.Dispose();
                ////return bytFile; 

                FileInfo fi = new FileInfo(fileFullName);
                if (!fi.Exists) return null;
                int length = Convert.ToInt32(fi.Length);

                //只有这样写，在文件被其它程序使用的时候，才不至于出错
                System.IO.FileStream s = new System.IO.FileStream(fileFullName, System.IO.FileMode.Open, System.IO.FileAccess.Read, FileShare.ReadWrite);
                byte[] bytFile = new byte[length];
                length = s.Read(bytFile, 0, bytFile.Length);
                s.Dispose();
                return bytFile;

            }
            catch//(Exception e)
            {
                //StringOperate.Alert(e.Message);
                byte[] byt = new byte[1];
                return byt;
            }
        }
        
        /// <summary>
        /// 把一个字符串编码为另一种编码的字符串
        /// </summary>
        internal string ConvertCode(string str, System.Text.Encoding SourceEncoding, System.Text.Encoding DestEncoding)
        {
            byte[] Sbyte = SourceEncoding.GetBytes(str.ToCharArray());

            byte[] Dbyte = Encoding.Convert(SourceEncoding, DestEncoding, Sbyte);

            return DestEncoding.GetString(Dbyte);
        }


        #endregion


        #region //读写文本文件

        /// <summary>
        /// 读取文本文件
        /// </summary>
        /// <param name="FullPath"></param>
        /// <returns></returns>
        internal string ReadTextFile(string FullPath)
        {
            FileInfo fi = new FileInfo(FullPath);
            if (!fi.Exists) return "";
            FileStream fs;
            fs = File.Open(FullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

            TextReader tr = new StreamReader(fs,Encoding.Default );
            string str = tr.ReadToEnd();


            tr.Close();
            tr.Dispose();
            fs.Close();
            fs.Dispose();

            return str;
            
            //byte[] bit = new byte[fi.Length];
            //int i = fs.Read(bit, 0, bit.Length);
            //string txt = Encoding.Default.GetString(bit, 0, i);
            //return txt;
        }

        /// <summary>
        /// 指定编码读取文件
        /// </summary>
        internal string ReadTextFile(string FullPath, Encoding encod)
        {
            FileInfo fi = new FileInfo(FullPath);
            if (!fi.Exists) return "";
            FileStream fs;
            fs = File.Open(FullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

            TextReader tr = new StreamReader(fs, encod);
            string str = tr.ReadToEnd();

            tr.Close();
            tr.Dispose();
            fs.Close();
            fs.Dispose();

            return str;

            //byte[] bit = new byte[fi.Length];
            //int i = fs.Read(bit, 0, bit.Length);
            //string txt = Encoding.Default.GetString(bit, 0, i);
            //return txt;
        }

        /// <summary>
        /// 读取文本文件,注意是读取的byte然后转换的,Encoding.UTF8，注意这里连换行符也一块读出来了，\r\n
        /// </summary>
        internal string ReadTextFileByByte(string FullPath)
        {
            if (!File.Exists(FullPath)) return "";
            FileInfo fi = new FileInfo(FullPath);
            int length = Convert.ToInt32(fi.Length);
            byte[] bytFile = new byte[length];
            try
            {
                //只有这样写，在文件被其它程序使用的时候，才不至于出错
                System.IO.FileStream s = new System.IO.FileStream(FullPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, FileShare.ReadWrite);
                length = s.Read(bytFile, 0, bytFile.Length);
                s.Dispose();
            }
            catch//(Exception e)
            {
                return "";
            }

            string str = System.Text.Encoding.UTF8.GetString(bytFile);
            return str;
        }


        /// <summary>
        /// 指定编码读取文件的每一行，返回一个数组，第三个参数是否包括空行，编码默认是GB2312,可以是null
        /// </summary>
        internal List<string> ReadTextFile(string FullPath, Encoding encod,bool inclode0LengthRows)
        {
            List<string> strArr= new List<string>();
            strArr.Add("");

            FileInfo fi = new FileInfo(FullPath);
            if (!fi.Exists) return strArr;

            if(encod ==null)
            {
                encod = Encoding.GetEncoding("GB2312");
            }

            FileStream fs = File.Open(FullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            TextReader tr = new StreamReader(fs, encod);

            string str = tr.ReadLine();
            while (str != null)
            {
                if (str.Length == 0)
                {
                    if (inclode0LengthRows) strArr.Add(str);
                }
                else
                {
                    strArr.Add(str);
                }
                str = tr.ReadLine();
            }

            tr.Close();
            tr.Dispose();
            fs.Close();
            fs.Dispose();

            return strArr;

            //byte[] bit = new byte[fi.Length];
            //int i = fs.Read(bit, 0, bit.Length);
            //string txt = Encoding.Default.GetString(bit, 0, i);
            //return txt;
        }

        /// <summary>
        /// 写入文本文件--创建//不需要编码信息
        /// </summary>
        internal bool WriteTextFileCreate(string FullPath, string str)
        {
            return this.WriteTextFileCreate(FullPath, str, null);
        }
        /// <summary>
        /// 写入文本文件--创建
        /// </summary>
        internal bool WriteTextFileCreate(string FullPath, string str, Encoding encoding)
        {
            string strPath = Path.GetDirectoryName(FullPath);
            if (!Directory.Exists(strPath))
            {
                Directory.CreateDirectory(strPath);
            }
            FileStream fs = File.Open(FullPath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
            TextWriter tw;
            if (encoding != null)
            {
                tw = new StreamWriter(fs, encoding);
            }
            else
            {
                tw = new StreamWriter(fs);
            }
            tw.WriteLine(str);
            tw.Close();
            tw.Dispose();
            fs.Close();
            fs.Dispose();

            return true;
        }


        /// <summary>
        /// 写入文本文件--追加
        /// </summary>
        internal bool WriteTextFileAppend(string FullPath, string str)
        {
            FileMode fmode = FileMode.Append;
            if (!File.Exists(FullPath))
            {
                fmode = FileMode.Create;
            }

            FileStream fs = File.Open(FullPath, fmode , FileAccess.Write, FileShare.Write);
            TextWriter tw = new StreamWriter(fs);
            tw.WriteLine(str);
            tw.Close();
            tw.Dispose();
            fs.Close();
            fs.Dispose();
            return true;
            //try
            //{
            //}
            //catch
            //{
            //    return false;
            //}
        }


        #endregion//////////////////////////////////////////////////////////////////////////////////////


        #region//高级字符串处理

        /// <summary>
        /// 得到一个字符串的绝对长度,中文的长度为2
        /// </summary>
        internal int GetStringLength(string str)
        {
            char[] charArr = str.ToCharArray();

            //区分中文还是英文，中文双字节
            int Length = 0;
            for (int i = 0; i < charArr.Length; i++)
            {
                int code = Convert.ToInt32(charArr[i]);
                if (code < 255)//英文
                {
                    Length += 1;
                }
                else//中文
                {
                    Length += 2;
                }
            }
            return Length;
        }

        /// <summary>
        /// 字符串处理，保留多少位小数//关键是不足这些位数的化后面要加0
        /// </summary>
        internal string BackStringByDouble(double val, int iLength)//假设iLength=4
        {
            string str = val.ToString();
            int index = str.IndexOf('.');
            if (index >= 0)//如：＋0.12
            {
                string sEnd = str.Substring(index + 1);//12
                int sEndLength = sEnd.Length;//2
                for (int i = sEndLength; i < iLength; i++)//2,3 
                {
                    str = str + "0";
                }
            }
            else if (index == -1)
            {
                str = str + ".";
                for (int i = 0; i < iLength; i++)
                {
                    str = str + "0";
                }
            }
            return str;
        }

        /// <summary>
        /// 返回小数位数为count 的字符串,等同于Math.Round()
        /// </summary>
        /// <param name="doubleValue">要处理的数</param>
        /// <param name="count">小数位数</param>
        /// <returns>处理完的字符串</returns>
        internal string BackStringByDouble(string doubleValue, int count)
        {
            try
            {
                if (count == 2)//保留两位小数
                {
                    float ff = Convert.ToSingle(doubleValue);
                    ff = ff + (float)0.00499999;
                    string s = ff.ToString();
                    try
                    {
                        s = s.Remove(s.IndexOf('.') + 3);
                    }
                    catch { }
                    return s;
                }
                if (count == 1)//保留一位小数
                {
                    float ff = Convert.ToSingle(doubleValue);
                    ff = ff + (float)0.04999999;
                    string s = ff.ToString();
                    try
                    {
                        s = s.Remove(s.IndexOf('.') + 2);
                    }
                    catch { }
                    return s;
                }
            }
            catch
            { }
            return doubleValue;
        }

        /// <summary>
        /// 处理文本，把一个长字符串剪切成固定长度的数组
        /// 其中，中文占两个长度，英文占一个长度
        /// </summary>
        internal string[] CutUnicodeString(string str, int LineLeth)
        {
            char[] charArr = str.ToCharArray();

            //区分中文还是英文，中文双字节
            int Length = 0;
            string sRet = "";
            for (int i = 0; i < charArr.Length; i++)
            {
                int code = Convert.ToInt32(charArr[i]);
                if (code < 255)//英文
                {
                    Length += 1;
                }
                else//中文
                {
                    Length += 2;
                }

                sRet = sRet + charArr[i].ToString();

                if (Length >= LineLeth)//完成一行了
                {
                    sRet += "ぺ";//用这个字符分割成数组
                    Length = 0;
                }
            }
            return sRet.Split(new char[] { 'ぺ' });
        }

        #endregion


        #region//正则表达式

        /// <summary>
        /// 正则表达式 去掉字符串中的数字
        /// </summary>
        internal static string RemoveNumber(string key)
        {
            return System.Text.RegularExpressions.Regex.Replace(key, @"\d", "");
        }
        /// <summary>
        /// 正则表达式 去掉字符串中的非数字
        /// </summary>
        internal static string RemoveNotNumber(string key)
        {
            return System.Text.RegularExpressions.Regex.Replace(key, @"[^\d]*", "");
        }

        /// <summary>
        /// 正则表达式 去掉所有非汉字，只返回汉字
        /// </summary>
        internal static string RemoveNotGB2312(string key)
        {
            //提取中文
            string pattern = @"[\u4e00-\u9fa5]+";

            //提取双引号之间信息
            //string pattern2 = "\"[^\"]*\"";

            Match result = Regex.Match(key, pattern,RegexOptions.IgnoreCase | RegexOptions.Multiline);
            return result.Value;
        }
        /// <summary>
        /// 正则表达式 去掉所有汉字
        /// </summary>
        internal static string RemoveGB2312(string key)
        {
            //提取中文
            string pattern = @"[\u4e00-\u9fa5]+";

            //提取双引号之间信息
            //string pattern2 = "\"[^\"]*\"";

            Regex regex = new Regex(pattern);
            return regex.Replace(key, "");
        }



        /// <summary>
        /// 正则表达式判断是否为中文
        /// </summary>
        internal static bool bIsChinese(string content) 
        { 
            string regexstr = @"[\u4e00-\u9fa5]";
            if (Regex.IsMatch(content, regexstr))
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        /// <summary>
        /// 判断是否为ASCII编码为255以下的字符组成的
        /// </summary>
        internal static bool bHasChinese(string content)
        {
            char[] charArr = content.ToCharArray();
            foreach (char c in charArr)
            {
                if (((int)c) > 255) return false;
            }
            return true;
        }


        /// <summary>
        /// 判断是不是全数字
        /// </summary>
        public static bool isInterger(string str)
        {
            foreach (char c in str)
            {
                if (!char.IsNumber(c)) return false;
            }
            return true;
        }


        /// <summary>
        /// 判断是否为ip
        /// </summary>
        public static bool IsIP(string ip)
        {
            return Regex.IsMatch(ip, @"^((2[0-4]\d|25[0-5]|[01]?\d\d?)\.){3}(2[0-4]\d|25[0-5]|[01]?\d\d?)$");
        }


        #endregion


        #region//简体中文到繁体中文之间的转换
        /// <summary>
        /// 中文字符工具类
        /// </summary>
        private const int LOCALE_SYSTEM_DEFAULT = 0x0800;
        private const int LCMAP_SIMPLIFIED_CHINESE = 0x02000000;
        private const int LCMAP_TRADITIONAL_CHINESE = 0x04000000;

        [DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int LCMapString(int Locale, int dwMapFlags, string lpSrcStr, int cchSrc, [Out] string lpDestStr, int cchDest);

        /// <summary>
        /// 将字符转换成简体中文
        /// </summary>
        /// <param name="s">输入要转换的字符串</param>
        /// <returns>转换完成后的字符串</returns>
        internal static string ToSimplified(string s)
        {
            String target = new String(' ', s.Length);
            int ret = LCMapString(LOCALE_SYSTEM_DEFAULT, LCMAP_SIMPLIFIED_CHINESE, s, s.Length, target, s.Length);
            return target;
        }

        /// <summary>
        /// 讲字符转换为繁体中文
        /// </summary>
        /// <param name="s">输入要转换的字符串</param>
        /// <returns>转换完成后的字符串</returns>
        internal static string ToTraditional(string s)
        {
            //如：拉伸1，后面带个1不好办。
            //if (htBig5Encoding == null)
            //{
            //    cfgInfo cinfo = new cfgInfo();
            //    htBig5Encoding = cinfo.getAllValue(AllData.StartUpPath + "GB2312-BIG5.cfg\\");
            //}
            //if (htBig5Encoding.ContainsKey(s))
            //{
            //    string aa=htBig5Encoding[s].ToString();
            //    if (aa.Length > 0)
            //    {
            //        return aa;
            //    }
            //}

            String target = new String(' ', s.Length);
            int ret = LCMapString(LOCALE_SYSTEM_DEFAULT, LCMAP_TRADITIONAL_CHINESE, s, s.Length, target, s.Length);
            return target;
        }



        /// <summary>
        /// 将字符串转换成繁体，（带判断是否需要转换）
        /// </summary>
        internal static string Big5Convert(string s)
        {
            if (s.Length == 0) return "";
            if (PSetUp.bIsBig5Encoding)
            {
                return ToTraditional(s);
            }
            return s;
        }
        /// <summary>
        /// 如果设置为繁体，把窗体的所有字符转换为繁体
        /// </summary>
        internal static void Big5Convert(System.Windows.Forms.Form frm)
        {
            if (frm == null) return;
            if (PSetUp.bIsBig5Encoding)
            {
                if(frm.Text.Length>0)frm.Text = StringOperate.ToTraditional(frm.Text);

                if (frm.Controls.ContainsKey("lblNotEncoding"))
                {
                    return;//说明此窗口禁止繁体编码
                }

                Big5ConvertForm_Ref(frm.Controls);
            }
        }
        private static void Big5ConvertForm_Ref(System.Windows.Forms.Control.ControlCollection ctlArr)
        {
            foreach (Control ctl in ctlArr)
            {
                Big5Convert(ctl);

                if (ctl.Controls.Count > 0)
                {
                    Big5ConvertForm_Ref(ctl.Controls);
                }
            }
        }

        /// <summary>
        /// 如果设置为繁体，把控件的所有字体转换为繁体,如果Tag为NoReEncoding就不转换，如果Tag为CanReEncoding就一定要转换，除此之外只转换Label等基本控件
        /// </summary>
        internal static void Big5Convert(Control ctl)
        {
            if (!PSetUp.bIsBig5Encoding) return;

            string strTag = (ctl.Tag != null ? ctl.Tag.ToString() : "");
            if (strTag == "NoReEncoding") return;

            //不要修改这些类型的控件，以免引起错误,除非有CanReEncoding
            if (strTag == "CanReEncoding")
            {
                if (ctl.GetType() == typeof(System.Windows.Forms.ComboBox))
                {
                    ComboBox cmb = (ComboBox)ctl;
                    cmb.Text = StringOperate.ToTraditional(cmb.Text);
                    for (int i = 0; i < cmb.Items.Count; i++)
                    {
                        cmb.Items[i] = StringOperate.ToTraditional(cmb.Items[i].ToString());
                    }
                    return;
                }
                else if (ctl.GetType() == typeof(System.Windows.Forms.DataGridView))
                {
                    Big5Convert(ctl as DataGridView); return;
                }
                else if (ctl.GetType() == typeof(System.Windows.Forms.TreeView))
                {
                    Big5Convert((TreeView)ctl); return;
                }
            }

            if (ctl.GetType() == typeof(System.Windows.Forms.ContextMenuStrip))
            {
                //继承关系 ContextMenuStrip : ToolStripDropDownMenu: ToolStripDropDown :ToolStrip
                ContextMenuStrip cmenu = (ContextMenuStrip)ctl;
                foreach (ToolStripItem item in cmenu.Items)
                {
                    if (item.GetType() == typeof(ToolStripSeparator)) continue;

                    if (item.Text.Length > 0) item.Text = StringOperate.ToTraditional(item.Text);
                    foreach (ToolStripItem subitem in (item as ToolStripMenuItem).DropDownItems)
                    {
                        if (subitem.Text.Length > 0) subitem.Text = StringOperate.ToTraditional(subitem.Text);
                    }
                }
            }
            else if (ctl.GetType() == typeof(System.Windows.Forms.MenuStrip))
            {
                MenuStrip menu = (MenuStrip)ctl;
                foreach (ToolStripItem item in menu.Items)
                {
                    if (item.GetType() == typeof(ToolStripSeparator)) continue;
                    if (item.GetType() == typeof(ToolStripComboBox)) continue;

                    if (item.Text.Length > 0) item.Text = StringOperate.ToTraditional(item.Text);
                    foreach (ToolStripItem subitem in (item as ToolStripMenuItem).DropDownItems)
                    {
                        if (subitem.Text.Length > 0) subitem.Text = StringOperate.ToTraditional(subitem.Text);
                    }
                }
            }
            else if (ctl.Text.Length > 0)
            {
                ctl.Text = StringOperate.ToTraditional(ctl.Text);
            }
        }





        /// <summary>
        /// 把一个表格转换为繁体
        /// </summary>
        internal static DataTable Big5Convert(DataTable dt)
        {
            if (!PSetUp.bIsBig5Encoding) return dt;
            foreach (DataColumn dc in dt.Columns)
            {
                int idx = dc.Ordinal;

                dc.ColumnName = StringOperate.Big5Convert(dc.ColumnName);
                dc.Caption = StringOperate.Big5Convert(dc.Caption);

                if (dc.DataType == typeof(System.String))
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dt.Rows[i][idx] = StringOperate.Big5Convert(dt.Rows[i][idx].ToString());
                    }
                }
            }

            return dt;
        }
        /// <summary>
        /// 把一个表格转换为繁体
        /// </summary>
        internal static void Big5Convert(DataGridView dgv)
        {
            if (!PSetUp.bIsBig5Encoding) return;
            foreach (DataGridViewColumn  dc in dgv.Columns)
            {
                int idx = dc.Index;

                dc.Name = StringOperate.Big5Convert(dc.Name);
                dc.HeaderText = StringOperate.Big5Convert(dc.HeaderText);

                //dc.ValueType其实获取不出任何信息来，以后再完善吧
                if (dc.ValueType == typeof(System.String))
                {
                    for (int i = 0; i < dgv.Rows.Count; i++)
                    {
                        DataGridViewCell cell = dgv.Rows[i].Cells[idx];
                        if (cell!=null && cell.Value != null)
                        {
                            cell.Value =StringOperate.Big5Convert(cell.Value.ToString());
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 把一个表格转换为繁体
        /// </summary>
        internal static void Big5Convert(TreeNodeCollection  nodeArr)
        {
            if (!PSetUp.bIsBig5Encoding) return;
            foreach(TreeNode tn in nodeArr)
            {
                tn.Text = ToTraditional(tn.Text);
                if (tn.Nodes.Count > 0)
                {
                    Big5Convert(tn.Nodes);
                }
            }
        }


        /// <summary>
        /// 把中文特征名称转换成英文，或转换成繁体中文。
        /// </summary>
        internal static string CEN(string cnName)
        {

            if (AllData.iSWLanguage <= 1)
            {
                return cnName;//中文无需转换
            }

            //去除数字，如【草图1>>草图】//如D1@阵列(圆周)1
            string[] strArr = cnName.Split(new char[] { '@' });

            if (AllData.iSWLanguage == 2)
            {
                if (htEnglishEncoding == null)
                {
                    cfgInfo cinfo = new cfgInfo();
                    htEnglishEncoding = cinfo.getAllValue(AllData.StartUpPath + "\\GB2312-English.cfg");
                }

                foreach (string SA in strArr)
                {
                    string SB = StringOperate.RemoveNumber(SA);//阵列(圆周)
                    if (htEnglishEncoding.ContainsKey(SB))
                    {
                        string val = htEnglishEncoding[SB].ToString();
                        if (val.Length > 0)
                        {
                            cnName = cnName.Replace(SB, val);
                        }
                    }
                }
            }
            else if (AllData.iSWLanguage == 3)
            {
                if (htBig5Encoding == null)
                {
                    cfgInfo cinfo = new cfgInfo();
                    htBig5Encoding = cinfo.getAllValue(AllData.StartUpPath + "\\GB2312-BIG5.cfg");
                }

                foreach (string SA in strArr)
                {
                    string SB = StringOperate.RemoveNumber(SA);//阵列(圆周)
                    if (htBig5Encoding.ContainsKey(SB))
                    {
                        string val = htBig5Encoding[SB].ToString();
                        if (val.Length > 0)
                        {
                            cnName = cnName.Replace(SB, val);
                        }
                    }
                    else
                    {
                        cnName = cnName.Replace(SB, Big5Convert(SB));
                    }
                }
            }

            return cnName;
        }


        /// <summary>
        /// 提示框，如果是繁体自动转换
        /// </summary>
        internal static void Alert(string text)
        {
            MessageBox.Show(Big5Convert(text));
        }
        /// <summary>
        /// 迈迪专用错误提示，如果是迈迪自己就会提示，否则不提示，写入到日志中
        /// </summary>
        internal static void AlertDebug(string text)
        {
            if (PSetUp.bIsMD)
            {
                MessageBox.Show("迈迪专用提示：" + text);
            }
            else
            {
                string s = DateTime.Now.ToString("G") + "   " + text ;
                StringOperate sop = new StringOperate();
                sop.WriteTextFileAppend(AllData.StartUpPath + "\\mdlog.txt", s);
            }
        }
        /// <summary>
        /// 提示框，如果是繁体自动转换
        /// </summary>
        internal static DialogResult Alert(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return MessageBox.Show(Big5Convert(text), caption, buttons, icon);
        }


        //英语翻译列表
        private static Hashtable htEnglishEncoding = null;
        //繁体翻译列表
        private static Hashtable htBig5Encoding = null;
        

        #endregion


    }


}
    

