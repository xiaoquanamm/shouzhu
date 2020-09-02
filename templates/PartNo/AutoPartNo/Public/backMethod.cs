using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;
using System.Collections;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace Interop.Office.Core
{
    internal  class backMethod
    {

        internal backMethod()
        {

        }


        #region //文本到表,表到文本的转换

        /// <summary>
        /// 把表转换成文本文件//不加密
        /// </summary>
        internal string ChangeTBToString(DataTable dt)
        {
            if (dt == null) return "";

            string str = "";
            foreach (DataColumn dc in dt.Columns)
            {
                str += dc.ColumnName + "♂" + dc.DataType.ToString() + "♂" + dc.ReadOnly.ToString() + "♂";
            }
            str = str.Remove(str.LastIndexOf('♂'));
            str += "≮";

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        str += dr[i].ToString() + "♂";
                    }
                }
                str = str.Remove(str.LastIndexOf('♂'));//这里必须有行才能执行。
            }

            return str;
        }
        /// <summary>
        /// 把dgv转换成文本文件//不加密
        /// </summary>
        internal string ChangeDGVToString(DataGridView dgv)
        {
            if (dgv == null) return "";

            string str = "";
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                str += col.Name + "♂" + col.ValueType.ToString() + "♂" + col.ReadOnly.ToString() + "♂";
            }
            str = str.Remove(str.LastIndexOf('♂'));
            str += "≮";

            if (dgv.Rows.Count > 0)
            {
                foreach (DataGridViewRow dr in dgv.Rows)
                {
                    for (int i = 0; i < dgv.Columns.Count; i++)
                    {
                        str += dr.Cells[i].Value.ToString() + "♂";
                    }
                }
                str = str.Remove(str.LastIndexOf('♂'));//这里必须有行才能执行。
            }

            return str;
        }


        /// <summary>
        /// 根据文件路径,把文本文件还原成表 ,以UTF8读取//有解密过程,(ref 是不是正版用户专用 false) 
        /// </summary>
        internal DataTable readDatFileToTB(string txtPath,ref bool bOnlyReg)
        {
            //读取并解密
            string str = ReadTextFile(txtPath, Encoding.UTF8, true );
            //这一句非常重要，去掉末尾的\r\n换行符，否则导致错误
            str = str.Trim();

            if (str.Length == 0) return null;

            //>=v3.5,有可能添加一列
            bool bAddCol = false;
            if (str.EndsWith("#注册提示:"))
            {
                str = str.Remove(str.LastIndexOf('#'));
                bAddCol = true;
                bOnlyReg = true;
            }

            int iStart = str.IndexOf('≮');
            string start = (iStart == -1) ? str : str.Remove(iStart);//表头
            string end = (iStart == -1) ? "" : str.Substring(iStart + 1);//表内容

            DataTable dt = new DataTable();
            string[] strArr = start.Split('♂');
            for (int i = 0; i < strArr.Length; i += 3)
            {
                string ColName = strArr[i].Trim();
                string DataType = strArr[i + 1];
                string ReadOnly = strArr[i + 2];
                Type T = Type.GetType(DataType);
                if (T == typeof(System.Object)) T = typeof(System.String);

                DataColumn dc = new DataColumn(ColName, T);
                dc.ReadOnly = Convert.ToBoolean(ReadOnly);

                if (!dt.Columns.Contains(ColName)) dt.Columns.Add(dc);
            }

            int ColCount = dt.Columns.Count;
            string[] ArrVal = end.Split('♂');
            try
            {
                int Mod = ArrVal.Length % ColCount;//获取余数
                int iLength = ArrVal.Length - Mod; //最后一行不全
                for (int i = 0; i < ArrVal.Length ; i += ColCount)
                {
                    DataRow dr = dt.NewRow();
                    dt.Rows.Add(dr);
                    for (int j = 0; j < ColCount; j++)
                    {
                        if (ArrVal[i + j] != "") dr[j] = (object)ArrVal[i + j];
                    }
                }
            }
            catch (Exception ea)
            {
                StringOperate.AlertDebug("表格转换出错：" + ea.Message);
            }

            //>=v3.5,有可能添加一列
            if (bAddCol)
            {
                if (!dt.Columns.Contains("注册提示:"))
                {
                    DataColumn dc = new DataColumn("注册提示:", typeof(System.String));
                    dc.ReadOnly = false;
                    dt.Columns.Add(dc);
                    foreach (DataRow row in dt.Rows)
                    {
                        row["注册提示:"] = "注册后才能使用";
                    }
                }
            }

            return dt;
        }


        #endregion



        #region //读写文本文件
        

        /// <summary>
        /// 读取文本文件//编码方式//要解密吗//注意这里连换行符也一块读出来了，\r\n
        /// </summary>
        internal string ReadTextFile(string FullPath, Encoding encoding, bool bReduce)
        {
            if(!File.Exists(FullPath)) return "";
            FileInfo fi = new FileInfo(FullPath);
            FileStream fs = File.Open(FullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

            TextReader tr = new StreamReader(fs, encoding);
            string str = tr.ReadToEnd();



            tr.Close();
            tr.Dispose();
            fs.Close();
            fs.Dispose();

            if (bReduce)//执行解密过程
            {
                return mReduce(str);
            }

            return str;
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

        #endregion



        #region//字符串加密

        //一般加密过程
        internal string mAdd(string str)
        {
            return  BeltT_XX.mAdd(str);
        }

        //特殊加密
        internal string mAdd2(string str)
        {
            return  BeltT_XX.mAdd2(str);
        }

        //解密过程,如果前两位不是char(1111),不执行解密
        internal string mReduce(string str)
        {
            return BeltT_XX.mReduce(str);
        }

        //判断文件当前是否是mAdd2加密
        internal bool bAdd2(string str)
        {
            return  BeltT_XX.bAdd2(str);
        }

        //加密方式0：没有加密，1：低级加密，2：高级加密
        internal int iAddType(string str,bool bIsPath)
        {
            if (bIsPath)
            {
                str = this.ReadTextFile(str,Encoding.UTF8 ,false);
            }
            return BeltT_XX.iAddType(str);
        }

        #endregion



    }

}
