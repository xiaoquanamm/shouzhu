using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.IO;

namespace Interop.Office.Core
{
    /// <summary>
    /// 所有的配置名称都从这里头获取
    /// </summary>
    enum SetupNames
    {
        strServerIP = 1,   //局域网服务器IP
        strServerPort = 2, //局域网服务器端口
        bAccess = 3,       //是否使用Access数据库

        UsingType = 5,//作为单机版使用StandAlone，作为客户端使用Client，作为服务器使用Server
        StandAlone = 6,
        Client = 7,
        Server = 8,

        BomTB1=9, //BOM表的相关设置
        BomTB2=10
    }

    /// <summary>
    /// 读写配置信息的类
    /// </summary>
    internal class cfgInfo
    {
        internal const string cfgFileName1 = "app.cfg";


        /// <summary>
        /// 检测cfgFileName 是文件名，还是全名称，并自动转换成全名称
        /// </summary>
        private string checkCfgFileName(string cfgFileName)
        {
            if (cfgFileName.Length == 0) cfgFileName = cfgFileName1;
            if (cfgFileName.IndexOf('\\') == -1 && cfgFileName.Length < 20)
            {
                cfgFileName = AllData.StartUpPath + "\\" + cfgFileName;
            }
            
            return cfgFileName;
        }


        /// <summary>
        /// 读取一个设置项(属性名称,配置文件全称或仅名称如"app.cfg","C:\\printer.cfg");
        /// </summary>
        internal string getValue(string AttName)
        {
            return getValue(AttName, cfgFileName1);
        }
        /// <summary>
        /// 读取一个设置项(属性名称,配置文件全称或仅名称如"app.cfg","C:\\printer.cfg");
        /// </summary>
        internal string getValue(string AttName, string cfgFileName)
        {
            cfgFileName= this.checkCfgFileName(cfgFileName);

            if (!File.Exists(cfgFileName)) return "";

            string AttValue = "";

            FileStream fs = File.Open(cfgFileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            TextReader tr = new StreamReader(fs, Encoding.UTF8);

            string str = tr.ReadLine();
            while (str != null)
            {
                if (str.IndexOf('=') != -1)
                {
                    string sName = str.Remove(str.IndexOf('='));
                    string sValue = str.Substring(str.IndexOf('=') + 1);

                    if (AttName == sName)
                    {
                        AttValue = sValue; break;
                    }
                }
                str = tr.ReadLine();//重新读取一行
            }

            tr.Close();
            tr.Dispose();
            fs.Close();
            fs.Dispose();

            return AttValue;
        }
        /// <summary>
        /// 读取所有的属性,返回一个哈希表[属性名称--属性值](配置文件全称或仅名称如"app.cfg","C:\\printer.cfg");
        /// </summary>
        internal Hashtable getAllValue(string cfgFileName)
        {
            cfgFileName = this.checkCfgFileName(cfgFileName);

            Hashtable ht = new Hashtable();
            if (!File.Exists(cfgFileName)) return ht;

            FileStream fs = File.Open(cfgFileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            TextReader tr = new StreamReader(fs, Encoding.UTF8);

            string str = tr.ReadLine();
            while (str != null)
            {
                if (str.IndexOf('=') != -1 && !str.StartsWith(";"))
                {
                    string sName = str.Remove(str.IndexOf('='));
                    string sValue = str.Substring(str.IndexOf('=') + 1);

                    if (!ht.ContainsKey(sName))
                    {
                        ht.Add(sName, sValue);
                    }
                }
                str = tr.ReadLine();//重新读取一行,
            }

            tr.Close();
            tr.Dispose();
            fs.Close();
            fs.Dispose();

            return ht;
        }

                                                       
        /// <summary>
        /// 设置一个属性的新值
        /// </summary>
        internal bool setValue(string AttName, string AttValue)
        {
            return setValue(AttName, AttValue, cfgFileName1);
        }
        /// <summary>
        /// 设置一个属性的新值(属性名，属性值，配置文件全称或仅名称如"app.cfg","C:\\printer.cfg");
        /// </summary>
        internal bool setValue(string AttName, string AttValue, string cfgFileName)
        {
            cfgFileName = this.checkCfgFileName(cfgFileName);
            bool bExists = File.Exists(cfgFileName);
            if (!bExists)
            {
                string pathstr= Path.GetDirectoryName(cfgFileName);
                if (!Directory.Exists(pathstr))
                {
                    Directory.CreateDirectory(pathstr);
                }
            }

            FileMode fmode = bExists ? FileMode.Open : FileMode.Create;
            FileStream fs = File.Open(cfgFileName, fmode, FileAccess.ReadWrite, FileShare.ReadWrite);
            TextReader tr = new StreamReader(fs, Encoding.UTF8);

            string newStr = "";//要写入的新

            bool bhasValue = false;//原先是否已经有这个属性

            string str = tr.ReadLine();
            while (str != null)
            {
                string s = str;
                if (s.IndexOf('=') != -1)
                {
                    string sName = s.Remove(s.IndexOf('='));
                    string sValue = s.Substring(s.IndexOf('=') + 1);

                    if (AttName == sName)
                    {
                        sValue = AttValue;
                        str = AttName + "=" + AttValue;
                        bhasValue = true;
                    }
                }

                newStr += str + System.Environment.NewLine;

                str = tr.ReadLine();//重新读取一行
            }

            if (!bhasValue)
            {
                newStr += AttName + "=" + AttValue + System.Environment.NewLine;
            }

            tr.Close();//首先关闭读取流
            tr.Dispose();
            fs.Close();//最后关闭文件流
            fs.Dispose();

            fs = File.Open(cfgFileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
            TextWriter tw = new StreamWriter(fs, Encoding.UTF8);
            tw.Write(newStr);//然后写入
            tw.Close();
            tw.Dispose();
            fs.Close();//最后关闭文件流
            fs.Dispose();

            return true;
        }
        /// <summary>
        /// 设置多个属性的值,这个函数性能不高,也是一个一个的设
        /// </summary>
        internal bool setAllValue(Hashtable htValues, string cfgFileName)
        {
            bool bok = true;
            foreach (object o in htValues.Keys)
            {
                bool b = setValue(o.ToString(), htValues[o].ToString(), cfgFileName);
                if (b == false)
                {
                    bok = false;
                } 
            }
            return bok;
        }

        /// <summary>
        /// 向多个值中添加一个值，如“11,22,33”中添加“44”，变成“11,22,33,44”
        /// </summary>
        internal bool addValue(string AttName, string AttValue)
        {
            string oldValue = getValue(AttName);
            string newVal = addOneValueToList(oldValue, AttValue);

            return setValue(AttName, newVal, cfgFileName1);
        }
        /// <summary>
        /// 向多个值中添加一个值，如“11,22,33”中添加“44”，变成“11,22,33,44”
        /// </summary>
        internal bool addValue(string AttName, string AttValue, string cfgFileName)
        {
            string oldValue = getValue(AttName,cfgFileName);
            string newVal = addOneValueToList(oldValue, AttValue);

            return setValue(AttName, newVal, cfgFileName);
        }


        /// <summary>
        /// 向多个值中添加一个值，如“11,22,33”中添加“44”，变成“11,22,33,44”
        /// </summary>
        /// <param name="oldValue">长值</param>
        /// <param name="newValue">要添加的值</param>
        /// <returns>新的值</returns>
        internal string addOneValueToList(string oldValue, string newValue)
        {
            string[] arrOld = oldValue.Split(new char[] { ',', '，' });
            string[] arrNew = newValue.Split(new char[] { ',', '，' });

            ArrayList ar = new ArrayList();
            foreach (string s in arrOld)
            {
                if (!ar.Contains(s)) ar.Add(s);
            }
            foreach (string s in arrNew)
            {
                if (!ar.Contains(s)) ar.Add(s);
            }

            string sret = "";
            foreach (object o in ar)
            {
                sret += o.ToString() + ",";
            }
            if (sret.Length > 0)
            {
                sret = sret.Remove(sret.Length - 1);
            }

            return sret;
        }



        /// <summary>
        /// 清空一个值，保存值的名称
        /// </summary>
        internal bool removeValue(string AttName, string cfgFileName)
        {
            cfgFileName = this.checkCfgFileName(cfgFileName);
            bool bExists = File.Exists(cfgFileName);
            if (!bExists)
            {
                return false;
            }

            FileMode fmode = bExists ? FileMode.Open : FileMode.Create;
            FileStream fs = File.Open(cfgFileName, fmode, FileAccess.ReadWrite, FileShare.ReadWrite);
            TextReader tr = new StreamReader(fs, Encoding.UTF8);

            string newStr = "";//要写入的新

            string str = tr.ReadLine();
            while (str != null)
            {
                if (str.TrimStart().StartsWith(AttName + "="))
                {
                }
                else
                {
                    newStr += str + System.Environment.NewLine;
                }
                str = tr.ReadLine();//重新读取一行
            }

            tr.Close();//首先关闭读取流
            tr.Dispose();
            fs.Close();//最后关闭文件流
            fs.Dispose();

            fs = File.Open(cfgFileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
            TextWriter tw = new StreamWriter(fs, Encoding.UTF8);
            tw.Write(newStr);//然后写入
            tw.Close();
            tw.Dispose();
            fs.Close();//最后关闭文件流
            fs.Dispose();

            return true;
        }


    }
}
