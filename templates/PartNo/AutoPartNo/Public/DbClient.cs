using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Drawing;
using System.Collections;
using System.Runtime.Serialization;
using Interop.Office.Core;
using System.IO;



/// <summary>
/// Database.OperateDB 的摘要说明。
/// </summary>
public class DbClient : DBMethod
{

    #region 初始化连接对象

    public DbClient(string myConnectionString)
    {
        string IP = "127.0.0.1";
        int iPort = 3547;

        //如果是作为客户端使用
        if (PSetUp.bUsingType_Client)
        {
            iPort = 3546;
            IP = PSetUp.fntWeb_ServerIP;

            myConnectionString = myConnectionString.Replace(AllData.StartUpPath, "[MD3DToolsBasePath]");
        }
        else
        {
            #region//检查外部数据提供程序是否已经启动MD3DToolsData.exe，并自动启动
            try
            {
                if (Sys.Is64BitOPSystem)
                {
                    System.Diagnostics.Process[] processArr = System.Diagnostics.Process.GetProcessesByName("MD3DToolsData");
                    if (processArr.Length == 0)
                    {
                        string exepath = AllData.StartUpPath + "\\MD3DToolsData.exe";
                        if (File.Exists(exepath))
                        {
                            System.Diagnostics.Process.Start(exepath);
                            System.Threading.Thread.Sleep(400);
                        }
                        else
                        {
                            StringOperate.Alert("检测到没有启动迈迪数据程序“MD3DtoolsData.exe”，这可能导致数据访问不正常！");
                        }
                    }
                }
            }
            catch (Exception ea)
            {
                //string str = "";
            }
            #endregion
        }

        this.ConnectionString = myConnectionString;
        this.tsp = new TransAccess(IP,iPort);
        this.Ser = new SerializeInfo();
    }

    /// <summary>
    /// 接口
    /// </summary>
    public string Connstr
    {
        get
        {
            return ConnectionString;
        }
        set
        {
            ConnectionString = value;
        }
    }
    private string ConnectionString = "";

    private TransAccess tsp = null;
    private SerializeInfo Ser = null;

    #endregion


    /// <summary>
    /// 得到第一行第一列的值,返回String结果
    /// </summary>
    public object FirstValue(string SqlString)
    {
        return FirstValue(SqlString, null, null);
    }
    public object FirstValue(string SqlString, string[] arrParamName, object[] arrParamValue)
    {
        string strcmd = "2FirstValueぺ" + this.Connstr + "ぺ" + SqlString + "ぺ";
        if (arrParamName != null)
        {
            for (int i = 0; i < arrParamName.Length; i++)
            {
                strcmd += (arrParamName[i] + "$" + arrParamValue[i].GetType().ToString() + "$" + arrParamValue[i].ToString() + "ぺ");
            }
        }
        strcmd = strcmd.TrimEnd(new char[] { 'ぺ' });

        string str = this.tsp.TCPString(strcmd);
        if (str.StartsWith("Error"))
        {
            StringOperate.Alert(str);
            return "";
        }
        else
        {
            return str;
        }
    }


    /// <summary>
    ///  执行Sql命令,返回受影响的行数
    /// </summary>
    public int Excute(string SqlString)
    {
        return this.Excute(SqlString, null, null);
    }
    public int Excute(string SqlString, string[] arrParamName, object[] arrParamValue)
    {
        string strcmd = "2Excuteぺ" + this.Connstr + "ぺ" + SqlString + "ぺ";

        if (arrParamName != null)
        {
            for (int i = 0; i < arrParamName.Length; i++)
            {
                if (arrParamName[i] == null || arrParamValue[i] == null) continue;
                strcmd += (arrParamName[i] + "$" + arrParamValue[i].GetType().ToString() + "$" + arrParamValue[i].ToString() + "ぺ");
            }
        }
        strcmd = strcmd.TrimEnd(new char[] { 'ぺ' });

        string str = this.tsp.TCPString(strcmd);
        if (str.StartsWith("Error"))
        {
            StringOperate.Alert(str);
            return 0;
        }
        else
        {
            return  Convert.ToInt16(str);
        }
    }
    /// <summary>
    /// 执行多个SQL语句,,成功返回“”，出错返回ErrorServer:
    /// </summary>
    public string Excute(string[] SqlArr)
    {
        string strcmd = "2ExcuteArrぺ" + this.Connstr + "ぺ";
        foreach (string s in SqlArr)
        {
            strcmd += (s + "ぺ");
        }
        strcmd = strcmd.TrimEnd(new char[] { 'ぺ' });

        string str = this.tsp.TCPString(strcmd);
        if (str.StartsWith("Error"))
        {
            StringOperate.Alert(str);
            return "";
        }
        else
        {
            return str;
        }
    }


    /// <summary>
    /// 执行Sql命令,返回一个数据表
    /// </summary>
    public DataTable DBTable(string SqlString)
    {
        return DBTable(SqlString, null, null);
    }
    public DataTable DBTable(string SqlString, string[] arrParamName, object[] arrParamValue)
    {
        string strcmd = "3DBTableぺ" + this.Connstr + "ぺ" + SqlString + "ぺ";
        if (arrParamName != null)
        {
            for (int i = 0; i < arrParamName.Length; i++)
            {
                strcmd += (arrParamName[i] + "$" + arrParamValue[i].GetType().ToString() + "$" + arrParamValue[i].ToString() + "ぺ");
            }
        }
        strcmd = strcmd.TrimEnd(new char[] { 'ぺ' });
        string strError = "";
        DataTable dt = this.tsp.TCPDataTable(strcmd, ref strError);
        if (strError.Length > 0)
        {
            StringOperate.Alert(strError);
        }
        return dt;
    }


    /// <summary>
    ///  得到一个数据库中的所有表名称
    /// </summary>
    public ArrayList GetAllTableNames()
    {

        System.Collections.ArrayList tableNames = new System.Collections.ArrayList();

        string strcmd = "2AllTablesぺ" + this.Connstr + "ぺ";
        string str = this.tsp.TCPString(strcmd);
        if (str.StartsWith("Error"))
        {
            StringOperate.Alert(str);
        }
        else
        {
            string[] strArr = str.Split(new char[] { 'ぺ' });
            foreach (string s in strArr)
            {
                if (s.Length > 0 && !tableNames.Contains(s))
                {
                    tableNames.Add(s);
                }
            }
        }

        return tableNames;
    }

    /// <summary>
    /// 把一个DataTable写入到数据库中(新建),返回插入行数，并添加新的标识列【ID20】因为ID列有可能和其它表关联，这个ID20是一个无关联的自动编号列，如果原先有ID20这一列，会删去这一列。
    /// </summary>
    public int CreateTable(DataTable dt, string tbName)
    {
        byte[] btArr = this.Ser.Serialize((object)dt);
        string strcmd = "4CreateTableぺ" + this.Connstr + "ぺ" + tbName + "ぺ" + btArr.Length.ToString();

        string str = this.tsp.TCPUpTable(strcmd, btArr);
        if(str.StartsWith("Error"))
        {
            StringOperate.Alert(str);
            return 0;
        }
        return Convert.ToInt16(str);
    }
    /// <summary>
    /// 把Access的习惯改成Sql
    /// </summary>
    public string checkSql(string sql)
    {
        sql = sql.Replace("int dientity(1,1)", "autoincrement(1)");
        return sql;
    }




    #region /---DBToolObj---/

    /// <summary>
    /// 根据查询语句从数据库中读取一个文件，并保存到指定的目录下,返回是否成功
    /// </summary>
    public bool ReadFile(string sql, string createFilePath)
    {
        string strcmd = "2ReadFileぺ" + this.Connstr + "ぺ" + sql + "ぺぺ" + createFilePath;
        string str = this.tsp.TCPString(strcmd);
        if (str.StartsWith("Error") || str.Length >0)
        {
            StringOperate.Alert(str);
            return false;
        }
        else
        {
            return true;
        }
    }
    /// <summary>
    /// 把一个文件存入数据库，（文件路径，sql语句，文件参数名@file ,要加密吗 ),返回受影响的行数
    /// </summary>
    public int WriteFile(string FilePath, string SqlString, string ParamName, bool Add)//有Param，如@aaa,这个参数代表文件
    {
        if (!System.IO.File.Exists(FilePath)) return 0;

        string strcmd = "2WriteFileぺ" + this.Connstr + "ぺ" + SqlString + "ぺ" + ParamName + "ぺ" + FilePath + "ぺ" + Add.ToString();
        string str = this.tsp.TCPString(strcmd);
        if (str.StartsWith("Error"))
        {
            StringOperate.Alert(str);
            return 0;
        }
        else
        {
            return Convert.ToInt16(str);
        }
    }


    /// <summary>
    /// 根据查询语句从数据库中读取一个文件流，并返回一张图片,出错返回null
    /// </summary>
    public System.Drawing.Image ReadFileToImage(string sql)
    {
        DataTable dt = this.DBTable(sql, null, null);
        if (dt == null || dt.Rows.Count ==0) return null;

        byte[] bFile = (byte[])dt.Rows[0][0];
        bFile = this.ReduceByteArr(bFile);//解密

        System.Drawing.ImageConverter imgC = new System.Drawing.ImageConverter();//将byte[]转换成图片
        System.Drawing.Image img = (Image)imgC.ConvertFrom(bFile);
        return img;
    }
    /// <summary>
    /// 根据查询语句从数据库中读取多个文件流，并返回一个图片数组,出错返回null
    /// </summary>
    internal System.Drawing.Image[] ReadFileToImageArr(string sql)
    {
        try
        {
            DataTable dt = this.DBTable(sql, null, null);
            if (dt == null || dt.Rows.Count == 0) return null;

            System.Drawing.ImageConverter imgC = new System.Drawing.ImageConverter();//将byte[]转换成图片
            Image[] imgArr = new Image[dt.Rows.Count];

            for (int j = 0; j < imgArr.Length; j++)
            {
                byte[] bFile = (byte[])dt.Rows[j][0];
                bFile = this.ReduceByteArr(bFile);//解密

                imgArr[j] = (Image)imgC.ConvertFrom(bFile);
            }

            return imgArr;
        }
        catch
        {
            return null;
        }
    }
    /// <summary>
    /// 根据查询语句从数据库中读取一个文件流，并返回String
    /// </summary>
    internal string ReadFileToString(string sql)
    {
        try
        {
            DataTable dt = this.DBTable(sql, null, null);
            if (dt == null || dt.Rows.Count == 0) return "";

            byte[] bFile = this.ReduceByteArr((byte[])dt.Rows[0][0]);

            System.Text.UTF8Encoding converter = new UTF8Encoding();
            string str = converter.GetString(bFile);
            return str;
        }
        catch
        {
            return "";
        }
    }
    /// <summary>
    /// 根据查询语句从数据库中读取一个文件流，并返回String
    /// </summary>
    internal string ReadFileToString(string sql, Encoding converter)
    {
        try
        {
            DataTable dt = this.DBTable(sql, null, null);
            if (dt == null || dt.Rows.Count == 0) return "";
            byte[] bFile = this.ReduceByteArr((byte[])dt.Rows[0][0]);

            //System.Text.UTF8Encoding converter2 = new UTF8Encoding();
            string str = converter.GetString(bFile);
            return str;
        }
        catch
        {
            return "";
        }
    }
    /// <summary>
    /// 根据查询语句从数据库中读取一个文件流，并返回byte[],出错返回null
    /// </summary>
    internal byte[] ReadFileToByte(string sql)
    {
        try
        {
            DataTable dt = this.DBTable(sql, null, null);
            if (dt == null || dt.Rows.Count == 0) return null;

            byte[] bFile = this.ReduceByteArr((byte[])dt.Rows[0][0]);
            return bFile;
        }
        catch
        {
            return null;
        }
    }
    /// <summary>
    /// 文件的解密
    /// </summary>
    private byte[] ReduceByteArr(byte[] bFile)
    {
        //如果前是个字节是3336669990,说明这个文件是加密的，要解密
        if (bFile[0] == 3 && bFile[1] == 3 && bFile[2] == 3)
        {
            if (bFile[3] == 6 && bFile[4] == 6 && bFile[5] == 6)
            {
                if (bFile[6] == 9 && bFile[7] == 9 && bFile[8] == 9)
                {
                    if (bFile[9] == 0)
                    {
                        byte[] bFile2 = new byte[bFile.Length - 10];//从第十个开始读
                        for (int i = 0; i < bFile2.Length; i++)
                        {
                            bFile2[i] = bFile[i + 10];
                        }
                        bFile = bFile2;
                    }
                }
            }
        }

        return bFile;
    }


    #endregion//一般的SQL命令


}

