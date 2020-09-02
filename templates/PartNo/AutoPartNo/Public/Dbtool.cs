using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Drawing;
using System.Collections;
using System.IO;
using Interop.Office.Core;


namespace Interop.Office.Core
{


    /// <summary>
    /// Database.OperateDB 的摘要说明。
    /// </summary>
    public class Dbtool : DBMethod
    {

        #region 初始化连接对象

        public Dbtool(string myConnectionString)
        {
            this.ConnectionString = myConnectionString;
        }

        /// <summary>
        /// 数据库的连接字符串
        /// </summary>
        internal string ConnectionString
        {
            get
            {
                return this.cn.ConnectionString;
            }
            set
            {
                this._cnString = value;
                this.cn = new OleDbConnection(value);
            }
        }
        private string _cnString = "";

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
                this.cn = new OleDbConnection(value);
            }
        }

        /// <summary>
        /// 连接超时最长时间
        /// </summary>
        internal int TimeOut
        {
            get
            {
                return _timeout;
            }
            set
            {
                _timeout = value;
            }
        }
        private int _timeout = 30;


        /// <summary>
        /// 数据库的连接对象
        /// </summary>
        private OleDbConnection cn = null;

        #endregion


        #region /---操作数据库---/

        /// <summary>
        /// 得到第一行第一列的值
        /// </summary>
        public object FirstValue(string SqlString)
        {
            return FirstValue(SqlString, null, null);
        }
        public object FirstValue(string SqlString, string[] arrParamName, object[] arrParamValue)
        {
            OleDbCommand cmd = new OleDbCommand(SqlString, this.cn);

            if (arrParamName != null)
            {
                for (int i = 0; i < arrParamName.Length; i++)
                {
                    cmd.Parameters.Add(new OleDbParameter(arrParamName[i], arrParamValue[i]));
                }
            }

            if (cn.State != ConnectionState.Open) cn.Open();

            object obj = cmd.ExecuteScalar();

            if (cn.State != ConnectionState.Closed) cn.Close();

            cmd.Dispose();

            return obj;
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
            OleDbDataAdapter da = new OleDbDataAdapter(SqlString, this.cn);

            if (arrParamName != null)
            {
                for (int i = 0; i < arrParamName.Length; i++)
                {
                    da.SelectCommand.Parameters.Add(new OleDbParameter(arrParamName[i], arrParamValue[i]));
                }
            }

            DataTable dt = new DataTable();

            da.Fill(dt);//填充数据表

            da.Dispose();//销毁数据对象

            if (cn.State != ConnectionState.Closed) cn.Close();

            return dt;
        }


        /// <summary>
        /// 执行Sql命令,返回数据流
        /// </summary>
        public OleDbDataReader DBReader(string SqlString)
        {
            OleDbCommand cmd = new OleDbCommand(SqlString, this.cn);
            return cmd.ExecuteReader();//返回数据流
        }
        /// <summary>
        /// 执行Sql命令返回数据集
        /// </summary>
        public DataSet DBDataSet(string SqlString)
        {
            OleDbDataAdapter da = new OleDbDataAdapter(SqlString, this.cn);

            DataSet ds = new DataSet("Table");

            da.Fill(ds);

            da.Dispose();

            return ds;
        }


        /// <summary>
        ///  执行Sql命令,返回受影响的行数
        /// </summary>
        public int Excute(string SqlString)
        {
            return Excute(SqlString, null, null);
        }
        public int Excute(string SqlString, string[] arrParamName, object[] arrParamValue)
        {
            int a = 0;
            OleDbCommand cmd = new OleDbCommand(SqlString, this.cn);

            if (arrParamName != null)
            {
                for (int i = 0; i < arrParamName.Length; i++)
                {
                    cmd.Parameters.Add(new OleDbParameter(arrParamName[i], arrParamValue[i]));
                }
            }

            if (cn.State != ConnectionState.Open) cn.Open();
            cmd.CommandTimeout = 500;
            a = cmd.ExecuteNonQuery();

            if (cn.State != ConnectionState.Closed) cn.Close();

            return a;
        }


        /// <summary>
        /// 执行多个SQL语句,
        /// </summary>
        public string Excute(string[] SqlArr)
        {
            string strMessage = "";
            try
            {
                if (cn.State != ConnectionState.Open) cn.Open();
                foreach (string s in SqlArr)
                {
                    OleDbCommand cmd = new OleDbCommand(s, this.cn);
                    cmd.CommandTimeout = 100;
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                strMessage = e.Message;
            }
            finally
            {
                if (cn.State != ConnectionState.Closed) cn.Close();
            }
            return strMessage;
        }


        #endregion//一般的SQL命令


        #region/---得到数据库中的所有的表---/

        /// <summary>
        ///  得到一个数据库中的所有表名称
        /// </summary>
        public ArrayList GetAllTableNames()
        {
            System.Collections.ArrayList tableNames = new System.Collections.ArrayList();

            //把所有的表找出来，添加到下拉列表中
            try
            {
                //先得到要导入的文件中有多少个表名，并将表名放入ArrayList:SheetNameList中，
                this.cn.Open();
                DataTable dtTableName = this.cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                //放入之前先清除原有内容

                for (int i = 0; i < dtTableName.Rows.Count; i++)
                {
                    tableNames.Add(dtTableName.Rows[i]["TABLE_NAME"].ToString());
                }

            }
            catch (Exception ex)
            {
                StringOperate.Alert(ex.Message);
            }
            finally
            {
                this.cn.Close();
            }

            return tableNames;
        }


        /// <summary>
        /// 把一个DataTable写入到数据库中(新建),返回插入行数，并添加新的标识列【ID20】因为ID列有可能和其它表关联，这个ID20是一个无关联的自动编号列，如果原先有ID20这一列，会删去这一列。
        /// </summary>
        public int CreateTable(DataTable dt, string tbName)
        {
            string sqlPM = "";
            string sqlPM2 = "";
            ArrayList arPMName = new ArrayList();
            SerializeInfo ser = new SerializeInfo();

            //先删除原来的表
            try
            {
                int dd = this.Excute("drop table [" + tbName + "]");
            }
            catch
            { }

            string sql1 = "Create Table [" + tbName + "](";
            foreach (DataColumn dc in dt.Columns)
            {
                string colName = dc.ColumnName.Replace('.', '_');
                if (colName.ToLower() == "id20") continue;

                string type = (ser.ChangeType(dc.DataType)).ToString();
                string nullb = (dc.AllowDBNull) ? "NULL" : "not null";
                //string identity = (dc.Unique) ? "identity(1,1)" : " ";

                sql1 += "[" + colName + "] " + type + " " + nullb + ",";
                sqlPM += "@" + colName + ",";
                sqlPM2 += "[" + colName + "],";
                arPMName.Add("@" + colName);
            }

            sql1 += " [ID20] autoincrement(1) ,constraint [ID20] primary key([ID20]) )";
            int b = this.Excute(sql1);


            int iOK = 0;
            sqlPM = sqlPM.TrimEnd(new char[] { ',' });
            sqlPM2 = sqlPM2.TrimEnd(new char[] { ',' });
            string sql = "insert into [" + tbName + "](" + sqlPM2 + ") values (" + sqlPM + ")";
            foreach (DataRow drRow in dt.Rows)
            {
                object[] objval = new object[dt.Columns.Count];
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    objval[i]= drRow[i];
                }

                int a = this.Excute(sql, ser.ArrayListToArr(arPMName),objval);
                if (a == 1)
                {
                    iOK++;
                }
                else
                {
                    //str += "------Error";
                }
            }

            return iOK;
        }

        /// <summary>
        /// 把Access的习惯改成Sql
        /// </summary>
        public string checkSql(string sql)
        {
            sql = sql.Replace("int dientity(1,1)", "autoincrement(1)");
            return sql;
        }

        #endregion//////////////


       
        
        #region /---DBToolObj---/

        /// <summary>
        /// 根据查询语句从数据库中读取一个文件，并保存到指定的目录下,返回是否成功
        /// </summary>
        public bool ReadFile(string sql, string createFilePath)
        {
            try
            {
                DataTable dt = this.DBTable(sql, null, null);
                if (dt == null || dt.Rows.Count == 0) return false;

                //创建目录
                if (!System.IO.Directory.Exists(Path.GetDirectoryName(createFilePath)))
                {
                    System.IO.Directory.CreateDirectory(Path.GetDirectoryName(createFilePath));
                }
                if (System.IO.File.Exists(createFilePath))
                {
                    System.IO.File.Delete(createFilePath);
                }

                byte[] bFile = this.ReduceByteArr((byte[])dt.Rows[0][0]);
                System.IO.FileStream oFile = new System.IO.FileStream(createFilePath, System.IO.FileMode.Create);
                oFile.Write(bFile, 0, bFile.Length);
                oFile.Flush();
                oFile.Close();
                oFile.Dispose();

                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 把一个文件存入数据库
        /// </summary>
        public int WriteFile(string FilePath, string SqlString, string ParamName, bool Add)
        {
            if (!System.IO.File.Exists(FilePath)) return 0;

            //存储和读取Access数据库的OLE对象一般是转换成byte数组进行处理
            System.IO.FileStream stream = new System.IO.FileStream(FilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            byte[] bData;
            int offset = 0;
            if (Add)
            {
                offset = 10;
                bData = new byte[stream.Length + 10];
                bData[0] = bData[1] = bData[2] = (byte)3;
                bData[3] = bData[4] = bData[5] = (byte)6;
                bData[6] = bData[7] = bData[8] = (byte)9;
                bData[9] = (byte)0;
            }
            else
            {
                bData = new byte[stream.Length];
            }
            stream.Read(bData, offset, (int)stream.Length);

            OleDbCommand cmd = new OleDbCommand(SqlString, this.cn);
            cmd.Parameters.Add(new OleDbParameter(ParamName, bData));

            if (cn.State != ConnectionState.Open) cn.Open();

            cmd.CommandTimeout = 500;
            int a = cmd.ExecuteNonQuery();
            cmd.Dispose();

            stream.Dispose();

            return a;
        }

        public System.Drawing.Image ReadFileToImage(string sql)
        {
            DataTable dt = this.DBTable(sql, null, null);
            if (dt == null || dt.Rows.Count == 0) return null;

            byte[] bFile = (byte[])dt.Rows[0][0];
            bFile = this.ReduceByteArr(bFile);//解密

            System.Drawing.ImageConverter imgC = new System.Drawing.ImageConverter();//将byte[]转换成图片
            System.Drawing.Image img = (Image)imgC.ConvertFrom(bFile);
            return img;
        }
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
        internal byte[] ReduceByteArr(byte[] bFile)
        {
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


    /// <summary>
    /// Database.OperateDB 的摘要说明。
    /// </summary>
    internal class Dbtool1
    {

        #region 初始化连接对象

        internal Dbtool1(string myConnectionString)
        {
            this.ConnectionString = myConnectionString;
        }

        /// <summary>
        /// 数据库的连接字符串
        /// </summary>
        internal string ConnectionString
        {
            get
            {
                return this.cn.ConnectionString;
            }
            set
            {
                this._cnString = value;
                this.cn = new OleDbConnection(value);
            }
        }
        private string _cnString = "";

        /// <summary>
        /// 连接超时最长时间
        /// </summary>
        internal int TimeOut
        {
            get
            {
                return _timeout;
            }
            set
            {
                _timeout = value;
            }
        }
        private int _timeout = 30;


        /// <summary>
        /// 数据库的连接对象
        /// </summary>
        private OleDbConnection cn = null;

        #endregion

        #region 打开或关闭数据库
        /// <summary>
        /// 打开数据库操作
        /// </summary>
        /// <returns>成功返回空,否则返回错误信息</returns>
        internal string Open()
        {
            try
            {
                this.cn.Open();//显示打开数据库连接

                return "";//成功打开,返回空字符串
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        /// <summary>
        /// 关闭数据库操作
        /// </summary>
        /// <returns>成功返回空,否则返回错误信息</returns>
        internal string Close()
        {
            try
            {
                if (cn.State != System.Data.ConnectionState.Closed)
                {
                    this.cn.Close();
                }

                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        #endregion

        #region /---操作数据库---/

        /// <summary>
        /// 得到第一行第一列的值
        /// </summary>
        internal object FirstValue(string SqlString)
        {
            OleDbCommand cmd = new OleDbCommand(SqlString, this.cn);

            if (cn.State != ConnectionState.Open) cn.Open();

            object obj = cmd.ExecuteScalar();

            if (cn.State != ConnectionState.Closed) cn.Close();

            cmd.Dispose();

            return obj;
        }

        /// <summary>
        /// 执行Sql命令,返回一个数据表
        /// </summary>
        internal DataTable DBTable(string SqlString)
        {
            OleDbDataAdapter da = new OleDbDataAdapter(SqlString, this.cn);

            DataTable dt = new DataTable();

            da.Fill(dt);//填充数据表

            da.Dispose();//销毁数据对象

            return dt;
        }

        /// <summary>
        /// 执行Sql命令,返回数据流
        /// </summary>
        internal OleDbDataReader DBReader(string SqlString)
        {
            OleDbCommand cmd = new OleDbCommand(SqlString, this.cn);

            return cmd.ExecuteReader();//返回数据流
        }

        /// <summary>
        /// 执行Sql命令
        /// </summary>
        internal void Excute(string SqlString)
        {
            OleDbCommand cmd = new OleDbCommand(SqlString, this.cn);

            if (cn.State != ConnectionState.Open) cn.Open();

            cmd.CommandTimeout = 500;
            cmd.ExecuteNonQuery();

            if (cn.State != ConnectionState.Closed) cn.Close();
        }

        /// <summary>
        /// 执行Sql命令//有Param，如@aaa,则ParamName=aaa;
        /// </summary>
        internal void Excute(string SqlString, string ParamValue, string ParamName)
        {
            OleDbCommand cmd = new OleDbCommand(SqlString, this.cn);
            cmd.Parameters.Add(new OleDbParameter(ParamName, ParamValue));

            if (cn.State != ConnectionState.Open) cn.Open();

            cmd.CommandTimeout = 500;
            cmd.ExecuteNonQuery();

            if (cn.State != ConnectionState.Closed) cn.Close();
        }


        /// <summary>
        /// 返回受影响的行数
        /// </summary>
        internal int Excute2(string SqlString)
        {
            int a = 0;
            OleDbCommand cmd = new OleDbCommand(SqlString, this.cn);

            if (cn.State != ConnectionState.Open) cn.Open();

            cmd.CommandTimeout = 500;
            a = cmd.ExecuteNonQuery();

            if (cn.State != ConnectionState.Closed) cn.Close();

            return a;
        }

        /// <summary>
        /// 执行Sql命令返回数据集
        /// </summary>
        internal DataSet DBDataSet(string SqlString)
        {
            OleDbDataAdapter da = new OleDbDataAdapter(SqlString, this.cn);

            DataSet ds = new DataSet("Table");

            da.Fill(ds);

            da.Dispose();

            return ds;
        }

        #endregion//一般的SQL命令

    }


    internal class Dbtool2
    {

        internal static System.Windows.Forms.Form strError
        {
            get
            {
                return BeltT_XX.l01100011101();
            }
        }

        internal static bool hasconn(bool showalert, Form frm)
        {
            if (false )
            {
            }
            else
            {
                //return BeltT_XX.l0111001110(showalert);
                return true;
            }
        }

        internal static bool hasclose(string connstr)
        {

            bool b = BeltT_XX.l0101001110();
            b = true;//去除注册
                if (!b)
                {
                    if (connstr.Length == 0)
                    {
                        connstr = "您不是正式注册用户，此功能无法使用！";
                    }
                    StringOperate.Alert(connstr);
                }
                return b;

        }
        internal static bool hasclose()
        {
            return BeltT_XX.l0101001110();
        }

        #region 初始化连接对象

        internal Dbtool2(string myConnectionString)
        {
        }

        /// <summary>
        /// 数据库的连接字符串
        /// </summary>
        internal string ConnectionString
        {
            get
            {
                return "";
            }
            set
            {
            }
        }

        /// <summary>
        /// 连接超时最长时间
        /// </summary>
        internal int TimeOut
        {
            get
            {
                return 0;
            }
            set
            {
            }
        }

        #endregion

        #region 打开或关闭数据库
        /// <summary>
        /// 打开数据库操作
        /// </summary>
        /// <returns>成功返回空,否则返回错误信息</returns>
        internal string Open()
        {
            return "";
        }
        /// <summary>
        /// 关闭数据库操作
        /// </summary>
        /// <returns>成功返回空,否则返回错误信息</returns>
        internal string Close()
        {
            return "";
        }
        #endregion
    }


    /// <summary>
    /// 执行与存储过程有关的SQL命令
    /// </summary>
    internal class DbtoolSP
    {

        #region 初始化连接对象

        internal DbtoolSP ChangeDB(string database)
        {
            this.cn.Open();

            this.cn.ChangeDatabase(database);

            this.cn.Close();

            return this;
        }

        internal DbtoolSP(string myConnectionString)
        {
            this.ConnectionString = myConnectionString;
        }

        internal DbtoolSP(OleDbConnection mySqlConnection)
        {
            this.cn = mySqlConnection;
        }

        private string _cnString = "";
        /// <summary>
        /// 数据库的连接字符串
        /// </summary>
        internal string ConnectionString
        {
            get
            {
                return this.cn.ConnectionString;
            }
            set
            {
                this._cnString = value;
                this.cn = new OleDbConnection(value);
            }
        }

        private int _timeout = 30;
        /// <summary>
        /// 连接超时最长时间
        /// </summary>
        internal int TimeOut
        {
            get
            {
                return _timeout;
            }
            set
            {
                _timeout = value;
            }
        }

        internal string _Database
        {
            get
            {
                return this.cn.Database;
            }
        }

        /// <summary>
        /// 获得数据库名称
        /// </summary>
        internal string _DataSource
        {
            get
            {
                return this.cn.DataSource;
            }
        }

        /// <summary>
        /// 数据库的连接对象
        /// </summary>
        private OleDbConnection cn = null;

        #endregion

        #region 打开或关闭数据库

        /// <summary>
        /// 打开数据库操作
        /// </summary>
        /// <returns>成功返回空,否则返回错误信息</returns>
        internal string Open()
        {
            try
            {
                this.cn.Open();//显示打开数据库连接

                return "";//成功打开,返回空字符串
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        /// <summary>
        /// 关闭数据库操作
        /// </summary>
        /// <returns>成功返回空,否则返回错误信息</returns>
        internal string Close()
        {
            try
            {
                if (cn.State != System.Data.ConnectionState.Closed)
                {
                    this.cn.Close();
                }

                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        #endregion

        #region//与存储过程有关的命令

        /// <summary>
        /// 执行存储过程,返回一个数据表
        /// </summary>
        /// <param name="myProc">存储过程名称</param>
        /// <param name="myParam">存储过程参数</param>
        /// <returns>数据集</returns>
        internal DataTable DBTable(string myProc, params OleDbParameter[] myParam)
        {
            OleDbDataAdapter da = new OleDbDataAdapter(myProc, this.cn);

            da.SelectCommand.CommandType = CommandType.StoredProcedure;

            foreach (OleDbParameter sp in myParam)
            {
                da.SelectCommand.Parameters.Add(sp);
            }

            DataTable dt = new DataTable();

            da.Fill(dt);

            da.Dispose();

            return dt;
        }

        /// <summary>
        /// 执行存储过程或Sql命令,返回数据表
        /// </summary>
        internal DataTable DBTable(string SqlString, CommandType CmdType, params OleDbParameter[] myParam)
        {
            OleDbDataAdapter da = new OleDbDataAdapter(SqlString, this.cn);

            da.SelectCommand.CommandType = CmdType;

            foreach (OleDbParameter sp in myParam)
            {
                da.SelectCommand.Parameters.Add(sp);
            }

            DataTable dt = new DataTable();

            da.Fill(dt);

            da.Dispose();

            return dt;
        }

        /// <summary>
        /// 执行存储过程,返回数据流
        /// </summary>
        /// <param name="myProc">存储过程名称</param>
        /// <param name="myParam">存储过程参数</param>
        /// <returns>数据流</returns>
        internal OleDbDataReader DBReader(string myProc, params OleDbParameter[] myParam)
        {
            OleDbCommand cmd = new OleDbCommand(myProc, this.cn);

            cmd.CommandType = CommandType.StoredProcedure;

            foreach (OleDbParameter sp in myParam)
            {
                cmd.Parameters.Add(sp);
            }

            return cmd.ExecuteReader();
        }

        /// <summary>
        /// 执行存储过程或Sql命令,返回数据流
        /// </summary>
        internal OleDbDataReader DBReader(string SqlString, CommandType CmdType, params OleDbParameter[] myParam)
        {
            OleDbCommand cmd = new OleDbCommand(SqlString, this.cn);

            cmd.CommandType = CmdType;

            foreach (OleDbParameter sp in myParam)
            {
                cmd.Parameters.Add(sp);
            }

            return cmd.ExecuteReader();
        }

        /// <summary>
        /// 执行存储过程,返回存储过程的返回值
        /// </summary>
        /// <param name="myProc">存储过程名称</param>
        /// <param name="index">存储过程返回值的索引</param>
        /// <param name="myParam">存储过程参数</param>
        /// <returns>返回值</returns>
        internal object ReturnValue(string myProc, int index, params OleDbParameter[] myParam)
        {
            OleDbCommand cmd = new OleDbCommand(myProc, this.cn);

            cmd.CommandType = CommandType.StoredProcedure;

            foreach (OleDbParameter sp in myParam)
            {
                cmd.Parameters.Add(sp);
            }

            cmd.Parameters[index].Direction = ParameterDirection.Output;

            cn.Open();

            cmd.ExecuteNonQuery();

            cn.Close();

            return cmd.Parameters[index].Value;
        }

        /// <summary>
        /// 执行存储过程,返回存储过程的返回值
        /// </summary>
        /// <param name="myProc">存储过程名称</param>
        /// <param name="index">存储过程返回值的名称</param>
        /// <param name="myParam">存储过程参数</param>
        /// <returns>返回值</returns>
        internal object ReturnValue(string myProc, string Param, params OleDbParameter[] myParam)
        {
            OleDbCommand cmd = new OleDbCommand(myProc, this.cn);

            cmd.CommandType = CommandType.StoredProcedure;

            foreach (OleDbParameter sp in myParam)
            {
                cmd.Parameters.Add(sp);
            }

            cmd.Parameters[Param].Direction = ParameterDirection.Output;

            cn.Open();

            cmd.ExecuteNonQuery();

            cn.Close();

            return cmd.Parameters[Param].Value;
        }

        /// <summary>
        /// 执行带占位符的Sql命令或存储过程,通过CommandType来判断
        /// </summary>
        internal void Excute(string myProc, CommandType myType, params OleDbParameter[] myParam)
        {
            OleDbCommand cmd = new OleDbCommand(myProc, this.cn);

            cmd.CommandType = myType;

            foreach (OleDbParameter sp in myParam)
            {
                cmd.Parameters.Add(sp);
            }

            cn.Open();

            cmd.ExecuteNonQuery();

            cn.Close();
        }

        /// <summary>
        /// 执行存储过程
        /// </summary>
        /// <param name="myProc"></param>
        /// <param name="myParam"></param>
        internal void Excute(string myProc, params OleDbParameter[] myParam)
        {
            OleDbCommand cmd = new OleDbCommand(myProc, this.cn);

            cmd.CommandTimeout = 500;
            cmd.CommandType = CommandType.StoredProcedure;

            foreach (OleDbParameter sp in myParam)
            {
                cmd.Parameters.Add(sp);
            }

            cn.Open();

            cmd.ExecuteNonQuery();

            cn.Close();
        }

        /// <summary>
        /// 执行存储过程返回数据集
        /// </summary>
        /// <param name="myProc"></param>
        /// <param name="myParam"></param>
        /// <returns></returns>
        internal DataSet DBDataSet(string myProc, params OleDbParameter[] myParam)
        {
            OleDbDataAdapter da = new OleDbDataAdapter(myProc, this.cn);

            da.SelectCommand.CommandType = CommandType.StoredProcedure;

            foreach (OleDbParameter sp in myParam)
            {
                da.SelectCommand.Parameters.Add(sp);
            }

            DataSet ds = new DataSet("Table");

            da.Fill(ds);

            da.Dispose();

            return ds;
        }

        /// <summary>
        /// 执行存储过程或Sql命令返回记录集
        /// </summary>
        /// <param name="SqlString"></param>
        /// <param name="CmdType"></param>
        /// <param name="myParam"></param>
        /// <returns></returns>
        internal DataSet DBDataSet(string SqlString, CommandType CmdType, params OleDbParameter[] myParam)
        {
            OleDbDataAdapter da = new OleDbDataAdapter(SqlString, this.cn);

            da.SelectCommand.CommandType = CmdType;

            foreach (OleDbParameter sp in myParam)
            {
                da.SelectCommand.Parameters.Add(sp);
            }

            DataSet ds = new DataSet("Table");

            da.Fill(ds);

            da.Dispose();

            return ds;
        }

        /// <summary>
        /// 执行存储过程或Sql命令,返回OBJECT类型的FirstValeu
        /// </summary>
        /// <param name="SqlString"></param>
        /// <param name="IsProc"></param>
        /// <param name="myParam"></param>
        /// <returns></returns>
        internal object FirstValue(string SqlString, bool IsProc, params OleDbParameter[] myParam)
        {
            OleDbCommand command1 = new OleDbCommand(SqlString, this.cn);

            if (IsProc)
            {
                command1.CommandType = CommandType.StoredProcedure;
            }
            else
            {
                command1.CommandType = CommandType.Text;
            }


            if (cn.State != ConnectionState.Open) cn.Open();

            foreach (OleDbParameter ObjPram in myParam)
            {
                command1.Parameters.Add(ObjPram);
            }

            object obj1 = command1.ExecuteScalar();

            if (cn.State != ConnectionState.Closed) cn.Close();

            return obj1;
        }

        #endregion
    }


    /// <summary>
    /// 执行与事务有关的操作
    /// </summary>
    internal class DbtoolTrans
    {

        #region 初始化连接对象

        internal DbtoolTrans ChangeDB(string database)
        {
            this.cn.Open();

            this.cn.ChangeDatabase(database);

            this.cn.Close();

            return this;
        }

        internal DbtoolTrans(string myConnectionString)
        {
            this.ConnectionString = myConnectionString;
        }

        internal DbtoolTrans(OleDbConnection mySqlConnection)
        {
            this.cn = mySqlConnection;
        }

        private string _cnString = "";
        /// <summary>
        /// 数据库的连接字符串
        /// </summary>
        internal string ConnectionString
        {
            get
            {
                return this.cn.ConnectionString;
            }
            set
            {
                this._cnString = value;
                this.cn = new OleDbConnection(value);
            }
        }

        private int _timeout = 30;
        /// <summary>
        /// 连接超时最长时间
        /// </summary>
        internal int TimeOut
        {
            get
            {
                return _timeout;
            }
            set
            {
                _timeout = value;
            }
        }

        internal string _Database
        {
            get
            {
                return this.cn.Database;
            }
        }

        /// <summary>
        /// 获得数据库名称
        /// </summary>
        internal string _DataSource
        {
            get
            {
                return this.cn.DataSource;
            }
        }

        /// <summary>
        /// 数据库的连接对象
        /// </summary>
        private OleDbConnection cn = null;
        #endregion


        #region 打开或关闭数据库
        /// <summary>
        /// 打开数据库操作
        /// </summary>
        /// <returns>成功返回空,否则返回错误信息</returns>
        internal string Open()
        {
            try
            {
                this.cn.Open();//显示打开数据库连接

                return "";//成功打开,返回空字符串
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        /// <summary>
        /// 关闭数据库操作
        /// </summary>
        /// <returns>成功返回空,否则返回错误信息</returns>
        internal string Close()
        {
            try
            {
                if (cn.State != System.Data.ConnectionState.Closed)
                {
                    this.cn.Close();
                }

                return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        #endregion


        #region /---执行有事务的数据库操作---/
        internal OleDbTransaction tran;
        /// <summary>
        /// 开始事务
        /// </summary>
        internal void BeginTran()
        {
            this.Open();
            tran = this.cn.BeginTransaction();
        }

        /// <summary>
        /// 保存事务
        /// </summary>
        internal void CommitTran()
        {
            tran.Commit();
            this.Close();
        }

        /// <summary>
        /// 回滚事务
        /// </summary>
        internal void RollbackTran()
        {
            tran.Rollback();
            this.Close();
        }

        /// <summary>
        /// 执行带有事务的数据库操作
        /// </summary>
        /// <param name="sql"></param>
        internal void ExcuteTran(string sqlString)
        {
            OleDbCommand cmd = new OleDbCommand(sqlString, this.cn);
            cmd.Transaction = tran;

            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 执行带有事务的数据库操作
        /// </summary>
        /// <param name="sqlString"></param>
        /// <param name="myParam"></param>
        internal void ExcuteTran(string sqlString, params OleDbParameter[] myParam)
        {
            OleDbCommand cmd = new OleDbCommand(sqlString, this.cn);
            cmd.Transaction = tran;

            foreach (OleDbParameter sp in myParam)
            {
                cmd.Parameters.Add(sp);
            }

            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 执行带有事务的操作，且返回一个数据表
        /// </summary>
        /// <param name="sqlString"></param>
        /// <returns></returns>
        internal DataTable DBTableTran(string sqlString)
        {
            OleDbCommand cmd = new OleDbCommand(sqlString, this.cn);
            cmd.Transaction = tran;

            OleDbDataAdapter da = new OleDbDataAdapter(cmd);

            DataTable dt = new DataTable();

            da.Fill(dt);
            da.Dispose();
            return dt;
        }

        /// <summary>
        /// 执行带有事务的操作，且返回一个数据表
        /// </summary>
        /// <param name="sqlString"></param>
        /// <returns></returns>
        internal DataTable DBTableTran(string sqlString, params OleDbParameter[] myParam)
        {
            OleDbCommand cmd = new OleDbCommand(sqlString, this.cn);
            cmd.Transaction = tran;

            foreach (OleDbParameter sp in myParam)
            {
                cmd.Parameters.Add(sp);
            }

            OleDbDataAdapter da = new OleDbDataAdapter(cmd);

            DataTable dt = new DataTable();

            da.Fill(dt);

            return dt;
        }

        /// <summary>
        /// 执行带有事务的操作, 且返回一个DataSet
        /// </summary>
        /// <param name="sqlString"></param>
        /// <returns></returns>
        internal DataSet DBDataSetTran(string sqlString)
        {
            OleDbCommand cmd = new OleDbCommand(sqlString, this.cn);
            cmd.Transaction = tran;

            OleDbDataAdapter da = new OleDbDataAdapter(cmd);

            DataSet ds = new DataSet("Table");
            da.Fill(ds);

            return ds;
        }

        #endregion
    }



}
