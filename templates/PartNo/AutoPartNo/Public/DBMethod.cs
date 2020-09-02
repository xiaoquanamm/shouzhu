using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Collections;


namespace Interop.Office.Core
{
    /// <summary>
    /// 统一Sql和Access 操作的接口
    /// </summary>
    public interface DBMethod
    {
        string Connstr { get; set; }

        object FirstValue(string SqlString);
        object FirstValue(string SqlString, string[] arrParamName, object[] arrParamValue);

        DataTable DBTable(string SqlString, string[] arrParamName, object[] arrParamValue);
        DataTable DBTable(string SqlString);

        int Excute(string SqlString);
        int Excute(string SqlString, string[] arrParamName, object[] arrParamValue);

        //DataSet DBDataSet(string SqlString);

        /// <summary>
        /// 检查Sql语句
        /// </summary>
        string checkSql(string sql);

        /// <summary>
        /// 得到数据库中所有的表
        /// </summary>
        ArrayList GetAllTableNames();

        /// <summary>
        /// 根据Datatable在数据库中创建一个表
        /// </summary>
        int CreateTable(DataTable dt, string tbName);


        System.Drawing.Image ReadFileToImage(string SqlString);

        /// <summary>
        /// 根据查询语句从数据库中读取一个文件，并保存到指定的目录下,返回是否成功
        /// </summary>
        bool ReadFile(string sql, string createFilePath);

        /// <summary>
        ///  把一个文件存入数据库，（文件路径，sql语句，文件参数名@file ,要加密吗 ),返回受影响的行数
        /// </summary>
        int WriteFile(string FilePath, string SqlString, string ParamName, bool Add);//有Param，如@aaa,这个参数代表文件
    }
}