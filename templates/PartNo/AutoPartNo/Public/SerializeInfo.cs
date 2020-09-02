using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using System.Collections;

/// <summary>
/// 类型转换类
/// </summary>
internal class SerializeInfo
{
    internal SerializeInfo()
    {
    }


    /// <summary>
    /// 序列化object对象//无法序列SqlParameter[]
    /// </summary>
    /// <param name="ds">object类型的参数</param>
    /// <returns></returns>
    internal byte[] Serialize(object ds)
    {
        MemoryStream ms = new MemoryStream();
        IFormatter formatter = new BinaryFormatter();
        //ds.RemotingFormat = SerializationFormat.Binary; 

        formatter.Serialize(ms, ds);
        byte[] binaryResult = ms.ToArray();
        ms.Close();
        ms.Dispose();
        return binaryResult;
    }
    /// <summary>
    /// 反序列化为object
    /// </summary>
    /// <param name="binaryData">byte[]参数</param>
    /// <returns></returns>
    internal object deSerialize(byte[] binaryData)
    {
        try
        {
            //创建内存流
            MemoryStream memStream = new MemoryStream(binaryData);
            //产生二进制序列化格式
            IFormatter formatter = new BinaryFormatter();
            //反串行化到内存中

            object obj = formatter.Deserialize(memStream);
            return obj;
        }
        catch
        {
            return null;
        }
    }



    /// <summary>
    /// 把ArrayList类型的参数转换成string[]
    /// </summary>
    internal string[] ArrayListToArr(ArrayList arr)
    {
        if (arr == null || arr.Count == 0) return null;

        string[] strArr = new string[arr.Count];
        for (int i = 0; i < arr.Count; i++)
        {
            strArr[i] = arr[i].ToString();
        }
        return strArr;
    }
    /// <summary>
    /// 把ArrayList类型的参数转换成string[]
    /// </summary>
    internal object[] ArrayListToObj(ArrayList arr)
    {
        if (arr == null || arr.Count == 0) return null;

        object[] objArr = new object[arr.Count];
        for (int i = 0; i < arr.Count; i++)
        {
            objArr[i] = arr[i];
        }
        return objArr;
    }


    /// <summary>
    /// 类型转换,返回"varchar (100)","int",等等字符串
    /// </summary>
    internal string ChangeType(Type type)
    {
        if (type.Equals(typeof(System.String))) return "varchar (200)";
        //return SqlDbType.VarChar;
        if (type.Equals(typeof(System.Byte))) return "SmallInt";
        if (type.Equals(typeof(System.Int16))) return "int";
        if (type.Equals(typeof(System.Int32))) return "int";
        if (type.Equals(typeof(System.Int64))) return "int";
        if (type.Equals(typeof(System.Single))) return "real";//access中real代表单精度
        if (type.Equals(typeof(System.Double))) return "real";//"float";防止自动加上小数
        // return SqlDbType.Int;
        if (type.Equals(typeof(System.DateTime))) return "datetime";
        //return SqlDbType.Timestamp;
        if (type.Equals(typeof(System.Decimal))) return "Decimal";
        // return SqlDbType.Decimal;
        if (type.Equals(typeof(System.Boolean))) return "bit";
        // return SqlDbType.Bit;
        return "varchar (100)";
    }


    /// <summary>
    /// 将字符串转换成指定类型的Object数据
    /// </summary>
    /// <param name="type"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static object ToType(Type type, string value)
    {
        return System.ComponentModel.TypeDescriptor.GetConverter(type).ConvertFrom(value);
    }





}
