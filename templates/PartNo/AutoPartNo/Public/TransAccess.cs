using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using System.Collections;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;
using System.IO;
using Interop.Office.Core;
using System.Drawing;

internal class TransAccess
{
    /// <summary>
    /// ErrorClient:客户端错误提示
    /// ErrorServer:服务器端错误提示
    /// </summary>
    internal TransAccess(string IP, int iport)
    {
        this.ipaddr = IP;
        this.port = iport;
    }

    private int port = 3547;//3547,3546
    private string ipaddr = "";
    private long LCache = 10000;//10KB的缓存,缓存不易设置过大
    private int iTimeOut = 5000;//超时长度5秒
    private SerializeInfo Ser = new SerializeInfo();




    /// <summary>
    /// 去String命令，返回string处理结果，没有具体缓存长度 (如果出错，以 ErrorServer：ErrorClient:  开头)
    /// </summary>
    internal string TCPString(string strIn)
    {
        string data = "";//返回的数

        IPEndPoint ipep = new IPEndPoint(IPAddress.Parse(ipaddr), port);
        Socket server = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        server.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendTimeout, iTimeOut);
        server.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveTimeout, iTimeOut);
        try
        {
            server.Connect(ipep); //试着连接服务器
            NetworkStream ns = new NetworkStream(server);
            StreamReader sr = new StreamReader(ns);
            StreamWriter sw = new StreamWriter(ns);
            sw.WriteLine(strIn);
            sw.Flush();
            data = sr.ReadLine();//如果服务器有错，返回的有可能是ErrorServer：信息
            sr.Close();
            sw.Close();
            ns.Close();
        }
        catch (Exception e)
        {
            data = "ErrorClient:" + e.Message;
        }
        if (server.Connected)
        {
            server.Shutdown(SocketShutdown.Both);
            server.Close();
        }
        return data;
    }


    /// <summary>
    /// 读取一个数据表,出错返回ErrorClient：或ErrorServer：
    /// </summary>
    internal DataTable TCPDataTable(string strIn, ref string strError)
    {
        byte[] btArr = null;//返回的数

        IPEndPoint ipep = new IPEndPoint(IPAddress.Parse(ipaddr), port);
        Socket server = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        server.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendTimeout, iTimeOut);
        server.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveTimeout, iTimeOut);
        try
        {
            server.Connect(ipep); //试着连接服务器
            NetworkStream ns = new NetworkStream(server);
            StreamReader sr = new StreamReader(ns);
            StreamWriter sw = new StreamWriter(ns);

            sw.WriteLine(strIn);
            sw.Flush();

            btArr = new byte[LCache];
            int length = ns.Read(btArr, 0, btArr.Length);
            ns.Flush();
            string strinfo = Encoding.Default.GetString(btArr, 0, ((length < 50) ? length : 50));
            if (strinfo.Trim().StartsWith("ErrorServer:"))
            {
                strError = Encoding.Default.GetString(btArr, 0, length);
                btArr = null;
            }
            if (strinfo.StartsWith("ErrorBig"))//缓冲区长度不够//重设长度//回"ErrorOK:"//重新接收
            {
                int serverleng = Convert.ToInt32(strinfo.Substring(strinfo.IndexOf(":") + 1).Trim());
                sw.WriteLine("ErrorOK:");
                sw.Flush();
                //服务器有可能没有输入流，要等一等
                btArr = new byte[serverleng];

                int iHas = 0;
                while (iHas < serverleng)
                {
                    int L = ns.Read(btArr, iHas, serverleng - iHas);
                    iHas+=L;
                }
            }

            sr.Close();
            sw.Close();
            ns.Close();
            server.Shutdown(SocketShutdown.Both);
            server.Close();
        }
        catch (Exception e)
        {
            strError = "ErrorClient:" +  e.Message;
        }
        if (server.Connected)
        {
            server.Shutdown(SocketShutdown.Both);
            server.Close();
        }

        DataTable dt = new DataTable();
        if (btArr != null && strError.Length ==0)
        {
            dt = (DataTable)this.Ser.deSerialize(btArr);
        }
        return dt;
    }


    /// <summary>
    /// 插入一个表，返回插入的行数，出错返回ErrorClient：或ErrorServer：
    /// </summary>
    internal string TCPUpTable(string strIn, byte[] btArr)
    {
        string data = "";//返回的数

        IPEndPoint ipep = new IPEndPoint(IPAddress.Parse(ipaddr), port);
        Socket server = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        server.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendTimeout, iTimeOut);
        server.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveTimeout, iTimeOut);
        try
        {
            server.Connect(ipep); //试着连接服务器
            NetworkStream ns = new NetworkStream(server);
            StreamReader sr = new StreamReader(ns);
            StreamWriter sw = new StreamWriter(ns);
            sw.WriteLine(strIn);
            sw.Flush();
            data = sr.ReadLine();
            if (data == "OK")
            {
                ns.Write(btArr, 0, btArr.Length);//发送表格
                ns.Flush();
                data = sr.ReadLine();//读取处理的结果
            }
            sr.Close();
            sw.Close();
            ns.Close();
        }
        catch (Exception e)
        {
            data = "ErrorClient:" + e.Message;
        }
        if (server.Connected)
        {
            server.Shutdown(SocketShutdown.Both);
            server.Close();
        }
        return data;
    }


}
