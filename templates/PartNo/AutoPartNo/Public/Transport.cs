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
using System.Drawing;
using System.DirectoryServices;
using System.Threading;
using Interop.Office.Core;


//在C#进行网络传输的时候buffer的大小不能太大才能保证安全，如果文件很大，
//需要连续传输这样的情况不能直接用一个循环连续发送和接收，因为这样可能在
//你收到的文件中间出现一些空字节，建议你在每传送一定大小的数据后，
//接收段给发送端一次握手信号，这要就可以保证传输正确了
//要是返回表,有可能返回null(出错)
//要是返回string ,有可能返回""(出错);
//要是返回受影响行数,有可能返回-1(出错);正常情况下是>=0的数
//查询最好用table,否则要是返回的正确结果也为"",或-1,就有可能判断为错误    
internal class Transport
{
    /// <summary>
    /// 从配置文件中读取IP 和 端口
    /// </summary>
    internal Transport()
    {
        this.ipaddr = PSetUp.fntWeb_ServerIP;
        this.port = PSetUp.fntWeb_ServerPort;
    }
    internal Transport(string ip, int port)
    {
        this.ipaddr = ip;
        this.port = port;
    }

    internal static int UserID = -1;//用户ID号:0--用户名或密码错,-1---还没有登陆,>=1的数-----成功登陆返回的UserID
    internal int port = 3548;//从注册表中读出端口号来,保存在此
    internal int portAccept = 8002;//接收数据的端口
    internal string ipaddr = "119.188.10.131";//保存IP地址
    internal static string IPAddr //自动获得本机的IP地址
    {
        get
        {
            IPAddress[] ip = Dns.GetHostAddresses(Dns.GetHostName());
            foreach (IPAddress addr in ip)
            {
                if (addr.ToString().Length < 16) return addr.ToString();
            }
            return "";
        }
    }


    #region//重要通信方法

    internal int iTimeOut = 5000;//5秒
    internal int LCache = 1000000;//缓冲区长度,可以随便定

    internal static int iAllLeng = 0;//需要判断百分比的，根据这两个值
    internal static int iHasLeng = 0;//目前还没有用

    /// <summary>
    /// 去String命令，返回string处理结果 (如果出错，以 ErrorServer：ErrorClient:  开头)
    /// 这些命令都是以2开头的：
    /// 2ConnStr:all
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
            data = sr.ReadLine();
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
    /// 从服务器下载文件,成功了返回空“”，否则返回错误原因 "路径中不能有: 如:C:\\要改成:C.\\"
    /// FilePath默认是【根目录】下面的目录开始的，如“MD3DParts\\文件说明.txt”
    /// </summary>
    internal string TCPDownFile(string FilePath, string SaveFilePath)//在主线程内的方法
    {
        string data = "";    //返回的数

        IPEndPoint ipep = new IPEndPoint(IPAddress.Parse(ipaddr), port);
        Socket server = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        server.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendTimeout, iTimeOut);
        server.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveTimeout, iTimeOut);
        NetworkStream ns = null;
        StreamReader sr = null;
        StreamWriter sw = null;
        try
        {
            server.Connect(ipep); //试着连接服务器
            ns = new NetworkStream(server);
            sr = new StreamReader(ns);
            sw = new StreamWriter(ns);

            //下载文件<<0ReadFile$准备长度:文件名:用户ID:文件类型:创建时间:修改时间:访问时间>>来<<byte[]>> or (((<<ErrorBig>> 去<<ErrorOK>> 来<<byte[]>>)))   
            sw.WriteLine("0ReadFile$0:" + FilePath + ":0:0");//去请求
            sw.Flush();
            string ret = sr.ReadLine();//"ErrorBig:" + s.Length.ToString()

            if (ret.Trim().StartsWith("ErrorServer"))//服务器端的数据库操作失败,返回失败信息
            {
                data = ret;
            }
            else if (ret.StartsWith("ErrorBig"))//缓冲区长度不够//重设长度//回"ErrorOK:"//重新接收
            {
                string[] ParamArr = ret.Split(new char[] { ':' }); //"ErrorBig:25458:0:0:0:2011-3-4 11.32.00:2011-3-4 11.32.00:2011-3-4 11.32.00"

                sw.WriteLine("ErrorOK:");
                sw.Flush();
                int hasleng = 0;
                int allleng = Convert.ToInt32(ParamArr[1]);

                //如果没有这个路径,就会出错,所以要新建路径
                string sPathName = Path.GetDirectoryName(SaveFilePath);
                if (!Directory.Exists(sPathName))
                {
                    Directory.CreateDirectory(sPathName);
                }

                Stream s = File.Open(SaveFilePath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                byte[] bytFile = new byte[LCache];
                while (hasleng < allleng)//没有读到预定的长度
                {
                    int thisL = ns.Read(bytFile, 0, bytFile.Length);
                    s.Position = hasleng;
                    s.Write(bytFile, 0, thisL);//将接收的信息写入文件
                    s.Flush();
                    hasleng += thisL;
                }
                s.Close();
                s.Dispose();

                //写入文件的属性
                if (ParamArr.Length > 7)
                {
                    if (ParamArr[5].Length > 5) File.SetCreationTime(SaveFilePath, Convert.ToDateTime(ParamArr[5].Replace(".", ":")));
                    if (ParamArr[6].Length > 5) File.SetLastWriteTime(SaveFilePath, Convert.ToDateTime(ParamArr[6].Replace(".", ":")));
                    if (ParamArr[7].Length > 5) File.SetLastAccessTime(SaveFilePath, Convert.ToDateTime(ParamArr[7].Replace(".", ":")));
                }
            }
        }
        catch (Exception e) //出错了,把Succeed设为false
        {
            data = "客户端错误提示:" + e.Message;
        }
        finally
        {
            if (sr != null) sr.Close();
            if (sw != null) sw.Close();
            if (ns != null) ns.Close();
            if (server.Connected)
            {
                server.Shutdown(SocketShutdown.Both);
            }
            server.Close();
        }
        return data;
    }

    /// <summary>
    /// 上传文件到服务器，成功了返回空“”，否则返回错误原因
    /// </summary>
    internal string TCPUpFile(string fileFullName, string SaveFilePath)//在主线程内的方法
    {
        if (!File.Exists(fileFullName)) return "文件不存在!";
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
            //只有这样写，在文件被其它程序使用的时候，才不至于出错
            System.IO.FileStream s = new System.IO.FileStream(fileFullName, System.IO.FileMode.Open, System.IO.FileAccess.Read, FileShare.ReadWrite);
            //上传文件//去<<9WriteFile$文件长度:文件名:用户ID:文件类型:创建时间:修改时间:访问时间>>来<<9OK>>去<<byte[]>>

            //设定文件的修改时间
            FileInfo fi = new FileInfo(fileFullName);
            string TimeInfo = fi.CreationTime.ToString().Replace(":", ".") + ":" + fi.LastWriteTime.ToString().Replace(":", ".") + ":" + fi.LastAccessTime.ToString().Replace(":", ".");

            sw.WriteLine("9WriteFile$" + s.Length.ToString() + ":" + SaveFilePath + ":0:0:" + TimeInfo);//去请求
            sw.Flush();
            string ret = sr.ReadLine();//"9OK"

            int ihasOk = 0;
            byte[] bytFile = new byte[LCache];
            while (ihasOk < s.Length)
            {
                s.Position = ihasOk;
                int thisL = s.Read(bytFile, 0, bytFile.Length);
                ns.Write(bytFile, 0, thisL);//发流
                ns.Flush();
                ihasOk += thisL;
            }
            s.Close();
            s.Dispose();

            sr.Close();
            sw.Close();
            ns.Close();
        }
        catch (Exception e)
        {
            data = "上传文件失败:" + e.Message;
        }
        if (server.Connected)
        {
            server.Shutdown(SocketShutdown.Both);
            server.Close();
        }
        return data;
    }


    #endregion//end重要通信方法


    #region//这里是一些附属的功能

    /// <summary>
    /// 是否可以联通到服务器,成功返回“”,失败返回错误原因
    /// </summary>
    internal string CanLinkServer
    {
        get
        {
            string str = this.TCPString("2ConnStr:all");
            if (str.StartsWith("Error"))
            {
                return str;
            }
            else
            {
                return "";
            }
        }
    }

    //注销返回“true","false"
    internal bool LogOut()
    {
        string str = this.TCPString("2LogOut:" + Transport.UserID.ToString());
        if (str == "true")
        {
            return true;
        }
        return false;
    }

    /// <summary>
    /// 创建目录,,目录前没有\\,目录已经存在也OK
    /// </summary>
    internal bool CreateFolder(string path)
    {
        string str = this.TCPString("2CreateFolder:" + path);
        if (str == "OK")
        {
            return true;
        }
        return false;
    }

    /// <summary>
    /// 检查服务器上的某个文件夹是否存在,（还需要完善，通信不成功按文件不存在处理）
    /// </summary>
    internal bool FolderExists(string path)
    {
        string str = this.TCPString("2FolderExists:" + path);
        if (str == "OK")
        {
            return true;
        }
        return false;
    }

    /// <summary>
    /// 检查服务器上的某个文件是否存在
    /// </summary>
    internal bool FileExists(string path)
    {
        string str = this.TCPString("2FileExists:" + path);
        if (str == "OK")
        {
            return true;
        }
        return false;
    }

    /// <summary>
    /// 检查服务器上的某个文件的信息 fi.LastWriteTime + "♂" + fi.Length + "♂"; 如果没有文件返回"Error:none" ,出错返回"Error:+e.message",客户端出错：ErrorClient:
    /// </summary>
    internal string FileInfo(string path)
    {
        return this.TCPString("2FileInfo:" + path);
    }

    /// <summary>
    /// 删除文件夹//和里面的文件//或删除文件//参数是服务器路径，如果是服务器全路径，以C.//开头
    /// </summary>
    internal bool DeleteFileAndFolder(string fullPath)
    {
        string str = this.TCPString("2DeleteFileAndFolder:" + fullPath);
        if (str == "OK")
        {
            return true;
        }
        return false;
    }

    /// <summary>
    /// 重命名服务器上的文件(或文件夹)成功返回"OK"否则返回错误信息
    /// </summary>
    internal string RenameFileOrFolder(string oldfile, string newfile)
    {
        string str = this.TCPString("2RenameFileOrFolder:0ぺ0ぺ" + oldfile + "ぺ" + newfile);
        if (str.StartsWith("Error"))
        {
            return str;//返回的是错误信息
        }
        else
        {
            if (str == "Has") return "该文件(或文件夹)已经存在";
            if (str == "OK") return "OK";
            if (str == "None") return "没有找到此文件(或文件夹),所以无法重命名";
            else return "服务器操作失败"; //(ret == "Error")
        }
    }

    /// <summary>
    /// 读取一个目录下的文件,文件夹名称,要是目录以:结尾
    /// </summary>
    internal string[] ReadFileAndFolderName(string Path)
    {
        string str = this.TCPString("2ReadFileAndFolderName:" + Path);
        if (str.StartsWith("Error"))
        {
            return null;//返回的是错误信息
        }
        else
        {
            if (str.Length < 3 || str == "None") return null;
            string[] strArr = str.Split('ぺ');
            if (strArr.Length > 0) return strArr;
        }
        return null;
    }

    /// <summary>
    /// 根据Exe文件的名称,从服务器读取这个文件的图标,失败返回null;
    /// </summary>
    internal System.Drawing.Icon GetExeIconFromServer(string FullPath, bool isLarge)
    {
        string str = this.TCPString("2BackIcon$0:" + FullPath + ":" + isLarge.ToString() + ":0");
        if (str.StartsWith("Error"))
        {
            return null;
        }
        else
        {
            byte[] btArr = System.Text.Encoding.ASCII.GetBytes(str);
            System.Drawing.IconConverter icc = new System.Drawing.IconConverter();
            System.Drawing.Icon _icon = icc.ConvertFrom(btArr) as System.Drawing.Icon;
            return _icon;
        }
    }

    /// <summary>
    /// 保存文件夹到服务器(如：C:\\ab\\cc, MD3Dtools\\ab\\dd,将[cc]里的所有文件复制到[dd]中),成功返回“”，否则返回错误提示
    /// </summary>
    internal string saveFolderToServer2(string ClientFullPath, string ServerFullPath, ref int iFileOKCount)
    {
        DirectoryInfo dir = new DirectoryInfo(ClientFullPath);
        if (!dir.Exists)
        {
            return "本机目录【" + dir.FullName + "】不存在" + Environment.NewLine;
        }

        if (!ServerFullPath.EndsWith("\\")) ServerFullPath = ServerFullPath + "\\";
        bool b = CreateFolder(ServerFullPath);//先创建一个文件夹,文件夹创建不成功就返回
        if (!b)
        {
            return "在服务器创建顶级目录【" + ServerFullPath + "】失败！" + Environment.NewLine;
        }

        string alertString = "";
        foreach (FileInfo fi in dir.GetFiles())
        {
            if (fi.Name == "Thumbs.db") continue;
            string sOK = this.TCPUpFile(fi.FullName, ServerFullPath + fi.Name);
            if (sOK.Length > 0)
            {
                alertString += sOK + Environment.NewLine;
            }
            else
            {
                iFileOKCount++;
            }
        }

        foreach (DirectoryInfo dirsub in dir.GetDirectories())
        {
            alertString += this.saveFolderToServer2(dirsub.FullName, ServerFullPath + dirsub.Name, ref iFileOKCount);
        }

        return alertString.Trim();
    }
    /// <summary>
    /// 保存文件夹到服务器(如：C:\\ab\\cc, MD3Dtools\\dd, 将[cc]复制到[dd]中),成功返回“”，否则返回错误提示
    /// </summary>
    internal string saveFolderToServer(string ClientFullPath, string ServerFullPath, ref int iFileOKCount)
    {
        DirectoryInfo dir = new DirectoryInfo(ClientFullPath);
        if (!dir.Exists)
        {
            return "本机目录【" + dir.FullName + "】不存在" + Environment.NewLine;
        }

        if (!ServerFullPath.EndsWith("\\")) ServerFullPath = ServerFullPath + "\\";
        bool b = CreateFolder(ServerFullPath + dir.Name);//先创建一个文件夹,文件夹创建不成功就返回
        if (!b)
        {
            return "在服务器创建顶级目录【" + ServerFullPath + dir.Name + "】失败！" + Environment.NewLine;
        }

        string alertString = "";
        foreach (FileInfo fi in dir.GetFiles())
        {
            if (fi.Name == "Thumbs.db") continue;
            string sOK = this.TCPUpFile(fi.FullName, ServerFullPath + dir.Name + "\\" + fi.Name);
            if (sOK.Length > 0)
            {
                alertString += sOK + Environment.NewLine;
            }
            else
            {
                iFileOKCount++;
            }
        }

        foreach (DirectoryInfo dirsub in dir.GetDirectories())
        {
            alertString += this.saveFolderToServer(dirsub.FullName, ServerFullPath + dir.Name + "\\", ref iFileOKCount);
        }

        return alertString.Trim();
    }

    /// <summary>
    /// 检查一个文件是否需要更新，如果需要就更新,更新成功“”，更新失败返回错误提示，不需要更新返回noneed
    /// </summary>
    internal string CheckAndDownFile(string ClientFullPath)
    {
        string serverPath = ClientFullPath.Substring(AllData.StartUpPath.Length + 1);
        bool bNeed = true;

        //如果文件不存在,一定需要下载//如果文件存在,检查新旧,如果旧的就需要更新,
        FileInfo fi = new FileInfo(ClientFullPath);
        if (fi.Exists && fi.Length > 0)
        {
            DateTime dtClient = File.GetLastWriteTime(ClientFullPath);
            string[] strArr = this.FileInfo(serverPath).Split(new char[] { '♂' });//fi.LastWriteTime + "♂" + fi.Length + "♂";
            if (strArr.Length > 2)
            {
                DateTime dtServer = Convert.ToDateTime(strArr[0]);
                if (dtClient >= dtServer)
                {
                    bNeed = false;
                }
            }

        }

        if (bNeed)
        {
            return TCPDownFile(serverPath, ClientFullPath);
        }

        return "noneed";
    }
    /// <summary>
    /// 检查一个文件是否比服务器上的新，如果新就上传,更新成功“”，更新失败返回错误提示，不需要更新返回noneed
    /// </summary>
    internal string CheckAndUPFile(string ClientFullPath)
    {
        FileInfo fi = new FileInfo(ClientFullPath);
        if (!fi.Exists || fi.Length == 0)
        {
            return "ErrorClient:客户端文件不存在或长度为0";
        }

        bool bNeed = false;
        DateTime dtClient = File.GetLastWriteTime(ClientFullPath);
        string serverPath = ClientFullPath.Substring(AllData.StartUpPath.Length + 1);
        string s = this.FileInfo(serverPath);
        if (s == "Error:none")
        {
            bNeed = true;
        }
        else
        {
            string[] strArr = s.Split(new char[] { '♂' });//fi.LastWriteTime + "♂" + fi.Length + "♂"
            if (strArr.Length > 2)
            {
                DateTime dtServer = Convert.ToDateTime(strArr[0]);
                if (dtClient > dtServer)
                {
                    bNeed = true;
                }
                else
                {
                    return "noneed";
                }
            }
            else
            {
                return "Error:无法读取服务器信息";
            }
        }

        if (bNeed)
        {
            return this.TCPUpFile(ClientFullPath, serverPath);
        }

        return "";
    }

    /// <summary>
    /// 文件浏览器使用,浏览指定用户ID的项目,从服务器读取一个文件,保存到"TempFile"目录下
    /// </summary>
    internal bool ReadFileFromServer(string FullPath, int kbyte)//后两项为0
    {
        if (Transport.UserID <= 0)
        {
            return false;//"您还没有登陆,请先登陆!!";
        }

        string patha = System.Windows.Forms.Application.StartupPath.ToString();
        if (FullPath.IndexOf("\\") != -1) FullPath = FullPath.Substring(FullPath.LastIndexOf("\\") + 1);
        string path = patha + "\\TempFile\\" + FullPath;
        try
        {
            System.IO.FileInfo fi = new System.IO.FileInfo(path);
            if (fi.Exists) fi.Delete();
        }
        catch { }
        string str = this.TCPDownFile(FullPath, path);
        if (str.StartsWith("Error"))
        {
            return false;
        }
        else
        {
            return true;
        }
    }

    /// <summary>
    /// 从服务器读取一个文件夹下的所有文件，保存到制定的目录下
    /// </summary>
    internal bool ReadFolderFromServer(string FullPath, string SavePath, int kbyte)
    {
        if (!FullPath.EndsWith("\\")) FullPath = FullPath + "\\";
        if (!SavePath.EndsWith("\\")) SavePath = SavePath + "\\";

        if (!Directory.Exists(SavePath)) Directory.CreateDirectory(SavePath);
        Interop.Office.Core.StringOperate pm = new Interop.Office.Core.StringOperate();

        string[] strArr = ReadFileAndFolderName(FullPath);
        foreach (string s in strArr)
        {
            if (s.EndsWith(":"))//文件夹
            {
                string FolderName = s.TrimEnd(new char[] { ':' });
                ReadFolderFromServer(FullPath + FolderName, SavePath + FolderName, 1);
            }
            else
            {
                string str = this.TCPDownFile(FullPath + s, SavePath);
            }
        }

        return true;
    }


    #endregion//end附属功能


    #region//获取局域网中的服务器IP地址，全部扫描一遍。

    /// <summary>
    /// 是否已经查到服务器
    /// </summary>
    private bool bHasGetServerIP = false;

    /// <summary>
    /// 当前有多少个线程正在运行
    /// </summary>
    private int iActiveThreadCount = 0;

    /// <summary>
    /// 扫描整个网络，检测哪台电脑的3548端口开放，若开放就返回此服务器的IP地址,否侧返回""
    /// </summary>
    internal string GetServerIPByScreen()
    {
        this.iActiveThreadCount = 0;
        this.bHasGetServerIP = false;

        DirectoryEntry root = new DirectoryEntry("WinNT:");
        DirectoryEntries domains = root.Children;
        domains.SchemaFilter.Add("domain");
        foreach (DirectoryEntry domain in domains)
        {
            DirectoryEntries computers = domain.Children;
            computers.SchemaFilter.Add("computer");
            foreach (DirectoryEntry computer in computers)
            {
                string[] arr = new string[3];
                IPHostEntry iphe = null;
                try
                {
                    iphe = Dns.GetHostEntry(computer.Name);

                    arr[0] = domain.Name;
                    arr[1] = computer.Name;
                    if (iphe != null)
                    {
                        arr[2] += iphe.AddressList[0].ToString();
                        if (bHasGetServerIP) return PSetUp.fntWeb_ServerIP;//已经得到了服务器的IP地址

                        //启动带参数的线程
                        Thread thd = new Thread(new ParameterizedThreadStart(TCPCanLine));
                        thd.Start(arr[2]);
                    }
                }
                catch (Exception ex)
                {
                    StringOperate.Alert(ex.Message);
                }
            }
        }

        int iTimes = 0;//这里没必要等很长时间，如果服务器开启，瞬间就会有反应，再等时间长了也没用
        while (this.iActiveThreadCount > 0)
        {
            if (bHasGetServerIP) return PSetUp.fntWeb_ServerIP;//已经得到了服务器的IP地址
            System.Threading.Thread.Sleep(100);
            iTimes++;
            if (iTimes > 60) break;
        }

        return "";
    }

    //测试连接
    private void TCPCanLine(object objIP)
    {
        this.iActiveThreadCount++;

        IPEndPoint ipep = new IPEndPoint(IPAddress.Parse(objIP.ToString()), PSetUp.fntWeb_ServerPort);
        Socket server = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        server.SendTimeout = 2000;
        server.ReceiveTimeout = 2000;
        try
        {
            server.Connect(ipep); //试着连接服务器
            server.Send(Encoding.Default.GetBytes("Test Connection"));

            //链接成功了
            PSetUp.fntWeb_ServerIP = objIP.ToString();
            this.bHasGetServerIP = true;
            StringOperate.Alert("检测到服务器的IP地址是：" + objIP.ToString());
        }
        catch (Exception e)
        { }

        iActiveThreadCount--;

        if (server.Connected)
        {
            server.Shutdown(SocketShutdown.Both);
        }
        server.Close();
    }

    #endregion


}

