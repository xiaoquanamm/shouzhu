using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Net;
using System.Data;
using System.Collections;
//using Interop.Office.Core.Exe;

namespace Interop.Office.Core
{
    /// <summary>
    /// 在启动程序时会自动检测一次是否联网,并读取在线配置。
    /// </summary>
    internal class TransWeb
    {
        internal TransWeb()
        {
        }


        #region//关于Web登录，保存Cookie


        /// <summary>
        /// 把Cookie添加到Webbrowser中
        /// </summary>
        /// <param name="webb">控件</param>
        /// <param name="cookies">Cookies集合</param>
        /// <param name="url">如【http://www.my3dparts.com】,为空时采用默认url</param>
        /// <returns>“”</returns>
        internal string AddCookiesToWebBrowser(WebBrowser webb, CookieContainer cookies, string url)
        {
            if (url.Length == 0)
            {
                url = "http://www.my3dparts.com";
            }

            foreach (Cookie cc in cookies.GetCookies(new Uri(url)))
            {
                InternetSetCookie(url, cc.Name, cc.Value);
            }

            return "";
        }

        //V6.02以后身份认证
        [System.Runtime.InteropServices.DllImport("wininet.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
        internal static extern bool InternetSetCookie(string lpszUrlName, string lbszCookieName, string lpszCookieData);


        /// <summary>
        /// 把WebBrowser中的Cookie 保存在一个CookieContainer中。
        /// </summary>
        internal CookieContainer GetCookiesByWebBrowser(WebBrowser webBr)
        {
            //在WebBrowser中登录cookie保存在WebBrowser.Document.Cookie中      
            CookieContainer cookies = new CookieContainer();

            //String 的Cookie　要转成　Cookie型的　并放入CookieContainer中  
            string cookieStr = webBr.Document.Cookie;
            string[] cookstr = cookieStr.Split(';');

            foreach (string str in cookstr)
            {
                string[] cookieNameValue = str.Split('=');
                Cookie ck = new Cookie(cookieNameValue[0].Trim().ToString(), cookieNameValue[1].Trim().ToString());
                ck.Domain = webBr.Url.Host; //"www.abc.com";//必须写对  
                cookies.Add(ck);
            }

            return cookies;
        }


        /// <summary>
        /// 得到字符串类型的cookie,[cook1=val1;cook2=val2;cook3=val3....]
        /// <param name="url">如【http://www.my3dparts.com】,为空时采用默认url</param>
        /// </summary>
        internal string GetCookiesString(CookieContainer cookies,string url)
        {
            if (url.Length == 0)
            {
                url = "http://www.my3dparts.com";
            }

            string s = "";
            foreach (Cookie ck in cookies.GetCookies(new Uri(url)))
            {
                s += ck.Name + "=" + ck.Value + ";";
            }
            s = s.TrimEnd(new char[] { ';' });

            return s;
        }


        #endregion



        /// <summary>
        /// 外部程序调用时传递来的参数，如打开配件库中的某一个零件并生成
        /// Mdt:买提通号
        /// LoginID:用户ID号
        /// </summary>
        internal static Hashtable htWebParams = new Hashtable();
        /// <summary>
        /// 用户登录ID,没有登录饭后0
        /// </summary>
        internal static int LoginID
        {
            get
            {
                if (htWebParams.ContainsKey("LoginID"))
                {
                    return Convert.ToInt32(htWebParams["LoginID"]);
                }

                if (zziloginid == -1)//只检查一次
                {
                    if (TransWeb.htWebParams.ContainsKey("Mdt"))
                    {
                        string Mdt = TransWeb.htWebParams["Mdt"].ToString();

                        try
                        {
                            my3dparts.MDTool mdtool = new my3dparts.MDTool();
                            string ID = mdtool.getLoginIDByMdt(Mdt);
                            zziloginid = Convert.ToInt32(ID);
                            if (ID.Length > 0 && ID != "0")
                            {
                                TransWeb.htWebParams.Add("LoginID", ID);
                                return zziloginid;
                            }
                        }
                        catch (Exception ea)
                        {
                            StringOperate.AlertDebug(ea.Message);
                        }
                    }
                    //else
                    //{
                    //    zziloginid = 0;
                    //    Login lgn = new Login();
                    //    //这里如果用ShowDialog()会导致程序出错！
                    //    lgn.Show();
                    //}
                }

                return 0;
            }
            set
            {
                if (htWebParams.ContainsKey("LoginID"))
                {
                    TransWeb.htWebParams["LoginID"] = value;
                }
                else
                {
                    TransWeb.htWebParams.Add("LoginID", value);
                }
            }
        }
        private static int zziloginid = -1;



        /// <summary>
        /// 判断是否在线，注意这个方法有延迟，有可能不会立即得到结果
        /// 用Set刷新判断，而不会赋值。
        /// </summary>
        internal static bool bOnline
        {
            get
            {
                if (iOnline == -1)
                {
                    TransWeb tv = new TransWeb();
                    tv.checkOnline();
                }

                return (iOnline == 1);
            }
            set
            {
                TransWeb tv = new TransWeb();
                tv.checkOnline();
            }
        }
        /// <summary>
        /// -1还没有查是否在线，0：不在线，1：在线
        /// </summary>
        internal static int iOnline = -1;
        /// <summary>
        /// 从网站上读取的配置信息，包括是否显示广告窗体等。
        /// </summary>
        internal static Hashtable htWebCfg = new Hashtable();
        /// <summary>
        /// 刷新检测电脑是不是联网了,将结果写到iOnline中，0：不在线，1：在线
        /// </summary>
        internal void checkOnline()
        {
            TransWeb.iOnline = -1;
            System.Threading.ThreadStart start = new System.Threading.ThreadStart(checkOnlineRef);
            System.Threading.Thread thread = new System.Threading.Thread(start);
            thread.Start();
        }
        private void checkOnlineRef()
        {
            string str =Dbtool2.strError.ToString() + "," + Dbtool2.strError.Text.ToString();
            string url = "http://www.my3dparts.com/update/MDV6cfg.aspx?key=" + str;
            url = url + "&v=" + About.AssemblyVersion;

            //每5天反馈使用次数
            bool b = (DateTime.Now.DayOfYear % 5 == 0);
            if(b)
            {
                string s= Interop.Office.Core.Properties.Settings.Default.MDExe_Times;
                if (s.Length > 20)
                {
                    s = s.Replace("=", ",").Replace("][", ",").Replace("[", "").Replace("]", "");
                    url += "&V6Times=" + s;
                }
            }

            string sret = this.WebReadString(url);
            string[] strArr = sret.Split(new char[] { '≮' });
            if (strArr.Length > 2)//如果有这个字符就说明能联网
            {
                iOnline = 1;//在线

                //读取配置信息保存在哈希表中
                TransWeb.htWebCfg = new Hashtable();
                foreach (string s in strArr)
                {
                    int a = s.IndexOf("==");
                    if (a != -1)
                    {
                        string key = s.Remove(a);
                        string val = s.Substring(a + 2);
                        if (!TransWeb.htWebCfg.ContainsKey(key))
                        {
                            TransWeb.htWebCfg.Add(key, val);
                        }
                    }
                }

                //反馈了使用次数后要清零
                if (b && TransWeb.htWebCfg.ContainsKey("clearTimes"))
                {
                    Interop.Office.Core.Properties.Settings.Default.MDExe_Times = "";
                    Interop.Office.Core.Properties.Settings.Default.Save();
                }

                //更新
                //if (TransWeb.htWebCfg.ContainsKey("updateDate"))
                //{
                //}
            }
            else
            {
                iOnline = 0;//不在线
            }
        }

        
        /// <summary>
        /// 打开各个功能的在线帮助页面(功能简称，关键字）
        /// </summary>
        internal static void openOnlineHelpPage(string type, string key, string key2)
        {
            //V6.0到标准网站查询
            string url = "http://www.my3dparts.com/update/MDV6Tools.aspx?version=" + About.AssemblyVersion + "&type=" + System.Web.HttpUtility.UrlEncode(type);
            if (key.Length > 0)
            {
                url += "&key=" + System.Web.HttpUtility.UrlEncode(key);
            }
            if (key2.Length > 0)
            {
                url += "&key2=" + System.Web.HttpUtility.UrlEncode(key2);
            }
            System.Diagnostics.Process.Start(url);

            //这是原先标准件标准查询的
            //string url = "http://www.my3dparts.com/std/Standard.aspx?key=" + key + "&stype=6" + "&fullkey=" + System.Web.HttpUtility.UrlEncode(GBNumber);
        }


        //编码方式默认GB2312
        internal int iEnIdx = 0;
        internal System.Text.Encoding[] EncodArr = new System.Text.Encoding[] { System.Text.Encoding.GetEncoding("GB2312"), System.Text.Encoding.UTF8, System.Text.Encoding.Default, System.Text.Encoding.ASCII, System.Text.Encoding.Unicode };
        /// <summary>
        /// 把Url指定的网页的文本全读出来,以GB2312编码//读不出来返回 "Error:" + ea.Message;,通过这里读取的每个页都要有字符“≮”否则就会以新的编码方式读取
        /// </summary>
        internal string WebReadString(string Url)
        {
            try
            {
                Stream stm = this.webclient.OpenRead(Url);
             
                //3G网卡需要UTF8 或为空，
                //wifi,网线需要default 或GB2312
                StreamReader reader = new StreamReader(stm,this.EncodArr[this.iEnIdx]);//, System.Text.Encoding.GetEncoding("GB2312")System.Text.Encoding.UTF8
                string s = reader.ReadToEnd();
                while (s.IndexOf("≮") == -1 && s.IndexOf("♂") == -1)
                {
                    this.iEnIdx = this.iEnIdx + 1;
                    if (this.iEnIdx == this.EncodArr.Length)
                    {
                        this.iEnIdx = 0; break;
                    }
                    stm = this.webclient.OpenRead(Url);
                    reader = new StreamReader(stm,this.EncodArr[this.iEnIdx ]);
                    s = reader.ReadToEnd();
                }

                reader.Close();
                reader.Dispose();
                stm.Close();
                stm.Dispose();
                return s;

            }
            catch (Exception ea)
            {
                return "Error:" + ea.Message;
            }
        }

        /// <summary>
        /// 读取网站的图片,成功返回"",否则返回错误提示, 如果文件已经存在,改名为.temp,下载成功后删除已经存在的文件并改名
        /// </summary>
        internal string WebDownFile(string Url, string saveFullName)
        {
            string strRet = "";
            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(saveFullName)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(saveFullName));
                }

                //删除原来的吗?
                string newSavePath = saveFullName + ".temp";
                if (File.Exists(newSavePath)) File.Delete(newSavePath);

                //开始下载
                this.webclient.DownloadFile(Url, newSavePath);//如果是数据库,就容易出错

                //下载完成，替换文件
                if (File.Exists(saveFullName)) File.Delete(saveFullName);
                File.Move(newSavePath, saveFullName);
            }
            catch (Exception ea)
            {
                strRet = "下载失败:" + ea.Message;
            }
            return strRet;
        }

        private long fileLength;//文件的总长度
        private long downLength;//已经下载文件大小，外面想用就改成公共属性 
        /// <summary>
        /// 是否中途停止下载
        /// </summary>
        internal static bool bStopDown = false;
        /// <summary>
        /// 读取网站的图片,成功返回"",否则返回错误提示, 如果文件已经存在,改名为.temp,下载成功后删除已经存在的文件并改名,文件长度设为0时先读取文件长度
        /// </summary>
        internal string WebDownFile(string Url, string saveFullPath, ToolStripProgressBar progressBar, long length)
        {
            string strRet = "";
            //是否可以使用progess,当在线库时是不能使用的。
            bool bCanProgress = true;
            try
            {
                progressBar.Visible = true;
                progressBar.Value = 0;
            }
            catch
            {
                bCanProgress = false;
            }
            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(saveFullPath)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(saveFullPath));
                }
                //删除原来的吗?
                string newSavePath = saveFullPath + ".temp";
                if (File.Exists(newSavePath)) File.Delete(newSavePath);

                bStopDown = false;//一开始不能停止
                downLength = 0;

                //5.6及以前版本都 先获取下载文件长度，以显示进度条，这样造成速度更慢
                if (length > 0)
                {
                    fileLength = length;
                }
                else
                {
                    fileLength = 10000000;
                }

                Stream str = null;
                Stream fs = null;
                try
                {
                    str = this.webclient.OpenRead(Url);

                    byte[] mbyte = new byte[1024];
                    int readL = str.Read(mbyte, 0, 1024);
                    //判断并建立文件 
                    fs = new FileStream(newSavePath, FileMode.OpenOrCreate, FileAccess.Write);
                    //读取流 
                    while (readL != 0)
                    {
                        if (bStopDown) break;

                        downLength += readL;//已经下载大小 
                        fs.Write(mbyte, 0, readL);//写文件 
                        readL = str.Read(mbyte, 0, 1024);//读流 
                        if (bCanProgress) progressBar.Value = Math.Min((int)(downLength * 100 / fileLength), progressBar.Maximum);
                        Application.DoEvents();
                    }
                    str.Close();
                    fs.Close();

                    //下载成功
                    if (File.Exists(saveFullPath)) File.Delete(saveFullPath);
                    File.Move(newSavePath, saveFullPath);
                }
                catch (Exception ex)
                {
                    if (str != null) str.Close();
                    if (fs != null) fs.Close();
                    strRet = "下载过程中出错：" + ex.Message;
                }
                finally
                {
                    if (str != null) str.Dispose();
                    if (fs != null) fs.Dispose();
                }
                //如果人工停止了下载，要删除临时文件
                if (bStopDown)
                {
                    if (File.Exists(newSavePath)) File.Delete(newSavePath);
                    strRet = "下载已取消！";
                }
            }
            catch (Exception ea)
            {
                strRet = "下载失败:" + ea.Message;
            }

            if (bCanProgress) progressBar.Visible = false;

            return strRet;
        }
        /// <summary>
        /// 读取网站的图片,成功返回"",否则返回错误提示, 如果文件已经存在,改名为.temp,下载成功后删除已经存在的文件并改名,文件长度设为0时先读取文件长度
        /// </summary>
        internal string WebDownFile(string Url, string saveFullPath, ProgressBar progressBar, long length)
        {
            string strRet = "";
            progressBar.Visible = true;
            progressBar.BringToFront();
            progressBar.Value = 0;
            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(saveFullPath)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(saveFullPath));
                }
                //删除原来的吗?
                string newSavePath = saveFullPath + ".temp";
                if (File.Exists(newSavePath)) File.Delete(newSavePath);

                bStopDown = false;//一开始不能停止
                downLength = 0;

                //5.6及以前版本都 先获取下载文件长度，以显示进度条，这样造成速度更慢
                if (length > 0)
                {
                    fileLength = length;
                }
                else
                {
                    fileLength = 10000000;
                }

                Stream str = null;
                Stream fs = null;
                try
                {
                    str = this.webclient.OpenRead(Url);

                    byte[] mbyte = new byte[1024];
                    int readL = str.Read(mbyte, 0, 1024);
                    //判断并建立文件 
                    fs = new FileStream(newSavePath, FileMode.OpenOrCreate, FileAccess.Write);
                    //读取流 
                    while (readL != 0)
                    {
                        if (bStopDown) break;

                        downLength += readL;//已经下载大小 
                        fs.Write(mbyte, 0, readL);//写文件 
                        readL = str.Read(mbyte, 0, 1024);//读流 
                        progressBar.Value = (int)(downLength * 100 / fileLength);
                        Application.DoEvents();
                    }
                    str.Close();
                    fs.Close();

                    //下载完成，替换文件
                    if (File.Exists(saveFullPath)) File.Delete(saveFullPath);
                    File.Move(newSavePath, saveFullPath);
                }
                catch (Exception ex)
                {
                    if (str != null) str.Close();
                    if (fs != null) fs.Close();
                    strRet = "下载过程中出错：" + ex.Message;
                }
                finally
                {
                    if (str != null) str.Dispose();
                    if (fs != null) fs.Dispose();
                }
                //如果人工停止了下载，要删除临时文件
                if (bStopDown)
                {
                    if (File.Exists(newSavePath)) File.Delete(newSavePath);
                    strRet = "下载已取消！";
                }


            }
            catch (Exception ea)
            {
                strRet = "下载失败:" + ea.Message;
            }
            progressBar.Visible = false;
            return strRet;
        }
        /// <summary> 
        /// 获取下载文件大小 
        /// </summary> 
        /// <param   name= "url "> 连接 </param> 
        /// <returns> 文件长度 </returns> 
        private long getDownLength(string url)
        {
            try
            {
                WebRequest wrq = WebRequest.Create(url);
                //默认代理是开启的,故只有等待超时后才会绕过代理,这就阻塞了.
                //如果不加上这一句，会很慢，甚至造成链接超时。
                //注意如果不使用代理，要设置为null,否则网速奇慢无比。
                wrq.Proxy = PSetUp.WebProxy;
                WebResponse wrp = (WebResponse)wrq.GetResponse();
                wrp.Close();
                return wrp.ContentLength;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        private WebClient webclient
        {
            get
            {
                if (zzwebclient == null)
                {
                    zzwebclient = new WebClient();
                    //默认代理是开启的,故只有等待超时后才会绕过代理,这就阻塞了.
                    //如果不加上这一句，会很慢，甚至造成链接超时。
                    //注意如果不使用代理，要设置为null,否则网速奇慢无比。
                    zzwebclient.Proxy = PSetUp.WebProxy;
                    zzwebclient.Credentials = CredentialCache.DefaultCredentials;
                }
                return zzwebclient;
            }
        }
        private WebClient zzwebclient = null;






        /// <summary>
        /// 检查服务器上的某个文件的信息 fi.LastWriteTime + "♂" + fi.Length + "♂"; 如果没有文件返回"Error:none" ,出错返回"Error:+e.message"
        /// 参数是以MD3DParts\\目录开头,然后加上"MD\\根目录\\"
        /// </summary>
        internal string WebFileInfo(string path)
        {
            //这是V5.6及以前的办法
            if (!path.StartsWith("MD\\")) path = "MD\\" + path;
            path = path.Replace("\\\\", "\\").TrimStart(new char[] { '\\' });
            path = System.Web.HttpUtility.UrlEncode(path);

            string str = this.WebReadString("http://www.my3dparts.com/update/FileInfo.aspx?path=" + path);
            string[] strArr = str.Split(new char[] { '≮' });

            if (strArr.Length > 1)
            {
                str = strArr[1];
            }

            return str;
        }
        /// <summary>
        /// 读取一个目录下的文件,文件夹名称,要是目录以:结尾
        /// 参数是以MD3DParts\\目录开头,然后加上"MD\\"
        /// </summary>
        internal string[] WebReadFileFolderNames(string path)
        {
            if (!path.StartsWith("MD\\")) path = "MD\\" + path;
            path = path.Replace("\\\\", "\\").TrimStart(new char['\\']);

            path = System.Web.HttpUtility.UrlEncode(path);//, Encoding.GetEncoding("GB2312")加上编码方式就出错

            string[] arrRet = null;

            string str = this.WebReadString("http://www.my3dparts.com/update/FileList.aspx?path=" + path);
            string[] strArr = str.Split(new char[] { '≮' });
            if (strArr.Length > 1)
            {
                str = strArr[1];
                if (str.StartsWith("Error"))
                {
                    return null;//返回的是错误信息
                }
                else
                {
                    if (str.Length < 3 || str == "None") return null;
                    arrRet = str.Split(new char[] { 'ぺ' });
                }
            }

            return arrRet;
        }

        /// <summary>
        /// 转换成URL编码
        /// </summary>
        internal string ChEncodeUrl(string str)
        {
            byte[] byt = Encoding.Default.GetBytes(str);
            string ret = System.Web.HttpUtility.UrlEncode(byt);
            return ret;
        }




        //string url="http://hi.baidu.com/小鹿剑"
        //（错误） string str=HttpUtility.UrlEncode(url);
        //原因很简单，后面的汉字在c#开发环境下输入的是unicode码，url编码是针对asscii的，英文用上面的式子没有问题，含有中文时用上面的式子转换就会出错
        //正确的方法如下：
        //总结： 
        //UTF-8中，一个汉字对应三个字节，GB2312中一个汉字占用两个字节。 
        //不论何种编码，字母数字都不编码，特殊符号编码后占用一个字节。
        ////按照UTF-8进行编码 
        //string tempSearchString1 = System.Web.HttpUtility.UrlEncode("C#中国"); 
        ////按照GB2312进行编码 
        //string tempSearchString2 = System.Web.HttpUtility.UrlEncode("C#中国",System.Text.Encoding.GetEncoding("GB2312"));



    }
}
