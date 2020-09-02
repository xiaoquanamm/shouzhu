using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Windows.Forms;

namespace Interop.Office.Core
{
    class Sys
    {

        /// <summary>
        /// 得到系统是64为还是32位,需添加引用System.Management
        /// </summary>
        internal static bool Is64BitOPSystem
        {
            get
            {
                //return IntPtr.Size == 8;//这是原先的方法，现在不行，不能正确判断

                //V6.1修改
                int a = 0;
                try
                {
                    a = Interop.Office.Core.Properties.Settings.Default.exeOSBit;//记录操作系统版本，初始值0， 或32，或64；
                }
                catch
                {
                    return true;
                }
                if (a == 64) return true;
                if (a == 32) return false;
                if (a == 0)
                {
                    string s = "";
                    try
                    {
                        ConnectionOptions mConnOption = new ConnectionOptions();
                        ManagementScope mMs = new ManagementScope("\\\\localhost", mConnOption);
                        ObjectQuery mQuery = new ObjectQuery("select AddressWidth from Win32_Processor");
                        ManagementObjectSearcher mSearcher = new ManagementObjectSearcher(mMs, mQuery);
                        ManagementObjectCollection mObjectCollection = mSearcher.Get();
                        foreach (ManagementObject mObject in mObjectCollection)
                        {
                            s = mObject["AddressWidth"].ToString();
                        }
                        if (s == "64" || s == "32")
                        {
                            Interop.Office.Core.Properties.Settings.Default.exeOSBit = Convert.ToInt16(s);
                            Interop.Office.Core.Properties.Settings.Default.Save();
                            return (s == "64");
                        }
                    }
                    catch
                    { }

                    if (StringOperate.Alert("您的操作系统是64位的吗？", "获取配置信息出错，需要人工确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        Interop.Office.Core.Properties.Settings.Default.exeOSBit = 64;
                        Interop.Office.Core.Properties.Settings.Default.Save();
                        return true;
                    }
                    else
                    {
                        Interop.Office.Core.Properties.Settings.Default.exeOSBit = 32;
                        Interop.Office.Core.Properties.Settings.Default.Save();
                        return false;
                    }
                }
                return true;
            }
        }




















    }
}
