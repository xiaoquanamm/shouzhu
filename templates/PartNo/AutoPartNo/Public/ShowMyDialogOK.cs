using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


internal partial class ShowMyDialogOK : Form
{
    internal ShowMyDialogOK(string title, string stralert)
    {
        InitializeComponent();
        this.strTitle = title;
        this.strAlert = stralert;
    }

    private string strTitle = "";
    private string strAlert = "";

    private void ShowMyDialogOK_Load(object sender, EventArgs e)
    {
        this.Text = strTitle;
        this.textBox1.Text = strAlert;
        this.timer1.Start();
    }

    //启动SolidWorks多少秒了
    private int iSeconds = 0;
    private void timer1_Tick(object sender, EventArgs e)
    {
        System.Diagnostics.Process[] processArr = System.Diagnostics.Process.GetProcessesByName("SLDWORKS");
        if (processArr.Length > 0)
        {
            iSeconds += Convert.ToInt16(timer1.Interval / 1000);
            if (iSeconds > 50)
            {
                this.button1.PerformClick();
            }
        }
    }

    private void ShowMyDialogOK_FormClosing(object sender, FormClosingEventArgs e)
    {
        this.timer1.Stop();
    }

    private void button2_Click(object sender, EventArgs e)
    {
        this.timer1.Stop();
    }
}



//private void HasOk()
//{
//}

//[System.Runtime.InteropServices.DllImport("kernel32", EntryPoint = "OpenProcess", ExactSpelling = true, CharSet = System.Runtime.InteropServices.CharSet.Ansi, SetLastError = true)]
//public static extern long OpenProcess(long dwDesiredAccess, long bInheritHandle, long dwProcessId);

//[System.Runtime.InteropServices.DllImport("kernel32", EntryPoint = "GetExitCodeProcess", ExactSpelling = true, CharSet = System.Runtime.InteropServices.CharSet.Ansi, SetLastError = true)]
//public static extern long GetExitCodeProcess(long hProcess, long lpExitCode);

//[System.Runtime.InteropServices.DllImport("kernel32", EntryPoint = "CloseHandle", ExactSpelling = true, CharSet = System.Runtime.InteropServices.CharSet.Ansi, SetLastError = true)]
//public static extern long CloseHandle(long hObject);

//public const int PROCESS_QUERY_INFORMATION = 0X400;
//public const int STATUS_PENDING = 0X103L;

////在需要的地方调用RunShell过程，如：res=RunShell("c:\windows\notepad.exe")
////返回值为真则程序结束 
//public bool RunShell(string cmdline)
//{
//    long hProcess = 0;
//    long ProcessId = 0;
//    long exitCode = 0;
//    ProcessId = Microsoft.VisualBasic.Interaction.Shell(cmdline, 1, false, -1);
//    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, false, ProcessId);
//    do
//    {
//        GetExitCodeProcess(hProcess, exitCode);
//        DoEvents;
//    } while (exitCode == STATUS_PENDING);
//    CloseHandle(hProcess);
//    return true;
//}

