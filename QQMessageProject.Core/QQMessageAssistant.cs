using AccessibleProject;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QQMessageProject.Core
{
    /// <summary>
    /// QQ消息助手
    /// 支持异常自动登录（暂需手动设置指定账号自动登录，且消息发送快捷键为Enter），多个本地QQ登录不影响使用
    /// 支持收取和发送消息
    /// 支持单击功能
    /// 支持自定义命令回调，执行后台程序后返回
    /// 用途：
    ///     1、在无人环境（如服务器）中，用小号给大号发送消息，监测软件实时运行情况、执行软件功能等
    ///     2、转发指定好友消息
    ///     3、远程执行本机命令（需自定义程序），如下班了提前打开直播等，避免孤单
    ///     4、自由发挥
    /// </summary>
    public class QQMessageAssistant
    {
        const string ExeName = "QQ";
        /// <summary>
        /// 通过URL   以{1}QQ 打开{0}好友的QQ会话窗口
        /// </summary>
        const string OpenWindowCommand = "tencent://Message/?menu=no&exe=QQ&uin={0}&fuin={1}&websiteName=halfasonYc.com&inf";

        /// <summary>
        /// 会话窗体类名，可通过AccLooker获取
        /// </summary>
        const string WindowClassName = "TXGuiFoundation";

        /// <summary>
        /// QQ注册表安装路径键名  HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node  32位节点下
        /// </summary>
        const string ExeInstallRegPath = @"SOFTWARE\TENCENT\QQ2009\";
        const string ExeInstallRegPathName = "Install";

        const string ExeSuffix = @"Bin\QQ.exe";
        const int WaitTimes = 10;
        const int MainWindowHeightMin = 600;
        const int MainWindowWidthMax = 500;

        const string MsgAccessibleName = "消息";
        const string InputAccessibleName = "输入";

        public const string DefaultTestMsg = "Hello World!";

        public static string InstallPath { get; private set; }

        /// <summary>
        /// 当前操作QQ窗口句柄
        /// </summary>
        [DisplayName("窗口句柄"), System.ComponentModel.Category("辅助信息"), Description("当前操作QQ窗口句柄")]
        public IntPtr Handle { get; set; }

        /// <summary>
        /// 操作QQ
        /// </summary>
        [DisplayName("操作QQ"), System.ComponentModel.Category("辅助信息"), Description("当前操作QQ，本人QQ")]
        public string QQ { get; set; }

        /// <summary>
        /// 自己昵称，用于区分消息发起方
        /// </summary>
        [DisplayName("操作昵称"), System.ComponentModel.Category("辅助信息"), Description("自己昵称，用于区分消息发起方")]
        public string Name { get; set; }

        /// <summary>
        /// 好友QQ
        /// </summary>
        [DisplayName("好友QQ"), System.ComponentModel.Category("辅助信息"), Description("被操作QQ，好友QQ")]
        public string SpecifyQQ { get; set; }

        /// <summary>
        /// 指定窗口名称，若有备注则为备注名，无备注则为QQ昵称
        /// </summary>
        [DisplayName("窗口名称"), System.ComponentModel.Category("辅助信息"), Description("指定窗口名称，若有备注则为备注名，无备注则为QQ昵称")]
        public string SpecifyName { get; set; }

        Func<string, string> messageCallBackHandler;
        public event Func<string, string> MessageCallBack
        {
            add
            {
                this.messageCallBackHandler += value;
            }
            remove
            {
                this.messageCallBackHandler -= value;
            }
        }

        /// <summary>
        /// 错误信息
        /// </summary>
        [DisplayName("错误信息"), System.ComponentModel.Category("辅助信息"), Description("错误信息，当操作失败时会置此属性")]
        public string LastError { get; set; }

        AccessibleHelper accessibleHelper;

        private QQMessageAssistant()
        {

        }

        static QQMessageAssistant()
        {
            try
            {
                RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).
                    OpenSubKey(ExeInstallRegPath, false);
                if (registry != null)
                {
                    InstallPath = registry.GetValue(ExeInstallRegPathName) as string;
                    if (!string.IsNullOrEmpty(InstallPath))
                    {
                        InstallPath = Path.Combine(InstallPath, ExeSuffix);
                    }
                }
            }
            catch (Exception)
            {
                InstallPath = string.Empty;
            }
        }

        public readonly static QQMessageAssistant _Empty = new QQMessageAssistant();
        /// <summary>
        /// 仅用于快捷绑定，没有实际意义
        /// </summary>
        public static QQMessageAssistant Empty
        {
            get
            {
                return _Empty;
            }
        }

        public static QQMessageAssistant FromHwnd(IntPtr hwnd, out string errorText)
        {
            QQMessageAssistant qqMessageAssistant = new QQMessageAssistant() { Handle = hwnd };
            if (qqMessageAssistant.ObtainVerifyHwnd())
            {
                errorText = string.Empty;
                return qqMessageAssistant;
            }
            errorText = qqMessageAssistant.LastError;
            return null;
        }

        public static QQMessageAssistant FromInformation(string qq, string specifyQQ, string windowName, string selfName, out string errorText)
        {
            if (string.IsNullOrWhiteSpace(qq) || string.IsNullOrWhiteSpace(specifyQQ) || string.IsNullOrWhiteSpace(windowName) || string.IsNullOrWhiteSpace(selfName))
            {
                errorText = "辅助信息不完整！";
            }
            else
            {
                QQMessageAssistant qqMessageAssistant = new QQMessageAssistant() { QQ = qq, SpecifyQQ = specifyQQ, SpecifyName = windowName, Name = selfName };
                if (qqMessageAssistant.ObtainVerifyHwnd())
                {
                    errorText = string.Empty;
                    return qqMessageAssistant;
                }
                errorText = qqMessageAssistant.LastError;
            }
            return null;
        }

        public void SendMessage(string text)
        {
            this.accessibleHelper?.PostText(text);
        }

        public void Click(Point relativeWindowLocation)
        {
            this.accessibleHelper?.Click(relativeWindowLocation);
        }

        /// <summary>
        /// 取对话窗口句柄并验证其是否有效
        /// </summary>
        /// <returns></returns>
        private bool ObtainVerifyHwnd()
        {
            if (Handle.ToInt32() > 0 && NativeMethods.IsWindow(Handle))
            {
                return true;
            }

            bool forceLogin = false;
            if (!IsProcessExist() || forceLogin)
            {
                if (string.IsNullOrEmpty(InstallPath))
                {
                    LastError = "获取QQ安装路径失败，请尝试以管理员身份运行！";
                    return false;
                }
                Process.Start(InstallPath);
                Task.Delay(2000).Wait();
                for (int i = 0; i < WaitTimes; i++)
                {
                    if (IsProcessExist())
                    {
                        break;
                    }
                    Task.Delay(2000).Wait();
                    if (i == WaitTimes)
                    {
                        LastError = "QQ进程启动失败";
                        return false;
                    }
                }

                //TODO: 非自动登录时手动选择账号（或置入账号）后登录
                //等待登录成功
                for (int i = 0; i < WaitTimes; i++)
                {
                    if (IsMainFormExist())
                    {
                        break;
                    }
                    Task.Delay(2000).Wait();
                    if (i == WaitTimes)
                    {
                        LastError = "等待QQ登录成功状态失败";
                        return false;
                    }
                }
            }
            if (!IsWindowExist())
            {
                string command = string.Format(OpenWindowCommand, SpecifyQQ, QQ);
                using (Process process = Process.Start(command))
                {
                    for (int i = 0; i < WaitTimes; i++)
                    {
                        if (IsWindowExist())
                        {
                            break;
                        }
                        Task.Delay(2000).Wait();
                        if (i == WaitTimes)
                        {
                            LastError = "等待会话窗体打开状态失败";
                            return false;
                        }
                    }
                }
            }

            if (Handle.ToInt32() > 0 && (this.accessibleHelper = AccessibleHelper.FromHwnd(Handle)) != null)
            {
                System.Diagnostics.Debug.WriteLine("QQ窗口已打开");
                LastError = string.Empty;
                return true;
            }
            LastError = "尚未启动对应QQ进程";
            return false;
        }

        /// <summary>
        /// 进程是否存在
        /// </summary>
        /// <returns></returns>
        private bool IsProcessExist()
        {
            Process[] processes = Process.GetProcessesByName(ExeName);
            return processes.Length > 0;
        }

        /// <summary>
        /// 已经登录成功，主界面已打开
        /// </summary>
        /// <returns></returns>
        private bool IsMainFormExist()
        {
            Handle = NativeMethods.FindWindowEx(IntPtr.Zero, IntPtr.Zero, WindowClassName, ExeName);
            bool handleValid = Handle.ToInt32() > 0;
            bool windowValid = NativeMethods.GetWindowRect(Handle, out NativeMethods.TagRect rectangle);
            bool rectValid = (rectangle.Bottom - rectangle.Top) > MainWindowHeightMin && (rectangle.Right - rectangle.Left) < MainWindowWidthMax;
            return handleValid && windowValid && rectValid;
        }

        /// <summary>
        /// 对话窗体是否存在
        /// </summary>
        /// <returns></returns>
        private bool IsWindowExist()
        {
            Handle = NativeMethods.FindWindowEx(IntPtr.Zero, IntPtr.Zero, WindowClassName, SpecifyName);
            return Handle.ToInt32() > 0;
        }
    }

    internal static partial class NativeMethods
    {
        [System.Runtime.InteropServices.StructLayoutAttribute(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct TagRect
        {

            /// LONG->int
            public int Left { get; set; }

            /// LONG->int
            public int Top { get; set; }

            /// LONG->int
            public int Right { get; set; }

            /// LONG->int
            public int Bottom { get; set; }
        }

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "IsWindow")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool IsWindow([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd);

        /// Return Type: HWND->HWND__*
        ///hWndParent: HWND->HWND__*
        ///hWndChildAfter: HWND->HWND__*
        ///lpszClass: LPCWSTR->WCHAR*
        ///lpszWindow: LPCWSTR->WCHAR*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "FindWindowEx")]
        public static extern System.IntPtr FindWindowEx([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWndParent,
            [System.Runtime.InteropServices.InAttribute()] System.IntPtr hWndChildAfter,
            [System.Runtime.InteropServices.InAttribute()] string lpszClass,
            [System.Runtime.InteropServices.InAttribute()] string lpszWindow);


        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        ///lpRect: LPRECT->tagRECT*      
        /// LONG->int
        ///public int left;
        /// LONG->int
        ///public int top;
        /// LONG->int
        ///public int right;
        /// LONG->int
        ///public int bottom;
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "GetWindowRect")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool GetWindowRect([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd,
            [System.Runtime.InteropServices.OutAttribute()] out TagRect lpRect);

        /// Return Type: BOOL->int
        ///param0: HWND->HWND__*
        ///param1: LPARAM->LONG_PTR->int
        [System.Runtime.InteropServices.UnmanagedFunctionPointerAttribute(System.Runtime.InteropServices.CallingConvention.StdCall)]
        public delegate int WndEnumProc(System.IntPtr param0, System.IntPtr param1);

        /// Return Type: BOOL->int
        ///lpEnumFunc: WNDENUMPROC
        ///lParam: LPARAM->LONG_PTR->int
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "EnumWindows")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool EnumWindows(WndEnumProc lpEnumFunc, IntPtr lParam);

        /// Return Type: DWORD->unsigned int
        ///hWnd: HWND->HWND__*
        ///lpdwProcessId: LPDWORD->DWORD*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "GetWindowThreadProcessId")]
        public static extern uint GetWindowThreadProcessId([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd, System.IntPtr lpdwProcessId);
    }
}
