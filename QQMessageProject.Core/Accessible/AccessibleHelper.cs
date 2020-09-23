using Accessibility;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

/* Description: AccessibleHelper，通过P/Invoke与IAccessible结合，完成UI解析动作
 *          By: HalfasonYc
 *        Time: 2020-09-22
 *       Email: Endless_yangc@foxmail.com
 *     Version: 1.0
 *     
 *     Modify: 
 *     Version: 1.1
 *     Content: 增加投递消息和鼠标点击命令
 */

namespace AccessibleProject
{
    public class AccessibleHelper
    {
        IAccessible _self = null;
        IntPtr _hwnd = IntPtr.Zero;
        int _id = 0;

        private AccessibleHelper()
        {

        }

        /// <summary>
        /// 失败返回null
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static AccessibleHelper FromHwnd(IntPtr hwnd)
        {
            Guid guid = NativeMethods.IAccessibleGuid;
            IAccessible accessible = null;
            if (NativeMethods.AccessibleObjectFromWindow(hwnd, NativeMethods.ObjId.Window, ref guid, ref accessible) == 0)
            {
                AccessibleHelper accessibleHelper = new AccessibleHelper();
                accessibleHelper._self = accessible;
                accessibleHelper._hwnd = hwnd;
                return accessibleHelper;
            }
            return null;
        }

        /// <summary>
        /// 失败返回null
        /// </summary>
        /// <param name="point"></param>
        /// <returns></returns>
        public static AccessibleHelper FromPoint(Point point)
        {
            if (NativeMethods.AccessibleObjectFromPoint(point, out IAccessible accessible, out object childId) == IntPtr.Zero)
            {
                AccessibleHelper accessibleHelper = FromObject(accessible);
                accessibleHelper._id = Convert.ToInt32(childId);
                return accessibleHelper;
            }
            return null;
        }

        public static AccessibleHelper FromObject(IAccessible accessible)
        {
            AccessibleHelper accessibleHelper = new AccessibleHelper();
            accessibleHelper._self = accessible;
            NativeMethods.WindowFromAccessibleObject(accessible, ref accessibleHelper._hwnd);
            return accessibleHelper;
        }

        public AccessibleInformation GetInformation()
        {
            AccessibleInformation infor = new AccessibleInformation();
            int x, y, width, height;
            int index = 0;
            //有些接口方法并未实现，有些属性无效，或者其他原因， 总之IAccessible的接口方法不安全，需要自行处理异常
  getValue: try
            {
                switch (index)
                {
                    case 0:
                        infor.AccDescription = _self.accDescription[_id];
                        break;
                    case 1:
                        infor.AccHelp = _self.accHelp[_id];
                        break;
                    case 2:
                        infor.AccDefaultAction = _self.accDefaultAction[_id];
                        break;
                    case 3:
                        infor.AccName = _self.accName[_id];
                        break;
                    case 4:
                        infor.AccRole = _self.accRole[_id];
                        break;
                    case 5:
                        infor.AccState = _self.accState[_id];
                        break;
                    case 6:
                        infor.AccValue = _self.accValue[_id];
                        break;
                    case 7:
                        infor.AccKeyboardShortcut = _self.accKeyboardShortcut[_id];
                        break;
                    case 8:
                        _self.accLocation(out x, out y, out width, out height, _id);
                        infor.AccLocation = new Rectangle(x, y, width, height);
                        break;
                    case 9:
                        infor.AccChildId = _id;
                        break;
                    case 10:
                        if (_id == 0)
                        {
                            infor.AccChildCount = _self.accChildCount;
                        }
                        else
                        {
                            infor.AccChildCount = (_self.accChild[_id] as IAccessible).accChildCount;
                        }
                        break;
                }
            }
            catch
            {
                //catch (System.NotImplementedException)
                //catch (System.ArgumentException)
                //catch (System.Runtime.InteropServices.COMException)
            }
            if (++index != 11)
            {
                goto getValue;
            }
            infor.HandleWindow = _hwnd;
            StringBuilder stringBuilder = new StringBuilder(200);
            if (NativeMethods.GetClassName(_hwnd, stringBuilder, 200) != 0)
            {
                infor.ClassNameWindow = stringBuilder.ToString();
            }

            int buffer = NativeMethods.GetWindowTextLength(_hwnd) + 1;
            stringBuilder = new StringBuilder(buffer);
            if (NativeMethods.GetWindowText(_hwnd, stringBuilder, buffer) != 0)
            {
                infor.TitleWindow = stringBuilder.ToString();
            }
            return infor;
        }

        public IAccessible Accessible
        {
            get
            {
                return _self;
            }
        }

        public IntPtr Hwnd
        {
            get
            {
                return _hwnd;
            }
        }

        public int Id
        {
            get
            {
                return _id;
            }
        }

        /// <summary>
        /// 在桌面实时绘画矩形，并自动清除
        /// </summary>
        /// <param name="accLocation"></param>
        public static void DrawBorderByRectangle(Rectangle accLocation)
        {
            //获取桌面设备场景
            IntPtr deskHwnd = NativeMethods.GetDesktopWindow();
            IntPtr deskDC = NativeMethods.GetWindowDC(deskHwnd);
            var oldR2 = NativeMethods.SetROP2(deskDC, NativeMethods.BinaryRasterOperations.R2_NOTXORPEN);

            IntPtr newPen = NativeMethods.CreatePen(NativeMethods.PenStyle.PS_SOLID, 2, (uint)ColorTranslator.ToWin32(Color.Red));
            IntPtr oldPen = NativeMethods.SelectObject(deskDC, newPen);
            //使用新笔画图
            NativeMethods.Rectangle(deskDC, accLocation.Left, accLocation.Top, accLocation.Right, accLocation.Bottom);
            Task.Delay(100).Wait();
            NativeMethods.Rectangle(deskDC, accLocation.Left, accLocation.Top, accLocation.Right, accLocation.Bottom);
            NativeMethods.SetROP2(deskDC, oldR2);
            //还原设备状态
            NativeMethods.SelectObject(deskDC, oldPen);
            NativeMethods.DeleteObject(newPen);
            NativeMethods.ReleaseDC(deskHwnd, deskDC);
        }

        /// <summary>
        /// 投递文本
        /// </summary>
        public void PostText(string text)
        {
            //必须先激活，否则可能因误操作导致最小化后不响应
            NativeMethods.OpenIcon(_hwnd);
            Task.Delay(100).Wait();
            byte[] bytes = Encoding.Default.GetBytes(text);
            foreach (var b in bytes)
            {
                System.Diagnostics.Debug.WriteLine(NativeMethods.PostMessage(_hwnd, NativeMethods.WinMsg.WM_CHAR, b, 0));
            }
            Task.Delay(100).Wait();
            const int keysEnter = 13;
            NativeMethods.PostMessage(_hwnd, NativeMethods.WinMsg.WM_KEYDOWN, keysEnter, 0); //Keys.Enter = 13
            Task.Delay(1000).Wait();
            NativeMethods.CloseWindow(_hwnd);
        }

        /// <summary>
        /// 鼠标单击
        /// </summary>
        /// <param name="point">相对窗口位置</param>
        public void Click(Point point)
        {
            //必须先激活，否则可能因误操作导致最小化后不响应
            NativeMethods.OpenIcon(_hwnd);
            Task.Delay(100).Wait();
            int pointValue = point.X + point.Y * 65536;
            NativeMethods.PostMessage(_hwnd, NativeMethods.WinMsg.WM_MOUSEFIRST, 2, pointValue);
            NativeMethods.PostMessage(_hwnd, NativeMethods.WinMsg.WM_LBUTTONDOWN, 1, pointValue);
            NativeMethods.PostMessage(_hwnd, NativeMethods.WinMsg.WM_LBUTTONUP, 0, pointValue);
        }
    }

    internal static partial class NativeMethods
    {
        //winable.h constants
        public enum ObjId : uint
        {
            Window = 0x00000000,
            SysMenu = 0xFFFFFFFF,
            TitleBar = 0xFFFFFFFE,
            Menu = 0xFFFFFFFD,
            Client = 0xFFFFFFFC,
            Vscroll = 0xFFFFFFFB,
            Hscroll = 0xFFFFFFFA,
            Sizegrip = 0xFFFFFFF9,
            Caret = 0xFFFFFFF8,
            Cursor = 0xFFFFFFF7,
            Alert = 0xFFFFFFF6,
            Sound = 0xFFFFFFF5,
        }
        public enum BinaryRasterOperations
        {

            R2_BLACK = 1,
            R2_NOTMERGEPEN = 2,
            R2_MASKNOTPEN = 3,
            R2_NOTCOPYPEN = 4,
            R2_MASKPENNOT = 5,
            R2_NOT = 6,
            R2_XORPEN = 7,
            R2_NOTMASKPEN = 8,
            R2_MASKPEN = 9,
            R2_NOTXORPEN = 10,
            R2_NOP = 11,
            R2_MERGENOTPEN = 12,
            R2_COPYPEN = 13,
            R2_MERGEPENNOT = 14,
            R2_MERGEPEN = 15,
            R2_WHITE = 16
        }
        public enum PenStyle : int
        {
            PS_SOLID = 0, //The pen is solid.
            PS_DASH = 1, //The pen is dashed.
            PS_DOT = 2, //The pen is dotted.
            PS_DASHDOT = 3, //The pen has alternating dashes and dots.
            PS_DASHDOTDOT = 4, //The pen has alternating dashes and double dots.
            PS_NULL = 5, //The pen is invisible.
            PS_INSIDEFRAME = 6,// Normally when the edge is drawn, it’s centred on the outer edge meaning that half the width of the pen is drawn
                               // outside the shape’s edge, half is inside the shape’s edge. When PS_INSIDEFRAME is specified the edge is drawn
                               //completely inside the outer edge of the shape.
            PS_USERSTYLE = 7,
            PS_ALTERNATE = 8,
            PS_STYLE_MASK = 0x0000000F,

            PS_ENDCAP_ROUND = 0x00000000,
            PS_ENDCAP_SQUARE = 0x00000100,
            PS_ENDCAP_FLAT = 0x00000200,
            PS_ENDCAP_MASK = 0x00000F00,

            PS_JOIN_ROUND = 0x00000000,
            PS_JOIN_BEVEL = 0x00001000,
            PS_JOIN_MITER = 0x00002000,
            PS_JOIN_MASK = 0x0000F000,

            PS_COSMETIC = 0x00000000,
            PS_GEOMETRIC = 0x00010000,
            PS_TYPE_MASK = 0x000F0000
        }
        public enum WinMsg
        {
            WM_NULL = 0x0000,
            WM_CREATE = 0x0001,
            WM_DESTROY = 0x0002,
            WM_MOVE = 0x0003,
            WM_SIZE = 0x0005,
            WM_ACTIVATE = 0x0006,
            WM_SETFOCUS = 0x0007,
            WM_KILLFOCUS = 0x0008,
            WM_ENABLE = 0x000A,
            WM_SETREDRAW = 0x000B,
            WM_SETTEXT = 0x000C,
            WM_GETTEXT = 0x000D,
            WM_GETTEXTLENGTH = 0x000E,
            WM_PAINT = 0x000F,
            WM_CLOSE = 0x0010,
            WM_QUERYENDSESSION = 0x0011,
            WM_QUERYOPEN = 0x0013,
            WM_ENDSESSION = 0x0016,
            WM_QUIT = 0x0012,
            WM_ERASEBKGND = 0x0014,
            WM_SYSCOLORCHANGE = 0x0015,
            WM_SHOWWINDOW = 0x0018,
            WM_WININICHANGE = 0x001A,
            WM_SETTINGCHANGE = 0x001A,
            WM_DEVMODECHANGE = 0x001B,
            WM_ACTIVATEAPP = 0x001C,
            WM_FONTCHANGE = 0x001D,
            WM_TIMECHANGE = 0x001E,
            WM_CANCELMODE = 0x001F,
            WM_SETCURSOR = 0x0020,
            WM_MOUSEACTIVATE = 0x0021,
            WM_CHILDACTIVATE = 0x0022,
            WM_QUEUESYNC = 0x0023,
            WM_GETMINMAXINFO = 0x0024,
            WM_PAINTICON = 0x0026,
            WM_ICONERASEBKGND = 0x0027,
            WM_NEXTDLGCTL = 0x0028,
            WM_SPOOLERSTATUS = 0x002A,
            WM_DRAWITEM = 0x002B,
            WM_MEASUREITEM = 0x002C,
            WM_DELETEITEM = 0x002D,
            WM_VKEYTOITEM = 0x002E,
            WM_CHARTOITEM = 0x002F,
            WM_SETFONT = 0x0030,
            WM_GETFONT = 0x0031,
            WM_SETHOTKEY = 0x0032,
            WM_GETHOTKEY = 0x0033,
            WM_QUERYDRAGICON = 0x0037,
            WM_COMPAREITEM = 0x0039,
            WM_GETOBJECT = 0x003D,
            WM_COMPACTING = 0x0041,
            WM_COMMNOTIFY = 0x0044,
            WM_WINDOWPOSCHANGING = 0x0046,
            WM_WINDOWPOSCHANGED = 0x0047,
            WM_POWER = 0x0048,
            WM_COPYDATA = 0x004A,
            WM_CANCELJOURNAL = 0x004B,
            WM_NOTIFY = 0x004E,
            WM_INPUTLANGCHANGEREQUEST = 0x0050,
            WM_INPUTLANGCHANGE = 0x0051,
            WM_TCARD = 0x0052,
            WM_HELP = 0x0053,
            WM_USERCHANGED = 0x0054,
            WM_NOTIFYFORMAT = 0x0055,
            WM_CONTEXTMENU = 0x007B,
            WM_STYLECHANGING = 0x007C,
            WM_STYLECHANGED = 0x007D,
            WM_DISPLAYCHANGE = 0x007E,
            WM_GETICON = 0x007F,
            WM_SETICON = 0x0080,
            WM_NCCREATE = 0x0081,
            WM_NCDESTROY = 0x0082,
            WM_NCCALCSIZE = 0x0083,
            WM_NCHITTEST = 0x0084,
            WM_NCPAINT = 0x0085,
            WM_NCACTIVATE = 0x0086,
            WM_GETDLGCODE = 0x0087,
            WM_SYNCPAINT = 0x0088,
            WM_NCMOUSEMOVE = 0x00A0,
            WM_NCLBUTTONDOWN = 0x00A1,
            WM_NCLBUTTONUP = 0x00A2,
            WM_NCLBUTTONDBLCLK = 0x00A3,
            WM_NCRBUTTONDOWN = 0x00A4,
            WM_NCRBUTTONUP = 0x00A5,
            WM_NCRBUTTONDBLCLK = 0x00A6,
            WM_NCMBUTTONDOWN = 0x00A7,
            WM_NCMBUTTONUP = 0x00A8,
            WM_NCMBUTTONDBLCLK = 0x00A9,
            WM_NCXBUTTONDOWN = 0x00AB,
            WM_NCXBUTTONUP = 0x00AC,
            WM_NCXBUTTONDBLCLK = 0x00AD,
            WM_INPUT = 0x00FF,
            WM_KEYFIRST = 0x0100,
            WM_KEYDOWN = 0x0100,
            WM_KEYUP = 0x0101,
            WM_CHAR = 0x0102,
            WM_DEADCHAR = 0x0103,
            WM_SYSKEYDOWN = 0x0104,
            WM_SYSKEYUP = 0x0105,
            WM_SYSCHAR = 0x0106,
            WM_SYSDEADCHAR = 0x0107,
            WM_UNICHAR = 0x0109,
            WM_KEYLAST = 0x0109,
            WM_IME_STARTCOMPOSITION = 0x010D,
            WM_IME_ENDCOMPOSITION = 0x010E,
            WM_IME_COMPOSITION = 0x010F,
            WM_IME_KEYLAST = 0x010F,
            WM_INITDIALOG = 0x0110,
            WM_COMMAND = 0x0111,
            WM_SYSCOMMAND = 0x0112,
            WM_TIMER = 0x0113,
            WM_HSCROLL = 0x0114,
            WM_VSCROLL = 0x0115,
            WM_INITMENU = 0x0116,
            WM_INITMENUPOPUP = 0x0117,
            WM_MENUSELECT = 0x011F,
            WM_MENUCHAR = 0x0120,
            WM_ENTERIDLE = 0x0121,
            WM_MENURBUTTONUP = 0x0122,
            WM_MENUDRAG = 0x0123,
            WM_MENUGETOBJECT = 0x0124,
            WM_UNINITMENUPOPUP = 0x0125,
            WM_MENUCOMMAND = 0x0126,
            WM_CHANGEUISTATE = 0x0127,
            WM_UPDATEUISTATE = 0x0128,
            WM_QUERYUISTATE = 0x0129,
            WM_CTLCOLORMSGBOX = 0x0132,
            WM_CTLCOLOREDIT = 0x0133,
            WM_CTLCOLORLISTBOX = 0x0134,
            WM_CTLCOLORBTN = 0x0135,
            WM_CTLCOLORDLG = 0x0136,
            WM_CTLCOLORSCROLLBAR = 0x0137,
            WM_CTLCOLORSTATIC = 0x0138,
            WM_MOUSEFIRST = 0x0200,
            WM_MOUSEMOVE = 0x0200,
            WM_LBUTTONDOWN = 0x0201,
            WM_LBUTTONUP = 0x0202,
            WM_LBUTTONDBLCLK = 0x0203,
            WM_RBUTTONDOWN = 0x0204,
            WM_RBUTTONUP = 0x0205,
            WM_RBUTTONDBLCLK = 0x0206,
            WM_MBUTTONDOWN = 0x0207,
            WM_MBUTTONUP = 0x0208,
            WM_MBUTTONDBLCLK = 0x0209,
            WM_MOUSEWHEEL = 0x020A,
            WM_XBUTTONDOWN = 0x020B,
            WM_XBUTTONUP = 0x020C,
            WM_XBUTTONDBLCLK = 0x020D,
            WM_MOUSELAST = 0x020D,
            WM_PARENTNOTIFY = 0x0210,
            WM_ENTERMENULOOP = 0x0211,
            WM_EXITMENULOOP = 0x0212,
            WM_NEXTMENU = 0x0213,
            WM_SIZING = 0x0214,
            WM_CAPTURECHANGED = 0x0215,
            WM_MOVING = 0x0216,
            WM_POWERBROADCAST = 0x0218,
            WM_DEVICECHANGE = 0x0219,
            WM_MDICREATE = 0x0220,
            WM_MDIDESTROY = 0x0221,
            WM_MDIACTIVATE = 0x0222,
            WM_MDIRESTORE = 0x0223,
            WM_MDINEXT = 0x0224,
            WM_MDIMAXIMIZE = 0x0225,
            WM_MDITILE = 0x0226,
            WM_MDICASCADE = 0x0227,
            WM_MDIICONARRANGE = 0x0228,
            WM_MDIGETACTIVE = 0x0229,
            WM_MDISETMENU = 0x0230,
            WM_ENTERSIZEMOVE = 0x0231,
            WM_EXITSIZEMOVE = 0x0232,
            WM_DROPFILES = 0x0233,
            WM_MDIREFRESHMENU = 0x0234,
            WM_IME_SETCONTEXT = 0x0281,
            WM_IME_NOTIFY = 0x0282,
            WM_IME_CONTROL = 0x0283,
            WM_IME_COMPOSITIONFULL = 0x0284,
            WM_IME_SELECT = 0x0285,
            WM_IME_CHAR = 0x0286,
            WM_IME_REQUEST = 0x0288,
            WM_IME_KEYDOWN = 0x0290,
            WM_IME_KEYUP = 0x0291,
            WM_MOUSEHOVER = 0x02A1,
            WM_MOUSELEAVE = 0x02A3,
            WM_NCMOUSEHOVER = 0x02A0,
            WM_NCMOUSELEAVE = 0x02A2,
            WM_WTSSESSION_CHANGE = 0x02B1,
            WM_TABLET_FIRST = 0x02c0,
            WM_TABLET_LAST = 0x02df,
            WM_CUT = 0x0300,
            WM_COPY = 0x0301,
            WM_PASTE = 0x0302,
            WM_CLEAR = 0x0303,
            WM_UNDO = 0x0304,
            WM_RENDERFORMAT = 0x0305,
            WM_RENDERALLFORMATS = 0x0306,
            WM_DESTROYCLIPBOARD = 0x0307,
            WM_DRAWCLIPBOARD = 0x0308,
            WM_PAINTCLIPBOARD = 0x0309,
            WM_VSCROLLCLIPBOARD = 0x030A,
            WM_SIZECLIPBOARD = 0x030B,
            WM_ASKCBFORMATNAME = 0x030C,
            WM_CHANGECBCHAIN = 0x030D,
            WM_HSCROLLCLIPBOARD = 0x030E,
            WM_QUERYNEWPALETTE = 0x030F,
            WM_PALETTEISCHANGING = 0x0310,
            WM_PALETTECHANGED = 0x0311,
            WM_HOTKEY = 0x0312,
            WM_PRINT = 0x0317,
            WM_PRINTCLIENT = 0x0318,
            WM_APPCOMMAND = 0x0319,
            WM_THEMECHANGED = 0x031A,
            WM_HANDHELDFIRST = 0x0358,
            WM_HANDHELDLAST = 0x035F,
            WM_AFXFIRST = 0x0360,
            WM_AFXLAST = 0x037F,
            WM_PENWINFIRST = 0x0380,
            WM_PENWINLAST = 0x038F,
            WM_APP = 0x8000,
            WM_USER = 0x0400,

        }

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

        //Guid obtained from OleAcc.idl from Platform SDK
        readonly static Guid _iAccessibleGuid = new Guid("618736e0-3c3d-11cf-810c-00aa00389b71");
        public static Guid IAccessibleGuid
        {
            get
            {
                return _iAccessibleGuid;
            }
        }

        /// Return Type: HWND->HWND__*
        ///hWndParent: HWND->HWND__*
        ///hWndChildAfter: HWND->HWND__*
        ///lpszClass: LPCSTR->CHAR*
        ///lpszWindow: LPCSTR->CHAR*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "FindWindowEx")]
        public static extern System.IntPtr FindWindowEx([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWndParent,
            [System.Runtime.InteropServices.InAttribute()] System.IntPtr hWndChildAfter,
            [System.Runtime.InteropServices.InAttribute()] string lpszClass,
            [System.Runtime.InteropServices.InAttribute()] string lpszWindow);

        /// <summary>
        /// AccessibleObjectFromWindow
        /// </summary>
        /// <param name="hwnd">Handle</param>
        /// <param name="id">ObjId</param>
        /// <param name="iid">IAccessibleGuid</param>
        /// <param name="ppvObject"></param>
        /// <returns></returns>
        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(IntPtr hwnd, ObjId id, ref Guid iid, ref IAccessible ppvObject);

        /// <summary>
        /// The AccessibleChildren function retrieves the child ID or IDispatch interface of each child within an accessible container object.
        /// </summary>
        /// <param name="paccContainer">[in] Pointer to the container object's IAccessible interface.</param>
        /// <param name="iChildStart">[in] Specifies the zero-based index of the first child retrieved. This parameter is an index, not a child ID. Typically, this parameter is set to zero (0).</param>
        /// <param name="cChildren">[in] Specifies the amount of children to retrieve.An application calls IAccessible.accChildCount to retrieve the current number of children.</param>
        /// <param name="rgvarChildren">[out] Pointer to an array of VARIANT structures that receives information about the container's children. If the vt member of an array element is VT_I4, then the lVal member for that element is the child ID. If the vt member of an array element is VT_DISPATCH, then the pdispVal member for that element is the address of the child object's IDispatch interface.</param>
        /// <param name="pcObtained">[out] Address of a variable that receives the number of elements in the rgvarChildren array filled in by the function. This value is the same as the cChildren parameter, unless you ask for more children than the number that exist. Then, this value will be less than cChildren.</param>
        /// <returns></returns>
        [DllImport("oleacc.dll")]
        public static extern uint AccessibleChildren(IAccessible paccContainer, int iChildStart, int cChildren, [Out] object[] rgvarChildren, out int pcObtained);

        [DllImport("oleacc.dll")]
        public static extern IntPtr AccessibleObjectFromPoint(Point pt, [Out, MarshalAs(UnmanagedType.Interface)] out IAccessible accObj, [Out] out object childID);


        [DllImport("oleacc.dll")]
        public static extern uint WindowFromAccessibleObject(IAccessible pacc, ref IntPtr phwnd);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        ///lpRect: LPRECT->tagRECT*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "GetWindowRect")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool GetWindowRect([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd, [System.Runtime.InteropServices.OutAttribute()] out TagRect lpRect);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        ///lpRect: RECT*
        ///bErase: BOOL->int
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "InvalidateRect")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool InvalidateRect([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd,
            [System.Runtime.InteropServices.InAttribute()] System.IntPtr lpRect,
            [System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)] bool bErase);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "UpdateWindow")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool UpdateWindow([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd);

        /// Return Type: HWND->HWND__*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "GetDesktopWindow")]
        public static extern System.IntPtr GetDesktopWindow();

        /// Return Type: HDC->HDC__*
        ///hWnd: HWND->HWND__*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "GetWindowDC")]
        public static extern System.IntPtr GetWindowDC([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd);

        /// Return Type: int
        ///hdc: HDC->HDC__*
        ///rop2: int
        [System.Runtime.InteropServices.DllImportAttribute("gdi32.dll", EntryPoint = "SetROP2")]
        public static extern BinaryRasterOperations SetROP2([System.Runtime.InteropServices.InAttribute()] System.IntPtr hdc, BinaryRasterOperations rop2);

        /// Return Type: HPEN->HPEN__*
        ///iStyle: int
        ///cWidth: int
        ///color: COLORREF->DWORD->unsigned int
        [System.Runtime.InteropServices.DllImportAttribute("gdi32.dll", EntryPoint = "CreatePen")]
        public static extern System.IntPtr CreatePen(PenStyle iStyle, int cWidth, uint color);

        /// Return Type: HGDIOBJ->void*
        ///hdc: HDC->HDC__*
        ///h: HGDIOBJ->void*
        [System.Runtime.InteropServices.DllImportAttribute("gdi32.dll", EntryPoint = "SelectObject")]
        public static extern System.IntPtr SelectObject([System.Runtime.InteropServices.InAttribute()] System.IntPtr hdc, [System.Runtime.InteropServices.InAttribute()] System.IntPtr h);

        /// Return Type: BOOL->int
        ///hdc: HDC->HDC__*
        ///left: int
        ///top: int
        ///right: int
        ///bottom: int
        [System.Runtime.InteropServices.DllImportAttribute("gdi32.dll", EntryPoint = "Rectangle")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool Rectangle([System.Runtime.InteropServices.InAttribute()] System.IntPtr hdc, int left, int top, int right, int bottom);

        /// Return Type: BOOL->int
        ///ho: HGDIOBJ->void*
        [System.Runtime.InteropServices.DllImportAttribute("gdi32.dll", EntryPoint = "DeleteObject")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool DeleteObject([System.Runtime.InteropServices.InAttribute()] System.IntPtr ho);

        /// Return Type: int
        ///hWnd: HWND->HWND__*
        ///hDC: HDC->HDC__*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "ReleaseDC")]
        public static extern int ReleaseDC([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd, [System.Runtime.InteropServices.InAttribute()] System.IntPtr hDC);

        /// Return Type: int
        ///hWnd: HWND->HWND__*
        ///lpClassName: LPWSTR->WCHAR*
        ///nMaxCount: int
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "GetClassName")]
        public static extern int GetClassName([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd,
            [System.Runtime.InteropServices.OutAttribute()] System.Text.StringBuilder lpClassName,
            int nMaxCount);

        /// Return Type: int
        ///hWnd: HWND->HWND__*
        ///lpString: LPSTR->CHAR*
        ///nMaxCount: int
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "GetWindowText")]
        public static extern int GetWindowText([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd,
            [System.Runtime.InteropServices.OutAttribute()] System.Text.StringBuilder lpString,
            int nMaxCount);

        /// Return Type: int
        ///hWnd: HWND->HWND__*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "GetWindowTextLength")]
        public static extern int GetWindowTextLength([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        ///Msg: UINT->unsigned int
        ///wParam: WPARAM->UINT_PTR->unsigned int
        ///lParam: LPARAM->LONG_PTR->int
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "PostMessage")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool PostMessage([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd, NativeMethods.WinMsg msg,
            int wParam, int lParam);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "CloseWindow")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool CloseWindow([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "OpenIcon")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool OpenIcon([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd);
    }
}
