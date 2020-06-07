using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HataRabo.Hook
{
    /// <summary>
    /// グローバルフックをするためのクラス
    /// </summary>
    public class GlobalHook_Mouse
    {
        #region 定数

        private const uint WM_MOUSEMOVE = 0x0200;
        private const uint WM_LBUTTONDOWN = 0x0201;
        private const uint WM_LBUTTONUP = 0x0202;
        private const uint WM_LBUTTONDBLCLK = 0x0203;
        private const uint WM_RBUTTONDOWN = 0x0204;
        private const uint WM_RBUTTONUP = 0x0205;
        private const uint WM_RBUTTONDBLCLK = 0x0206;
        private const uint WM_MBUTTONDOWN = 0x0207;
        private const uint WM_MBUTTONUP = 0x0208;
        private const uint WM_MBUTTONDBLCLK = 0x0209;
        private const uint WM_MOUSEWHEEL = 0x020A;
        private const uint WM_XBUTTONDOWN = 0x020B;
        private const uint WM_XBUTTONUP = 0x020C;
        private const uint WM_XBUTTONDBLCLK = 0x020D;
        private const uint WM_MOUSEHWHEEL = 0x020E;

        #endregion

        #region 列挙体

        /// <summary>
        /// マウスボタン
        /// </summary>
        [Flags]
        public enum MouseButton : int
        {
            /// <summary>
            /// 
            /// </summary>
            None = 0x00,

            /// <summary>
            /// 左
            /// </summary>
            Left = 0x01,

            /// <summary>
            /// 右
            /// </summary>
            Right = 0x02,

            /// <summary>
            /// 中央
            /// </summary>
            Middle = 0x04,

            /// <summary>
            /// ホイール上に回す
            /// </summary>
            WheelUp = 0x08,

            /// <summary>
            /// ホイール下に回す
            /// </summary>
            WheelDown = 0x10,

            /// <summary>
            /// XButton1
            /// </summary>
            XButton1 = 0x20,

            /// <summary>
            /// XButton2
            /// </summary>
            XButton2 = 0x40,

            /// <summary>
            /// Unknown
            /// </summary>
            Unknown = 0xFF,
        }

        /// <summary>
        /// マウス操作
        /// </summary>
        [Flags]
        public enum MouseAction
        {
            /// <summary>
            /// 
            /// </summary>
            None = 0x00,

            /// <summary>
            /// 移動
            /// </summary>
            Move = 0x01,

            /// <summary>
            /// ダウン
            /// </summary>
            Down = 0x02,

            /// <summary>
            /// アップ
            /// </summary>
            Up = 0x04,

            /// <summary>
            /// クリック
            /// </summary>
            Click = 0x08,

            /// <summary>
            /// ダブルクリック
            /// </summary>
            DoubleClick = 0x10,

            /// <summary>
            /// ドラッグ＆ドロップ
            /// </summary>
            DragAndDrop = 0x20
        }

        /// <summary>
        /// 修飾キー
        /// </summary>
        [Flags]
        public enum ModifiedKey
        {
            /// <summary>
            /// 
            /// </summary>
            None = 0x00,

            /// <summary>
            /// Shift
            /// </summary>
            Shift = 0x01,

            /// <summary>
            /// Ctrl
            /// </summary>
            Ctrl = 0x02,

            /// <summary>
            /// Alt
            /// </summary>
            Alt = 0x04,

            /// <summary>
            /// Windows
            /// </summary>
            Windows = 0x08
        }

        #endregion

        #region デリゲート

        /// <summary>
        /// フックイベントハンドラのデリゲート
        /// </summary>
        /// <param name="state"></param>
        public delegate void HookHandler(ref MouseState state);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hookCode"></param>
        /// <param name="msg"></param>
        /// <param name="mslHookStruct"></param>
        /// <returns></returns>
        private delegate IntPtr MouseHookCallBack(int hookCode, uint msg, ref MSLHooKStruct mslHookStruct);

        #endregion


        #region 構造体

        /// <summary>
        /// マウスの状態
        /// </summary>
        public struct MouseState
        {
            public MouseButton button;
            public MouseAction action;
            public int x;
            public int y;
            public uint data;
            public uint flags;
            public uint time;
            public IntPtr extraInfo;
        }

        /// <summary>
        /// マウス座標
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct Point
        {
            public int x;
            public int y;
        }

        /// <summary>
        /// Win32API用パラメータ
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct MSLHooKStruct
        {
            public Point point;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr extraInfo;
        }

        #endregion


        #region フィールド

        /// <summary>
        /// マウスの状態
        /// </summary>
        private static MouseState state;

        /// <summary>
        /// フックしたプロセスのハンドル
        /// </summary>
        private static IntPtr hookProcedureHandle;

        /// <summary>
        /// フック処理停止フラグ
        /// </summary>
        private static bool isCancel;

        /// <summary>
        /// フックイベントハンドラのコレクション
        /// </summary>
        private static List<HookHandler> events;

        /// <summary>
        /// フックイベント
        /// </summary>
        private static event HookHandler hookEvent;

        /// <summary>
        /// フック処理関数実行用デリゲート
        /// </summary>
        private static event MouseHookCallBack hookCallBack;

        #endregion


        #region プロパティ

        /// <summary>
        /// フック状態フラグ
        /// </summary>
        public static bool IsHooking { get; private set; }

        /// <summary>
        /// ポーズ状態フラグ
        /// </summary>
        public static bool IsPaused { get; set; }

        #endregion


        #region メソッド

        /// <summary>
        /// マウスのグローバルフックの開始
        /// </summary>
        public static void Start()
        {
            if (GlobalHook_Mouse.IsHooking)
            {
                return;
            }

            GlobalHook_Mouse.IsHooking = true;
            GlobalHook_Mouse.IsPaused = false;
            GlobalHook_Mouse.hookCallBack += GlobalHook_Mouse.HookProcedure;

            IntPtr handle = GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName);

            int WH_MOUSE_LL = 14;
            GlobalHook_Mouse.hookProcedureHandle = SetWindowsHookEx(WH_MOUSE_LL, GlobalHook_Mouse.hookCallBack, handle, 0);

            if (GlobalHook_Mouse.hookProcedureHandle == IntPtr.Zero)
            {
                GlobalHook_Mouse.IsHooking = false;
                GlobalHook_Mouse.IsPaused = true;

                throw new System.ComponentModel.Win32Exception();
            }
        }

        /// <summary>
        /// マウスのグローバルフックの終了
        /// </summary>
        public static void Stop()
        {
            if (!GlobalHook_Mouse.IsHooking)
            {
                return;
            }

            if (hookProcedureHandle != IntPtr.Zero)
            {
                GlobalHook_Mouse.IsHooking = false;
                GlobalHook_Mouse.IsPaused = true;

                //GlobalHook_Mouse.ClearEvent();

                UnhookWindowsHookEx(GlobalHook_Mouse.hookProcedureHandle);
                GlobalHook_Mouse.hookProcedureHandle = IntPtr.Zero;
                GlobalHook_Mouse.hookCallBack -= GlobalHook_Mouse.HookProcedure;
            }
        }

        /// <summary>
        /// キャンセル
        /// </summary>
        public static void Cancel()
        {
            GlobalHook_Mouse.isCancel = true;
        }

        /// <summary>
        /// フックイベントの追加
        /// </summary>
        /// <param name="hookHandler"></param>
        public static void AddEvent(HookHandler hookHandler)
        {
            if (GlobalHook_Mouse.events == null)
            {
                GlobalHook_Mouse.events = new List<HookHandler>();
            }

            GlobalHook_Mouse.events.Add(hookHandler);
            GlobalHook_Mouse.hookEvent += GlobalHook_Mouse.events.Last();
        }

        /// <summary>
        /// フックイベントの削除
        /// </summary>
        /// <param name="hookHandler"></param>
        public static void RemoveEvent(HookHandler hookHandler)
        {
            if (GlobalHook_Mouse.events == null)
            {
                return;
            }

            GlobalHook_Mouse.hookEvent -= hookHandler;
            GlobalHook_Mouse.events.Remove(hookHandler);
        }

        /// <summary>
        /// フックイベントの全削除
        /// </summary>
        public static void ClearEvent()
        {
            if (GlobalHook_Mouse.events == null)
            {
                return;
            }

            foreach (var handler in GlobalHook_Mouse.events)
            {
                GlobalHook_Mouse.hookEvent -= handler;
            }

            GlobalHook_Mouse.events.Clear();
        }

        /// <summary>
        /// フック手続き
        /// </summary>
        /// <param name="hookCode"></param>
        /// <param name="msg"></param>
        /// <param name="hookStruct"></param>
        /// <returns></returns>
        private static IntPtr HookProcedure(int hookCode, uint msg, ref MSLHooKStruct hookStruct)
        {
            try
            {
                if (0 <= hookCode && GlobalHook_Mouse.hookEvent != null && !GlobalHook_Mouse.IsPaused)
                {
                    var stroke = GlobalHook_Mouse.ConverManageStroke(msg, ref hookStruct);
                    state.button = stroke.Item1;
                    state.action = stroke.Item2;
                    state.x = hookStruct.point.x;
                    state.y = hookStruct.point.y;
                    state.data = hookStruct.mouseData;
                    state.flags = hookStruct.flags;
                    state.time = hookStruct.time;
                    state.extraInfo = hookStruct.extraInfo;

                    GlobalHook_Mouse.hookEvent(ref state);

                    if (GlobalHook_Mouse.isCancel)
                    {
                        GlobalHook_Mouse.isCancel = false;
                        return new IntPtr(1);
                    }
                }

                return CallNextHookEx(GlobalHook_Mouse.hookProcedureHandle, hookCode, msg, ref hookStruct);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 修飾キーを取得する
        /// </summary>
        /// <param name="modifierKey"></param>
        /// <returns></returns>
        public static Keys GetModifiedKeys(ModifiedKey modifierKey)
        {
            Keys keys = Keys.None;

            if ((ModifiedKey.Alt & modifierKey) == ModifiedKey.Alt)
            {
                keys |= Keys.Alt;
            }

            if ((ModifiedKey.Shift & modifierKey) == ModifiedKey.Shift)
            {
                keys |= Keys.ShiftKey;
            }

            if ((ModifiedKey.Ctrl & modifierKey) == ModifiedKey.Ctrl)
            {
                keys |= Keys.ControlKey;
            }

            return keys;
        }

        /// <summary>
        /// マウスの操作状況の取得
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="hookStruct"></param>
        /// <returns></returns>
        private static Tuple<MouseButton, MouseAction> ConverManageStroke(uint msg, ref MSLHooKStruct hookStruct)
        {
            MouseButton button;
            MouseAction action;

            switch (msg)
            {
                case WM_MOUSEMOVE:
                    button = MouseButton.None;
                    action = MouseAction.Move;
                    break;
                case WM_LBUTTONDOWN:
                    button = MouseButton.Left;
                    action = MouseAction.Down;
                    break;
                case WM_LBUTTONUP:
                    button = MouseButton.Left;
                    action = MouseAction.Up;
                    break;
                case WM_RBUTTONDOWN:
                    button = MouseButton.Right;
                    action = MouseAction.Down;
                    break;
                case WM_RBUTTONUP:
                    button = MouseButton.Right;
                    action = MouseAction.Up;
                    break;
                case WM_MBUTTONDOWN:
                    button = MouseButton.Middle;
                    action = MouseAction.Down;
                    break;
                case WM_MBUTTONUP:
                    button = MouseButton.Middle;
                    action = MouseAction.Up;
                    break;
                case WM_MOUSEHWHEEL:
                    action = MouseAction.None;
                    if (0 < (hookStruct.mouseData >> 16 & 0xFFFF))
                    {
                        button = MouseButton.WheelUp;
                    }
                    else
                    {
                        button = MouseButton.WheelDown;
                        break;
                    }
                    break;
                case WM_XBUTTONDOWN:
                    action = MouseAction.Down;
                    switch (hookStruct.mouseData >> 16)
                    {
                        case 1:
                            button = MouseButton.XButton1;
                            break;
                        case 2:
                            button = MouseButton.XButton2;
                            break;
                        default:
                            button = MouseButton.Unknown;
                            break;
                    }
                    break;
                case WM_XBUTTONUP:
                    action = MouseAction.Up;
                    switch (hookStruct.mouseData >> 16)
                    {
                        case 1:
                            button = MouseButton.XButton1;
                            break;
                        case 2:
                            button = MouseButton.XButton2;
                            break;
                        default:
                            button = MouseButton.Unknown;
                            break;
                    }
                    break;
                default:
                    action = MouseAction.None;
                    button = MouseButton.Unknown;
                    break;
            }

            return new Tuple<MouseButton, MouseAction>(button, action);
        }

        #region Win32API

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hookType"></param>
        /// <param name="hookProcedure"></param>
        /// <param name="hMod"></param>
        /// <param name="threadId"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int hookType, MouseHookCallBack hookProcedure, IntPtr hMod, uint threadId);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hookHandle"></param>
        /// <param name="hookCode"></param>
        /// <param name="msg"></param>
        /// <param name="mslHookStruct"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hookHandle, int hookCode, uint msg, ref MSLHooKStruct mslHookStruct);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hookHandle"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hookHandle);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="lpModuleName"></param>
        /// <returns></returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        #endregion


        #endregion
    }
}
