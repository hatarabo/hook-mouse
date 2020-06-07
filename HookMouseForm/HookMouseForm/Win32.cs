using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace HookMouseForm
{
    internal class Win32
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct Point
        {
            public int x;
            public int y;
        }

        public struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            public Rect(int left, int top, int right, int bottom)
            {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }

            public Rect(System.Drawing.Rectangle r) : this(r.Left, r.Top, r.Right, r.Bottom) { }

            public int X
            {
                get
                {
                    return Left;
                }
                set
                {
                    Right -= (Left - value); Left = value;
                }
            }

            public int Y
            {
                get
                {
                    return Top;
                }
                set
                {
                    Bottom -= (Top - value); Top = value;
                }
            }

            public int Height
            {
                get
                {
                    return Bottom - Top;
                }
                set
                {
                    Bottom = value + Top;
                }
            }

            public int Width
            {
                get
                {
                    return Right - Left;
                }
                set
                {
                    Right = value + Left;
                }
            }

            public System.Drawing.Point Location
            {
                get
                {
                    return new System.Drawing.Point(Left, Top);
                }
                set
                {
                    X = value.X;
                    Y = value.Y;
                }
            }

            public System.Drawing.Size Size
            {
                get
                {
                    return new System.Drawing.Size(Width, Height);
                }
                set
                {
                    Width = value.Width; Height = value.Height;
                }
            }

        }


        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(Point p);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetWindowRect(IntPtr hwnd, out Rect lpRect);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, string lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowTextLength(IntPtr hwnd);
    }
}
