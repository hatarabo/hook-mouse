using HataRabo.Hook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HookMouseForm
{
    public partial class Form1 : Form
    {
        private const string LocationFormat = "Corsor Locaiton = [{0}, {1}]";

        private class Data
        {
            public string Title { get; private set; }
            public Rectangle Rectangle { get; private set; }

            public Data(string title, Rectangle rectangle)
            {
                this.Title = title;
                this.Rectangle = rectangle;
            }
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Loadイベント
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            GlobalHook_Mouse.AddEvent(this.HookMouse);
        }

        /// <summary>
        /// Closedイベント
        /// </summary>
        /// <param name="e"></param>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            GlobalHook_Mouse.RemoveEvent(this.HookMouse);
        }

        /// <summary>
        /// 開始ボタン＿クリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void startButton_Click(object sender, EventArgs e)
        {
            this.startButton.Enabled = false;
            GlobalHook_Mouse.Start();
            this.stopButton.Enabled = true;
        }

        /// <summary>
        /// 停止ボタン＿クリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void stopButton_Click(object sender, EventArgs e)
        {
            this.stopButton.Enabled = false;
            GlobalHook_Mouse.Stop();
            this.startButton.Enabled = true;
        }

        /// <summary>
        /// マウスフック実処理
        /// </summary>
        /// <param name="state"></param>
        private void HookMouse(ref GlobalHook_Mouse.MouseState state)
        {
            GlobalHook_Mouse.IsPaused = true;

            try
            {
                var mouseButton = state.button;
                var mouseAction = state.action;

                if (mouseAction == GlobalHook_Mouse.MouseAction.Move)
                {
                    var locationText = string.Format(LocationFormat, state.x, state.y);

                    this.Invoke((MethodInvoker)delegate
                    {
                        this.locationLabel.Text = locationText;
                    });
                }
                else if (mouseButton == GlobalHook_Mouse.MouseButton.Left
                        && mouseAction == GlobalHook_Mouse.MouseAction.Down)
                {
                    var point = new Win32.Point() { x = state.x, y = state.y };
                    var handle = Win32.WindowFromPoint(point);
                    var rect = new Win32.Rect();

                    bool result = true;
                    result &= Win32.GetWindowRect(handle, out rect);

                    var length = Win32.GetWindowTextLength(handle);
                    var title = new string('\0', length + 1);
                    result &= 0 != Win32.GetWindowText(handle, title, title.Length);

                    Data data = null;

                    if (result)
                    {
                        data = new Data(title, new Rectangle(rect.Location, rect.Size));
                    }

                    this.Invoke((MethodInvoker)delegate
                    {
                        this.propertyGrid.SelectedObject = data;
                    });
                }
            }
            finally
            {
                GlobalHook_Mouse.IsPaused = false;
            }
        }
    }
}
