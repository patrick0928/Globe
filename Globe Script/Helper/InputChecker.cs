using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Globe_Script.Helper
{
    public class InputChecker
    {
        [DllImport("User32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, ref Rect rectangle);

        public struct Rect
        {
            public int Left { get; set; }
            public int Top { get; set; }
            public int Right { get; set; }
            public int Bottom { get; set; }
        }

        public void RunChecker()
        {
            Process process = Process.GetProcessesByName("abacusworkspace").FirstOrDefault();
            IntPtr ptr = process.MainWindowHandle;

            Rect sabreWorkspace = new Rect();
            GetWindowRect(ptr, ref sabreWorkspace);

            ClickOnPoint click = new ClickOnPoint();
            click.OnClick(process.MainWindowHandle, new System.Drawing.Point(500, sabreWorkspace.Bottom - 90), "left");

            Task.Delay(300).ContinueWith(t => mouseMove());
        }

        public void mouseMove()
        {
            Process process = Process.GetProcessesByName("abacusworkspace").FirstOrDefault();
            ClickOnPoint click = new ClickOnPoint();
            click.OnClick(process.MainWindowHandle, new System.Drawing.Point(10, 15), "mouseMove");

            Task.Delay(300).ContinueWith(t => mouseRightClick());
        }

        public void mouseRightClick()
        {
            Process process = Process.GetProcessesByName("abacusworkspace").FirstOrDefault();
            ClickOnPoint click = new ClickOnPoint();
            click.OnClick(process.MainWindowHandle, new System.Drawing.Point(10, 150), "right");

            Task.Delay(300).ContinueWith(t => copyText());
        }

        public void copyText()
        {
            Process process = Process.GetProcessesByName("abacusworkspace").FirstOrDefault();
            ClickOnPoint click = new ClickOnPoint();
            click.OnClick(process.MainWindowHandle, new System.Drawing.Point(35, 165), "copyText");
        }
    }
}
