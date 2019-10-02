using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Globe_Script.Helper
{
    public class ClickOnPoint
    {
        [DllImport("User32")]
        static extern bool ClientToScreen(IntPtr hWnd, ref Point lnPoint);

        [DllImport("User32")]
        internal static extern uint SendInput(uint nInput, [MarshalAs(UnmanagedType.LPArray), In] INPUT[] pinputs, int cbSize);

        internal struct INPUT
        {
            public UInt32 type;
            public MOUSEKEYBDHARWAREINPUT Data;
        }

        [StructLayout(LayoutKind.Explicit)]
        internal struct MOUSEKEYBDHARWAREINPUT
        {
            [FieldOffset(0)]
            public MOUSEINPUT Mouse;
        }

        internal struct MOUSEINPUT
        {
            public Int32 X;
            public Int32 Y;
            public UInt32 MouseData;
            public UInt32 Flags;
            public UInt32 Time;
            public IntPtr Extrainfo;
        }

        public void OnClick(IntPtr wndHandle, Point clientPoint, string action)
        {
            var oldPos = System.Windows.Forms.Cursor.Position;

            ClientToScreen(wndHandle, ref clientPoint);
            System.Windows.Forms.Cursor.Position = new Point(clientPoint.X, clientPoint.Y);

            var inputMouseDown = new INPUT();
            var inputMouseUp = new INPUT();
            if(action == "right")
            {
                inputMouseDown.type = 0;
                inputMouseDown.Data.Mouse.Flags = 0x0008;

                inputMouseUp.type = 0;
                inputMouseUp.Data.Mouse.Flags = 0x0010;
            }
            else if(action == "left")
            {
                inputMouseDown.type = 0;
                inputMouseDown.Data.Mouse.Flags = 0x0002;
            }
            else if(action == "mouseMove")
            {
                inputMouseDown.type = 0;
                inputMouseDown.Data.Mouse.Flags = 0x0001;

                inputMouseUp.type = 0;
                inputMouseUp.Data.Mouse.Flags = 0x0004;
            }
            else if(action == "copyText")
            {
                inputMouseDown.type = 0;
                inputMouseDown.Data.Mouse.Flags = 0x0002;

                inputMouseUp.type = 0;
                inputMouseUp.Data.Mouse.Flags = 0x0004;
            }

            var inputs = new INPUT[] { inputMouseDown, inputMouseUp };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));

            System.Windows.Forms.Cursor.Position = oldPos;
        }
    }
}
