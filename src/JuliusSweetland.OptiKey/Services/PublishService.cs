using System;
using System.Linq;
using System.Windows;
using WindowsInput;
using WindowsInput.Native;
using JuliusSweetland.OptiKey.Static;
using log4net;
using System.Runtime.InteropServices;


namespace JuliusSweetland.OptiKey.Services
{
    public class PublishService : IPublishService
    {
        const int INPUT_MOUSE = 0;
        const int INPUT_KEYBOARD = 1;
        const int INPUT_HARDWARE = 2;
        const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        const uint KEYEVENTF_KEYUP = 0x0002;
        const uint KEYEVENTF_UNICODE = 0x0004;
        const uint KEYEVENTF_SCANCODE = 0x0008;

        struct INPUT
        {
            public int type;
            public InputUnion u;
        }

        [StructLayout(LayoutKind.Explicit)]
        struct InputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;
            [FieldOffset(0)]
            public KEYBDINPUT ki;
            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT
        {
            /*Virtual Key code.  Must be from 1-254.  If the dwFlags member specifies KEYEVENTF_UNICODE, wVk must be 0.*/
            public ushort wVk;
            /*A hardware scan code for the key. If dwFlags specifies KEYEVENTF_UNICODE, wScan specifies a Unicode character which is to be sent to the foreground application.*/
            public ushort wScan;
            /*Specifies various aspects of a keystroke.  See the KEYEVENTF_ constants for more information.*/
            public uint dwFlags;
            /*The time stamp for the event, in milliseconds. If this parameter is zero, the system will provide its own time stamp.*/
            public uint time;
            /*An additional value associated with the keystroke. Use the GetMessageExtraInfo function to obtain this information.*/
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        [DllImport("user32.dll")]
        static extern IntPtr GetMessageExtraInfo();

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern short VkKeyScan(char ch);


        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        private readonly InputSimulator inputSimulator;
        private readonly WindowsInputDeviceStateAdaptor inputDeviceStateAdaptor;

        public event EventHandler<Exception> Error;
        
        public PublishService()
        {
            inputSimulator = new WindowsInput.InputSimulator();
            inputDeviceStateAdaptor = new WindowsInput.WindowsInputDeviceStateAdaptor();
        }

        public void ReleaseAllDownKeys()
        {
            try
            {
                Log.InfoFormat("Releasing all keys (with virtual key codes) which are down.");
                foreach (var virtualKeyCode in Enum.GetValues(typeof(VirtualKeyCode)).Cast<VirtualKeyCode>())
                {
                    if (inputDeviceStateAdaptor.IsHardwareKeyDown(virtualKeyCode))
                    {
                        Log.DebugFormat("{0} is down - calling KeyUp", virtualKeyCode);
                        KeyUp(virtualKeyCode);
                    }
                }
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void KeyDown(VirtualKeyCode virtualKeyCode)
        {
            try
            {
                Log.DebugFormat("Simulating key down {0}", virtualKeyCode);
                inputSimulator.Keyboard.KeyDown(virtualKeyCode);
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void KeyUp(VirtualKeyCode virtualKeyCode)
        {
            try
            {
                Log.DebugFormat("Simulating key up: {0}", virtualKeyCode);
                inputSimulator.Keyboard.KeyUp(virtualKeyCode);
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void KeyDownUp(VirtualKeyCode virtualKeyCode)
        {
            try
            {
                Log.DebugFormat("Simulating key press (down & up): {0}", virtualKeyCode);
                //inputSimulator.Keyboard.KeyPress(virtualKeyCode);

                //char keychar = clickedButton.Text[0];
                ushort vkey = (ushort) virtualKeyCode;

                INPUT[] inputs = new INPUT[]
                {
                new INPUT
                {
                    type = INPUT_KEYBOARD,
                    u = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = vkey,
                            wScan = 0,
                            dwFlags = 0,
                            dwExtraInfo = GetMessageExtraInfo(),
                        }
                    }
                }
                };

                SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));

            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void TypeText(string text)
        {
            try
            {
                Log.DebugFormat("Simulating typing text '{0}'", text);
                inputSimulator.Keyboard.TextEntry(text);
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void MouseMouseToPoint(Point point)
        {
            try
            {
                Log.DebugFormat("Simulating moving mouse to point '{0}'", point);

                var virtualScreenWidthInPixels = SystemParameters.VirtualScreenWidth * Graphics.DipScalingFactorX;
                var virtualScreenHeightInPixels = SystemParameters.VirtualScreenHeight * Graphics.DipScalingFactorY;

                //N.B. InputSimulator does not deal in pixels. The position should be a scaled point between 0 and 65535. 
                //https://inputsimulator.codeplex.com/discussions/86530
                inputSimulator.Mouse.MoveMouseToPositionOnVirtualDesktop(
                    Math.Ceiling(65535 * (point.X / virtualScreenWidthInPixels)),
                    Math.Ceiling(65535 * (point.Y / virtualScreenHeightInPixels)));
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void LeftMouseButtonClick()
        {
            try
            {
                Log.Info("Simulating clicking the left mouse button click");
                inputSimulator.Mouse.LeftButtonClick();
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void LeftMouseButtonDoubleClick()
        {
            try
            {
                Log.Info("Simulating pressing the left mouse button down twice");
                inputSimulator.Mouse.LeftButtonDoubleClick();
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void LeftMouseButtonDown()
        {
            try
            {
                Log.Info("Simulating pressing the left mouse button down");
                inputSimulator.Mouse.LeftButtonDown();
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void LeftMouseButtonUp()
        {
            try
            {
                Log.Info("Simulating releasing the left mouse button down");
                inputSimulator.Mouse.LeftButtonUp();
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void MiddleMouseButtonClick()
        {
            try
            {
                Log.Info("Simulating clicking the middle mouse button click");
                inputSimulator.Mouse.MiddleButtonClick();
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void MiddleMouseButtonDown()
        {
            try
            {
                Log.Info("Simulating pressing the middle mouse button down");
                inputSimulator.Mouse.MiddleButtonDown();
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void MiddleMouseButtonUp()
        {
            try
            {
                Log.Info("Simulating releasing the middle mouse button down");
                inputSimulator.Mouse.MiddleButtonUp();
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void RightMouseButtonClick()
        {
            try
            {
                Log.Info("Simulating pressing the right mouse button down");
                inputSimulator.Mouse.RightButtonClick();
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void RightMouseButtonDown()
        {
            try
            {
                Log.Info("Simulating pressing the right mouse button down");
                inputSimulator.Mouse.RightButtonDown();
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void RightMouseButtonUp()
        {
            try
            {
                Log.Info("Simulating releasing the right mouse button down");
                inputSimulator.Mouse.RightButtonUp();
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void ScrollMouseWheelUp(int clicks)
        {
            try
            {
                Log.DebugFormat("Simulating scrolling the vertical mouse wheel up by {0} clicks", clicks);
                inputSimulator.Mouse.VerticalScroll(clicks);
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void ScrollMouseWheelDown(int clicks)
        {
            try
            {
                Log.DebugFormat("Simulating scrolling the vertical mouse wheel down by {0} clicks", clicks);
                inputSimulator.Mouse.VerticalScroll(0 - clicks);
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void ScrollMouseWheelLeft(int clicks)
        {
            try
            {
                Log.DebugFormat("Simulating scrolling the horizontal mouse wheel left by {0} clicks", clicks);
                inputSimulator.Mouse.HorizontalScroll(0 - clicks);
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        public void ScrollMouseWheelRight(int clicks)
        {
            try
            {
                Log.DebugFormat("Simulating scrolling the horizontal mouse wheel right by {0} clicks", clicks);
                inputSimulator.Mouse.HorizontalScroll(clicks);
            }
            catch (Exception exception)
            {
                PublishError(this, exception);
            }
        }

        private void PublishError(object sender, Exception ex)
        {
            Log.Error("Publishing Error event (if there are any listeners)", ex);
            if (Error != null)
            {
                Error(sender, ex);
            }
        }
    }
}
