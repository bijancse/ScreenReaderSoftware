using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SpeechBuilder;
using Gma.UserActivityMonitor;
using Accessibility;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading;

namespace DummyScreenReader
{
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00020400-0000-0000-C000-000000000046")]
    public interface IDispatch
    {
    }
    public partial class Form1 : Form
    {
        Thread thread = null;


        private Boolean releaseAlt = false;
        private Boolean ctrl = false;
        private Boolean alter = false;
        private Boolean shift = false;
        private Boolean insrt = false;


        public ManagementEventWatcher mgmtWtch;

        private enum SystemEvents : uint
        {
            EVENT_SYSTEM_SOUND = 0x0001,
            EVENT_SYSTEM_DESTROY = 0x8001,
            EVENT_SYSTEM_MINIMIZESTART = 0x0016,
            EVENT_SYSTEM_MINIMIZEEND = 0x0017,
            EVENT_SYSTEM_FOREGROUND = 0x0003,
            EVENT_SYSTEM_MENUSTART = 0x0004,
            EVENT_OBJECT_FOCUS = 0x8005,
            EVENT_OBJECT_PARENTCHANGE = 0x800F,
        }
        public const uint WINEVENT_OUTOFCONTEXT = 0x0000;
        public const uint WINEVENT_SKIPOWNTHREAD = 0x0001;
        public const uint WINEVENT_SKIPOWNPROCESS = 0x0002;
        public const uint WINEVENT_INCONTEXT = 0x0004;

        #region APIs
        public delegate bool EnumChildCallback(int hwnd, ref int lParam);

        [DllImport("ole32.dll")]
        static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string
           lpszProgID, out Guid pclsid);

        [DllImport("ole32.dll", ExactSpelling = true)]
        public static extern int CoRegisterMessageFilter(int newFilter, ref int oldMsgFilter);


        // AccessibleObjectFromWindow gets the IDispatch pointer of an object
        // that supports IAccessible, which allows us to get to the native OM.       
        //[DllImport("Oleacc.dll")]
        //private static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid, ref PowerPoint.DocumentWindow ptr);

        [DllImport("Oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            int hwnd, uint dwObjectID,
            byte[] riid,
             ref IntPtr exW);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern void SHChangeNotify(UInt32 wEventId, UInt32 uFlags, IntPtr dwItem1, IntPtr dwItem2);

        static Accessibility.IAccessible iAccessible;//interface: Accessibility namespace
        static object ChildId;

        [DllImport("oleacc.dll")]
        public static extern uint WindowFromAccessibleObject(IAccessible pacc, ref IntPtr phwnd);

        [DllImport("oleacc.dll")]
        private static extern IntPtr AccessibleObjectFromEvent(IntPtr hwnd, uint dwObjectID, uint dwChildID,
            out IAccessible ppacc, [MarshalAs(UnmanagedType.Struct)] out object pvarChild);
        [DllImport("oleacc.dll")]
        public static extern uint AccessibleChildren(IAccessible paccContainer, int iChildStart, int cChildren, [Out] object[] rgvarChildren, out int pcObtained);

        [DllImport("user32.dll")]
        static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr
        hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess,
        uint idThread, uint dwFlags);
        delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, uint idObject, uint idChild, uint dwEventThread, uint dwmsEventTime);
        #endregion

        private WinEventDelegate dEvent;
        private IntPtr pHook;
        private Boolean flag;
        private SpeechControl speaker;
        //      Win32API Win32API = new Win32API();
        public Form1(SpeechControl speaker)
        {
            InitializeComponent();


            HookManager.KeyUp += HookManager_KeyUp;
            HookManager.KeyDown += HookManager_KeyDown;

            this.speaker = speaker;

            flag = true;
            dEvent = this.WinEvent;
            pHook = SetWinEventHook(
                    (uint)SystemEvents.EVENT_SYSTEM_FOREGROUND,
                    (uint)SystemEvents.EVENT_OBJECT_FOCUS,
                    IntPtr.Zero,
                    dEvent,
                    (uint)0,
                    (uint)0,
                    WINEVENT_OUTOFCONTEXT
                    );
        }
        public static IntPtr GetControlHandlerFromEvent(IntPtr hWnd, uint idObject, uint idChild)
        {
            //IntPtr hwnd = GetFocusedWindow();
            IntPtr handler = IntPtr.Zero;
            //IAccessible accWindow = null;
            //object objChild;
            handler = AccessibleObjectFromEvent(hWnd, idObject, idChild, out iAccessible, out ChildId);
            WindowFromAccessibleObject(iAccessible, ref handler);
            return handler;

        }
        public string Gettext()
        {
            try
            {
                if (iAccessible != null && ChildId != null)
                {
                    return iAccessible.get_accName(ChildId);
                }
                else return " ";
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        private void Speak()
        {
            speaker.speak(Gettext());
        }
        private void WinEvent(IntPtr hWinEventHook, uint eventType, IntPtr hWnd, uint idObject, uint idChild, uint dwEventThread, uint dwmsEventTime)
        {
            //MessageBox.Show("fdkjfkdj");
            //i++;
            //Console.WriteLine("Object"+i);            
            if (eventType == (uint)SystemEvents.EVENT_OBJECT_FOCUS)
            {
                //Console.WriteLine("On focus change event");
                GetControlHandlerFromEvent(hWnd, idObject, idChild);
                if (thread != null) { thread.Abort(); thread = null; }
                thread = new Thread(new ThreadStart(Speak));
                thread.Start();
            }
            if (eventType == (uint)SystemEvents.EVENT_SYSTEM_FOREGROUND)
            {
                //Console.WriteLine("On Foreground change event");

            }
        }

        private void HookManager_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {

                if (!isControllingVolume(e.KeyCode.ToString()))
                {
                    speaker.stop();
                }

            }
            catch (Exception)
            {
                //MessageBox.Show("Keydown exception of first try");
            }

            if (e.KeyCode.ToString().Equals("LControlKey") || e.KeyCode.ToString().Equals("RControlKey"))
            {
                ctrl = true;
                speaker.speak(e.KeyData.ToString());
            }

            else if (e.KeyCode.ToString().Equals("Insert")) ///////////////////////////////////////
            {
                insrt = true;
            }

            else if (e.KeyCode.ToString().Equals("LMenu") || e.KeyCode.ToString().Equals("RMenu"))
            {
                releaseAlt = true;
                alter = true;
                //NarratorRunOrNotCheck();   
            }

        }
        //browser selected hotkey closed

        private void HookManager_KeyUp(object sender, KeyEventArgs e) ///// HookManager ?? ???? ??????
        {
            int i = 0;
            if (e.KeyData.ToString() == "Return")
            {
                //   MessageBox.Show("test");
            }

            else if (e.KeyCode.ToString().Equals("Tab"))
            {
                if (releaseAlt) releaseAlt = false;
            }
            else if (e.KeyCode.ToString().Equals("LControlKey") || e.KeyCode.ToString().Equals("RControlKey"))
            {
                ctrl = false;
            }
            else if (e.KeyCode.ToString().Equals("Insert")) //////////////////////////////////////////////
            {
                insrt = false;
            }

            else if (e.KeyCode.ToString().Equals("LMenu") || e.KeyCode.ToString().Equals("RMenu"))
            {
                releaseAlt = false;
                alter = false;
            }

            else if (e.KeyData.ToString().Equals("Escape"))
            {
                if (alter)
                {
                    try
                    {
                        Application.Exit();
                    }

                    catch (Exception ex)
                    {
                        //Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
                    }
                }

            }

            try
            {

                if (!isControllingVolume(e.KeyCode.ToString()))
                {
                    if (e.KeyData.ToString().Equals("Capital"))
                    {
                        if (Control.IsKeyLocked(Keys.CapsLock))
                        {
                            speaker.speak("caps lock turns on");
                        }
                        else
                            speaker.speak("caps lock turns off");
                    }
                    else if (!e.KeyCode.ToString().Equals("LMenu") && !e.KeyCode.ToString().Equals("RMenu") && !e.KeyCode.ToString().Equals("LControlKey") && !e.KeyCode.ToString().Equals("RControlKey"))
                    {
                        speaker.speak(e.KeyData.ToString());
                    }
                }
            }
            catch (Exception ex) { }

        }


        private Boolean isControllingVolume(String key)
        {
            if (key.Equals("Add") || key.Equals("Subtract")
                || key.Equals("Multiply") || key.Equals("Divide")
                || key.Equals("Space"))
            {
                return true;
            }
            return false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            WaitForProcess(); 
        }


        private void WaitForProcess()
        {
            try
            {
                WqlEventQuery query1 = new WqlEventQuery("__InstanceCreationEvent", new TimeSpan(0, 0, 3), "TargetInstance isa \"Win32_Process\"");
                WqlEventQuery query2 = new WqlEventQuery("__InstancedeletionEvent", new TimeSpan(0, 0, 3), "TargetInstance isa \"Win32_Process\"");
                ManagementEventWatcher startWatch1 = new ManagementEventWatcher(query1);
                ManagementEventWatcher startWatch2 = new ManagementEventWatcher(query2);
                startWatch1.EventArrived
                                    += new EventArrivedEventHandler(startWatch_EventArrived);
                startWatch2.EventArrived
                                    += new EventArrivedEventHandler(endWatch_EventArrived);
                startWatch1.Start();
                startWatch2.Start();

            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }

        }

          private void startWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {            
            ManagementBaseObject bobj = ((ManagementBaseObject)e.NewEvent["TargetInstance"]);
            //MessageBox.Show(bobj["Name"].ToString());//return process name ex:notepad.exe          
            String processName = bobj["Name"].ToString();
            Console.WriteLine(processName);           
            try
            {
                //System.Threading.Thread.Sleep(6000);
                if (processName == "WINWORD.EXE")
                {
                    speaker.speak("OPEN WORD DOCUMENT");                  
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("on Event arrived" + ex.ToString());
                //excel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

            }
        }
        private void endWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject mBaseObj = ((ManagementBaseObject)e.NewEvent["TargetInstance"]);
            //MessageBox.Show(bobj["Name"].ToString());//return process name ex:notepad.exe          
            String endProcessName = mBaseObj["Name"].ToString();
            Console.WriteLine("endProcessName"+endProcessName);
            if (endProcessName == "WINWORD.EXE")
            {
                speaker.speak("Close WORD DOCUMENT");  
            }                    
        }

    }
}
