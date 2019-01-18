using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rainmeter;

using System.Diagnostics;
using System.Management;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.IO;
using System.Text;


//This is the source code for a custom dll being used by the rainmeter skin Distiller Block
//It tracks current running application in the operation system and convert icons into png file for display purpose
//Author: Innonion

namespace PluginHandleProcess
{
    //Necessary Information and data about a process, including:
    //WindowName: Title of the main window
    //ProcessNamePng: Process Name in the form of png ( e.g. firefox.png )
    //ExeAddress: Location of the exe that executes the process ( used for getting icon )
    public struct ProcessData
    {
        public string WindowName;
        public string ProcessNamePng;
        public string ExeAddress;

        //Constructor
        public ProcessData( string WN = "", string PN = "Default.png", string address = "")
        {
          WindowName = WN;
          ProcessNamePng = PN;
          ExeAddress = address;
        }

        //Copy Constructor
        public ProcessData(ProcessData copyobj)
        {
            WindowName = copyobj.WindowName;
            ProcessNamePng = copyobj.ProcessNamePng;
            ExeAddress = copyobj.ExeAddress;
        }

        public override bool Equals(object obj)
        {
            if (!this.GetType().Equals(obj.GetType()))
                return false;
            else
            {
                ProcessData UWPObj = (ProcessData)obj;
                return (this.WindowName == UWPObj.WindowName && this.ProcessNamePng == UWPObj.ProcessNamePng) ? true : false;
            }
        }
    }


    class Measure
    {
        //Write values to .ini file
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);

        static public implicit operator Measure(IntPtr data)
        {
            return (Measure)GCHandle.FromIntPtr(data).Target;
        }
        public IntPtr buffer = IntPtr.Zero;

        //Delay detecting Number Of UWP for each time the skin refresh, due to weird false-detection of double-detecting UWP when refreshing a rainmeter skin
        public int DelayCycle = 3;

        //Max Size of the program
        public const int MaxSize = 15;
     
        //Total number of all running processes, including those that are background processes
        public int TotalProcSize = 0;

        //Actual total number of previous running applications
        public int PrevSize = 0;

        //Total number of UWP ( as running application ) 
        public int UWPCount = 0;

        //Total number of UWP in PreviousUpdate Cycle
        public int PrevUWPCount = 0;

        //Total number of Windows Setting file
        public int WinSettingCount = 0;

        //Total number of Windows Setting file in PreviousUpdate Cycle
        public int PrevWinSettingCount = 0;


        //Previous string list of running applications ( in .png)
        public HashSet<string> PreSysProcesses = new HashSet<string>();

        //Current list of running applications
        public ProcessData[] TraySysProcesses = new ProcessData[MaxSize];


        //Previous Special list for holding UWP application
        public HashSet<string> PrevUWPList = new HashSet<string>();


        //Special list for holding UWP applications ( in .png)
        public ProcessData[] UWPList = new ProcessData[MaxSize];

        //Determines if a change in application processes are made
        public static int IsUpdate = 0;

        //iniPath
        public static string iniPath = string.Concat(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString(), "\\Rainmeter\\Skins\\Distiller_Block\\Display\\Display.ini");


        //IsRunningApplication determines whether the process is actually a real application, i.e. Application seen in alt-tab and under TaskManager Section
        //0: it is not a running application
        //1: it is a normal Windows running application
        internal int IsRunningApplication(Process service, ProcessData[] Applist = null, int ApplistSize = 0)
        {
            Process p = service;
            string name = p.ProcessName;
            string nameInPNG = string.Concat(name, ".png");

            //Exclude specific running process(es):
            //Rainmeter:            No need to self-check itself
            if (name == "Rainmeter"  )
            {
                return 0;
            }

            //Check if it has a graphical interface
            if (p.MainWindowHandle == IntPtr.Zero)
                return 0;


            try { string test = p.MainModule.FileName; }
            catch (Exception e)
            {
                //If the program run into errors getting ExecutionFile Location, it is probably Native Windows Setting ( like Task Manager )
                WinSettingCount++;
                return 0;
            }

            
            //Check Redundancy
            for ( int i = 0; i < ApplistSize; i++)
            {
                if (nameInPNG == Applist[i].ProcessNamePng )
                    return 0;
            }
            return 1;
        }

        //Set and Write Variable to ini File
        internal void SetVariable( string KeyName, string Value, IntPtr rm)
        {
            WritePrivateProfileString("Variables", KeyName, Value, Measure.iniPath);

            Rainmeter.API api = (Rainmeter.API)rm;
            string Command = string.Concat("[!SetVariable ", KeyName, " \"", Value, "\"]");
            API.Execute(api.GetSkin(), Command);
            return;
        }
    }

    public class Plugin
    {
        //Read Value from .ini file
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        static extern int GetPrivateProfileString(string section, string key, string defaultValue,
                                                  [In, Out] char[] value, int size, string filePath);

        //Write values to .ini file
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);

        // return windows text
        [DllImport("user32.dll", EntryPoint = "GetWindowText",
        ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpWindowText, int nMaxCount);

        // check if windows visible       
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        //Get Window Style
        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

        //Get Process ID through Window Handle
        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, ref uint lpdwProcessId);

        //enumarator on all desktop windows
        [DllImport("user32.dll", EntryPoint = "EnumDesktopWindows",
        ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool EnumDesktopWindows(IntPtr hDesktop, EnumDelegate lpEnumCallbackFunction, IntPtr lParam);

        [DllImport("dwmapi.dll")]
        public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out IntPtr pvAttribute, int cbAttribute);

        public enum DWMWINDOWATTRIBUTE
        {
            DWMWA_NCRENDERING_ENABLED = 1,
            DWMWA_NCRENDERING_POLICY,
            DWMWA_TRANSITIONS_FORCEDISABLED,
            DWMWA_ALLOW_NCPAINT,
            DWMWA_CAPTION_BUTTON_BOUNDS,
            DWMWA_NONCLIENT_RTL_LAYOUT,
            DWMWA_FORCE_ICONIC_REPRESENTATION,
            DWMWA_FLIP3D_POLICY,
            DWMWA_EXTENDED_FRAME_BOUNDS,
            DWMWA_HAS_ICONIC_BITMAP,
            DWMWA_DISALLOW_PEEK,
            DWMWA_EXCLUDED_FROM_PEEK,
            DWMWA_CLOAK,
            DWMWA_CLOAKED,
            DWMWA_FREEZE_REPRESENTATION,
            DWMWA_LAST
        }


        [DllExport]
        public static void Initialize(ref IntPtr data, IntPtr rm)
        {

            Measure measure = new Measure();

            //Initialize PrevSysProcess List first, based on reading PValues inside ini files            
            for (int Count = 0; Count < Measure.MaxSize; Count++)
            {
                char[] PrevPValueCharArray = new char[256];
                GetPrivateProfileString("Variables", string.Concat("PValue", Count.ToString()), "Default.png", PrevPValueCharArray, 256, Measure.iniPath);
                string PrevPValue = new string(PrevPValueCharArray);
                PrevPValue = System.Text.RegularExpressions.Regex.Replace(PrevPValue, @"\0", "");

                if (PrevPValue == "WinAppAndUWP.png")
                    continue;

                if (PrevPValue != "Default.png")
                    measure.PrevSize++;

                measure.PreSysProcesses.Add(PrevPValue);
            }

            //Initialize PrevUWPCount
            char[] PrevUWPCount = new char[256];
            GetPrivateProfileString("Variables", "WinAppCount", "0", PrevUWPCount, 256, Measure.iniPath);
            string Value = new string(PrevUWPCount);
            Value = System.Text.RegularExpressions.Regex.Replace(Value, @"\0", "");
            measure.PrevUWPCount = Int32.Parse(Value);



            data = GCHandle.ToIntPtr(GCHandle.Alloc(measure));
        }


        [DllExport]
        public static void Finalize(IntPtr data)
        {
            Measure measure = (Measure)data;

            //Release Data
            if (measure.buffer != IntPtr.Zero)
                Marshal.FreeHGlobal(measure.buffer);
            GCHandle.FromIntPtr(data).Free();
        }


        //delegate
        public delegate bool EnumDelegate(IntPtr hWnd, int lParam);

        [DllExport]
        public static void Reload(IntPtr data, IntPtr rm, ref double maxValue)
        {
            Measure measure = (Measure)data;
            Rainmeter.API api = (Rainmeter.API)rm;

           
            //Filter all background and redundant processes, keeping all application process stored in TraySysProcesses
            // TotalSize       = total no of normal running application after filtering  
            // ProcessesRecord = a temporary record of measure.TraySysProcesses in form of string name
            // UWPRecord       = a temporary record of measure.UWPRecord in form of UWP's Windows Title
            int TotalSize = 0;
            HashSet<string> ProcessesRecord = new HashSet<string>();
            HashSet<string> UWPRecord = new HashSet<string>();

            EnumDelegate filter = delegate (IntPtr hWnd, int lParam)
            {
                //Get Window Title
                StringBuilder strbTitle = new StringBuilder(255);
                int nLength = GetWindowText(hWnd, strbTitle, strbTitle.Capacity + 1);
                string strTitle = strbTitle.ToString();


                if (IsWindowVisible(hWnd))
                {
                    Int64 style = (Int64)GetWindowLongPtr(hWnd, -16); // GWL_STYLE
                    Int64 style2 = (Int64)GetWindowLongPtr(hWnd, -20); // GWL_EXSTYLE

                    //Does not have a WindowCaption
                    if ((style & 0x00C00000L) != 0x00C00000L)
                        return true;

                    //Is Tool Window
                    if ((style2 & 0x00000080L) == 0x00000080L)
                        return true;

                    //Get Target Process
                    uint ProcessID = 0;
                    GetWindowThreadProcessId(hWnd, ref ProcessID);
                    Process Process = Process.GetProcessById((int)ProcessID);

                    //Reaching Limit
                    if (TotalSize >= Measure.MaxSize)
                        return true;

                    int IsApp = measure.IsRunningApplication(Process, measure.TraySysProcesses, TotalSize);

                    //Detect if it is Universal Window App 
                    if ( IsApp == 1 && (Process.ProcessName == "ApplicationFrameHost" || Process.ProcessName=="WWAHost")  )
                    {
                        //Check if it is not a coated window ( important for determining UWP being visible or not )
                        IntPtr attribute = new IntPtr();
                        int result = DwmGetWindowAttribute(hWnd, (int)DWMWINDOWATTRIBUTE.DWMWA_CLOAKED, out attribute, 256);

                        //Actual Running Application, put it inside UWPList which will be dealt with later on 
                        if (attribute.ToInt32() == 0)
                        {
                            UWPRecord.Add(strTitle);
                            measure.UWPList[measure.UWPCount] = new ProcessData(strTitle, "Default.png", "");
                            measure.UWPCount++;
                            return true;
                        }
                    }
                    //Detect if it is windows application
                    else if ( IsApp == 1 && !String.IsNullOrEmpty(strTitle) )
                    {
                        string ProcessNameinPNG = string.Concat(Process.ProcessName, ".png");

                        ProcessesRecord.Add(ProcessNameinPNG);
                        measure.TraySysProcesses[TotalSize] = new ProcessData(strTitle, ProcessNameinPNG, "");
                        try { measure.TraySysProcesses[TotalSize].ExeAddress = Process.MainModule.FileName; } catch { measure.TraySysProcesses[TotalSize].ExeAddress = ""; }
       
                        //Check if a new application is running
                        if (measure.PreSysProcesses.Contains(ProcessNameinPNG) == false)
                        {
                            Measure.IsUpdate = 1;
                        }
                        TotalSize = TotalSize + 1;
                    }
                    // Not a Windows application
                    else { return true; }
                }
                

                return true;
            };
            EnumDesktopWindows(IntPtr.Zero, filter, IntPtr.Zero);


            //Check if any Application(s) being deleted
            if (TotalSize < measure.PrevSize )
                Measure.IsUpdate = 1;


            if (measure.DelayCycle > 0 )
                measure.DelayCycle--;
            if (measure.PrevUWPCount + measure.PrevWinSettingCount != measure.UWPCount + measure.WinSettingCount   && measure.DelayCycle <= 0)
                Measure.IsUpdate = 1;


            //Record current processes list, which can be used to compare next update later
            measure.PreSysProcesses = new HashSet<string>(ProcessesRecord);
            measure.PrevSize = TotalSize;
            if (measure.DelayCycle <= 0)
            {
                measure.PrevUWPList = new HashSet<string>(UWPRecord);
                measure.PrevUWPCount = measure.UWPCount;
                measure.PrevWinSettingCount = measure.WinSettingCount;
            }


            //Ignore Extract Icon process if no need to update
            if (Measure.IsUpdate == 0)
                return;


            //Write UWPCount in ini file, and display it if there is/are UWP(s) running
            measure.SetVariable("WinAppCount", (measure.UWPCount+measure.WinSettingCount).ToString(), rm);
            int start = ( (measure.UWPCount + measure.WinSettingCount) > 0) ? 1 : 0;
            if (start == 1)
                measure.SetVariable("PValue0", "WinAppAndUWP.png", rm);

         
            //Extract Icon from Process and saved it inside icon folder
            for (int pcount = start; pcount < measure.TraySysProcesses.Length; pcount++)
            {
                ProcessData ServiceData = measure.TraySysProcesses[pcount-start];
                string KeyName = string.Concat("PValue", pcount.ToString());

                // reaching empty items
                if (ServiceData.ProcessNamePng == "Default.png" || String.IsNullOrEmpty(ServiceData.ProcessNamePng) )
                {
                    //write to ini file, Set Variable
                    measure.SetVariable(KeyName, "Default.png", rm);
                    continue;
                }

                //write to ini file, Set Variable
                measure.SetVariable(KeyName, ServiceData.ProcessNamePng, rm);

                //icon path
                string path = string.Concat(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString(), "\\Rainmeter\\Skins\\Distiller_Block\\@Resources\\Icon\\", ServiceData.ProcessNamePng);

                //avoid re-saving the same icon picture
                if (File.Exists(path))
                {
                    continue;
                }
                Measure.IsUpdate = 2;


                // Saving the icon picture into @Resource\Icon Folder
                try
                {
                    using (Icon ico = Icon.ExtractAssociatedIcon(ServiceData.ExeAddress))
                    {
                        using (Bitmap btmap = ico.ToBitmap())
                        {
                            btmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                        }
                    }

                }
                catch (Exception e)
                {
                    //failure in getting the icon, use Default.png instead
                    measure.SetVariable(KeyName, "Default.png", rm);
                }

            }


        }

        [DllExport]
        public static double Update(IntPtr data)
        {
            Measure measure = (Measure)data;
            double result = (double)Measure.IsUpdate;
            
            //Flushing Data
            Measure.IsUpdate = 0;
            measure.UWPCount = 0;
            measure.WinSettingCount = 0;
            Array.Clear(measure.TraySysProcesses, 0, measure.TraySysProcesses.Length);
            Array.Clear(measure.UWPList, 0, measure.UWPList.Length);

            return result;
        }

    }
}


