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


//This is the source code for a custom dll being used by the rainmeter skin Distiller Block
//It tracks current running application in the operation system and convert icons into png file for display purpose
//Author: Innonion

namespace PluginHandleProcess
{

    class Measure
    {


        static public implicit operator Measure(IntPtr data)
        {
            return (Measure)GCHandle.FromIntPtr(data).Target;
        }
        public IntPtr buffer = IntPtr.Zero;

        //Max Size of the program
        public const int MaxSize = 15;


        //Previous string list of running applications
        public string[] PrevSysProcesses = new string[MaxSize];


        //Actual total number of previous running applications
        public int PrevSize = 0;

        //Current list of running applications
        public Process[] TraySysProcesses = new Process[MaxSize];

        //Determines if a change in application processes are made
        public static int IsUpdate = 0;

        //iniPath
        public static string iniPath = string.Concat(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString(), "\\Rainmeter\\Skins\\Distiller_Block\\Display\\Display.ini");



        internal string ExtractName(string service_name)
        {
            string name = service_name;

            //Wild Guess: service_name ="System.Diagnostics.Process (XXXXXXXXXXXX)"
            if (service_name[27] == '(')
            {
                name = name.Substring(28);
                return name.Remove(name.Length - 1);
            }
            else
            {
                return name;
            }

        }


        internal bool IsRunningApplication(Process p, Process[] list)
        {
            //name of the process
            string name = ExtractName(p.ToString());


            //Exclude specific running process:
            //ApplicationFrameHost: defines Frameworks of Microsoft Apps, doesn't really have a visible windows of its own but still considered as graphical interace
            //Rainmeter:            No need to self-check itself
            //MicrosoftEdgeCP:      Running along side with Microsoft Edge CP, redundancy, also being mis-identified as running process
            //SystemSettings:       Being mis-identified as running application even when it is not opened ( probably running in the background )
            //explorer:             Being mis-identified as running application even when it is not opened ( probably running in the background )
            if (name == "ApplicationFrameHost" || name == "Rainmeter" || name == "MicrosoftEdge" || name == "MicrosoftEdgeCP" ||
                name == "SystemSettings" || name == "explorer")
                return false;


            //check if it has a graphical interface
            if (p.MainWindowHandle == IntPtr.Zero)
                return false;


            //Exclude Windows Default Apps and Special Manager Program(e.g. TaskManager) that cannot get the filename 
            try
            {
                if (p.MainModule.FileName.Contains("WindowsApps"))
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                return false;
            }

            return true;
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

                if (PrevPValue != "Default.png")
                    measure.PrevSize++;

                measure.PrevSysProcesses[Count] = PrevPValue;
            }



            data = GCHandle.ToIntPtr(GCHandle.Alloc(measure));
        }


        [DllExport]
        public static void Finalize(IntPtr data)
        {
            Measure measure = (Measure)data;
            if (measure.buffer != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(measure.buffer);
            }
            GCHandle.FromIntPtr(data).Free();
        }


        [DllExport]
        public static void Reload(IntPtr data, IntPtr rm, ref double maxValue)
        {
            Measure measure = (Measure)data;
            Rainmeter.API api = (Rainmeter.API)rm;

            //Collect All processes
            Process[] AllSysProcesses = Process.GetProcesses();


            //Filter all background and redundant processes, keeping all application process stored in TraySysProcesses
            // TotalSize       = total no of running application after filtering  
            // ProcessesRecord = a temporary record of measure.TraySysProcesses in form of string name
            int TotalSize = 0;
            string[] ProcessesRecord = new string[Measure.MaxSize];

            foreach (Process service in AllSysProcesses)
            {
                if (measure.IsRunningApplication(service, measure.TraySysProcesses) && TotalSize < Measure.MaxSize)
                {
                    //Add the application to the processes list
                    measure.TraySysProcesses[TotalSize] = service;

                    string PValueName = string.Concat(measure.ExtractName(service.ToString()), ".png");
                    ProcessesRecord[TotalSize] = PValueName;

                    // Decide whether we need to update if there is a change ( only one change discovery is needed )
                    if (Measure.IsUpdate == 0)
                    {
                        //Check if a new application is running
                        if (Array.IndexOf(measure.PrevSysProcesses, PValueName) == -1)
                        {
                            Measure.IsUpdate = 1;
                        }
                    }

                    TotalSize = TotalSize + 1;
                }
            }

            //Chekck if Application(s) being deleted
            if (TotalSize < measure.PrevSize)
                Measure.IsUpdate = 1;


            //Record current processes list, which can be used to compare next update later
            Array.Copy(ProcessesRecord, measure.PrevSysProcesses, Measure.MaxSize);
            measure.PrevSize = TotalSize;



            //Ignore Extract Icon process if no need to update
            if (Measure.IsUpdate == 0)
                return;

            //Extract Icon from Process and saved it inside icon folder
            for (int pcount = 0; pcount < measure.TraySysProcesses.Length; pcount++)
            {
                Process service = measure.TraySysProcesses[pcount];

                //ini file setting
                string KeyName = string.Concat("PValue", pcount.ToString());


                // reaching empty items
                if (ProcessesRecord[pcount] == null)
                {
                    WritePrivateProfileString("Variables", KeyName, "Default.png", Measure.iniPath);
                    continue;
                }

                string PValueName = string.Concat(measure.ExtractName(service.ToString()), ".png");
                //write to ini file
                WritePrivateProfileString("Variables", KeyName, PValueName, Measure.iniPath);

                //icon path
                string path = string.Concat(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString(), "\\Rainmeter\\Skins\\Distiller_Block\\@Resources\\Icon\\", PValueName);

                //avoid re-saving the same icon picture
                if (File.Exists(path))
                {
                    continue;
                }

                // Saving the icon picture into @Resource\Icon Folder
                try
                {
                    using (Icon ico = Icon.ExtractAssociatedIcon(service.MainModule.FileName))
                    {
                        using (Bitmap btmap = ico.ToBitmap())
                        {
                            btmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                        }
                    }

                }
                catch (Exception e)
                { }

            }


        }

        [DllExport]
        public static double Update(IntPtr data)
        {
            Measure measure = (Measure)data;
            double result = (double)Measure.IsUpdate;
            Measure.IsUpdate = 0;

            return result;
        }

    }
}


