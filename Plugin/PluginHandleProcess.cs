using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rainmeter;

using System.Diagnostics;
using System.Management;
using System.Drawing;
using System.IO;



// Overview: This is a blank canvas on which to build your plugin.

// Note: GetString, ExecuteBang and an unnamed function for use as a section variable
// have been commented out. If you need GetString, ExecuteBang, and/or section variables 
// and you have read what they are used for from the SDK docs, uncomment the function(s)
// and/or add a function name to use for the section variable function(s). 
// Otherwise leave them commented out (or get rid of them)!

namespace PluginHandleProcess
{
    
    class Measure
    {
        static public implicit operator Measure(IntPtr data)
        {
            TraySysProcesses = new Process[MaxSize];
            return (Measure)GCHandle.FromIntPtr(data).Target;
        }
        public IntPtr buffer = IntPtr.Zero;
        

        public const int MaxSize = 15;
        public static Process[] TraySysProcesses;

        
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

            //Special Case: System Settings
            if (name == "SystemSettings")
                return true;

            //ignore specific running process:
            //ApplicationFrameHost: defines Frameworks of Microsoft Apps, doesn't really have a visible windows of its own but still considered as graphical interace
            //Rainmeter:            No need to self-check itself
            //MicrosoftEdgeCP:      Running along side with Microsoft Edge, redundancy
            if (name == "ApplicationFrameHost" || name == "Rainmeter" || name == "MicrosoftEdgeCP")
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
          
            /*
            //check redundancy, e.g.  multiple MicrosoftEdgeCP 
            foreach (Process element in list)
            {
                if (element == null)
                    break;

                string listname = element.ToString();
                if (listname.Contains(name))
                    return false;
            }
            */

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
            data = GCHandle.ToIntPtr(GCHandle.Alloc(new Measure()));
            Rainmeter.API api = (Rainmeter.API)rm;
            
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


            //Collect an array of PValues in ini files before update
            //Also Calculate no of total application running before the cycle
            string[] PrevPValues = new string[Measure.MaxSize];
            int totalPrevProcesses = 0;

            for ( int Count = 0; Count < Measure.MaxSize; Count++ )
            {
                char[] PrevPValueCharArray = new char[64];
                GetPrivateProfileString("Variables", string.Concat("PValue", Count), "Default.png", PrevPValueCharArray, 64, Measure.iniPath);
                string PrevPValue = new String(PrevPValueCharArray);
                PrevPValue = System.Text.RegularExpressions.Regex.Replace(PrevPValue, @"\0", "");

                if (PrevPValue != "Default.png")
                    totalPrevProcesses++;

                PrevPValues[Count] = PrevPValue;
            }
            

            //Collect All processes and filter all background and redundant processes, keeping all application process stored in TraySysProcesses
            // size = total no of application processes ( after filtering )
            int TotalSize = 0;
            Process[] AllSysProcesses = Process.GetProcesses();

            
            foreach (Process service in AllSysProcesses)
            {
                if (measure.IsRunningApplication(service, Measure.TraySysProcesses) && TotalSize < Measure.MaxSize)
                {
                    Measure.TraySysProcesses[TotalSize] = service;

                    // Compare Previous PValue with Current One
                    if ( Measure.IsUpdate == 0 )
                    {
                        //upcoming (Current) PValue
                        string CurPValue = string.Concat(measure.ExtractName(Measure.TraySysProcesses[TotalSize].ToString()), ".png");

                        // A New application is running
                        if ( Array.IndexOf( PrevPValues ,CurPValue) == -1 ) 
                        {
                            Measure.IsUpdate = 1;
                        }
                    }                    
                    TotalSize = TotalSize + 1;
                }
            }

            //Chekck if Application(s) being deleted
            if (TotalSize < totalPrevProcesses)
            {
                Measure.IsUpdate = 1;
            }

            //Ignore Extract Icon process if no need to update
            if (Measure.IsUpdate == 0)
                return;

            //Extract Icon from Process and saved it inside icon folder
            for ( int pcount = 0; pcount < Measure.TraySysProcesses.Length; pcount++ )
            {
                Process service = Measure.TraySysProcesses[pcount];

                //ini file setting
                string KeyName = string.Concat("PValue", pcount.ToString());
                

                // reaching empty items
                if (service == null)
                {
                    WritePrivateProfileString("Variables", KeyName, "Default.png", Measure.iniPath);
                    continue;
                }

                string name = measure.ExtractName(service.ToString());
                string value = string.Concat(name, ".png");
                //write to ini file
                WritePrivateProfileString("Variables", KeyName, value, Measure.iniPath );

                //icon path
                string path = string.Concat(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString(), "\\Rainmeter\\Skins\\Distiller_Block\\@Resources\\Icon\\", name, ".png");

                //avoid re-saving the same icon picture
                if (File.Exists(path))   
                {
                    Console.WriteLine("{0} is already existed!", name);
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

