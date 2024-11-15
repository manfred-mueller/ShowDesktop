using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.Concurrent;
using System.Windows.Forms;
using Version = System.Version;
using System.Runtime.InteropServices.ComTypes;


namespace ShowDesktop
{
    static class Program
    {

        // Import AttachConsole and FreeConsole from Kernel32.dll
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AttachConsole(int dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool FreeConsole();

        // Import WriteConsoleInput to simulate input
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool WriteConsoleInput(IntPtr hConsoleInput, [In] INPUT_RECORD[] lpBuffer, uint nLength, out uint lpNumberOfEventsWritten);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GetStdHandle(int nStdHandle);

        const int STD_INPUT_HANDLE = -10;
        const int ATTACH_PARENT_PROCESS = -1;

        // Define INPUT_RECORD struct for input events
        struct INPUT_RECORD
        {
            public ushort EventType;
            public KEY_EVENT_RECORD KeyEvent;
        }

        // Define KEY_EVENT_RECORD struct
        struct KEY_EVENT_RECORD
        {
            public bool bKeyDown;
            public ushort wVirtualKeyCode;
            public ushort wVirtualScanCode;
            public char uChar;
        }

        [STAThread]
        static void Main()
        {
            // Get the application name from the assembly info
            string appName = Assembly.GetExecutingAssembly().GetName().Name;

            // Define the path to the user's TaskBar Quick Launch folder
            string taskBarFolderPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                @"Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar");

            // Define the full path to the expected shortcut file (appName + ".lnk")
            string shortcutPath = Path.Combine(taskBarFolderPath, appName + ".lnk");
            string appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);

            // Versioning information for display
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            string appVersion = $"{version.Major}.{version.Minor}.{version.Build}";
            string helpStringCmd = 
                   String.Format("{3}{0} {1}, {2} {3}{3}", Assembly.GetExecutingAssembly().GetName().Name, appVersion, ShowDesktop.Properties.Resources.Copyright, Environment.NewLine) +
                   String.Format(ShowDesktop.Properties.Resources.HelpTextCmd, appPath, Environment.NewLine, System.Diagnostics.Process.GetCurrentProcess().ProcessName) + "\n";
            string helpStringGui = 
                   String.Format("{0} {1}{2}{2}", Assembly.GetExecutingAssembly().GetName().Name, appVersion, Environment.NewLine) +
                   String.Format(ShowDesktop.Properties.Resources.HelpTextGui, Environment.NewLine, System.Diagnostics.Process.GetCurrentProcess().ProcessName) + "\n";

            // Check if the shortcut exists
            if (!File.Exists(shortcutPath))
            {
                // Attach to parent console (e.g., command line if launched from there)
                AttachConsole(ATTACH_PARENT_PROCESS);
                if (IsParentProcessShell())
                {
                    Console.WriteLine(helpStringCmd);
                    // Emulate Enter key press to simulate user pressing "Enter"
                    SendKeys.SendWait("{ENTER}");
                    FreeConsole();
                    // Exit the application cleanly
                    Application.Exit();
                }
                else
                {
                    MessageBox.Show(helpStringGui, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Process.Start(appPath);
                    string filePath = Path.Combine(appPath, "nircmd.exe");
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                    Application.Exit();
                }
            }
            else
            {
                Type typeShell = Type.GetTypeFromProgID("Shell.Application");
                var objShell = Activator.CreateInstance(typeShell);
                typeShell.InvokeMember("ToggleDesktop", System.Reflection.BindingFlags.InvokeMethod, null, objShell, null);
            }
        }
        static bool IsParentProcessShell()
        {
            try
            {
                // Hole den Parent-Prozess des aktuellen Prozesses
                Process parentProcess = GetParentProcess(Process.GetCurrentProcess().Id);
                return parentProcess != null && (parentProcess.ProcessName.Equals("cmd", StringComparison.OrdinalIgnoreCase) || parentProcess.ProcessName.Equals("powershell", StringComparison.OrdinalIgnoreCase));
            }
            catch
            {
                return false;
            }
        }

        static Process GetParentProcess(int processId)
        {
            var process = new Process();
            var pbi = new PROCESS_BASIC_INFORMATION();
            int returnLength;

            var handle = Process.GetProcessById(processId).Handle;

            // Hole Informationen über den Parent-Prozess
            if (NtQueryInformationProcess(handle, 0, ref pbi, Marshal.SizeOf(pbi), out returnLength) == 0)
            {
                try
                {
                    return Process.GetProcessById(pbi.InheritedFromUniqueProcessId.ToInt32());
                }
                catch
                {
                    return null;
                }
            }
            return null;
        }

        // Externe Windows-API-Funktion
        [DllImport("ntdll.dll")]
        private static extern int NtQueryInformationProcess(IntPtr hProcess, int processInformationClass, ref PROCESS_BASIC_INFORMATION processInformation, int processInformationLength, out int returnLength);

        [StructLayout(LayoutKind.Sequential)]
        private struct PROCESS_BASIC_INFORMATION
        {
            public IntPtr Reserved1;
            public IntPtr PebBaseAddress;
            public IntPtr Reserved2_0;
            public IntPtr Reserved2_1;
            public IntPtr UniqueProcessId;
            public IntPtr InheritedFromUniqueProcessId;
        }
    }
}