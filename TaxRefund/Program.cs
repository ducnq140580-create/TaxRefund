using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TaxRefund
{
    static class Program
    {   
        private static Mutex mutex = null;

        [STAThread]
        static void Main()
        {
            const string appMutex = "Global\\TaxRefundSystemTRS";
            bool createdNew;

            mutex = new Mutex(true, appMutex, out createdNew);

            if (!createdNew)
            {
                // Application is already running
                MessageBox.Show("Tax Refund System is already running!\n\nPlease check your system tray or taskbar.",
                    "TRS - Already Running",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // Optional: Bring existing instance to front
                BringExistingInstanceToFront();

                return;
            }

            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                //// Test codes
                //Application.Run(new frmTracuu());

                bool keepRunning = true;

                while (keepRunning)
                {
                    // Show login form and handle login process
                    using (frmLogin loginForm = new frmLogin())
                    {
                        if (loginForm.ShowDialog() == DialogResult.OK)
                        {
                            // Use the main form instance created by login form
                            Application.Run(loginForm.MainForm);

                            // If we reach here, MDIParentMain has closed. 
                            // If it closed because of Application.Restart(), the 'while' loop starts over.
                            // If it closed because the user hit 'X' or 'Exit', we stop.
                            keepRunning = false;
                        }
                        else
                        {
                            // Login failed or cancelled
                            Application.Exit();

                            // User cancelled login or hit 'X' on login form
                            keepRunning = false;
                        }
                    }

                    // Keep mutex alive until application exits
                    GC.KeepAlive(mutex);

                }
            }
            finally
            {
                mutex?.ReleaseMutex();
            }           
        }
    

    private static void BringExistingInstanceToFront()
        {
            Process current = Process.GetCurrentProcess();
            Process[] processes = Process.GetProcessesByName(current.ProcessName);

            foreach (Process process in processes)
            {
                if (process.Id != current.Id)
                {
                    IntPtr handle = process.MainWindowHandle;
                    if (handle != IntPtr.Zero)
                    {
                        // Restore if minimized
                        if (NativeMethods.IsIconic(handle))
                        {
                            NativeMethods.ShowWindow(handle, NativeMethods.SW_RESTORE);
                        }

                        // Bring to front
                        NativeMethods.SetForegroundWindow(handle);
                    }
                    break;
                }
            }
        }
    }

    // Native methods helper class
    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        internal static extern bool IsIconic(IntPtr hWnd);

        internal const int SW_RESTORE = 9;
        internal const int SW_SHOW = 5;
    }
}
   

