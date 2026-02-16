using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TaxRefund
{
    // SessionManager.cs
    public static class SessionManager
    {
        private static System.Windows.Forms.Timer _idleTimer;
        private static DateTime _lastActivityTime;
        private static Form _mainForm;

        public static void StartSession(Form mainForm)
        {
            _mainForm = mainForm;
            _lastActivityTime = DateTime.Now;

            _idleTimer = new System.Windows.Forms.Timer();
            _idleTimer.Interval = 30000;
            _idleTimer.Tick += IdleTimer_Tick;
            _idleTimer.Start();

            // Hook into application events
            Application.AddMessageFilter(new ActivityMessageFilter());
        }

        private static void IdleTimer_Tick(object sender, EventArgs e)
        {
            if ((DateTime.Now - _lastActivityTime).TotalMinutes >= 5)
            {
                ForceLogout();
            }
        }

        public static void ForceLogout()
        {
            _idleTimer?.Stop();
            _mainForm?.Invoke((MethodInvoker)delegate {
                MessageBox.Show("Session expired due to inactivity");
                _mainForm.Close(); // This will close the application
            });
        }

        public static void RecordActivity()
        {
            _lastActivityTime = DateTime.Now;
        }
    }

    // Message filter to detect user activity
    public class ActivityMessageFilter : IMessageFilter
    {
        public bool PreFilterMessage(ref Message m)
        {
            // Detect mouse and keyboard messages
            if (m.Msg >= 0x100 && m.Msg <= 0x209) // WM_KEYFIRST to WM_KEYLAST
            {
                SessionManager.RecordActivity();
            }
            return false;
        }
    }
}
