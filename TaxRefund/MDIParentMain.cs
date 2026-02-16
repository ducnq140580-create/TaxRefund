using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Marshal = System.Runtime.InteropServices.Marshal;
using Newtonsoft.Json.Linq; // Add this at the top

namespace TaxRefund
{
    public partial class MDIParentMain : Form
    {

        frmTracuu fTracuu;
        frmChangePassword fChangePassword;
        frmRegister fRegister;
        private bool needForceExit;

        private System.Windows.Forms.Timer _idleTimer;
        private DateTime _lastActivityTime;

        private const int IDLE_TIMEOUT_SECONDS = 30; // Changed to 30 seconds
        private bool _isLoggedIn = true;
        private bool _isLoggingOut = false;
        private bool _warningShown = false;


        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private string loginName;
        private string password;
        private string tencc;
        private bool decentralization;
        // 1. Declare the timer
        System.Windows.Forms.Timer weatherTimer = new System.Windows.Forms.Timer();

        public MDIParentMain()
        {
            InitializeComponent();
            InitializeIdleTimer();
            SetupActivityTracking();
            StartIdleTimer();

            InitializeTrayIcon();

            // Handle form closing event
            this.FormClosing += MDIParentMain_FormClosing;

            SetupWeatherTimer();
        }

        private void SetupWeatherTimer()
        {
            // 2. Set the interval to 1 hour (3,600,000 ms)
            weatherTimer.Interval = 3600000;
            weatherTimer.Tick += new EventHandler(WeatherTimer_Tick);

            // 3. Start the timer
            weatherTimer.Start();

            //// 4. Run it immediately once so the user doesn't wait an hour for the first load
            UpdateWeather();
        }

        private void WeatherTimer_Tick(object sender, EventArgs e)
        {
            UpdateWeather();
        }

        private void UpdateWeather()
        {
            try
            {
                string apiKey = "d4ebf19a79c9d42a81fc4bbf4bb6a73e";
                string url = $"https://api.openweathermap.org/data/2.5/weather?q=ho chi minh city,vn&appid={apiKey}&units=metric";

                using (System.Net.WebClient webClient = new System.Net.WebClient())
                {
                    string json_text = webClient.DownloadString(url);
                    JObject data = JObject.Parse(json_text);

                    // Existing data
                    double temp = (double)data["main"]["temp"];
                    string description = (string)data["weather"][0]["description"];

                    // New data: humidity
                    int humidity = (int)data["main"]["humidity"];

                    // New data: wind speed and direction
                    double windSpeed = (double)data["wind"]["speed"];
                    int? windDeg = (int?)data["wind"]["deg"]; // Can be null in some cases

                    // New data: sunset time (Unix timestamp)
                    long sunsetUnix = (long)data["sys"]["sunset"];

                    // Convert Unix timestamp to DateTime
                    DateTimeOffset sunsetDateTime = DateTimeOffset.FromUnixTimeSeconds(sunsetUnix);

                    // Format sunset time to local time
                    string sunsetTime = sunsetDateTime.LocalDateTime.ToString("hh:mm tt");

                    // Get wind direction if available
                    string windDirection = "N/A";
                    if (windDeg.HasValue)
                    {
                        windDirection = GetWindDirection(windDeg.Value);
                    }

                    // Update UI with all data
                    lblTemperature.Text = temp.ToString("N0") + "°C";
                    lblDescription.Text = char.ToUpper(description[0]) + description.Substring(1);

                    // New UI elements (you need to add these labels to your form)
                    lblHumidity.Text = "Humidity: " + humidity.ToString() + "%";
                    lblWind.Text = $"Wind Speed: {windSpeed:N1} m/s {windDirection}";
                    lblSunset.Text = "Sunset: " + sunsetTime;

                    //MessageBox.Show(json_text);
                }
            }
            catch (Exception ex)
            {
                // Log error so the app doesn't crash if the internet is down
                MessageBox.Show("Weather update failed: " + ex.Message);
            }
        }

        // Helper method to convert wind degrees to compass direction
        private string GetWindDirection(int degrees)
        {
            string[] directions = { "N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
                          "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW" };

            int index = (int)((degrees + 11.25) / 22.5) % 16;
            return directions[index];
        }

        private void MDIParentMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                ShowCloseDialog();
            }         
        }

        private void InitializeTrayIcon()
        {
            // Create context menu first
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Show", null, OnShow);
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("Exit", null, OnExit);

            // Create tray icon
            trayIcon = new NotifyIcon();
            //trayIcon.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); // Use app icon

            // Add a custom icon file to your project (Properties → Resources)
            trayIcon.Icon = Properties.Resources.vat_refund_icon;

            trayIcon.Text = "TRS - Tax Refund System";
            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.Visible = true; // Make visible immediately

            // Add double-click event to show form
            trayIcon.DoubleClick += OnShow;

            // Optional: Add balloon tip notification
            trayIcon.ShowBalloonTip(1000, "TRS", "Application is running in system tray", ToolTipIcon.Info);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
        }

        private void ShowCloseDialog()
        {
            var result = MessageBox.Show(
                "Do you want to minimize TRS to system tray?",
                "TRS Options",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button3
            );

            switch (result)
            {
                case DialogResult.Yes: // Minimize to tray
                    if (trayIcon == null)
                    {
                        InitializeTrayIcon();
                    }
                    MinimizeToTray();
                    break;
                case DialogResult.No: // Close application
                    ReallyClose();
                    break;
                case DialogResult.Cancel: // Keep form open
                    break;
            }
        }

        private void MinimizeToTray()
        {
            // Ensure tray icon exists before using it
            if (trayIcon == null)
            {
                InitializeTrayIcon(); // Initialize if null
            }

            this.Hide();
            this.WindowState = FormWindowState.Minimized;
            trayIcon.Visible = true;
            trayIcon.ShowBalloonTip(1000, "TRS", "Minimized to system tray", ToolTipIcon.Info);
        }

        private void ReallyClose()
        {
            // Clean up tray icon before closing
            if (trayIcon != null)
            {
                trayIcon.Visible = false;
                trayIcon.Dispose();
                trayMenu.Dispose();
            }
            _isLoggedIn = false;
            //trayIcon.Dispose();
            //trayMenu.Dispose();
            Application.Exit();
        }

        private void OnShow(object sender, EventArgs e)
        {
            // Show the form when user clicks "Show" in tray menu or double-clicks icon
            this.Show();
            this.WindowState = FormWindowState.Maximized;
            this.BringToFront();
            trayIcon.Visible = false; // Hide tray icon when form is shown
        }

        // Exit application when user clicks "Exit" in tray menu
        private void OnExit(object sender, EventArgs e)
        {
            ReallyClose();
        }

        private void InitializeIdleTimer()
        {
            _idleTimer = new System.Windows.Forms.Timer();
            _idleTimer.Interval = 1000; // Check every 1 second for more precise timing
            _idleTimer.Tick += IdleTimer_Tick;
        }      

        private void IdleTimer_Tick(object sender, EventArgs e)
        {
            Console.WriteLine($"Timer tick - Logged in: {_isLoggedIn}");
            if (!_isLoggedIn) return;

            TimeSpan idleTime = DateTime.Now - _lastActivityTime;
            Console.WriteLine($"Idle time: {idleTime.TotalSeconds:F0} seconds");

            if (idleTime.TotalSeconds >= IDLE_TIMEOUT_SECONDS)
            {
                Console.WriteLine("Auto logout triggered");
                AutoLogout();
            }
            else if (idleTime.TotalSeconds >= IDLE_TIMEOUT_SECONDS - 10 && !_warningShown)
            {
                // Show warning 10 seconds before logout
                Console.WriteLine("Warning triggered");
                ShowLogoutWarning((int)(IDLE_TIMEOUT_SECONDS - idleTime.TotalSeconds));
            }
        }
        private void ShowLogoutWarning(int secondsLeft)
        {
            if (secondsLeft > 0 && !_warningShown)
            {
                _warningShown = true;
                MessageBox.Show($"Warning: You will be automatically logged out in {secondsLeft} second(s) due to inactivity.",
                       "Inactivity Warning",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Warning);
            }
        }
        private void StartIdleTimer()
        {
            _lastActivityTime = DateTime.Now;
            _idleTimer.Start();
        }

        private void StopIdleTimer()
        {
            _idleTimer.Stop();
        }

      
        private void SetupActivityTracking()
        {
            // FORM-LEVEL EVENTS
            this.MouseMove += TrackUserActivity_Mouse;
            this.KeyDown += TrackUserActivity_Key;
            this.MouseClick += TrackUserActivity_Mouse;
            this.Activated += Form_Activated;

            // RECURSIVELY SETUP ALL CONTROLS
            SetupControlActivityTracking(this);

            // MDI SPECIFIC: Track activation of child forms
            this.MdiChildActivate += MDIParentMain_MdiChildActivate;
        }

        private void MDIParentMain_MdiChildActivate(object sender, EventArgs e)
        {
            TrackUserActivity(); // Reset when child forms are activated
        }

        private void SetupControlActivityTracking(Control parentControl)
        {
            foreach (Control control in parentControl.Controls)
            {
                // Subscribe to events with correct signatures
                control.MouseMove += TrackUserActivity_Mouse;
                control.MouseClick += TrackUserActivity_Mouse;
                control.KeyDown += TrackUserActivity_Key;
                control.Click += TrackUserActivity;

                // Recursively setup for child controls
                if (control.HasChildren)
                {
                    SetupControlActivityTracking(control);
                }
            }
        }

        // Event handler for form activation
        private void Form_Activated(object sender, EventArgs e)
        {
            TrackUserActivity();
        }

        // CORRECTED EVENT HANDLERS WITH PROPER SIGNATURES
        private void TrackUserActivity(object sender, EventArgs e)
        {
            TrackUserActivity();
        }

        private void TrackUserActivity_Mouse(object sender, MouseEventArgs e)
        {
            TrackUserActivity();
        }

        private void TrackUserActivity_Key(object sender, KeyEventArgs e)
        {
            TrackUserActivity();
        }

        // Parameterless version for manual calls
        private void TrackUserActivity()
        {
            Console.WriteLine("Activity detected - resetting timer");
            _lastActivityTime = DateTime.Now;
            _warningShown = false; // Reset warning flag on any activity
        }
       
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            //StopIdleTimer();

            // Properly unsubscribe from all events
            this.MouseMove -= TrackUserActivity_Mouse;
            this.KeyDown -= TrackUserActivity_Key;
            this.MouseClick -= TrackUserActivity_Mouse;
            this.Activated -= Form_Activated;
            this.MdiChildActivate -= MDIParentMain_MdiChildActivate;

            // Recursively unsubscribe from all controls
            UnsubscribeControlEvents(this);

            base.OnFormClosing(e);
        }

        private void UnsubscribeControlEvents(Control parentControl)
        {
            foreach (Control control in parentControl.Controls)
            {
                control.MouseMove -= TrackUserActivity_Mouse;
                control.MouseClick -= TrackUserActivity_Mouse;
                control.KeyDown -= TrackUserActivity_Key;
                control.Click -= TrackUserActivity;

                if (control.HasChildren)
                {
                    UnsubscribeControlEvents(control);
                }
            }
        }

        private void DisposeAllResources()
        {
            // Force process termination if needed
            if (needForceExit)
            {
                Process.GetCurrentProcess().Kill();
            }
            Application.Exit();
        }        
       
        public MDIParentMain(string loginName, string password)
        {
            this.loginName = loginName;
            this.password = password;
        }

        public MDIParentMain(string loginName, string tencc, string password, bool decentralization) : this(loginName, password)
        {
            this.tencc = tencc;
            this.decentralization = decentralization;
        }


        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //SaveFileDialog saveFileDialog = new SaveFileDialog();
            //saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            //saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            //if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            //{
            //    string FileName = saveFileDialog.FileName;
            //}
        }
        private void releaseObject(ref object obj) // note ref!
        {
            // Do not catch an exception from this.
            // You may want to remove these guards depending on
            // what you think the semantics should be.
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
            // Since passed "by ref" this assingment will be useful
            // (It was not useful in the original, and neither was the
            //  GC.Collect.)
            obj = null;
        }
        private void ExitToolsStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult tb = MessageBox.Show("Are you sure to quit the Tax Refund System (TRS)?", "Notice", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                if (tb == DialogResult.OK)
                {
                    DisposeAllResources();
                    _isLoggedIn = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void CutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form activeChildForm = this.ActiveMdiChild;
            if (activeChildForm != null)
            {
                TextBox RichtxtEditor = activeChildForm.ActiveControl as TextBox;
                if (RichtxtEditor != null)
                {
                    if (RichtxtEditor.SelectionLength > 0)
                        RichtxtEditor.Cut();
                }
            }
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form activeChildForm = this.ActiveMdiChild;
            if (activeChildForm != null)
            {
                TextBox RichtxtEditor = activeChildForm.ActiveControl as TextBox;
                if (RichtxtEditor != null)
                {
                    if (RichtxtEditor.SelectionLength > 0)
                        RichtxtEditor.Copy();
                }
            }
        }

        private void PasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form activeChildForm = this.ActiveMdiChild;
            if (activeChildForm != null)
            {
                TextBox RichtxtEditor = activeChildForm.ActiveControl as TextBox;
                if (RichtxtEditor != null)
                {
                    RichtxtEditor.Paste();
                }
            }
        }

        private void ToolBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStrip.Visible = toolBarToolStripMenuItem.Checked;
        }

        private void StatusBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            statusStrip.Visible = statusBarToolStripMenuItem.Checked;
        }

        private void CascadeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.Cascade);
        }

        private void TileVerticalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileVertical);
        }

        private void TileHorizontalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void ArrangeIconsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.ArrangeIcons);
        }

        private void CloseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form childForm in MdiChildren)
            {
                childForm.Close();
            }
        }

        private void MDIParentMain_Load(object sender, EventArgs e)
        {
            Utility ut = new Utility();
            var conn = ut.OpenDB();

            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();

                    IPHostEntry host;
                    host = Dns.GetHostEntry(Dns.GetHostName());

                    foreach (IPAddress ip in host.AddressList)
                    {
                        if (ip.AddressFamily == AddressFamily.InterNetwork)
                        {
                            lblIPaddress.Text = ip.ToString();
                        }
                    }
                    if (lblstatus.Text != "True")
                    {
                        registerToolStripMenuItem.Enabled = false;
                    }
                }
                conn.Close();
                conn.Dispose();
             
                UpdateWeather(); 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private bool CheckOpened(string name)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm.Text == name)
                {
                    return true;
                }
            }
            return false;
        }

        private void loginToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                //// Close the current Main Form
                //this.Close();

                frmLogin fLogin = new frmLogin
                {
                    StartPosition = FormStartPosition.CenterParent,
                    Dock = DockStyle.Fill,
                };
                //fLogin.StartPosition = FormStartPosition.CenterParent;
                //fLogin.Dock = DockStyle.Fill;
                fLogin.Show();
                fLogin.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void registerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                fRegister = new frmRegister();

                if (CheckOpened(fRegister.Text) == true)
                {
                    fRegister.MdiParent = this;
                    fRegister.FormClosed += new FormClosedEventHandler(FRegister_FormClosed);
                    fRegister.StartPosition = FormStartPosition.CenterScreen;
                    //fRegister.Dock = DockStyle.Fill;
                    fRegister.Show();
                    //fRegister.WindowState = FormWindowState.Normal;
                    fRegister.BringToFront();
                }
                else
                {
                    MessageBox.Show("The request window is currently opened. Break Ctrl + F6 to change the windows.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    fRegister.Activate();
                    fRegister.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void FRegister_FormClosed(object sender, FormClosedEventArgs e)
        {
            //throw new NotImplementedException();
        }

        private void changePasswordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fChangePassword = new frmChangePassword();

            if (CheckOpened(fChangePassword.Text) == true)
            {
                fChangePassword.MdiParent = this;
                fChangePassword.FormClosed += new FormClosedEventHandler(FChangePassword_FormClosed);
                //fChangePassword.Dock = DockStyle.None;                
                fChangePassword.StartPosition = FormStartPosition.CenterScreen;
                fChangePassword.Show();
                fChangePassword.WindowState = FormWindowState.Normal;
                fChangePassword.BringToFront();
            }
            else
            {
                MessageBox.Show("Cửa sổ bạn chọn hiện đang mở. Nhấn Ctrl + F6 để chuyển màn hình nghiệp vụ.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                fChangePassword.Activate();
                fChangePassword.Focus();
            }
        }

        private void FChangePassword_FormClosed(object sender, FormClosedEventArgs e)
        {
            //fChangePassword = null;
            //throw new NotImplementedException();
        }

        private void searchTaxRefundInfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fTracuu = new frmTracuu();

            if (!CheckOpened(fTracuu.Text) == true)
            {
                fTracuu.MdiParent = this;
                fTracuu.Dock = DockStyle.None;
                fTracuu.Show();
                fTracuu.StartPosition = FormStartPosition.CenterParent;
                fTracuu.WindowState = FormWindowState.Maximized;
                fTracuu.BringToFront();
            }
            else
            {
                MessageBox.Show("Cửa sổ bạn chọn hiện đang mở. Nhấn Ctrl + F6 để chuyển màn hình nghiệp vụ.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                fTracuu.Activate();
                //this.fTracuuTK.WindowState = FormWindowState.Maximized;
                //this.fTracuuTK.Focus();
                //this.fTracuuTK.BringToFront();
            }
        }

        private void helpToolStripButton_Click(object sender, EventArgs e)
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TRS Guidelines.txt");

            try
            {
                Process.Start("notepad.exe", filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private void lblTenCC_Click(object sender, EventArgs e)
        {
            // Ensure the menu exists before trying to show it
            if (lblTenCC.ContextMenuStrip != null)
            {
                // Show the menu directly under the label
                lblTenCC.ContextMenuStrip.Show(lblTenCC, new Point(0, lblTenCC.Height));
            }
        }

        // Add this method to create a proper Logout function
        private void Logout()
        {
            try
            {
                // 1. Ask for confirmation
                DialogResult confirm = MessageBox.Show(
                    "Are you sure you want to log out?",
                    "Confirm Logout",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (confirm != DialogResult.Yes)
                    return;

                // 2. Set logging out flag to prevent close dialog from showing
                _isLoggingOut = true;

                // 3. Stop all timers

                //StopIdleTimer();
                if (weatherTimer != null)
                {
                    weatherTimer.Stop();
                    weatherTimer.Dispose();
                }

                // 4. Update login status
                _isLoggedIn = false;

                // 5. Clean up tray icon
                if (trayIcon != null)
                {
                    trayIcon.Visible = false;
                    trayIcon.Dispose();
                    trayIcon = null;
                }

                if (trayMenu != null)
                {
                    trayMenu.Dispose();
                    trayMenu = null;
                }

                // 6. Close all child forms
                foreach (Form childForm in MdiChildren)
                {
                    childForm.Close();
                }

                // 7. Show login form
                frmLogin loginForm = new frmLogin();
                loginForm.StartPosition = FormStartPosition.CenterScreen;

                // 8. Hide current form
                this.Hide();

                // 9. Show login form
                loginForm.ShowDialog();

                // 10. If login was successful, the login form should open a new main form
                // If we get here and login was cancelled, close the application
                if (!_isLoggedIn)
                {
                    Application.Exit();
                }
                else
                {
                    // If for some reason we're still logged in, close this form
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during logout: {ex.Message}", "Logout Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        // Alternative method that restarts the application (similar to your current approach)
        private void LogoutWithRestart()
        {
            try
            {
                // Ask for confirmation
                DialogResult confirm = MessageBox.Show(
                    "Are you sure you want to log out?",
                    "Confirm Logout",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (confirm != DialogResult.Yes)
                    return;

                // Set logging out flag
                _isLoggingOut = true;

                // Stop timers
                StopIdleTimer();
                if (weatherTimer != null)
                {
                    weatherTimer.Stop();
                }

                // Update status
                _isLoggedIn = false;

                // Clean up tray icon
                if (trayIcon != null)
                {
                    trayIcon.Visible = false;
                    trayIcon.Dispose();
                }

                // Restart application
                Application.Restart();

                // Close current instance
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during logout: {ex.Message}", "Logout Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        // Add this method to handle logout from menu item
        private void logoutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Logout(); // or LogoutWithRestart() depending on your preference
        }

        // Update your constructor to add the logout menu item
        public MDIParentMain(string loginname, string tencc, string matkhau, string role)
        {
            InitializeComponent();

            lblLoginName.Text = loginname;
            lblTenCC.Text = "You are logged in as: " + tencc;
            lblUserName.Text = tencc;
            lblLoginPassword.Text = matkhau;
            lblstatus.Text = role;

            // 1. Create the ContextMenuStrip
            ContextMenuStrip dropdownMenu = new ContextMenuStrip();

            // 2. Add items to the menu
            dropdownMenu.Items.Add("Change Password", null, (s, e) =>
            {
                fChangePassword = new frmChangePassword();

                if (CheckOpened(fChangePassword.Text))
                {
                    fChangePassword.MdiParent = this;
                    fChangePassword.FormClosed += new FormClosedEventHandler(FChangePassword_FormClosed);
                    fChangePassword.StartPosition = FormStartPosition.CenterScreen;
                    fChangePassword.Show();
                    fChangePassword.WindowState = FormWindowState.Normal;
                    fChangePassword.BringToFront();
                }
                else
                {
                    MessageBox.Show("Cửa sổ bạn chọn hiện đang mở. Nhấn Ctrl + F6 để chuyển màn hình nghiệp vụ.",
                        "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    fChangePassword.Activate();
                    fChangePassword.Focus();
                }
            });

            dropdownMenu.Items.Add("-"); // Separator

            dropdownMenu.Items.Add("Logout", null, (s, e) =>
            {
                Logout(); // Call the logout function
            });

            // 3. Attach it to the label
            lblTenCC.ContextMenuStrip = dropdownMenu;

            //// Add logout to main menu if needed
            //AddLogoutToMainMenu();
        }

        // Method to add logout to the main menu
        private void AddLogoutToMainMenu()
        {
            // Check if logout menu item already exists
            bool logoutExists = false;
            foreach (ToolStripItem item in menuStrip.Items)
            {
                if (item.Text == "Logout")
                {
                    logoutExists = true;
                    break;
                }
            }

            // Add logout to file menu if it doesn't exist
            if (!logoutExists)
            {
                // Find the File menu or create one
                ToolStripMenuItem fileMenu = null;
                foreach (ToolStripItem item in menuStrip.Items)
                {
                    if (item.Text.Contains("File") || item.Text.Contains("Tệp"))
                    {
                        fileMenu = item as ToolStripMenuItem;
                        break;
                    }
                }

                if (fileMenu == null)
                {
                    fileMenu = new ToolStripMenuItem("File");
                    menuStrip.Items.Insert(0, fileMenu);
                }

                // Add separator if needed
                if (fileMenu.DropDownItems.Count > 0)
                {
                    fileMenu.DropDownItems.Add(new ToolStripSeparator());
                }

                // Add logout item
                ToolStripMenuItem logoutMenuItem = new ToolStripMenuItem("Logout");
                //logoutMenuItem.Image = Properties.Resources.logout_icon; // Add an icon resource if available
                logoutMenuItem.ShortcutKeys = Keys.Control | Keys.L;
                logoutMenuItem.Click += logoutToolStripMenuItem_Click;
                fileMenu.DropDownItems.Add(logoutMenuItem);
            }
        }

        // Update the PerformRestart method to use the new Logout function
        private void PerformRestart()
        {
            LogoutWithRestart();
        }

        // Update the ExecuteLogout method to use the new Logout function
        private void ExecuteLogout()
        {
            LogoutWithRestart();
        }

        // Add a public method to allow external logout calls
        public void ForceLogout(string reason = "")
        {
            if (!string.IsNullOrEmpty(reason))
            {
                MessageBox.Show(reason, "Session Ended", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            LogoutWithRestart();
        }

        // Update the AutoLogout method to use ForceLogout
        private void AutoLogout()
        {
            StopIdleTimer();
            _isLoggedIn = false;

            DialogResult result = MessageBox.Show(
                "You have been idle for too long. You will be logged out for security reasons.",
                "Auto Logout",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );

            ForceLogout("Session terminated due to inactivity.");

        }
    }
}
