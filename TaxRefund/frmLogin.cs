using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace TaxRefund
{
    public partial class frmLogin : Form
    {
        SqlCommand cmd;
        SqlDataAdapter sda;
        DataTable dt;
        Utility ut;
        private System.Windows.Forms.ToolTip toolTip1;  // Declare the ToolTip variable

        // Add these class-level variables
        private CancellationTokenSource _autoCompleteCts;
        private DateTime _lastTypingTime;

        private int loginAttempts = 0;
        private const int MAX_LOGIN_ATTEMPTS = 3;
        private int lockoutSeconds = 10;
        private System.Windows.Forms.Timer lockoutTimer;
        private System.Windows.Forms.Label lblLockoutMessage;

        public frmLogin()
        {
            InitializeComponent();
            InitializeLockoutControls();

            this.FormBorderStyle = FormBorderStyle.None;
            this.Paint += StylishForm_Paint;
            this.Resize += (s, e) => this.Invalidate(); // Redraw on resize
            ApplyRoundedCorners();

            toolTip1 = new System.Windows.Forms.ToolTip()
            {
                // Set default properties
                AutoPopDelay = 2000,     // How long tooltip stays visible
                InitialDelay = 100,      // Time before tooltip appears
                ReshowDelay = 500,       // Time between tooltips when moving mouse
                ShowAlways = true,        // Show even when form isn't active

                // Optional visual customization
                ToolTipTitle = "Hint",
                ToolTipIcon = ToolTipIcon.Info,
                IsBalloon = true
            };

            // Set tooltip for your PictureBox
            toolTip1.SetToolTip(pictureBoxToggleShowPassword, "Click to show/hide password");
            toolTip1.SetToolTip(btnLogin, "Click to login the TRS");

            // Set textbox properties
            txtUsername.Multiline = true;
            txtUsername.MinimumSize = new Size(0, txtUsername.Height);

            txtPassword.Multiline = true;
            txtPassword.MinimumSize = new Size(0, txtPassword.Height);

            txtPassword.UseSystemPasswordChar = true;  // Uses the system's default password char
            txtPassword.PasswordChar = '*';

            // Initialize password field with placeholder
            txtPassword.Text = "Password";
            txtPassword.ForeColor = Color.Gray;
            txtPassword.UseSystemPasswordChar = false; // Important for placeholder visibility
            txtPassword.Tag = "placeholder"; // Mark as placeholder

            // Initialize username field with placeholder
            txtUsername.Text = "Username";
            txtUsername.ForeColor = Color.Gray;
            txtUsername.Tag = "placeholder"; // Mark as placeholder

            // Add event handlers
            txtUsername.Enter += Username_Enter;
            txtUsername.Leave += Username_Leave;
            txtPassword.Enter += Password_Enter;
            txtPassword.Leave += Password_Leave;           

            pictureBoxToggleShowPassword.Image = Properties.Resources.icons8_eye_opened_30_1;
            pictureBoxToggleShowPassword.Cursor = Cursors.Hand;
            pictureBoxToggleShowPassword.BackColor = Color.Transparent;
            pictureBoxToggleShowPassword.SizeMode = PictureBoxSizeMode.StretchImage;

            // Position toggle eye open/ eye closed inside textbox
            pictureBoxToggleShowPassword.Location = new Point(
                txtPassword.Right - pictureBoxToggleShowPassword.Width - 5,
                txtPassword.Top + (txtPassword.Height - pictureBoxToggleShowPassword.Height) / 2
            );

            // Position toggle username/ password inside textbox            
            pictureBoxToggleUsername.Image = Properties.Resources.ic_account_child_128_28130;
            pictureBoxToggleUsername.Location = new Point(
            txtUsername.Left + 1, // Small padding from left edge
            txtUsername.Top + (txtUsername.Height - pictureBoxToggleUsername.Height) / 2
            );

            pictureBoxTogglePassword.Image = Properties.Resources.internet_lock_locked_padlock_password_secure_security_icon_127100;
            pictureBoxTogglePassword.Location = new Point(
            txtPassword.Left + 1, // Small padding from left edge
            txtPassword.Top + (txtPassword.Height - pictureBoxTogglePassword.Height) / 2
            );

            // Set TextBox padding to prevent typing under icon
            txtUsername.Padding = new Padding(pictureBoxToggleUsername.Width + 10, 0, 0, 0);

            // Handle TextBox resize to keep icon properly positioned
            txtUsername.Resize += (s, e) =>
            {
                pictureBoxToggleUsername.Location = new Point(
                    txtUsername.Left + 5,
                    txtUsername.Top + (txtUsername.Height - pictureBoxToggleUsername.Height) / 2);
            };

            //Adjust textbox padding to prevent text overlapping with icons
            txtUsername.Padding = new Padding(
                pictureBoxToggleUsername.Width + 10, 8, pictureBoxToggleShowPassword.Width + 10, txtUsername.Height / 2);

            txtPassword.Padding = new Padding(
                pictureBoxTogglePassword.Width + 10, txtPassword.Height / 2, pictureBoxToggleShowPassword.Width + 10, txtPassword.Height / 2);

            // Bring picturebox to front            
            pictureBoxToggleUsername.BringToFront();
            pictureBoxTogglePassword.BringToFront();
            pictureBoxToggleShowPassword.BringToFront();
        }

        private void InitializeLockoutControls()
        {
            // Create lockout timer
            lockoutTimer = new System.Windows.Forms.Timer();
            lockoutTimer.Interval = 1000; // 1 second
            lockoutTimer.Tick += LockoutTimer_Tick;

            // Create lockout message label
            lblLockoutMessage = new Label();
            lblLockoutMessage.Text = "";
            lblLockoutMessage.ForeColor = Color.Red;
            lblLockoutMessage.Font = new Font("Microsoft Sans Serif", 8, FontStyle.Bold);
            lblLockoutMessage.TextAlign = ContentAlignment.MiddleCenter;            
            lblLockoutMessage.Height = 25;           
            lblLockoutMessage.Visible = false;

            // Manual positioning 15px from bottom
            lblLockoutMessage.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            lblLockoutMessage.Width = this.ClientSize.Width - 30;
            lblLockoutMessage.Top = this.ClientSize.Height - lblLockoutMessage.Height - 15;
            lblLockoutMessage.Left = 15;

            // Add to form
            this.Controls.Add(lblLockoutMessage);            
        }

        private void LockoutTimer_Tick(object sender, EventArgs e)
        {
            lockoutSeconds--;

            if (lockoutSeconds > 0)
            {
                lblLockoutMessage.Text = $"Too many failed attempts. Please try again in {lockoutSeconds} seconds.";
            }
            else
            {
                // Lockout period over
                lockoutTimer.Stop();
                lblLockoutMessage.Visible = false;
                EnableLoginControls(true);
                // Don't reset loginAttempts here - keep them until successful login
                //loginAttempts = 0; // Reset attempts after successful wait period
            }
        }
        private void EnableLoginControls(bool enable)
        {
            txtUsername.Enabled = enable;
            txtPassword.Enabled = enable;
            btnLogin.Enabled = enable;

            if (enable)
            {
                txtPassword.Focus();
                lblLockoutMessage.Visible = false;
            }
        }
        private void InitializeAutoComplete()
        {
            txtUsername.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtUsername.AutoCompleteSource = AutoCompleteSource.CustomSource;

            // Preload common usernames or leave empty to populate dynamically
            txtUsername.AutoCompleteCustomSource = new AutoCompleteStringCollection();
        }
        private void Username_Enter(object sender, EventArgs e)
        {
            if ((string)txtUsername.Tag == "placeholder")
            {
                txtUsername.Text = "";
                txtUsername.ForeColor = Color.DeepSkyBlue;
                txtUsername.Font = new Font(txtUsername.Font.FontFamily, 18f); // Normal font size
                txtUsername.Tag = null;
                txtUsername.TextAlign = HorizontalAlignment.Center;
            }
        }

        private void Username_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUsername.Text))
            {
                txtUsername.Text = "Username";
                txtUsername.ForeColor = Color.Gray;
                //txtUsername.Font = new Font(txtUsername.Font.FontFamily, 12); // Normal font size
                txtUsername.Tag = "placeholder";
            }
        }

        private void Password_Enter(object sender, EventArgs e)
        {
            if ((string)txtPassword.Tag == "placeholder")
            {
                txtPassword.Text = "";
                txtPassword.ForeColor = Color.DeepSkyBlue;
                txtPassword.UseSystemPasswordChar = true;
                txtPassword.Tag = null;
            }
        }

        private void Password_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPassword.Text) && !this.ContainsFocus)
            {
                txtPassword.UseSystemPasswordChar = false;
                txtPassword.Text = "Password";
                txtPassword.ForeColor = Color.Gray;
                txtPassword.Tag = "placeholder";
            }
        }
        private void ToolTip1_Draw(object sender, DrawToolTipEventArgs e)
        {
            // Custom drawing code
            e.DrawBackground();
            e.DrawBorder();
            e.DrawText();
        }

        private void ApplyRoundedCorners()
        {
            int cornerRadius = 20; // Adjust for roundness
            GraphicsPath path = new GraphicsPath();

            // Create a rounded rectangle path
            path.AddArc(0, 0, cornerRadius, cornerRadius, 180, 90); // Top-left
            path.AddArc(this.Width - cornerRadius, 0, cornerRadius, cornerRadius, 270, 90); // Top-right
            path.AddArc(this.Width - cornerRadius, this.Height - cornerRadius, cornerRadius, cornerRadius, 0, 90); // Bottom-right
            path.AddArc(0, this.Height - cornerRadius, cornerRadius, cornerRadius, 90, 90); // Bottom-left
            path.CloseFigure();

            this.Region = new Region(path); // Apply the rounded shape
        }
        private void StylishForm_Paint(object sender, PaintEventArgs e)
        {
            int borderWidth = 10;
            Rectangle rect = new Rectangle(0, 0, this.Width - 1, this.Height - 1);

            // Gradient: Color.DeepSkyBlue, Color.Cyan
            using (var brush = new LinearGradientBrush(rect, Color.DodgerBlue, Color.Cyan, LinearGradientMode.ForwardDiagonal))
            using (var pen = new Pen(brush, borderWidth))
            {
                e.Graphics.DrawRectangle(pen, rect);
            }
        }
        // Optional: Allow dragging the form
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button == MouseButtons.Left)
            {
                this.Capture = false;
                Message msg = Message.Create(this.Handle, 0xA1, new IntPtr(2), IntPtr.Zero);
                this.WndProc(ref msg);
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }

        static string Encrypt(string value)
        {
            string hash = "f0xle@rn";
            byte[] data = UTF8Encoding.UTF8.GetBytes(value);
            using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
            {
                byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                using (TripleDESCryptoServiceProvider tripDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                {
                    ICryptoTransform transform = tripDes.CreateEncryptor();
                    byte[] results = transform.TransformFinalBlock(data, 0, data.Length);
                    return Convert.ToBase64String(results, 0, results.Length);
                }
            }
        }       

        static string Decrypt(string value)
        {
            string hash = "f0xle@rn";
            byte[] data = Convert.FromBase64String(value);
            using (MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
            {
                byte[] keys = md5.ComputeHash(UTF8Encoding.UTF8.GetBytes(hash));
                using (TripleDESCryptoServiceProvider tripDes = new TripleDESCryptoServiceProvider() { Key = keys, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 })
                {
                    ICryptoTransform transform = tripDes.CreateDecryptor();
                    byte[] results = transform.TransformFinalBlock(data, 0, data.Length);
                    //return Convert.ToBase64String(results, 0, results.Length);
                    return UTF8Encoding.UTF8.GetString(results, 0, results.Length);
                }
            }
        }
        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (Login())
            {
                // Login successful - set DialogResult and form will close
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                // Login failed - explicitly set to None to prevent form closure
                this.DialogResult = DialogResult.None;
            }

            //// Set DialogResult based on login success
            //this.DialogResult = Login() ? DialogResult.OK : DialogResult.Cancel;

            // Old codes
            #region
            //ut = new Utility();
            //var conn = ut.OpenDB();
            //DateTime strngaynhapht = DateTime.Now;
            //string ngaynhapht = strngaynhapht.ToString("yyyy-MM-dd hh:mm:ss tt");

            //try
            //{
            //    if (ValidInput() == true)
            //    {
            //        if (conn.State == ConnectionState.Closed)
            //        {
            //            conn.Open();

            //            string username = txtUsername.Text.Trim();
            //            string password = Encrypt(txtPassword.Text.Trim());

            //            cmd = new SqlCommand("", conn);
            //            sda = new SqlDataAdapter(@"Select LoginPassword From Login Where LoginName = N'" + username + "'", conn);
            //            dt = new DataTable();
            //            sda.Fill(dt);

            //            if (dt.Rows[0][0].ToString() != txtPassword.Text.Trim())
            //            {
            //                cmd = new SqlCommand("", conn);
            //                sda = new SqlDataAdapter(@"Select * From Login Where LoginName = N'" + username + "' and LoginPassword = N'" + password + "'", conn);
            //                dt = new DataTable();
            //                sda.Fill(dt);

            //                if (dt.Rows.Count == 1)
            //                {
            //                    this.Hide();
            //                    MDIParentMain frmMDI = new MDIParentMain(dt.Rows[0][0].ToString(), dt.Rows[0][1].ToString(), dt.Rows[0][2].ToString(), (bool)dt.Rows[0][3]);

            //                    frmMDI.Show();
            //                    frmMDI.StartPosition = FormStartPosition.CenterScreen;
            //                    frmMDI.WindowState = FormWindowState.Maximized;
            //                    frmMDI.Refresh();
            //                    cmd = new SqlCommand("", conn);

            //                    if (dt.Rows[0][3].ToString() == "True")
            //                    {
            //                        this.Hide();

            //                        cmd.CommandText = "Update Login Set NgaynhapHT = @NgaynhapHT Where LoginName = N'" + txtUsername.Text.Trim() + "'";
            //                        cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;

            //                        cmd.CommandType = CommandType.Text;
            //                        cmd.ExecuteNonQuery();
            //                        cmd.Parameters.Clear();
            //                    }
            //                    else
            //                    {
            //                        this.Hide();
            //                        frmMDI.Show();
            //                        frmMDI.StartPosition = FormStartPosition.CenterScreen;
            //                        frmMDI.WindowState = FormWindowState.Maximized;
            //                        frmMDI.Refresh();
            //                        cmd.CommandText = "Update Login Set NgaynhapHT = @NgaynhapHT Where LoginName = N'" + txtUsername.Text.Trim() + "'";
            //                        cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;

            //                        cmd.CommandType = CommandType.Text;
            //                        cmd.ExecuteNonQuery();
            //                        cmd.Parameters.Clear();
            //                    }
            //                }
            //                else
            //                {
            //                    MessageBox.Show("Login failed. Please check the username and password.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //                }
            //            }
            //            else
            //            {
            //                //string mk = txtLoginPassword.Text.Trim();
            //                cmd = new SqlCommand("", conn);
            //                SqlDataAdapter sda = new SqlDataAdapter("Select * From Login " +
            //                    "Where LoginName = N'" + username + "' and LoginPassword = N'" + password + "'", conn);
            //                DataTable dt = new DataTable();
            //                sda.Fill(dt);

            //                if (dt.Rows.Count == 1)
            //                {
            //                    this.Hide();
            //                    MDIParentMain frmMDI = new MDIParentMain(dt.Rows[0][0].ToString(), dt.Rows[0][1].ToString(), dt.Rows[0][2].ToString(), (bool)dt.Rows[0][3]);

            //                    frmMDI.Show();
            //                    frmMDI.StartPosition = FormStartPosition.CenterScreen;
            //                    frmMDI.WindowState = FormWindowState.Maximized;
            //                    frmMDI.Refresh();
            //                    cmd = new SqlCommand("", conn);

            //                    if (dt.Rows[0][3].ToString() == "True")
            //                    {
            //                        this.Hide();

            //                        cmd.CommandText = "Update Login Set NgaynhapHT = @NgaynhapHT Where LoginName = N'" + txtUsername.Text.Trim() + "'";
            //                        cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;

            //                        cmd.CommandType = CommandType.Text;
            //                        cmd.ExecuteNonQuery();
            //                        cmd.Parameters.Clear();
            //                    }
            //                    else
            //                    {
            //                        this.Hide();
            //                        frmMDI.Show();
            //                        frmMDI.StartPosition = FormStartPosition.CenterScreen;
            //                        frmMDI.WindowState = FormWindowState.Maximized;
            //                        frmMDI.Refresh();
            //                        cmd.CommandText = "Update Login Set NgaynhapHT = @NgaynhapHT Where LoginName = N'" + txtUsername.Text.Trim() + "'";
            //                        cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;

            //                        cmd.CommandType = CommandType.Text;
            //                        cmd.ExecuteNonQuery();
            //                        cmd.Parameters.Clear();
            //                    }
            //                }
            //                else
            //                {
            //                    MessageBox.Show("Login failed. Please check the username and password.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //                }
            //            }
            //        }
            //        conn.Close();
            //        conn.Dispose();

            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            #endregion
        }
        private bool Login()
        {
            // Check if currently locked out
            if (lockoutTimer.Enabled)
            {
                MessageBox.Show("Please wait for the lockout period to end before trying again.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // Check login attempts
            if (loginAttempts >= MAX_LOGIN_ATTEMPTS)
            {
                StartLockoutPeriod();
                return false;
            }

            if (!ValidInput())
            {
                MessageBox.Show("Please enter both username and password.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtPassword.Text = "";
                txtPassword.Focus();
                loginAttempts++;

                // Check if reached max attempts after this failure
                if (loginAttempts >= MAX_LOGIN_ATTEMPTS)
                {
                    StartLockoutPeriod();
                }
                return false;
            }

            try
            {
                ut = new Utility();
                using (var conn = ut.OpenDB())
                {
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                    }
                    string username = txtUsername.Text.Trim();
                    string password = Encrypt(txtPassword.Text.Trim());
                    DateTime currentDate = DateTime.Now;

                    string query = @"SELECT * FROM Login WHERE LoginName = @Username AND LoginPassword = @Password";

                    using (cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Username", username);
                        cmd.Parameters.AddWithValue("@Password", password);

                        using (sda = new SqlDataAdapter(cmd))
                        {
                            dt = new DataTable();
                            sda.Fill(dt);

                            if (dt.Rows.Count == 1)
                            {
                                // Login successful - reset attempts
                                loginAttempts = 0; // Reset only on successful login
                                DataRow userData = dt.Rows[0];
                                UpdateLastLoginDate(conn, username, currentDate);
                                ShowMainForm(userData);
                                return true;
                            }
                            else
                            {
                                if (CheckUnencryptedPassword(conn, username, txtPassword.Text.Trim()))
                                {
                                    UpdatePasswordToEncrypted(conn, username, password);
                                    UpdateLastLoginDate(conn, username, currentDate);

                                    using (var cmd2 = new SqlCommand(query, conn))
                                    {
                                        cmd2.Parameters.AddWithValue("@Username", username);
                                        cmd2.Parameters.AddWithValue("@Password", password);

                                        using (var sda2 = new SqlDataAdapter(cmd2))
                                        {
                                            dt = new DataTable();
                                            sda2.Fill(dt);

                                            if (dt.Rows.Count == 1)
                                            {
                                                // Login successful - reset attempts
                                                loginAttempts = 0; // Reset only on successful login
                                                ShowMainForm(dt.Rows[0]);
                                                return true;
                                            }
                                        }
                                    }
                                }

                                MessageBox.Show("Login failed. Please check the username and password.",
                                    "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                txtPassword.Text = "";
                                txtPassword.Focus();
                                loginAttempts++;

                                // Check if reached max attempts after this failure
                                if (loginAttempts >= MAX_LOGIN_ATTEMPTS)
                                {
                                    StartLockoutPeriod();
                                }
                                return false;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Login error: {ex.Message}", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);

            if (this.Visible && !lockoutTimer.Enabled)
            {
                // Reset attempts when form becomes visible and not in lockout
                loginAttempts = 0;
            }
        }

        #region
        //private bool Login()
        //{
        //    // Check if currently locked out
        //    if (lockoutTimer.Enabled)
        //    {
        //        MessageBox.Show("Please wait for the lockout period to end before trying again.", "Notice",
        //            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return false;
        //    }

        //    // Check login attempts
        //    if (loginAttempts >= MAX_LOGIN_ATTEMPTS)
        //    {
        //        StartLockoutPeriod();
        //        return false;
        //    }

        //    if (!ValidInput())
        //    {
        //        MessageBox.Show("Please enter both username and password.", "Notice",
        //            MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        txtPassword.Text = "";
        //        txtPassword.Focus();
        //        loginAttempts++;

        //        // Check if reached max attempts after this failure
        //        if (loginAttempts >= MAX_LOGIN_ATTEMPTS)
        //        {
        //            StartLockoutPeriod();
        //        }
        //        return false;
        //    }

        //    try
        //    {
        //        ut = new Utility();
        //        using (var conn = ut.OpenDB())
        //        {
        //            if (conn.State == ConnectionState.Closed)
        //            {
        //                conn.Open();
        //            }
        //            string username = txtUsername.Text.Trim();
        //            string password = Encrypt(txtPassword.Text.Trim());
        //            DateTime currentDate = DateTime.Now;

        //            string query = @"SELECT * FROM Login WHERE LoginName = @Username AND LoginPassword = @Password";

        //            using (cmd = new SqlCommand(query, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@Username", username);
        //                cmd.Parameters.AddWithValue("@Password", password);

        //                using (sda = new SqlDataAdapter(cmd))
        //                {
        //                    dt = new DataTable();
        //                    sda.Fill(dt);

        //                    if (dt.Rows.Count == 1)
        //                    {
        //                        // Login successful - reset attempts
        //                        loginAttempts = 0;
        //                        DataRow userData = dt.Rows[0];
        //                        UpdateLastLoginDate(conn, username, currentDate);
        //                        ShowMainForm(userData);
        //                        return true;
        //                    }
        //                    else
        //                    {
        //                        if (CheckUnencryptedPassword(conn, username, txtPassword.Text.Trim()))
        //                        {
        //                            UpdatePasswordToEncrypted(conn, username, password);
        //                            UpdateLastLoginDate(conn, username, currentDate);

        //                            using (var cmd2 = new SqlCommand(query, conn))
        //                            {
        //                                cmd2.Parameters.AddWithValue("@Username", username);
        //                                cmd2.Parameters.AddWithValue("@Password", password);

        //                                using (var sda2 = new SqlDataAdapter(cmd2))
        //                                {
        //                                    dt = new DataTable();
        //                                    sda2.Fill(dt);

        //                                    if (dt.Rows.Count == 1)
        //                                    {
        //                                        // Login successful - reset attempts
        //                                        loginAttempts = 0;
        //                                        ShowMainForm(dt.Rows[0]);
        //                                        return true;
        //                                    }
        //                                }
        //                            }
        //                        }

        //                        MessageBox.Show("Login failed. Please check the username and password.",
        //                            "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                        txtPassword.Text = "";
        //                        txtPassword.Focus();
        //                        loginAttempts++;

        //                        // Check if reached max attempts after this failure
        //                        if (loginAttempts >= MAX_LOGIN_ATTEMPTS)
        //                        {
        //                            StartLockoutPeriod();
        //                        }
        //                        return false;
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Login error: {ex.Message}", "Notice",
        //            MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        return false;
        //    }
        //}
        #endregion

        private void StartLockoutPeriod()
        {
            // Double the lockout time for subsequent lockouts (10, 20, 40, etc.)
            if (loginAttempts > MAX_LOGIN_ATTEMPTS)
            {
                lockoutSeconds *= 2;
            }
            else
            {
                lockoutSeconds = 10; // Initial lockout time
            }

            // Start lockout
            EnableLoginControls(false);
            lblLockoutMessage.Visible = true;
            lblLockoutMessage.Text = $"Too many failed attempts. Please try again in {lockoutSeconds} seconds.";
            lblLockoutMessage.BringToFront(); // Ensure it's visible
            lockoutTimer.Start();
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            lockoutTimer?.Stop();
            lockoutTimer?.Dispose();
            base.OnFormClosing(e);
        }
        public MDIParentMain MainForm { get; private set; }
        private void ShowMainFormWithParameters(DataRow userData)
        {
            // Extract all required parameters
            string loginName = userData["LoginName"].ToString();
            string tencc = userData["TenCC"].ToString();
            string matkhau = userData["LoginPassword"].ToString();

            // Handle role - you can use different approaches:
            string role = userData["Role"].ToString();
            // OR convert boolean to string:
            // string role = Convert.ToBoolean(userData["IsAdmin"]) ? "Admin" : "User";

            // Create main form using parameterized constructor
            MDIParentMain mainForm = new MDIParentMain(loginName, tencc, matkhau, role);

            this.Hide();
            mainForm.StartPosition = FormStartPosition.CenterScreen;
            mainForm.WindowState = FormWindowState.Maximized;
            mainForm.Show();
        }


        private bool CheckUnencryptedPassword(SqlConnection conn, string username, string plainPassword)
        {
            string query = "SELECT COUNT(*) FROM Login WHERE LoginName = @Username AND LoginPassword = @Password";

            using (var cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@Username", username);
                cmd.Parameters.AddWithValue("@Password", plainPassword);

                int count = (int)cmd.ExecuteScalar();
                return count > 0;
            }
        }

        private void UpdatePasswordToEncrypted(SqlConnection conn, string username, string encryptedPassword)
        {
            string query = "UPDATE Login SET LoginPassword = @Password WHERE LoginName = @Username";

            using (var cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@Username", username);
                cmd.Parameters.AddWithValue("@Password", encryptedPassword);
                cmd.ExecuteNonQuery();
            }
        }

        private void UpdateLastLoginDate(SqlConnection conn, string username, DateTime loginDate)
        {
            string query = "UPDATE Login SET NgaynhapHT = @LoginDate WHERE LoginName = @Username";

            using (var cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@Username", username);
                cmd.Parameters.AddWithValue("@LoginDate", loginDate);
                cmd.ExecuteNonQuery();
            }
        }

        private void ShowMainForm(DataRow userData)
        {
            // Old codes
            #region 
            //// Extract user information
            //string loginName = userData["LoginName"].ToString();
            //string tencc = userData["TenCC"].ToString();
            //string password = userData["LoginPassword"].ToString();
            ////string decentralization = userData["Decentralization"].ToString();
            //bool isAdmin = Convert.ToBoolean(userData["Decentralization"]);

            //string roleString = isAdmin ? "Administrator" : "User";
            //// Create main form
            //MDIParentMain mainForm = new MDIParentMain(loginName, tencc, password, roleString);

            //// Hide login form
            //this.Hide();

            //// Show main form           
            //mainForm.StartPosition = FormStartPosition.CenterScreen;
            //mainForm.WindowState = FormWindowState.Maximized;
            //mainForm.Show();
            #endregion

            // Extract all required parameters
            string loginName = userData["LoginName"].ToString();
            string tencc = userData["TenCC"].ToString();
            string matkhau = userData["LoginPassword"].ToString();

            // Handle role - you can use different approaches:
            string role = userData["Decentralization"].ToString();
            // OR convert boolean to string:
            // string role = Convert.ToBoolean(userData["IsAdmin"]) ? "Admin" : "User";

            // Create main form using parameterized constructor
            MDIParentMain mainForm = new MDIParentMain(loginName, tencc, matkhau, role);

            this.Hide();
            mainForm.StartPosition = FormStartPosition.CenterScreen;
            mainForm.WindowState = FormWindowState.Maximized;
            mainForm.Show();
        }

        //private bool ValidInput()
        //{
        //    return !string.IsNullOrWhiteSpace(txtUsername.Text) &&
        //           !string.IsNullOrWhiteSpace(txtPassword.Text);
        //}
        private bool ValidInput()
        {
            string username = txtUsername.Text.Trim();
            string password = txtPassword.Text.Trim();

            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password) || username.Length != 9)
            {
                MessageBox.Show("Login name or password is invalid.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            //if (txtUsername.Text.Trim() == "" || txtUsername.TextLength > 9 || txtUsername.TextLength < 9 || txtPassword.Text.Trim() == "")
            //{
            //    MessageBox.Show("Login name or password is invalid.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return false;
            //}
            return true;
        }

        //private string Encrypt(string plainText)
        //{
        //    // Your encryption logic here
        //    // Example: return YourEncryptionUtility.Encrypt(plainText);
        //    return plainText; // Replace with actual encryption
        //}

        // Login method codes
        #region
        //private void Login()
        //{
        //    ut = new Utility();
        //    var conn = ut.OpenDB();
        //    DateTime strngaynhapht = DateTime.Now;
        //    string ngaynhapht = strngaynhapht.ToString("yyyy-MM-dd hh:mm:ss tt");

        //    try
        //    {
        //        if (ValidInput() == true)
        //        {
        //            if (conn.State == ConnectionState.Closed)
        //            {
        //                conn.Open();

        //                string username = txtUsername.Text.Trim();
        //                string password = Encrypt(txtPassword.Text.Trim());

        //                cmd = new SqlCommand("", conn);
        //                sda = new SqlDataAdapter("Select LoginPassword From Login Where LoginName = N'" + username + "'", conn);
        //                dt = new DataTable("tblLoginPassword");
        //                sda.Fill(dt);

        //                if (dt.Rows[0][0].ToString() != txtPassword.Text.Trim())
        //                {
        //                    cmd = new SqlCommand("", conn);
        //                    sda = new SqlDataAdapter(@"Select * From Login Where LoginName = N'" + username + "' and LoginPassword = N'" + password + "'", conn);
        //                    dt = new DataTable();
        //                    sda.Fill(dt);

        //                    if (dt.Rows.Count == 1)
        //                    {
        //                        this.Hide();
        //                        //this.Close();
        //                        MDIParentMain frmMDI = new MDIParentMain(dt.Rows[0][0].ToString(), dt.Rows[0][1].ToString(), dt.Rows[0][2].ToString(), (bool)dt.Rows[0][3]);

        //                        frmMDI.Show();
        //                        frmMDI.StartPosition = FormStartPosition.CenterScreen;
        //                        frmMDI.WindowState = FormWindowState.Maximized;
        //                        frmMDI.Refresh();

        //                        if (dt.Rows[0][3].ToString() == "True")
        //                        {
        //                            //this.Hide();

        //                            cmd.CommandText = "Update Login Set NgaynhapHT = @NgaynhapHT Where LoginName = N'" + username + "'";
        //                            cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;

        //                            cmd.CommandType = CommandType.Text;
        //                            cmd.ExecuteNonQuery();
        //                            cmd.Parameters.Clear();
        //                        }
        //                        else
        //                        {
        //                            //this.Hide();

        //                            frmMDI.Show();
        //                            frmMDI.StartPosition = FormStartPosition.CenterScreen;
        //                            frmMDI.WindowState = FormWindowState.Maximized;
        //                            frmMDI.Refresh();
        //                            cmd.CommandText = "Update Login Set NgaynhapHT = @NgaynhapHT Where LoginName = N'" + username + "'";
        //                            cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;

        //                            cmd.CommandType = CommandType.Text;
        //                            cmd.ExecuteNonQuery();
        //                            cmd.Parameters.Clear();
        //                        }
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Login failed. Please check the username and password.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                    }
        //                }
        //                else
        //                {
        //                    cmd = new SqlCommand("", conn);
        //                    SqlDataAdapter sda = new SqlDataAdapter("Select * From Login " +
        //                        "Where LoginName = N'" + username + "' and LoginPassword = N'" + password + "'", conn);
        //                    DataTable dt = new DataTable();
        //                    sda.Fill(dt);

        //                    if (dt.Rows.Count == 1)
        //                    {
        //                        this.Hide();
        //                        //this.Close();
        //                        MDIParentMain frmMDI = new MDIParentMain(dt.Rows[0][0].ToString(), dt.Rows[0][1].ToString(), dt.Rows[0][2].ToString(), (bool)dt.Rows[0][3]);

        //                        frmMDI.Show();
        //                        frmMDI.StartPosition = FormStartPosition.CenterScreen;
        //                        frmMDI.WindowState = FormWindowState.Maximized;
        //                        frmMDI.Refresh();
        //                        cmd = new SqlCommand("", conn);

        //                        if (dt.Rows[0][3].ToString() == "True")
        //                        {
        //                            //this.Hide();

        //                            cmd.CommandText = "Update Login Set NgaynhapHT = @NgaynhapHT Where LoginName = N'" + username + "'";
        //                            cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;

        //                            cmd.CommandType = CommandType.Text;
        //                            cmd.ExecuteNonQuery();
        //                            cmd.Parameters.Clear();
        //                        }
        //                        else
        //                        {
        //                            //this.Hide();
        //                            frmMDI.Show();
        //                            frmMDI.StartPosition = FormStartPosition.CenterScreen;
        //                            frmMDI.WindowState = FormWindowState.Maximized;
        //                            frmMDI.Refresh();
        //                            cmd.CommandText = "Update Login Set NgaynhapHT = @NgaynhapHT Where LoginName = N'" + username + "'";
        //                            cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;

        //                            cmd.CommandType = CommandType.Text;
        //                            cmd.ExecuteNonQuery();
        //                            cmd.Parameters.Clear();
        //                        }
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Login failed. Please check the username and password.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                    }
        //                }
        //            }
        //            // Close Connecttion
        //            conn.Close();
        //            conn.Dispose();
        //            conn = null;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }
        //}
        #endregion

        private bool VerifyPassword(string inputPassword, string storedHash, string salt)
        {
            if (string.IsNullOrEmpty(storedHash) || string.IsNullOrEmpty(salt))
                return false;

            // Hash the input password with the same salt
            string inputHash = HashPassword(inputPassword, salt);

            // Compare hashes (use secure comparison to prevent timing attacks)
            return SecureCompare(inputHash, storedHash);
        }

        private string HashPassword(string password, string salt)
        {
            using (var sha256 = SHA256.Create())
            {
                byte[] saltedPassword = Encoding.UTF8.GetBytes(password + salt);
                byte[] hash = sha256.ComputeHash(saltedPassword);
                return Convert.ToBase64String(hash);
            }
        }

        private bool SecureCompare(string a, string b)
        {
            // Prevents timing attacks by comparing all characters
            if (a.Length != b.Length) return false;

            int result = 0;
            for (int i = 0; i < a.Length; i++)
            {
                result |= a[i] ^ b[i];
            }
            return result == 0;
        }
        private void frmLogin_Load(object sender, EventArgs e)
        {
            Utility ut = new Utility();
            var conn = ut.OpenDB();

            InitializeAutoComplete();
            InitializeLoginPasswordControls();
            txtPassword.PasswordChar = '*';
            txtPassword.UseSystemPasswordChar = true;
            txtPassword.Text = ""; // Clear and reset

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
                }
                conn.Close();
                conn.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void TogglePasswordVisibility()
        {
            if (txtPassword.UseSystemPasswordChar == true)
            {
                // Show password
                txtPassword.UseSystemPasswordChar = false;
                pictureBoxToggleShowPassword.Image = Properties.Resources.icons8_eye_opened_30_1;
            }
            else
            {
                // Hide password
                txtPassword.UseSystemPasswordChar = true;
                pictureBoxToggleShowPassword.Image = Properties.Resources.icons8_eye_closed_30_1;
            }
            // Keep focus on password field
            txtPassword.Focus();
        }

        private void pictureBoxClose_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }

        #region
        //private void txtLoginName_TextChanged(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        Utility ut = new Utility();
        //        var conn = ut.OpenDB();

        //        txtUsername.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
        //        txtUsername.AutoCompleteSource = AutoCompleteSource.CustomSource;

        //        //txtLoginName.AutoCompleteMode = AutoCompleteMode.None;
        //        //txtLoginName.AutoCompleteSource = AutoCompleteSource.None;
        //        //txtLoginName.ImeMode = ImeMode.Disable;

        //        //txtPassword.AutoCompleteMode = AutoCompleteMode.None;
        //        //txtPassword.AutoCompleteSource = AutoCompleteSource.None;
        //        //txtPassword.ImeMode = ImeMode.Disable;

        //        string loginname = txtUsername.Text.Trim();

        //        //System.Windows.Forms.TextBox textBox = (System.Windows.Forms.TextBox)sender;
        //        //textBox.Text = textBox.Text.ToUpper();
        //        //textBox.SelectionStart = textBox.Text.Length; // Keep cursor at the end

        //        if (txtUsername.TextLength == 9 && txtPassword.Text != null)
        //        {
        //            btnLogin.ForeColor = Color.Blue;
        //            btnLogin.Cursor = Cursors.Hand;
        //            btnLogin.Enabled = true;
        //        }
        //        else
        //        {
        //            //Refresh();
        //            btnLogin.ForeColor = Color.Blue;
        //            btnLogin.Cursor = Cursors.No;
        //            btnLogin.Enabled = false;
        //        }

        //        if (conn.State == ConnectionState.Closed)
        //        {
        //            conn.Open();

        //            if (txtUsername.TextLength > 2)
        //            {
        //                SqlCommand cmd = new SqlCommand("select LoginName From Login Where LoginName Like @LoginName", conn);
        //                cmd.Parameters.Add(new SqlParameter("LoginName", "%" + loginname + "%"));

        //                SqlDataReader dr = cmd.ExecuteReader();
        //                AutoCompleteStringCollection suggestions = new AutoCompleteStringCollection();

        //                while (dr.Read())
        //                {
        //                    suggestions.Add(dr.GetString(0));
        //                }

        //                txtUsername.AutoCompleteCustomSource = suggestions;
        //            }
        //        }
        //        //Close connection
        //        conn.Close();
        //        conn.Dispose();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //    }
        //}
        #endregion

        private void InitializeLoginPasswordControls()
        {
            // Username textbox
            txtUsername.Text = "Username";
            txtUsername.ForeColor = Color.Gray;
            txtUsername.Tag = "placeholder"; // Mark as placeholder text

            txtUsername.Enter += (sender, e) => {
                if ((string)txtUsername.Tag == "placeholder")
                {
                    txtUsername.Text = "";
                    txtUsername.ForeColor = Color.DeepSkyBlue;
                    txtUsername.Tag = null;
                }
            };

            txtUsername.Leave += (sender, e) => {
                if (string.IsNullOrWhiteSpace(txtUsername.Text))
                {
                    txtUsername.Text = "Username";
                    txtUsername.ForeColor = Color.Gray;
                    txtUsername.Tag = "placeholder";
                }
            };

            // Password textbox
            txtPassword.Text = "Password";
            txtPassword.ForeColor = Color.Gray;
            txtPassword.UseSystemPasswordChar = false;
            txtPassword.Tag = "placeholder"; // Mark as placeholder text

            txtPassword.Enter += (sender, e) => {
                if ((string)txtPassword.Tag == "placeholder")
                {
                    txtPassword.Text = "";
                    txtPassword.ForeColor = Color.DeepSkyBlue;
                    txtPassword.UseSystemPasswordChar = true;
                    txtPassword.Tag = null;
                }
            };

            txtPassword.Leave += (sender, e) => {
                // Only restore placeholder if we're not switching to another control
                if (string.IsNullOrWhiteSpace(txtPassword.Text) && !this.ContainsFocus)
                {
                    txtPassword.UseSystemPasswordChar = false;
                    txtPassword.Text = "Password";
                    txtPassword.ForeColor = Color.Gray;
                    txtPassword.Tag = "placeholder";
                }
            };

            // Add these events to the form
            this.Deactivate += (sender, e) => {
                // When form loses focus, check if we need to restore placeholders
                if (string.IsNullOrWhiteSpace(txtUsername.Text))
                {
                    txtUsername.Text = "Username";
                    txtUsername.ForeColor = Color.Gray;
                    txtUsername.Tag = "placeholder";
                }

                if (string.IsNullOrWhiteSpace(txtPassword.Text))
                {
                    txtPassword.UseSystemPasswordChar = false;
                    txtPassword.Text = "Password";
                    txtPassword.ForeColor = Color.Gray;
                    txtPassword.Tag = "placeholder";
                }
            };
        }

        //private void txtUsername_TextChanged(object sender, EventArgs e)
        //{
        //    string username = txtUsername.Text.Trim();
        //    ConfigureAutoCompleteSettings();

        //    // Handle login button state
        //    UpdateLoginButtonState();

        //    // Only query when we have enough characters
        //    if (username.Length >= 8)
        //    {
        //        UpdateAutoCompleteSuggestions(username);
        //    }
        //}

        //private async void txtUsername_TextChanged(object sender, EventArgs e)
        //{
        //    // Cancel any pending operation
        //    _autoCompleteCts?.Cancel();
        //    _autoCompleteCts = new CancellationTokenSource();

        //    // Update last typing time
        //    _lastTypingTime = DateTime.Now;

        //    ConfigureAutoCompleteSettings();
        //    UpdateLoginButtonState();

        //    // Only query after a brief delay when typing stops
        //    if (txtUsername.TextLength >= 3)
        //    {
        //        try
        //        {
        //            await Task.Delay(300, _autoCompleteCts.Token); // Wait for typing to pause

        //            if ((DateTime.Now - _lastTypingTime).TotalMilliseconds >= 300)
        //            {
        //                await LoadAutoCompleteSuggestionsAsync(_autoCompleteCts.Token);
        //            }
        //        }
        //        catch (TaskCanceledException)
        //        {
        //            // Expected cancellation -no action needed
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }
        //    }
        //}

        private async Task LoadAutoCompleteSuggestionsAsync(CancellationToken ct)
        {
            try
            {
                ut = new Utility();
                using (var conn = ut.OpenDB())
                {
                    if (conn.State == ConnectionState.Closed)
                    {
                        await conn.OpenAsync(ct);
                    }

                    string username = txtUsername.Text.Trim();
                    string query = "SELECT LoginName FROM Login WHERE LoginName LIKE @LoginName";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@LoginName", "%" + username + "%");

                        using (var dr = await cmd.ExecuteReaderAsync(ct))
                        {
                            var suggestions = new AutoCompleteStringCollection();

                            while (await dr.ReadAsync(ct))
                            {
                                ct.ThrowIfCancellationRequested();

                                if (!dr.IsDBNull(0))
                                {
                                    suggestions.Add(dr.GetString(0));
                                }
                            }

                            // Update UI on the main thread
                            if (!ct.IsCancellationRequested)
                            {
                                this.Invoke((MethodInvoker)delegate
                                {
                                    txtUsername.AutoCompleteCustomSource = suggestions;
                                });
                            }
                        }
                    }
                }
            }
            catch (SqlException ex) when (ex.Number == -2) // Timeout
            {
                // Handle timeout specifically if needed
            }
        }

        private void ConfigureAutoCompleteSettings()
        {
            txtUsername.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtUsername.AutoCompleteSource = AutoCompleteSource.CustomSource;

            // Disable auto-complete for password field (uncomment if needed)
            txtPassword.AutoCompleteMode = AutoCompleteMode.None;
            txtPassword.AutoCompleteSource = AutoCompleteSource.None;
            txtPassword.ImeMode = ImeMode.Disable;
        }

        private void UpdateLoginButtonState()
        {
            bool isValid = txtUsername.TextLength == 9 && !string.IsNullOrEmpty(txtPassword.Text);
            btnLogin.Enabled = isValid;
            btnLogin.ForeColor = isValid ? Color.Blue : Color.Gray;
            btnLogin.Cursor = isValid ? Cursors.Hand : Cursors.Default;
        }

        private void LoadAutoCompleteSuggestions()
        {
            try
            {
                ut = new Utility();
                using (var conn = ut.OpenDB())
                {
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                    }

                    string query = "SELECT LoginName FROM Login WHERE LoginName LIKE @LoginName";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@LoginName", "%" + txtUsername.Text.Trim() + "%");

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            var suggestions = new AutoCompleteStringCollection();

                            while (dr.Read())
                            {
                                if (!dr.IsDBNull(0))
                                {
                                    suggestions.Add(dr.GetString(0));
                                }
                            }

                            txtUsername.AutoCompleteCustomSource = suggestions;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading suggestions: {ex.Message}",
                               "Database Error",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Warning);
            }
        }

        private void pictureBoxTogglePassword_MouseEnter(object sender, EventArgs e)
        {
            pictureBoxToggleShowPassword.BackColor = Color.FromArgb(240, 240, 240);
        }

        private void pictureBoxTogglePassword_MouseLeave(object sender, EventArgs e)
        {
            pictureBoxToggleShowPassword.BackColor = Color.Transparent;
        }
        private void pictureBoxLogin_Click(object sender, EventArgs e)
        {
            // Set DialogResult based on login success
            this.DialogResult = Login() ? DialogResult.OK : DialogResult.Cancel;

            //Login();
        }

        private void UpdateAutoCompleteSuggestions(string partialUsername)
        {
            try
            {
                Utility ut = new Utility();
                using (var conn = ut.OpenDB())
                {
                    if (conn.State == ConnectionState.Closed)
                        conn.Open();

                    using (SqlCommand cmd = new SqlCommand(
                        "SELECT LoginName FROM Login WHERE LoginName LIKE @LoginName", conn))
                    {
                        cmd.Parameters.AddWithValue("@LoginName", partialUsername + "%");

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            AutoCompleteStringCollection coll = new AutoCompleteStringCollection();

                            while (dr.Read())
                            {
                                if (!dr.IsDBNull(0))
                                    coll.Add(dr.GetString(0));
                            }

                            // This is thread-sensitive - ensure we're on UI thread
                            if (txtUsername.InvokeRequired)
                            {
                                txtUsername.Invoke(new Action(() =>
                                    txtUsername.AutoCompleteCustomSource = coll));
                            }
                            else
                            {
                                txtUsername.AutoCompleteCustomSource = coll;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //private void txtUsername_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    e.KeyChar = char.ToUpper(e.KeyChar);

        //    //// Get current cursor position
        //    //int cursorPos = txtUsername.SelectionStart;

        //    //// Get character position from cursor position
        //    //Point charPos = txtUsername.GetPositionFromCharIndex(cursorPos);

        //    //// If trying to type in protected area (left of icon)
        //    //if (charPos.X < pictureBoxToggleUsername.Right - txtUsername.Left)
        //    //{
        //    //    // Move cursor to right of icon
        //    //    txtUsername.SelectionStart = txtUsername.TextLength;
        //    //    e.Handled = true;
        //    //}
        //}

        private void pictureBoxToggleShowPassword_Click(object sender, EventArgs e)
        {
            TogglePasswordVisibility();
        }

        private void txtUsername_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = char.ToUpper(e.KeyChar);
        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {
            if (txtPassword.TextLength != 0)
            {
                txtPassword.UseSystemPasswordChar = true;
            }
            else
            {
                btnLogin.Enabled = false;
            }
            btnLogin.Enabled = txtPassword.TextLength != 0 ? true : false;
            TogglePasswordVisibility();
        }

        private void txtUsername_TextChanged(object sender, EventArgs e)
        {
            string username = txtUsername.Text.Trim();
            ConfigureAutoCompleteSettings();

            // Handle login button state
            UpdateLoginButtonState();

            // Only query when we have enough characters
            if (username.Length >= 8)
            {
                UpdateAutoCompleteSuggestions(username);
            }
        }
    
    }
}
