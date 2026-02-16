using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TaxRefund
{
    public partial class frmChangePassword : Form
    {
        SqlCommand cmd;
        SqlDataAdapter sda;
        DataTable dtb;

        public frmChangePassword()
        {
            InitializeComponent();

            this.FormBorderStyle = FormBorderStyle.None;
            this.Paint += StylishForm_Paint;
            this.Resize += (s, e) => this.Invalidate(); // Redraw on resize
            ApplyRoundedCorners();
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
            int borderWidth = 7;
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

        private void frmChangePassword_Load(object sender, EventArgs e)
        {
            lblStatus.Text = ((Form)this.MdiParent).Controls["lblStatus"].Text;
            lblLoginName.Text = ((Form)this.MdiParent).Controls["lblLoginName"].Text;

            //txtLoginName.Text = ((Form)this.MdiParent).Controls["lblLoginName"].Text;
            //lblCurrentPassword.Text = ((Form)this.MdiParent).Controls["lblLoginPassword"].Text;

            lblNewPassword.Enabled = false;
            txtNewPassword.Enabled = false;
            lblPasswordConfirm.Enabled = false;
            txtPasswordconfirm.Enabled = false;
            btnSubmit.Enabled = false;

            //frmChangePassword fChangePassword = new frmChangePassword();
            //fChangePassword.StartPosition = FormStartPosition.CenterParent;
            //fChangePassword.WindowState = FormWindowState.Normal;

            Utility ut = new Utility();
            var conn = ut.OpenDB();

            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();


                    cmd = new SqlCommand("", conn);
                    //cmd.ExecuteNonQuery();

                    sda = new SqlDataAdapter("SELECT LoginName, LoginPassword FROM Login  Where LoginName = N'" + lblLoginName.Text.Trim() + "'", conn);
                    dtb = new DataTable();
                    sda.Fill(dtb);

                    //txtLoginName.DataBindings.Clear();
                    //txtLoginName.DataBindings.Add("Text", dtb, "LoginName");

                    txtLoginName.Text = dtb.Rows[0][0].ToString();
                    lblCurrentPassword.Text = dtb.Rows[0][1].ToString();

                }
                conn.Close();
                conn.Dispose();
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

        private void txtLoginPassword_TextChanged(object sender, EventArgs e)
        {
            Utility ut = new Utility();
            var conn = ut.OpenDB();

            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();

                    string tk = txtLoginName.Text.Trim();

                    cmd = new SqlCommand("", conn);
                    SqlDataAdapter da = new SqlDataAdapter("Select LoginPassword From Login Where LoginName = N'" + tk + "'", conn);
                    DataTable dtb = new DataTable("tblLoginPassword");
                    da.Fill(dtb);

                    if (dtb.Rows[0][0].ToString() != txtLoginPassword.Text.Trim())
                    {
                        string mk = Encrypt(txtLoginPassword.Text.Trim());
                        cmd = new SqlCommand("Select Count(LoginName) From Login Where " +
                       "LoginName = N'" + tk + "' and LoginPassword = N'" + mk + "'", conn);
                        Int32 count = Convert.ToInt32(cmd.ExecuteScalar());

                        if (count > 0)
                        {                           
                            lblNewPassword.Enabled = true;
                            txtNewPassword.Enabled = true;
                            lblPasswordConfirm.Enabled = true;
                            txtPasswordconfirm.Enabled = true;
                        }
                        else
                        {
                            lblNewPassword.Enabled = false;
                            txtNewPassword.Enabled = false;
                            lblPasswordConfirm.Enabled = false;
                            txtPasswordconfirm.Enabled = false;
                        }
                    }
                    else
                    {
                        string mk = txtLoginPassword.Text.Trim();
                        cmd = new SqlCommand("Select Count(LoginName) From Login Where " +
                       "LoginName = N'" + tk + "' and LoginPassword = N'" + mk + "'", conn);
                        Int32 count = Convert.ToInt32(cmd.ExecuteScalar());

                        if (count > 0)
                        {
                            lblNewPassword.Enabled = true;
                            txtNewPassword.Enabled = true;
                            lblPasswordConfirm.Enabled = true;
                            txtPasswordconfirm.Enabled = true;
                        }
                        else
                        {
                            lblNewPassword.Enabled = false;
                            txtNewPassword.Enabled = false;
                            lblPasswordConfirm.Enabled = false;
                            txtPasswordconfirm.Enabled = false;
                        }
                    }
                }
                //Conn Close
                conn.Close();
                conn.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnChangePassword_Click(object sender, EventArgs e)
        {
            Utility ut = new Utility();
            var conn = ut.OpenDB();

            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();

                    string tk = txtLoginName.Text.Trim();

                    cmd = new SqlCommand();
                    sda = new SqlDataAdapter("Select LoginPassword From Login Where LoginName = N'" + tk + "'", conn);
                    dtb = new DataTable();
                    sda.Fill(dtb);

                    string currentPassword = dtb.Rows[0][0].ToString();

                    if (currentPassword != txtLoginPassword.Text.Trim())
                    {
                        string mk = Encrypt(txtLoginPassword.Text.Trim());
                        cmd = new SqlCommand("", conn);
                        SqlDataAdapter da = new SqlDataAdapter("Select * From Login Where LoginName = N'" + tk + "' and LoginPassword = N'" + mk + "'", conn);
                        DataTable dt = new DataTable();
                        da.Fill(dt);

                        if (dt.Rows.Count == 1)
                        {
                            if (txtPasswordconfirm.Text.Contains(txtNewPassword.Text))
                            {
                                string pwConfirmEncript = Encrypt(txtPasswordconfirm.Text.Trim());

                                //SqlCommand cmd_change = new SqlCommand("Update Login Set LoginPassword = N'" + txtPasswordconfirm.Text.Trim() + "' Where LoginName = N'" + tk + "'", conn);                           
                                SqlCommand cmd_change = new SqlCommand("Update Login Set LoginPassword = N'" + pwConfirmEncript + "' Where LoginName = N'" + tk + "'", conn);
                                cmd_change.ExecuteNonQuery();
                                this.Refresh();

                                //Kiem tra xem thay doi mat khau thanh cong khong?
                                string xacnhanmk = pwConfirmEncript;

                                cmd = new SqlCommand("Select Count(LoginName) From Login Where LoginName = N'" + tk + "' and LoginPassword = N'" + xacnhanmk + "'", conn);
                                Int32 count = Convert.ToInt32(cmd.ExecuteScalar());

                                if (count == 1)
                                {
                                    MessageBox.Show("Password changed successfully.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    this.Hide();
                                }
                                else
                                {
                                    MessageBox.Show("Password cannot be changed. Please check again.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Confirm password does not match new password. Please enter confirm password.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Incorrect login password. Please enter the current used password.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    else
                    {
                        string mk = txtLoginPassword.Text.Trim();
                        cmd = new SqlCommand("", conn);
                        SqlDataAdapter da = new SqlDataAdapter("Select * From Login Where LoginName = N'" + tk + "' and LoginPassword = N'" + mk + "'", conn);
                        DataTable dt = new DataTable();
                        da.Fill(dt);

                        if (dt.Rows.Count == 1)
                        {
                            if (txtPasswordconfirm.Text.Contains(txtNewPassword.Text))
                            {
                                string pwConfirmEncript = Encrypt(txtPasswordconfirm.Text.Trim());
                                //string pwConfirmDecript = Decrypt(pwConfirmEncript);

                                //SqlCommand cmd_change = new SqlCommand("Update Login Set LoginPassword = N'" + txtPasswordconfirm.Text.Trim() + "' Where LoginName = N'" + tk + "'", conn);                           
                                SqlCommand cmd_change = new SqlCommand("Update Login Set LoginPassword = N'" + pwConfirmEncript + "' Where LoginName = N'" + tk + "'", conn);
                                cmd_change.ExecuteNonQuery();
                                this.Refresh();

                                //Kiem tra xem thay doi mat khau thanh cong khong?
                                string xacnhanmk = pwConfirmEncript;

                                cmd = new SqlCommand("Select Count(LoginName) From Login Where LoginName = N'" + tk + "' and LoginPassword = N'" + xacnhanmk + "'", conn);
                                Int32 count = Convert.ToInt32(cmd.ExecuteScalar());

                                if (count == 1)
                                {
                                    MessageBox.Show("Password changed successfully.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    this.Hide();
                                }
                                else
                                {
                                    MessageBox.Show("Password cannot be changed. Please check again.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Confirm password does not match new password. Please enter confirm password.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Incorrect login password. Please enter the current used password.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

        private void btnExit_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult tb = MessageBox.Show("Bạn có chắc muốn đóng cửa sổ hiện hành không?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (tb == DialogResult.Yes)
                {
                    //Application.Exit();
                    this.Close();
                }
                else
                {
                    //Do Nothing
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void txtNewPassword_TextChanged(object sender, EventArgs e)
        {
            string currentPassword = txtLoginPassword.Text.Trim();
            string newPassword = txtNewPassword.Text.Trim();

            if (string.IsNullOrEmpty(newPassword))
            {
                lblPasswordConfirm.Enabled = false;
                txtPasswordconfirm.Enabled = false;                
            }
            else if (newPassword == currentPassword)
            {
                lblPasswordConfirm.Enabled = false;
                txtPasswordconfirm.Enabled = false;
                MessageBox.Show("The password you entered has already been used.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);                
            }
            else
            {
                lblPasswordConfirm.Enabled = true;
                txtPasswordconfirm.Enabled = true;
            }
        }

        private void txtLoginName_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            if (textBox.Text != textBox.Text.ToUpper())
            {
                int cursorPos = textBox.SelectionStart; // Save cursor position
                textBox.Text = textBox.Text.ToUpper();
                textBox.SelectionStart = cursorPos; // Restore cursor position
            }
        }

        private void txtPasswordconfirm_TextChanged(object sender, EventArgs e)
        {
            if(txtNewPassword.TextLength > 0) 
            {
                btnSubmit.Enabled = true;
            }
        }

        private void pictureBoxClose_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult tb = MessageBox.Show("Are you sure to close current window?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (tb == DialogResult.Yes)
                {
                    //Application.Exit();
                    this.Close();
                }
                else
                {
                    //Do Nothing
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}
