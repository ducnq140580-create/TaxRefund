using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Globalization;
using System.Security.Cryptography;

namespace TaxRefund
{
    public partial class frmRegister : Form
    {
        SqlCommand cmd;
        SqlDataAdapter sda;
        DataTable dt;
        Utility ut;
        //SqlConnection conn;

        public frmRegister()
        {
            InitializeComponent();
        }
        private void SetupDataGridView()
        {
            dgvUserInfo.ColumnHeadersDefaultCellStyle.BackColor = Color.Orange;
            dgvUserInfo.ColumnHeadersDefaultCellStyle.ForeColor = Color.Blue;
            dgvUserInfo.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 9, FontStyle.Bold);
            dgvUserInfo.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvUserInfo.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 9);
            dgvUserInfo.DefaultCellStyle.BackColor = Color.AntiqueWhite;
            dgvUserInfo.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgvUserInfo.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgvUserInfo.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dgvUserInfo.MultiSelect = true;
        }

        private void LoadData()
        {
            ut = new Utility();
            var conn = ut.OpenDB();

            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();

                    if (((Form)this.MdiParent).Controls["lblStatus"].Text == "True")
                    {
                        cmd = new SqlCommand(@"Select ROW_NUMBER() OVER (ORDER BY NgaynhapHT DESC) AS [STT], LoginName as [User], TenCC as [Officer Name], 
                            LoginPassword as [Password], Decentralization, NgaynhapHT 
                            From Login", conn);

                        sda = new SqlDataAdapter(cmd);
                        dt = new System.Data.DataTable();
                        sda.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            dgvUserInfo.DataSource = dt;
                            SetupDataGridView();

                            // Data binding
                            txtLoginName.DataBindings.Clear();
                            txtLoginName.DataBindings.Add("Text", dgvUserInfo.DataSource, "User");
                            txtTenCC.DataBindings.Clear();
                            txtTenCC.DataBindings.Add("Text", dgvUserInfo.DataSource, "Officer Name");
                            txtLoginPassword.DataBindings.Clear();
                            txtLoginPassword.DataBindings.Add("Text", dgvUserInfo.DataSource, "Password");
                            txtConfirmedPassword.DataBindings.Clear();
                            txtConfirmedPassword.DataBindings.Add("Text", dgvUserInfo.DataSource, "Password");
                        }
                        else
                        {
                            MessageBox.Show("You do not have permission to register a new account. Please contact with system administrator.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                //Close connection
                conn.Close();
                conn.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void frmRegister_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void btnRegister_Click(object sender, EventArgs e)
        {
            Utility ut = new Utility();
            var conn = ut.OpenDB();

            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();

                    if (((Form)this.MdiParent).Controls["lblStatus"].Text == "True")
                    {
                        DialogResult tb = MessageBox.Show("Are you sure to register a new account?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (tb == DialogResult.Yes)
                        {
                            txtLoginName.Text = "HQ";
                            txtTenCC.Text = null;
                            txtLoginPassword.Text = "123456";
                            txtConfirmedPassword.Text = "123456";
                            lblPasswordConfirm.Enabled = true;
                            txtConfirmedPassword.Enabled = true;
                            btnSave.Enabled = true;
                        }
                    }
                    else
                    {
                        MessageBox.Show("You do not have permission to register a new account. Please contact with system administrator.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                //Dong ket noi
                conn.Close();
                conn.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            ut = new Utility();
            var conn = ut.OpenDB();

            string strloginame = txtLoginName.Text.Trim();
            string strtenCC = txtTenCC.Text.Trim();
            string strmatkhau = Encrypt(txtLoginPassword.Text.Trim());
            bool strDecentralization = false;

            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();

                    cmd = new SqlCommand("Select Count(LoginName) From Login Where LoginName  = '" + strloginame + "'", conn);
                    int count_loginname = Convert.ToInt32(cmd.ExecuteScalar());

                    if (count_loginname > 0)
                    {
                        MessageBox.Show("User name is existed.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        if (txtLoginPassword.Text == txtConfirmedPassword.Text)
                        {
                            if (txtLoginName.Text.Contains("HQ") != true && txtLoginName.Text.Contains("-") != true && txtLoginName.TextLength != 8)
                            {
                                MessageBox.Show("Username is registered failed. Username is not in the right format.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                            else
                            {                                
                                cmd = new SqlCommand("Insert Into Login (LoginName, TenCC, LoginPassword, Decentralization) " +
                                    "Values (@LoginName, @TenCC, @LoginPassword, @Decentralization)", conn);
                                
                                cmd.Parameters.Add("@LoginName", SqlDbType.VarChar).Value = strloginame;
                                cmd.Parameters.Add("@TenCC", SqlDbType.NVarChar).Value = strtenCC;
                                cmd.Parameters.Add("@LoginPassword", SqlDbType.NVarChar).Value = strmatkhau;
                                cmd.Parameters.Add("@Decentralization", SqlDbType.Bit).Value = strDecentralization; //Convert.ToBoolean(strDecentralization);

                                // Execute query
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();                                
                                cmd.Parameters.Clear();

                                // To check the existence of username
                                cmd = new SqlCommand("Select Count(LoginName) From Login Where LoginName  = '" + strloginame + "'", conn);
                                int count = Convert.ToInt32(cmd.ExecuteScalar());

                                if (count > 0)
                                {
                                    MessageBox.Show("Username is registered successfully. Initial password is: 123456.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                    cmd = new SqlCommand(@"Select ROW_NUMBER() OVER (ORDER BY LoginName DESC) AS [STT], LoginName as [User], TenCC as [Officer Name], 
                                            LoginPassword as [Password], Decentralization, NgaynhapHT 
                                            From Login", conn);

                                    sda = new SqlDataAdapter(cmd);
                                    dt = new System.Data.DataTable();
                                    sda.Fill(dt);

                                    if (dt.Rows.Count > 0)
                                    {
                                        dgvUserInfo.DataSource = dt;
                                        SetupDataGridView();
                                        this.Refresh();

                                        // Data binding
                                        txtLoginName.DataBindings.Clear();
                                        txtLoginName.DataBindings.Add("Text", dgvUserInfo.DataSource, "User");
                                        txtTenCC.DataBindings.Clear();
                                        txtTenCC.DataBindings.Add("Text", dgvUserInfo.DataSource, "Officer Name");
                                        txtLoginPassword.DataBindings.Clear();
                                        txtLoginPassword.DataBindings.Add("Text", dgvUserInfo.DataSource, "Password");
                                        txtConfirmedPassword.DataBindings.Clear();
                                        txtConfirmedPassword.DataBindings.Add("Text", dgvUserInfo.DataSource, "Password");
                                    }
                                    else
                                    {
                                        MessageBox.Show("Username is registered failed. Please contact with system administrator.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                    }
                                    dgvUserInfo.Refresh();
                                }
                                else
                                {
                                    MessageBox.Show("Username is registered failed. Please contact with system administrator.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Confirmation password does not match. Please check again.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    conn.Close();
                    conn.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        //private void btnSave_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (!ValidateInputs())
        //        {
        //            return;
        //        }

        //        ut = new Utility();
        //        using (var conn = ut.OpenDB())
        //        {
        //            if (conn.State == ConnectionState.Closed)
        //            {
        //                conn.Open();
        //            }

        //            if (IsUsernameExists(conn, txtLoginName.Text.Trim()))
        //            {
        //                ShowWarning("User name already exists.");
        //                return;
        //            }

        //            if (SaveNewUser(conn))
        //            {
        //                RefreshUserDataGrid(conn);
        //                ShowSuccess("Username registered successfully. Initial password is: 123456.");
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ShowError(ex.Message);
        //    }
        //}

        private bool ValidateInputs()
        {
            if (txtLoginPassword.Text != txtConfirmedPassword.Text)
            {
                ShowWarning("Confirmation password does not match. Please check again.");
                return false;
            }

            if (!IsUsernameFormatValid(txtLoginName.Text))
            {
                ShowWarning("Username format is invalid. Must contain 'HQ' or '-', and be 8 characters long.");
                return false;
            }

            return true;
        }

        private bool IsUsernameFormatValid(string username)
        {
            return (username.Contains("HQ") || username.Contains("-")) && username.Length == 8;
        }

        private bool IsUsernameExists(SqlConnection conn, string username)
        {
            const string query = "SELECT COUNT(LoginName) FROM Login WHERE LoginName = @LoginName";

            using (var cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@LoginName", username);
                return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
            }
        }

        private bool SaveNewUser(SqlConnection conn)
        {
            const string insertQuery = @"
                INSERT INTO Login (LoginName, TenCC, LoginPassword, Decentralization) 
                VALUES (@LoginName, @TenCC, @LoginPassword, @Decentralization)";

            using (var cmd = new SqlCommand(insertQuery, conn))
            {
                cmd.Parameters.AddWithValue("@LoginName", txtLoginName.Text.Trim());
                cmd.Parameters.AddWithValue("@TenCC", txtTenCC.Text.Trim());
                cmd.Parameters.AddWithValue("@LoginPassword", Encrypt(txtLoginPassword.Text.Trim()));
                cmd.Parameters.AddWithValue("@Decentralization", false);

                int rowsAffected = cmd.ExecuteNonQuery();
                return rowsAffected > 0;
            }
        }

        private void RefreshUserDataGrid(SqlConnection conn)
        {
            const string selectQuery = @"
                SELECT 
                    ROW_NUMBER() OVER (ORDER BY LoginName DESC) AS [STT], 
                    LoginName as [User], 
                    TenCC as [Officer Name], 
                    LoginPassword as [Password], 
                    Decentralization, 
                    NgaynhapHT 
                FROM Login";

            using (var cmd = new SqlCommand(selectQuery, conn))
            using (var sda = new SqlDataAdapter(cmd))
            {
                var dt = new DataTable();
                sda.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    dgvUserInfo.DataSource = dt;
                    SetupDataGridView();
                    BindFormFields();
                }
            }
        }

        private void BindFormFields()
        {
            txtLoginName.DataBindings.Clear();
            txtLoginName.DataBindings.Add("Text", dgvUserInfo.DataSource, "User");

            txtTenCC.DataBindings.Clear();
            txtTenCC.DataBindings.Add("Text", dgvUserInfo.DataSource, "Officer Name");

            txtLoginPassword.DataBindings.Clear();
            txtLoginPassword.DataBindings.Add("Text", dgvUserInfo.DataSource, "Password");

            txtConfirmedPassword.DataBindings.Clear();
            txtConfirmedPassword.DataBindings.Add("Text", dgvUserInfo.DataSource, "Password");
        }

        private void ShowSuccess(string message)
        {
            MessageBox.Show(message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowWarning(string message)
        {
            MessageBox.Show(message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        private bool ValidInput()
        {
            if (txtLoginName.Text.Trim() == "" || txtLoginPassword.Text.Trim() == "")
            {
                MessageBox.Show("Please input username and password.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
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
            ut = new Utility();
            var conn = ut.OpenDB();
            DateTime strngaynhapht = DateTime.Now;
            string ngaynhapht = strngaynhapht.ToString("yyyy-MM-dd hh:mm:ss tt");

            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                    ValidInput();
                    string username = txtLoginName.Text.Trim();
                    string password = Encrypt(txtLoginPassword.Text.Trim());

                    cmd = new SqlCommand("", conn);
                    sda = new SqlDataAdapter("Select * From Login " +
                        "Where LoginName = N'" + username + "' and LoginPassword = N'" + password + "'", conn);
                    dt = new DataTable();
                    sda.Fill(dt);

                    if (dt.Rows.Count == 1)
                    {
                        //frmMain opMfrm = new frmMain(dt.Rows[0][0].ToString(), dt.Rows[0][3].ToString());
                        //opMfrm.StartPosition = FormStartPosition.CenterScreen;
                        //opMfrm.Show();
                        //Visible = false;
                        //this.Hide();

                        this.Hide();
                        MDIParentMain frmMDI = new MDIParentMain(dt.Rows[0][0].ToString(), dt.Rows[0][1].ToString(), dt.Rows[0][2].ToString(), dt.Rows[0][3].ToString());

                        //frmMDI.Show();
                        //frmMDI.StartPosition = FormStartPosition.CenterScreen;
                        //frmMDI.WindowState = FormWindowState.Maximized;
                        frmMDI.Refresh();
                        cmd = new SqlCommand("", conn);

                        if (dt.Rows[0][3].ToString() == "True")
                        {
                            this.Hide();

                            //this.Show();
                            //lblPasswordConfirm.Enabled = true;
                            //btnRegister.Enabled = true;
                            //btnSave.Enabled = true;

                            cmd.CommandText = "Update Login Set NgaynhapHT = @NgaynhapHT Where LoginName = N'" + username + "'";
                            cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;

                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                        }
                        else
                        {
                            this.Hide();
                            //frmMDI.Show();
                            //frmMDI.StartPosition = FormStartPosition.CenterScreen;
                            //frmMDI.WindowState = FormWindowState.Maximized;
                            frmMDI.Refresh();
                            cmd.CommandText = "Update Login Set NgaynhapHT = @NgaynhapHT Where LoginName = N'" + username + "'";
                            cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;

                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Login failed. Please check username and password.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                conn.Close();
                conn.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();           
        }

        private void txtTaikhoan_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            textBox.Text = textBox.Text.ToUpper();
            textBox.SelectionStart = textBox.Text.Length; // Keep cursor at the end

            try
            {
                if (txtLoginName.TextLength > 0)
                {
                    txtLoginPassword.Enabled = true;
                }
                else {txtLoginName.Enabled = false; }   
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txtMatkhau_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtLoginPassword.TextLength > 0)
                {
                    txtConfirmedPassword.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txtLoginName_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = char.ToUpper(e.KeyChar);
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            ut = new Utility();
            var conn = ut.OpenDB();

            try
            {
                DialogResult dr = MessageBox.Show("Are you sure to reset passcode?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dr == DialogResult.Yes)
                {
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                        string username = txtLoginName.Text.Trim();
                        string password = Encrypt(txtLoginPassword.Text.Trim());
                        string newpassword = Encrypt(123456.ToString());

                        cmd = new SqlCommand("", conn);
                        sda = new SqlDataAdapter("Select User From Login Where LoginName = N'" + username + "'", conn);
                        dt = new DataTable();
                        sda.Fill(dt);

                        if (dt.Rows.Count == 1)
                        {
                            // Data binding
                            txtLoginName.DataBindings.Clear();
                            txtLoginName.DataBindings.Add("Text", dgvUserInfo.DataSource, "User");
                            txtTenCC.DataBindings.Clear();
                            txtTenCC.DataBindings.Add("Text", dgvUserInfo.DataSource, "Officer Name");
                            txtLoginPassword.DataBindings.Clear();
                            txtLoginPassword.DataBindings.Add("Text", dgvUserInfo.DataSource, "Password");
                            txtConfirmedPassword.DataBindings.Clear();
                            txtConfirmedPassword.DataBindings.Add("Text", dgvUserInfo.DataSource, "Password");

                            string tencc = txtTenCC.Text.Trim();

                            cmd.CommandText = "Update Login Set LoginPassword = @LoginPassword, TenCC = @TenCC " +
                                "Where LoginName = N'" + username + "'";
                            cmd.Parameters.Add("@LoginPassword", SqlDbType.NChar).Value = newpassword;
                            cmd.Parameters.Add("@TenCC", SqlDbType.NVarChar).Value = tencc;

                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                            this.Refresh();
                        }

                        cmd = new SqlCommand("Select distinct Count(LoginName) From Login " +
                            "Where LoginName = N'" + username + "' and LoginPassword = N'" + newpassword + "'", conn);
                        int count = Convert.ToInt32(cmd.ExecuteScalar());

                        if (count > 0)
                        {
                            MessageBox.Show("Passcode reset successfully.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LoadData();
                            this.Refresh();
                        }
                        else
                        {
                            MessageBox.Show("Passcode reset failed.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    conn.Close();
                    conn.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            ut = new Utility();
            var conn = ut.OpenDB();

            try
            {
                DialogResult dr = MessageBox.Show("Are you sure to delete this user account?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dr == DialogResult.Yes)
                {
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();

                        string username = txtLoginName.Text.Trim();

                        cmd = new SqlCommand("Select distinct count(LoginName) From Login", conn);
                        int count_beforedel = Convert.ToInt32(cmd.ExecuteScalar());
                        
                        cmd.CommandText = "Delete From Login " +
                                "Where LoginName = @LoginName";
                        cmd.Parameters.Add("@LoginName", SqlDbType.NChar).Value = username;
                        cmd.CommandType = CommandType.Text;
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                        this.Refresh();

                        cmd = new SqlCommand("Select distinct count(LoginName) From Login", conn);
                        int count_afterdel = Convert.ToInt32(cmd.ExecuteScalar());
                        int del_row = count_beforedel - count_afterdel;

                        if (del_row > 0)
                        {
                            MessageBox.Show("User acoount deleted successfully.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            
                            cmd = new SqlCommand(@"Select ROW_NUMBER() OVER (ORDER BY LoginName DESC) AS [STT], LoginName as [User], TenCC as [Officer Name], 
                                            LoginPassword as [Password], Decentralization, NgaynhapHT 
                                            From Login", conn);

                            sda = new SqlDataAdapter(cmd);
                            dt = new System.Data.DataTable();
                            sda.Fill(dt);

                            if (dt.Rows.Count > 0)
                            {
                                dgvUserInfo.DataSource = dt;
                                SetupDataGridView();

                                // Data binding
                                txtLoginName.DataBindings.Clear();
                                txtLoginName.DataBindings.Add("Text", dgvUserInfo.DataSource, "User");
                                txtTenCC.DataBindings.Clear();
                                txtTenCC.DataBindings.Add("Text", dgvUserInfo.DataSource, "Officer Name");
                                txtLoginPassword.DataBindings.Clear();
                                txtLoginPassword.DataBindings.Add("Text", dgvUserInfo.DataSource, "Password");
                                txtConfirmedPassword.DataBindings.Clear();
                                txtConfirmedPassword.DataBindings.Add("Text", dgvUserInfo.DataSource, "Password");
                            }
                            else
                            {
                                MessageBox.Show("Username is registered failed. Please contact with system administrator.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            }
                            dgvUserInfo.Refresh();
                        }
                        else 
                        {
                            MessageBox.Show("User acoount deleted unsuccessfully.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    // Close connection
                    conn.Close();
                    conn.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
