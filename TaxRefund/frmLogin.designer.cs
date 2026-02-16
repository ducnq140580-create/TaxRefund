
namespace TaxRefund
{
    partial class frmLogin
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmLogin));
            this.label1 = new System.Windows.Forms.Label();
            this.lblIPaddress = new System.Windows.Forms.Label();
            this.txtUsername = new System.Windows.Forms.TextBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.pictureBoxClose = new System.Windows.Forms.PictureBox();
            this.pictureBoxTogglePassword = new System.Windows.Forms.PictureBox();
            this.pictureBoxToggleUsername = new System.Windows.Forms.PictureBox();
            this.pictureBoxToggleShowPassword = new System.Windows.Forms.PictureBox();
            this.btnLogin = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxClose)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxTogglePassword)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxToggleUsername)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxToggleShowPassword)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label1.ForeColor = System.Drawing.Color.DeepSkyBlue;
            this.label1.Location = new System.Drawing.Point(38, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(129, 25);
            this.label1.TabIndex = 121;
            this.label1.Text = "TRS Login:";
            // 
            // lblIPaddress
            // 
            this.lblIPaddress.AutoSize = true;
            this.lblIPaddress.Location = new System.Drawing.Point(227, 50);
            this.lblIPaddress.Name = "lblIPaddress";
            this.lblIPaddress.Size = new System.Drawing.Size(64, 13);
            this.lblIPaddress.TabIndex = 127;
            this.lblIPaddress.Text = "lblIPaddress";
            this.lblIPaddress.Visible = false;
            // 
            // txtUsername
            // 
            this.txtUsername.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUsername.Location = new System.Drawing.Point(110, 85);
            this.txtUsername.Multiline = true;
            this.txtUsername.Name = "txtUsername";
            this.txtUsername.Size = new System.Drawing.Size(252, 40);
            this.txtUsername.TabIndex = 128;
            this.txtUsername.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtUsername.TextChanged += new System.EventHandler(this.txtUsername_TextChanged);
            this.txtUsername.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtUsername_KeyPress);
            // 
            // txtPassword
            // 
            this.txtPassword.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPassword.Location = new System.Drawing.Point(110, 145);
            this.txtPassword.Multiline = true;
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(252, 40);
            this.txtPassword.TabIndex = 129;
            this.txtPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtPassword.UseSystemPasswordChar = true;
            this.txtPassword.TextChanged += new System.EventHandler(this.txtPassword_TextChanged);
            // 
            // pictureBoxClose
            // 
            this.pictureBoxClose.Cursor = System.Windows.Forms.Cursors.Default;
            this.pictureBoxClose.Image = global::TaxRefund.Properties.Resources.icons8_close_48;
            this.pictureBoxClose.Location = new System.Drawing.Point(432, 9);
            this.pictureBoxClose.Name = "pictureBoxClose";
            this.pictureBoxClose.Size = new System.Drawing.Size(19, 20);
            this.pictureBoxClose.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBoxClose.TabIndex = 126;
            this.pictureBoxClose.TabStop = false;
            this.pictureBoxClose.Click += new System.EventHandler(this.pictureBoxClose_Click);
            // 
            // pictureBoxTogglePassword
            // 
            this.pictureBoxTogglePassword.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxTogglePassword.Cursor = System.Windows.Forms.Cursors.Default;
            this.pictureBoxTogglePassword.Image = global::TaxRefund.Properties.Resources.internet_lock_locked_padlock_password_secure_security_icon_127100;
            this.pictureBoxTogglePassword.Location = new System.Drawing.Point(64, 148);
            this.pictureBoxTogglePassword.Name = "pictureBoxTogglePassword";
            this.pictureBoxTogglePassword.Size = new System.Drawing.Size(40, 34);
            this.pictureBoxTogglePassword.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBoxTogglePassword.TabIndex = 125;
            this.pictureBoxTogglePassword.TabStop = false;
            // 
            // pictureBoxToggleUsername
            // 
            this.pictureBoxToggleUsername.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxToggleUsername.Cursor = System.Windows.Forms.Cursors.Default;
            this.pictureBoxToggleUsername.Image = global::TaxRefund.Properties.Resources.ic_account_child_128_28130;
            this.pictureBoxToggleUsername.Location = new System.Drawing.Point(64, 88);
            this.pictureBoxToggleUsername.Name = "pictureBoxToggleUsername";
            this.pictureBoxToggleUsername.Size = new System.Drawing.Size(40, 34);
            this.pictureBoxToggleUsername.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBoxToggleUsername.TabIndex = 124;
            this.pictureBoxToggleUsername.TabStop = false;
            // 
            // pictureBoxToggleShowPassword
            // 
            this.pictureBoxToggleShowPassword.Image = global::TaxRefund.Properties.Resources.icons8_eye_opened_30_1;
            this.pictureBoxToggleShowPassword.Location = new System.Drawing.Point(368, 160);
            this.pictureBoxToggleShowPassword.Name = "pictureBoxToggleShowPassword";
            this.pictureBoxToggleShowPassword.Size = new System.Drawing.Size(28, 20);
            this.pictureBoxToggleShowPassword.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBoxToggleShowPassword.TabIndex = 123;
            this.pictureBoxToggleShowPassword.TabStop = false;
            this.pictureBoxToggleShowPassword.Click += new System.EventHandler(this.pictureBoxToggleShowPassword_Click);
            this.pictureBoxToggleShowPassword.MouseEnter += new System.EventHandler(this.pictureBoxTogglePassword_MouseEnter);
            this.pictureBoxToggleShowPassword.MouseLeave += new System.EventHandler(this.pictureBoxTogglePassword_MouseLeave);
            // 
            // btnLogin
            // 
            this.btnLogin.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnLogin.BackColor = System.Drawing.Color.Transparent;
            this.btnLogin.BackgroundImage = global::TaxRefund.Properties.Resources.icons8_login_48;
            this.btnLogin.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.btnLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btnLogin.ForeColor = System.Drawing.Color.Blue;
            this.btnLogin.Location = new System.Drawing.Point(170, 216);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(132, 49);
            this.btnLogin.TabIndex = 3;
            this.btnLogin.UseVisualStyleBackColor = false;
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // frmLogin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FloralWhite;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(460, 318);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.txtUsername);
            this.Controls.Add(this.lblIPaddress);
            this.Controls.Add(this.pictureBoxClose);
            this.Controls.Add(this.pictureBoxTogglePassword);
            this.Controls.Add(this.pictureBoxToggleUsername);
            this.Controls.Add(this.pictureBoxToggleShowPassword);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnLogin);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmLogin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Login";
            this.Load += new System.EventHandler(this.frmLogin_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxClose)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxTogglePassword)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxToggleUsername)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxToggleShowPassword)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnLogin;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBoxToggleShowPassword;
        private System.Windows.Forms.PictureBox pictureBoxToggleUsername;
        private System.Windows.Forms.PictureBox pictureBoxTogglePassword;
        private System.Windows.Forms.PictureBox pictureBoxClose;
        private System.Windows.Forms.Label lblIPaddress;
        private System.Windows.Forms.TextBox txtUsername;
        private System.Windows.Forms.TextBox txtPassword;
    }
}