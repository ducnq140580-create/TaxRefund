
namespace TaxRefund
{
    partial class frmChangePassword
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
            this.lblCurrentPassword = new System.Windows.Forms.Label();
            this.lblLoginName = new System.Windows.Forms.Label();
            this.lblNewPassword = new System.Windows.Forms.Label();
            this.txtNewPassword = new System.Windows.Forms.TextBox();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblPasswordConfirm = new System.Windows.Forms.Label();
            this.txtPasswordconfirm = new System.Windows.Forms.TextBox();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.lblMatkhau = new System.Windows.Forms.Label();
            this.txtLoginPassword = new System.Windows.Forms.TextBox();
            this.lblUserName = new System.Windows.Forms.Label();
            this.txtLoginName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBoxClose = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxClose)).BeginInit();
            this.SuspendLayout();
            // 
            // lblCurrentPassword
            // 
            this.lblCurrentPassword.AutoSize = true;
            this.lblCurrentPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCurrentPassword.Location = new System.Drawing.Point(59, 242);
            this.lblCurrentPassword.Name = "lblCurrentPassword";
            this.lblCurrentPassword.Size = new System.Drawing.Size(131, 20);
            this.lblCurrentPassword.TabIndex = 161;
            this.lblCurrentPassword.Text = "CurrentPassword";
            this.lblCurrentPassword.Visible = false;
            // 
            // lblLoginName
            // 
            this.lblLoginName.AutoSize = true;
            this.lblLoginName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLoginName.Location = new System.Drawing.Point(59, 211);
            this.lblLoginName.Name = "lblLoginName";
            this.lblLoginName.Size = new System.Drawing.Size(90, 20);
            this.lblLoginName.TabIndex = 160;
            this.lblLoginName.Text = "LoginName";
            this.lblLoginName.Visible = false;
            // 
            // lblNewPassword
            // 
            this.lblNewPassword.AutoSize = true;
            this.lblNewPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lblNewPassword.ForeColor = System.Drawing.Color.DeepSkyBlue;
            this.lblNewPassword.Location = new System.Drawing.Point(116, 144);
            this.lblNewPassword.Name = "lblNewPassword";
            this.lblNewPassword.Size = new System.Drawing.Size(117, 20);
            this.lblNewPassword.TabIndex = 159;
            this.lblNewPassword.Text = "New Password:";
            // 
            // txtNewPassword
            // 
            this.txtNewPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.txtNewPassword.Location = new System.Drawing.Point(240, 141);
            this.txtNewPassword.Name = "txtNewPassword";
            this.txtNewPassword.PasswordChar = '*';
            this.txtNewPassword.Size = new System.Drawing.Size(181, 26);
            this.txtNewPassword.TabIndex = 152;
            this.txtNewPassword.TextChanged += new System.EventHandler(this.txtNewPassword_TextChanged);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatus.Location = new System.Drawing.Point(297, 34);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(56, 20);
            this.lblStatus.TabIndex = 158;
            this.lblStatus.Text = "Status";
            this.lblStatus.Visible = false;
            // 
            // lblPasswordConfirm
            // 
            this.lblPasswordConfirm.AutoSize = true;
            this.lblPasswordConfirm.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lblPasswordConfirm.ForeColor = System.Drawing.Color.DeepSkyBlue;
            this.lblPasswordConfirm.Location = new System.Drawing.Point(92, 176);
            this.lblPasswordConfirm.Name = "lblPasswordConfirm";
            this.lblPasswordConfirm.Size = new System.Drawing.Size(141, 20);
            this.lblPasswordConfirm.TabIndex = 157;
            this.lblPasswordConfirm.Text = "Password Confirm:";
            // 
            // txtPasswordconfirm
            // 
            this.txtPasswordconfirm.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.txtPasswordconfirm.Location = new System.Drawing.Point(239, 173);
            this.txtPasswordconfirm.Name = "txtPasswordconfirm";
            this.txtPasswordconfirm.PasswordChar = '*';
            this.txtPasswordconfirm.Size = new System.Drawing.Size(181, 26);
            this.txtPasswordconfirm.TabIndex = 153;
            this.txtPasswordconfirm.TextChanged += new System.EventHandler(this.txtPasswordconfirm_TextChanged);
            // 
            // btnSubmit
            // 
            this.btnSubmit.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btnSubmit.ForeColor = System.Drawing.Color.DeepSkyBlue;
            this.btnSubmit.Location = new System.Drawing.Point(299, 229);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(121, 35);
            this.btnSubmit.TabIndex = 155;
            this.btnSubmit.Text = "Submit";
            this.btnSubmit.UseVisualStyleBackColor = true;
            this.btnSubmit.Click += new System.EventHandler(this.btnChangePassword_Click);
            // 
            // lblMatkhau
            // 
            this.lblMatkhau.AutoSize = true;
            this.lblMatkhau.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lblMatkhau.ForeColor = System.Drawing.Color.DeepSkyBlue;
            this.lblMatkhau.Location = new System.Drawing.Point(92, 115);
            this.lblMatkhau.Name = "lblMatkhau";
            this.lblMatkhau.Size = new System.Drawing.Size(139, 20);
            this.lblMatkhau.TabIndex = 154;
            this.lblMatkhau.Text = "Current Password:";
            // 
            // txtLoginPassword
            // 
            this.txtLoginPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.txtLoginPassword.Location = new System.Drawing.Point(240, 109);
            this.txtLoginPassword.Name = "txtLoginPassword";
            this.txtLoginPassword.PasswordChar = '*';
            this.txtLoginPassword.Size = new System.Drawing.Size(181, 26);
            this.txtLoginPassword.TabIndex = 150;
            this.txtLoginPassword.TextChanged += new System.EventHandler(this.txtLoginPassword_TextChanged);
            // 
            // lblUserName
            // 
            this.lblUserName.AutoSize = true;
            this.lblUserName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lblUserName.ForeColor = System.Drawing.Color.DeepSkyBlue;
            this.lblUserName.Location = new System.Drawing.Point(141, 80);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(93, 20);
            this.lblUserName.TabIndex = 151;
            this.lblUserName.Text = "User Name:";
            // 
            // txtLoginName
            // 
            this.txtLoginName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.txtLoginName.Location = new System.Drawing.Point(240, 77);
            this.txtLoginName.Name = "txtLoginName";
            this.txtLoginName.Size = new System.Drawing.Size(181, 26);
            this.txtLoginName.TabIndex = 149;
            this.txtLoginName.TextChanged += new System.EventHandler(this.txtLoginName_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.DeepSkyBlue;
            this.label1.Location = new System.Drawing.Point(12, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(232, 29);
            this.label1.TabIndex = 148;
            this.label1.Text = "Change Password:";
            // 
            // pictureBoxClose
            // 
            this.pictureBoxClose.Cursor = System.Windows.Forms.Cursors.Default;
            this.pictureBoxClose.Image = global::TaxRefund.Properties.Resources.icons8_close_48;
            this.pictureBoxClose.Location = new System.Drawing.Point(488, 6);
            this.pictureBoxClose.Name = "pictureBoxClose";
            this.pictureBoxClose.Size = new System.Drawing.Size(19, 20);
            this.pictureBoxClose.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBoxClose.TabIndex = 162;
            this.pictureBoxClose.TabStop = false;
            this.pictureBoxClose.Click += new System.EventHandler(this.pictureBoxClose_Click);
            // 
            // frmChangePassword
            // 
            this.AcceptButton = this.btnSubmit;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FloralWhite;
            this.ClientSize = new System.Drawing.Size(514, 286);
            this.Controls.Add(this.pictureBoxClose);
            this.Controls.Add(this.lblCurrentPassword);
            this.Controls.Add(this.lblLoginName);
            this.Controls.Add(this.lblNewPassword);
            this.Controls.Add(this.txtNewPassword);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.lblPasswordConfirm);
            this.Controls.Add(this.txtPasswordconfirm);
            this.Controls.Add(this.btnSubmit);
            this.Controls.Add(this.lblMatkhau);
            this.Controls.Add(this.txtLoginPassword);
            this.Controls.Add(this.lblUserName);
            this.Controls.Add(this.txtLoginName);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frmChangePassword";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Load += new System.EventHandler(this.frmChangePassword_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxClose)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblCurrentPassword;
        private System.Windows.Forms.Label lblLoginName;
        private System.Windows.Forms.Label lblNewPassword;
        private System.Windows.Forms.TextBox txtNewPassword;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblPasswordConfirm;
        private System.Windows.Forms.TextBox txtPasswordconfirm;
        private System.Windows.Forms.Button btnSubmit;
        private System.Windows.Forms.Label lblMatkhau;
        private System.Windows.Forms.TextBox txtLoginPassword;
        private System.Windows.Forms.Label lblUserName;
        private System.Windows.Forms.TextBox txtLoginName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBoxClose;
    }
}