namespace TaxRefund
{
    partial class frmRegister
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
            this.btnSave = new System.Windows.Forms.Button();
            this.lblPasswordConfirm = new System.Windows.Forms.Label();
            this.txtConfirmedPassword = new System.Windows.Forms.TextBox();
            this.btnRegister = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnLogin = new System.Windows.Forms.Button();
            this.lblMatkhau = new System.Windows.Forms.Label();
            this.txtLoginPassword = new System.Windows.Forms.TextBox();
            this.lblLoginName = new System.Windows.Forms.Label();
            this.txtLoginName = new System.Windows.Forms.TextBox();
            this.lblstatus = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtTenCC = new System.Windows.Forms.TextBox();
            this.dgvUserInfo = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnReset = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvUserInfo)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnSave
            // 
            this.btnSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btnSave.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.btnSave.Location = new System.Drawing.Point(123, 246);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(104, 35);
            this.btnSave.TabIndex = 127;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lblPasswordConfirm
            // 
            this.lblPasswordConfirm.AutoSize = true;
            this.lblPasswordConfirm.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lblPasswordConfirm.ForeColor = System.Drawing.Color.Blue;
            this.lblPasswordConfirm.Location = new System.Drawing.Point(5, 196);
            this.lblPasswordConfirm.Name = "lblPasswordConfirm";
            this.lblPasswordConfirm.Size = new System.Drawing.Size(142, 20);
            this.lblPasswordConfirm.TabIndex = 126;
            this.lblPasswordConfirm.Text = "Re-type Password:";
            // 
            // txtConfirmedPassword
            // 
            this.txtConfirmedPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.txtConfirmedPassword.Location = new System.Drawing.Point(152, 193);
            this.txtConfirmedPassword.Name = "txtConfirmedPassword";
            this.txtConfirmedPassword.PasswordChar = '*';
            this.txtConfirmedPassword.Size = new System.Drawing.Size(181, 26);
            this.txtConfirmedPassword.TabIndex = 120;
            // 
            // btnRegister
            // 
            this.btnRegister.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btnRegister.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.btnRegister.Location = new System.Drawing.Point(231, 246);
            this.btnRegister.Name = "btnRegister";
            this.btnRegister.Size = new System.Drawing.Size(104, 35);
            this.btnRegister.TabIndex = 124;
            this.btnRegister.Text = "Register";
            this.btnRegister.UseVisualStyleBackColor = true;
            this.btnRegister.Click += new System.EventHandler(this.btnRegister_Click);
            // 
            // btnExit
            // 
            this.btnExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btnExit.ForeColor = System.Drawing.Color.Red;
            this.btnExit.Location = new System.Drawing.Point(229, 397);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(104, 35);
            this.btnExit.TabIndex = 122;
            this.btnExit.Text = "Close";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnLogin
            // 
            this.btnLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btnLogin.ForeColor = System.Drawing.Color.Blue;
            this.btnLogin.Location = new System.Drawing.Point(121, 397);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(104, 35);
            this.btnLogin.TabIndex = 121;
            this.btnLogin.Text = "Login";
            this.btnLogin.UseVisualStyleBackColor = true;
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // lblMatkhau
            // 
            this.lblMatkhau.AutoSize = true;
            this.lblMatkhau.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lblMatkhau.ForeColor = System.Drawing.Color.Blue;
            this.lblMatkhau.Location = new System.Drawing.Point(65, 162);
            this.lblMatkhau.Name = "lblMatkhau";
            this.lblMatkhau.Size = new System.Drawing.Size(82, 20);
            this.lblMatkhau.TabIndex = 120;
            this.lblMatkhau.Text = "Password:";
            // 
            // txtLoginPassword
            // 
            this.txtLoginPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.txtLoginPassword.Location = new System.Drawing.Point(153, 159);
            this.txtLoginPassword.Name = "txtLoginPassword";
            this.txtLoginPassword.PasswordChar = '*';
            this.txtLoginPassword.Size = new System.Drawing.Size(181, 26);
            this.txtLoginPassword.TabIndex = 119;
            this.txtLoginPassword.TextChanged += new System.EventHandler(this.txtMatkhau_TextChanged);
            // 
            // lblLoginName
            // 
            this.lblLoginName.AutoSize = true;
            this.lblLoginName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lblLoginName.ForeColor = System.Drawing.Color.Blue;
            this.lblLoginName.Location = new System.Drawing.Point(54, 93);
            this.lblLoginName.Name = "lblLoginName";
            this.lblLoginName.Size = new System.Drawing.Size(93, 20);
            this.lblLoginName.TabIndex = 118;
            this.lblLoginName.Text = "User Name:";
            // 
            // txtLoginName
            // 
            this.txtLoginName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.txtLoginName.Location = new System.Drawing.Point(153, 90);
            this.txtLoginName.Name = "txtLoginName";
            this.txtLoginName.Size = new System.Drawing.Size(181, 26);
            this.txtLoginName.TabIndex = 117;
            this.txtLoginName.TextChanged += new System.EventHandler(this.txtTaikhoan_TextChanged);
            this.txtLoginName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtLoginName_KeyPress);
            // 
            // lblstatus
            // 
            this.lblstatus.AutoSize = true;
            this.lblstatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblstatus.Location = new System.Drawing.Point(26, 9);
            this.lblstatus.Name = "lblstatus";
            this.lblstatus.Size = new System.Drawing.Size(68, 20);
            this.lblstatus.TabIndex = 132;
            this.lblstatus.Text = "lblstatus";
            this.lblstatus.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label1.ForeColor = System.Drawing.Color.Blue;
            this.label1.Location = new System.Drawing.Point(40, 127);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(106, 20);
            this.label1.TabIndex = 135;
            this.label1.Text = "Officer Name:";
            // 
            // txtTenCC
            // 
            this.txtTenCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.txtTenCC.Location = new System.Drawing.Point(152, 124);
            this.txtTenCC.Name = "txtTenCC";
            this.txtTenCC.Size = new System.Drawing.Size(181, 26);
            this.txtTenCC.TabIndex = 118;
            // 
            // dgvUserInfo
            // 
            this.dgvUserInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvUserInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvUserInfo.Location = new System.Drawing.Point(372, 48);
            this.dgvUserInfo.Name = "dgvUserInfo";
            this.dgvUserInfo.Size = new System.Drawing.Size(694, 457);
            this.dgvUserInfo.TabIndex = 136;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnReset);
            this.panel1.Controls.Add(this.btnDel);
            this.panel1.Controls.Add(this.lblLoginName);
            this.panel1.Controls.Add(this.txtLoginName);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.txtLoginPassword);
            this.panel1.Controls.Add(this.txtTenCC);
            this.panel1.Controls.Add(this.lblMatkhau);
            this.panel1.Controls.Add(this.btnLogin);
            this.panel1.Controls.Add(this.btnExit);
            this.panel1.Controls.Add(this.btnSave);
            this.panel1.Controls.Add(this.lblPasswordConfirm);
            this.panel1.Controls.Add(this.btnRegister);
            this.panel1.Controls.Add(this.txtConfirmedPassword);
            this.panel1.Location = new System.Drawing.Point(12, 49);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(354, 456);
            this.panel1.TabIndex = 137;
            // 
            // btnReset
            // 
            this.btnReset.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btnReset.ForeColor = System.Drawing.Color.SeaGreen;
            this.btnReset.Location = new System.Drawing.Point(123, 287);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(104, 35);
            this.btnReset.TabIndex = 136;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // btnDel
            // 
            this.btnDel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btnDel.ForeColor = System.Drawing.Color.Brown;
            this.btnDel.Location = new System.Drawing.Point(231, 287);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(104, 35);
            this.btnDel.TabIndex = 137;
            this.btnDel.Text = "Delete";
            this.btnDel.UseVisualStyleBackColor = true;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // frmRegister
            // 
            this.AcceptButton = this.btnSave;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(1090, 532);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dgvUserInfo);
            this.Controls.Add(this.lblstatus);
            this.Name = "frmRegister";
            this.Load += new System.EventHandler(this.frmRegister_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvUserInfo)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label lblPasswordConfirm;
        private System.Windows.Forms.TextBox txtConfirmedPassword;
        private System.Windows.Forms.Button btnRegister;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnLogin;
        private System.Windows.Forms.Label lblMatkhau;
        private System.Windows.Forms.TextBox txtLoginPassword;
        private System.Windows.Forms.Label lblLoginName;
        private System.Windows.Forms.TextBox txtLoginName;
        private System.Windows.Forms.Label lblstatus;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtTenCC;
        private System.Windows.Forms.DataGridView dgvUserInfo;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.Button btnDel;
    }
}