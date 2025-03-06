
namespace Echeck_LogFile_Process
{
    partial class Login
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Login));
            this.LogoPictureBox = new System.Windows.Forms.PictureBox();
            this.txtUsername = new System.Windows.Forms.TextBox();
            this.Cancel = new System.Windows.Forms.Button();
            this.UsernameLabel = new System.Windows.Forms.Label();
            this.OK = new System.Windows.Forms.Button();
            this.PasswordLabel = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.GBUserInfo = new System.Windows.Forms.GroupBox();
            this.GBDept = new System.Windows.Forms.GroupBox();
            this.rbPE = new System.Windows.Forms.RadioButton();
            this.rbPro = new System.Windows.Forms.RadioButton();
            this.rbNPI = new System.Windows.Forms.RadioButton();
            this.rbDE = new System.Windows.Forms.RadioButton();
            this.rbQA = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.LogoPictureBox)).BeginInit();
            this.GBUserInfo.SuspendLayout();
            this.GBDept.SuspendLayout();
            this.SuspendLayout();
            // 
            // LogoPictureBox
            // 
            this.LogoPictureBox.Image = ((System.Drawing.Image)(resources.GetObject("LogoPictureBox.Image")));
            this.LogoPictureBox.Location = new System.Drawing.Point(11, 11);
            this.LogoPictureBox.Margin = new System.Windows.Forms.Padding(2);
            this.LogoPictureBox.Name = "LogoPictureBox";
            this.LogoPictureBox.Size = new System.Drawing.Size(192, 208);
            this.LogoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.LogoPictureBox.TabIndex = 9;
            this.LogoPictureBox.TabStop = false;
            // 
            // txtUsername
            // 
            this.txtUsername.Font = new System.Drawing.Font("Arial", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUsername.Location = new System.Drawing.Point(20, 55);
            this.txtUsername.Margin = new System.Windows.Forms.Padding(2);
            this.txtUsername.Name = "txtUsername";
            this.txtUsername.Size = new System.Drawing.Size(166, 23);
            this.txtUsername.TabIndex = 1;
            this.txtUsername.Text = "Admin";
            // 
            // Cancel
            // 
            this.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Cancel.Font = new System.Drawing.Font("Arial", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Cancel.Location = new System.Drawing.Point(112, 168);
            this.Cancel.Margin = new System.Windows.Forms.Padding(2);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(70, 26);
            this.Cancel.TabIndex = 5;
            this.Cancel.Text = "&Cancel";
            this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // UsernameLabel
            // 
            this.UsernameLabel.Font = new System.Drawing.Font("Arial", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UsernameLabel.Location = new System.Drawing.Point(19, 25);
            this.UsernameLabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.UsernameLabel.Name = "UsernameLabel";
            this.UsernameLabel.Size = new System.Drawing.Size(165, 19);
            this.UsernameLabel.TabIndex = 0;
            this.UsernameLabel.Text = "&User ID";
            this.UsernameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // OK
            // 
            this.OK.Font = new System.Drawing.Font("Arial", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OK.Location = new System.Drawing.Point(19, 168);
            this.OK.Margin = new System.Windows.Forms.Padding(2);
            this.OK.Name = "OK";
            this.OK.Size = new System.Drawing.Size(70, 26);
            this.OK.TabIndex = 4;
            this.OK.Text = "&OK";
            this.OK.Click += new System.EventHandler(this.OK_Click);
            // 
            // PasswordLabel
            // 
            this.PasswordLabel.Font = new System.Drawing.Font("Arial", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PasswordLabel.Location = new System.Drawing.Point(17, 98);
            this.PasswordLabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.PasswordLabel.Name = "PasswordLabel";
            this.PasswordLabel.Size = new System.Drawing.Size(165, 19);
            this.PasswordLabel.TabIndex = 2;
            this.PasswordLabel.Text = "&Password";
            this.PasswordLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtPassword
            // 
            this.txtPassword.Font = new System.Drawing.Font("Arial", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPassword.Location = new System.Drawing.Point(19, 125);
            this.txtPassword.Margin = new System.Windows.Forms.Padding(2);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(166, 23);
            this.txtPassword.TabIndex = 3;
            // 
            // GBUserInfo
            // 
            this.GBUserInfo.Controls.Add(this.GBDept);
            this.GBUserInfo.Controls.Add(this.txtUsername);
            this.GBUserInfo.Controls.Add(this.Cancel);
            this.GBUserInfo.Controls.Add(this.UsernameLabel);
            this.GBUserInfo.Controls.Add(this.OK);
            this.GBUserInfo.Controls.Add(this.PasswordLabel);
            this.GBUserInfo.Controls.Add(this.txtPassword);
            this.GBUserInfo.Font = new System.Drawing.Font("Arial", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GBUserInfo.Location = new System.Drawing.Point(208, 11);
            this.GBUserInfo.Margin = new System.Windows.Forms.Padding(2);
            this.GBUserInfo.Name = "GBUserInfo";
            this.GBUserInfo.Padding = new System.Windows.Forms.Padding(2);
            this.GBUserInfo.Size = new System.Drawing.Size(334, 208);
            this.GBUserInfo.TabIndex = 10;
            this.GBUserInfo.TabStop = false;
            this.GBUserInfo.Text = "User Information";
            // 
            // GBDept
            // 
            this.GBDept.Controls.Add(this.rbPE);
            this.GBDept.Controls.Add(this.rbPro);
            this.GBDept.Controls.Add(this.rbNPI);
            this.GBDept.Controls.Add(this.rbDE);
            this.GBDept.Controls.Add(this.rbQA);
            this.GBDept.Location = new System.Drawing.Point(204, 25);
            this.GBDept.Name = "GBDept";
            this.GBDept.Size = new System.Drawing.Size(110, 169);
            this.GBDept.TabIndex = 6;
            this.GBDept.TabStop = false;
            this.GBDept.Text = "Department";
            // 
            // rbPE
            // 
            this.rbPE.AutoSize = true;
            this.rbPE.Location = new System.Drawing.Point(18, 137);
            this.rbPE.Name = "rbPE";
            this.rbPE.Size = new System.Drawing.Size(44, 20);
            this.rbPE.TabIndex = 4;
            this.rbPE.TabStop = true;
            this.rbPE.Text = "PE";
            this.rbPE.UseVisualStyleBackColor = true;
            this.rbPE.CheckedChanged += new System.EventHandler(this.rbPE_CheckedChanged);
            // 
            // rbPro
            // 
            this.rbPro.AutoSize = true;
            this.rbPro.Location = new System.Drawing.Point(18, 111);
            this.rbPro.Name = "rbPro";
            this.rbPro.Size = new System.Drawing.Size(48, 20);
            this.rbPro.TabIndex = 3;
            this.rbPro.TabStop = true;
            this.rbPro.Text = "Pro";
            this.rbPro.UseVisualStyleBackColor = true;
            this.rbPro.CheckedChanged += new System.EventHandler(this.rbPro_CheckedChanged);
            // 
            // rbNPI
            // 
            this.rbNPI.AutoSize = true;
            this.rbNPI.Location = new System.Drawing.Point(18, 85);
            this.rbNPI.Name = "rbNPI";
            this.rbNPI.Size = new System.Drawing.Size(47, 20);
            this.rbNPI.TabIndex = 2;
            this.rbNPI.TabStop = true;
            this.rbNPI.Text = "NPI";
            this.rbNPI.UseVisualStyleBackColor = true;
            this.rbNPI.CheckedChanged += new System.EventHandler(this.rbNPI_CheckedChanged);
            // 
            // rbDE
            // 
            this.rbDE.AutoSize = true;
            this.rbDE.Location = new System.Drawing.Point(18, 59);
            this.rbDE.Name = "rbDE";
            this.rbDE.Size = new System.Drawing.Size(45, 20);
            this.rbDE.TabIndex = 1;
            this.rbDE.TabStop = true;
            this.rbDE.Text = "DE";
            this.rbDE.UseVisualStyleBackColor = true;
            this.rbDE.CheckedChanged += new System.EventHandler(this.rbDE_CheckedChanged);
            // 
            // rbQA
            // 
            this.rbQA.AutoSize = true;
            this.rbQA.Location = new System.Drawing.Point(18, 33);
            this.rbQA.Name = "rbQA";
            this.rbQA.Size = new System.Drawing.Size(46, 20);
            this.rbQA.TabIndex = 0;
            this.rbQA.Text = "QA";
            this.rbQA.UseVisualStyleBackColor = true;
            this.rbQA.CheckedChanged += new System.EventHandler(this.rbQA_CheckedChanged);
            // 
            // Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(551, 229);
            this.Controls.Add(this.LogoPictureBox);
            this.Controls.Add(this.GBUserInfo);
            this.Name = "Login";
            this.Text = "Login";
            this.Load += new System.EventHandler(this.Login_Load);
            ((System.ComponentModel.ISupportInitialize)(this.LogoPictureBox)).EndInit();
            this.GBUserInfo.ResumeLayout(false);
            this.GBUserInfo.PerformLayout();
            this.GBDept.ResumeLayout(false);
            this.GBDept.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.PictureBox LogoPictureBox;
        internal System.Windows.Forms.TextBox txtUsername;
        internal System.Windows.Forms.Button Cancel;
        internal System.Windows.Forms.Label UsernameLabel;
        internal System.Windows.Forms.Button OK;
        internal System.Windows.Forms.Label PasswordLabel;
        internal System.Windows.Forms.TextBox txtPassword;
        internal System.Windows.Forms.GroupBox GBUserInfo;
        private System.Windows.Forms.GroupBox GBDept;
        private System.Windows.Forms.RadioButton rbPE;
        private System.Windows.Forms.RadioButton rbPro;
        private System.Windows.Forms.RadioButton rbNPI;
        private System.Windows.Forms.RadioButton rbDE;
        private System.Windows.Forms.RadioButton rbQA;
    }
}