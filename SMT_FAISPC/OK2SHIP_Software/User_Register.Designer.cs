namespace OK2SHIP_Software
{
    partial class User_Register
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(User_Register));
            this.LogoPictureBox = new System.Windows.Forms.PictureBox();
            this.txtUsername = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.UsernameLabel = new System.Windows.Forms.Label();
            this.btnAdd = new System.Windows.Forms.Button();
            this.PasswordLabel = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.GBUserInfo = new System.Windows.Forms.GroupBox();
            this.btnDelete = new System.Windows.Forms.Button();
            this.Label2 = new System.Windows.Forms.Label();
            this.txtReTypePass = new System.Windows.Forms.TextBox();
            this.txtUserID = new System.Windows.Forms.TextBox();
            this.Label1 = new System.Windows.Forms.Label();
            this.DGV_User = new System.Windows.Forms.DataGridView();
            this.GBDept = new System.Windows.Forms.GroupBox();
            this.rbPE = new System.Windows.Forms.RadioButton();
            this.rbPro = new System.Windows.Forms.RadioButton();
            this.rbNPI = new System.Windows.Forms.RadioButton();
            this.rbDE = new System.Windows.Forms.RadioButton();
            this.rbQA = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.LogoPictureBox)).BeginInit();
            this.GBUserInfo.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_User)).BeginInit();
            this.GBDept.SuspendLayout();
            this.SuspendLayout();
            // 
            // LogoPictureBox
            // 
            this.LogoPictureBox.Image = ((System.Drawing.Image)(resources.GetObject("LogoPictureBox.Image")));
            this.LogoPictureBox.Location = new System.Drawing.Point(2, 12);
            this.LogoPictureBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.LogoPictureBox.Name = "LogoPictureBox";
            this.LogoPictureBox.Size = new System.Drawing.Size(271, 265);
            this.LogoPictureBox.TabIndex = 8;
            this.LogoPictureBox.TabStop = false;
            // 
            // txtUsername
            // 
            this.txtUsername.Location = new System.Drawing.Point(6, 48);
            this.txtUsername.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtUsername.Name = "txtUsername";
            this.txtUsername.Size = new System.Drawing.Size(180, 20);
            this.txtUsername.TabIndex = 1;
            this.txtUsername.TextChanged += new System.EventHandler(this.txtUsername_TextChanged);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(313, 204);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(106, 40);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // UsernameLabel
            // 
            this.UsernameLabel.Location = new System.Drawing.Point(4, 15);
            this.UsernameLabel.Name = "UsernameLabel";
            this.UsernameLabel.Size = new System.Drawing.Size(182, 27);
            this.UsernameLabel.TabIndex = 0;
            this.UsernameLabel.Text = "User name";
            this.UsernameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(313, 20);
            this.btnAdd.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(106, 40);
            this.btnAdd.TabIndex = 4;
            this.btnAdd.Text = "Add User";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // PasswordLabel
            // 
            this.PasswordLabel.Location = new System.Drawing.Point(5, 133);
            this.PasswordLabel.Name = "PasswordLabel";
            this.PasswordLabel.Size = new System.Drawing.Size(182, 27);
            this.PasswordLabel.TabIndex = 2;
            this.PasswordLabel.Text = "Password";
            this.PasswordLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(7, 166);
            this.txtPassword.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(180, 20);
            this.txtPassword.TabIndex = 3;
            this.txtPassword.TextChanged += new System.EventHandler(this.txtPassword_TextChanged);
            // 
            // GBUserInfo
            // 
            this.GBUserInfo.Controls.Add(this.GBDept);
            this.GBUserInfo.Controls.Add(this.btnDelete);
            this.GBUserInfo.Controls.Add(this.Label2);
            this.GBUserInfo.Controls.Add(this.txtReTypePass);
            this.GBUserInfo.Controls.Add(this.txtUserID);
            this.GBUserInfo.Controls.Add(this.Label1);
            this.GBUserInfo.Controls.Add(this.txtUsername);
            this.GBUserInfo.Controls.Add(this.btnCancel);
            this.GBUserInfo.Controls.Add(this.UsernameLabel);
            this.GBUserInfo.Controls.Add(this.btnAdd);
            this.GBUserInfo.Controls.Add(this.PasswordLabel);
            this.GBUserInfo.Controls.Add(this.txtPassword);
            this.GBUserInfo.Location = new System.Drawing.Point(279, 12);
            this.GBUserInfo.Name = "GBUserInfo";
            this.GBUserInfo.Size = new System.Drawing.Size(454, 265);
            this.GBUserInfo.TabIndex = 9;
            this.GBUserInfo.TabStop = false;
            this.GBUserInfo.Text = "User Information";
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(313, 112);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(106, 40);
            this.btnDelete.TabIndex = 10;
            this.btnDelete.Text = "Delete User";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // Label2
            // 
            this.Label2.Location = new System.Drawing.Point(5, 192);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(182, 27);
            this.Label2.TabIndex = 8;
            this.Label2.Text = "Re-Type Password";
            this.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtReTypePass
            // 
            this.txtReTypePass.Location = new System.Drawing.Point(7, 225);
            this.txtReTypePass.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtReTypePass.Name = "txtReTypePass";
            this.txtReTypePass.PasswordChar = '*';
            this.txtReTypePass.Size = new System.Drawing.Size(180, 20);
            this.txtReTypePass.TabIndex = 9;
            this.txtReTypePass.TextChanged += new System.EventHandler(this.txtReTypePass_TextChanged);
            // 
            // txtUserID
            // 
            this.txtUserID.Location = new System.Drawing.Point(6, 107);
            this.txtUserID.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtUserID.Name = "txtUserID";
            this.txtUserID.Size = new System.Drawing.Size(180, 20);
            this.txtUserID.TabIndex = 7;
            this.txtUserID.TextChanged += new System.EventHandler(this.txtUserID_TextChanged);
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(4, 74);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(182, 27);
            this.Label1.TabIndex = 6;
            this.Label1.Text = "User ID";
            this.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // DGV_User
            // 
            this.DGV_User.AllowUserToAddRows = false;
            this.DGV_User.AllowUserToDeleteRows = false;
            this.DGV_User.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_User.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.DGV_User.Location = new System.Drawing.Point(2, 284);
            this.DGV_User.Name = "DGV_User";
            this.DGV_User.RowTemplate.Height = 24;
            this.DGV_User.Size = new System.Drawing.Size(731, 262);
            this.DGV_User.TabIndex = 10;
            this.DGV_User.RowHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_User_RowHeaderMouseDoubleClick);
            // 
            // GBDept
            // 
            this.GBDept.Controls.Add(this.rbPE);
            this.GBDept.Controls.Add(this.rbPro);
            this.GBDept.Controls.Add(this.rbNPI);
            this.GBDept.Controls.Add(this.rbDE);
            this.GBDept.Controls.Add(this.rbQA);
            this.GBDept.Location = new System.Drawing.Point(194, 19);
            this.GBDept.Name = "GBDept";
            this.GBDept.Size = new System.Drawing.Size(101, 226);
            this.GBDept.TabIndex = 11;
            this.GBDept.TabStop = false;
            this.GBDept.Text = "Department";
            // 
            // rbPE
            // 
            this.rbPE.AutoSize = true;
            this.rbPE.Location = new System.Drawing.Point(18, 190);
            this.rbPE.Name = "rbPE";
            this.rbPE.Size = new System.Drawing.Size(39, 17);
            this.rbPE.TabIndex = 4;
            this.rbPE.TabStop = true;
            this.rbPE.Text = "PE";
            this.rbPE.UseVisualStyleBackColor = true;
            this.rbPE.CheckedChanged += new System.EventHandler(this.rbPE_CheckedChanged);
            // 
            // rbPro
            // 
            this.rbPro.AutoSize = true;
            this.rbPro.Location = new System.Drawing.Point(18, 152);
            this.rbPro.Name = "rbPro";
            this.rbPro.Size = new System.Drawing.Size(41, 17);
            this.rbPro.TabIndex = 3;
            this.rbPro.TabStop = true;
            this.rbPro.Text = "Pro";
            this.rbPro.UseVisualStyleBackColor = true;
            this.rbPro.CheckedChanged += new System.EventHandler(this.rbPro_CheckedChanged);
            // 
            // rbNPI
            // 
            this.rbNPI.AutoSize = true;
            this.rbNPI.Location = new System.Drawing.Point(18, 114);
            this.rbNPI.Name = "rbNPI";
            this.rbNPI.Size = new System.Drawing.Size(43, 17);
            this.rbNPI.TabIndex = 2;
            this.rbNPI.TabStop = true;
            this.rbNPI.Text = "NPI";
            this.rbNPI.UseVisualStyleBackColor = true;
            this.rbNPI.CheckedChanged += new System.EventHandler(this.rbNPI_CheckedChanged);
            // 
            // rbDE
            // 
            this.rbDE.AutoSize = true;
            this.rbDE.Location = new System.Drawing.Point(18, 76);
            this.rbDE.Name = "rbDE";
            this.rbDE.Size = new System.Drawing.Size(40, 17);
            this.rbDE.TabIndex = 1;
            this.rbDE.TabStop = true;
            this.rbDE.Text = "DE";
            this.rbDE.UseVisualStyleBackColor = true;
            this.rbDE.CheckedChanged += new System.EventHandler(this.rbDE_CheckedChanged);
            // 
            // rbQA
            // 
            this.rbQA.AutoSize = true;
            this.rbQA.Location = new System.Drawing.Point(18, 38);
            this.rbQA.Name = "rbQA";
            this.rbQA.Size = new System.Drawing.Size(40, 17);
            this.rbQA.TabIndex = 0;
            this.rbQA.TabStop = true;
            this.rbQA.Text = "QA";
            this.rbQA.UseVisualStyleBackColor = true;
            this.rbQA.CheckedChanged += new System.EventHandler(this.rbQA_CheckedChanged);
            // 
            // User_Register
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(744, 558);
            this.Controls.Add(this.LogoPictureBox);
            this.Controls.Add(this.GBUserInfo);
            this.Controls.Add(this.DGV_User);
            this.Name = "User_Register";
            this.Text = "User_Register";
            this.Load += new System.EventHandler(this.User_Register_Load);
            ((System.ComponentModel.ISupportInitialize)(this.LogoPictureBox)).EndInit();
            this.GBUserInfo.ResumeLayout(false);
            this.GBUserInfo.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_User)).EndInit();
            this.GBDept.ResumeLayout(false);
            this.GBDept.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.PictureBox LogoPictureBox;
        internal System.Windows.Forms.TextBox txtUsername;
        internal System.Windows.Forms.Button btnCancel;
        internal System.Windows.Forms.Label UsernameLabel;
        internal System.Windows.Forms.Button btnAdd;
        internal System.Windows.Forms.Label PasswordLabel;
        internal System.Windows.Forms.TextBox txtPassword;
        internal System.Windows.Forms.GroupBox GBUserInfo;
        internal System.Windows.Forms.Button btnDelete;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.TextBox txtReTypePass;
        internal System.Windows.Forms.TextBox txtUserID;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.DataGridView DGV_User;
        private System.Windows.Forms.GroupBox GBDept;
        private System.Windows.Forms.RadioButton rbPE;
        private System.Windows.Forms.RadioButton rbPro;
        private System.Windows.Forms.RadioButton rbNPI;
        private System.Windows.Forms.RadioButton rbDE;
        private System.Windows.Forms.RadioButton rbQA;
    }
}