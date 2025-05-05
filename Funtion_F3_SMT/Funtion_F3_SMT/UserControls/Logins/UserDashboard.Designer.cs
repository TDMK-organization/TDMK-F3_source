namespace OK2SHIP_SMT.UserControls.Logins
{
    partial class UserDashboard
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tp_AccountManager = new System.Windows.Forms.TabPage();
            this.tp_EditProfile = new System.Windows.Forms.TabPage();
            this.tabControl.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.tp_AccountManager);
            this.tabControl.Controls.Add(this.tp_EditProfile);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(1251, 610);
            this.tabControl.TabIndex = 0;
            // 
            // tp_AccountManager
            // 
            this.tp_AccountManager.Location = new System.Drawing.Point(4, 22);
            this.tp_AccountManager.Name = "tp_AccountManager";
            this.tp_AccountManager.Padding = new System.Windows.Forms.Padding(3);
            this.tp_AccountManager.Size = new System.Drawing.Size(1243, 584);
            this.tp_AccountManager.TabIndex = 0;
            this.tp_AccountManager.Text = "Account Manager";
            this.tp_AccountManager.UseVisualStyleBackColor = true;
            // 
            // tp_EditProfile
            // 
            this.tp_EditProfile.Location = new System.Drawing.Point(4, 22);
            this.tp_EditProfile.Name = "tp_EditProfile";
            this.tp_EditProfile.Padding = new System.Windows.Forms.Padding(3);
            this.tp_EditProfile.Size = new System.Drawing.Size(1243, 584);
            this.tp_EditProfile.TabIndex = 1;
            this.tp_EditProfile.Text = "Edit Profile";
            this.tp_EditProfile.UseVisualStyleBackColor = true;
            // 
            // UserDashboard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl);
            this.Name = "UserDashboard";
            this.Size = new System.Drawing.Size(1251, 610);
            this.tabControl.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tp_AccountManager;
        private System.Windows.Forms.TabPage tp_EditProfile;
    }
}
