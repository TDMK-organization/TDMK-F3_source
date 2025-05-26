namespace OK2SHIP_Software
{
    partial class MainScreen
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainScreen));
            this.lblSearchData = new System.Windows.Forms.Label();
            this.lblExit = new System.Windows.Forms.Label();
            this.lblExport = new System.Windows.Forms.Label();
            this.lblUserRegister = new System.Windows.Forms.Label();
            this.lblFormatSetup = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblSearchData
            // 
            this.lblSearchData.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblSearchData.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSearchData.Image = ((System.Drawing.Image)(resources.GetObject("lblSearchData.Image")));
            this.lblSearchData.Location = new System.Drawing.Point(241, 275);
            this.lblSearchData.Name = "lblSearchData";
            this.lblSearchData.Size = new System.Drawing.Size(193, 202);
            this.lblSearchData.TabIndex = 15;
            this.lblSearchData.Text = "Analysis Data";
            this.lblSearchData.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblSearchData.Click += new System.EventHandler(this.lblSearchData_Click);
            // 
            // lblExit
            // 
            this.lblExit.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblExit.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblExit.Image = ((System.Drawing.Image)(resources.GetObject("lblExit.Image")));
            this.lblExit.Location = new System.Drawing.Point(472, 275);
            this.lblExit.Name = "lblExit";
            this.lblExit.Size = new System.Drawing.Size(193, 211);
            this.lblExit.TabIndex = 14;
            this.lblExit.Text = "Exit";
            this.lblExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblExit.Click += new System.EventHandler(this.lblExit_Click);
            // 
            // lblExport
            // 
            this.lblExport.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblExport.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblExport.Image = ((System.Drawing.Image)(resources.GetObject("lblExport.Image")));
            this.lblExport.Location = new System.Drawing.Point(26, 266);
            this.lblExport.Name = "lblExport";
            this.lblExport.Size = new System.Drawing.Size(193, 211);
            this.lblExport.TabIndex = 11;
            this.lblExport.Text = "Export to Report";
            this.lblExport.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblExport.Click += new System.EventHandler(this.lblExport_Click);
            // 
            // lblUserRegister
            // 
            this.lblUserRegister.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblUserRegister.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUserRegister.Image = ((System.Drawing.Image)(resources.GetObject("lblUserRegister.Image")));
            this.lblUserRegister.Location = new System.Drawing.Point(472, 25);
            this.lblUserRegister.Name = "lblUserRegister";
            this.lblUserRegister.Size = new System.Drawing.Size(193, 202);
            this.lblUserRegister.TabIndex = 10;
            this.lblUserRegister.Text = "User Register";
            this.lblUserRegister.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblUserRegister.Click += new System.EventHandler(this.lblUserRegister_Click);
            // 
            // lblFormatSetup
            // 
            this.lblFormatSetup.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblFormatSetup.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFormatSetup.Image = ((System.Drawing.Image)(resources.GetObject("lblFormatSetup.Image")));
            this.lblFormatSetup.Location = new System.Drawing.Point(26, 25);
            this.lblFormatSetup.Name = "lblFormatSetup";
            this.lblFormatSetup.Size = new System.Drawing.Size(193, 202);
            this.lblFormatSetup.TabIndex = 17;
            this.lblFormatSetup.Text = "Format Setup";
            this.lblFormatSetup.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblFormatSetup.Click += new System.EventHandler(this.lblFormatSetup_Click);
            // 
            // label1
            // 
            this.label1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.label1.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Image = ((System.Drawing.Image)(resources.GetObject("label1.Image")));
            this.label1.Location = new System.Drawing.Point(241, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(193, 202);
            this.label1.TabIndex = 18;
            this.label1.Text = "Edit Items";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // MainScreen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(679, 502);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblFormatSetup);
            this.Controls.Add(this.lblSearchData);
            this.Controls.Add(this.lblExit);
            this.Controls.Add(this.lblExport);
            this.Controls.Add(this.lblUserRegister);
            this.Name = "MainScreen";
            this.Text = "OK2SHIP Automation System";
            this.Load += new System.EventHandler(this.MainScreen_Load);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Label lblSearchData;
        internal System.Windows.Forms.Label lblExit;
        internal System.Windows.Forms.Label lblExport;
        internal System.Windows.Forms.Label lblUserRegister;
        internal System.Windows.Forms.Label lblFormatSetup;
        internal System.Windows.Forms.Label label1;
    }
}