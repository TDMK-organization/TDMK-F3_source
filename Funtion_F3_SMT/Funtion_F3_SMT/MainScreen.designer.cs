namespace Funtion_F3_SMT
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
            this.lblFormatSetup = new System.Windows.Forms.Label();
            this.lblInputdata = new System.Windows.Forms.Label();
            this.btn_TableOfContent = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblFormatSetup
            // 
            this.lblFormatSetup.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblFormatSetup.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFormatSetup.Image = ((System.Drawing.Image)(resources.GetObject("lblFormatSetup.Image")));
            this.lblFormatSetup.Location = new System.Drawing.Point(33, 54);
            this.lblFormatSetup.Name = "lblFormatSetup";
            this.lblFormatSetup.Size = new System.Drawing.Size(205, 186);
            this.lblFormatSetup.TabIndex = 17;
            this.lblFormatSetup.Text = "Cài đặt số chân Peel/Pull/Shear";
            this.lblFormatSetup.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblFormatSetup.Click += new System.EventHandler(this.lblFormatSetup_Click);
            // 
            // lblInputdata
            // 
            this.lblInputdata.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblInputdata.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInputdata.Image = ((System.Drawing.Image)(resources.GetObject("lblInputdata.Image")));
            this.lblInputdata.Location = new System.Drawing.Point(306, 29);
            this.lblInputdata.Name = "lblInputdata";
            this.lblInputdata.Size = new System.Drawing.Size(193, 211);
            this.lblInputdata.TabIndex = 19;
            this.lblInputdata.Text = "Input Data";
            this.lblInputdata.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.lblInputdata.Click += new System.EventHandler(this.lblInputdata_Click);
            // 
            // btn_TableOfContent
            // 
            this.btn_TableOfContent.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_TableOfContent.Location = new System.Drawing.Point(310, 264);
            this.btn_TableOfContent.Name = "btn_TableOfContent";
            this.btn_TableOfContent.Size = new System.Drawing.Size(203, 219);
            this.btn_TableOfContent.TabIndex = 20;
            this.btn_TableOfContent.Text = "Table Of Content";
            this.btn_TableOfContent.UseVisualStyleBackColor = true;
            this.btn_TableOfContent.Click += new System.EventHandler(this.button1_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(37, 264);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(201, 219);
            this.button1.TabIndex = 21;
            this.button1.Text = "Account Manager";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // MainScreen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(558, 514);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btn_TableOfContent);
            this.Controls.Add(this.lblInputdata);
            this.Controls.Add(this.lblFormatSetup);
            this.Name = "MainScreen";
            this.Text = "OK2SHIP SMT";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainScreen_FormClosed);
            this.Load += new System.EventHandler(this.MainScreen_Load);
            this.ResumeLayout(false);

        }

        #endregion
        internal System.Windows.Forms.Label lblFormatSetup;
        internal System.Windows.Forms.Label lblInputdata;
        private System.Windows.Forms.Button btn_TableOfContent;
        private System.Windows.Forms.Button button1;
    }
}