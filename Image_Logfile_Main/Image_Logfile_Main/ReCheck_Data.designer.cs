
namespace OK2SHIP
{
    partial class ReCheck_Data
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
            this.DGV_Details = new System.Windows.Forms.DataGridView();
            this.DGV_Recheck = new System.Windows.Forms.DataGridView();
            this.txtLogfile = new System.Windows.Forms.TextBox();
            this.numPCS = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.cbMachine = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnReplace = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.tblMain = new System.Windows.Forms.TableLayoutPanel();
            this.panelControl = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Details)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Recheck)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPCS)).BeginInit();
            this.tblMain.SuspendLayout();
            this.panelControl.SuspendLayout();
            this.SuspendLayout();
            // 
            // DGV_Details
            // 
            this.DGV_Details.AllowUserToAddRows = false;
            this.DGV_Details.AllowUserToDeleteRows = false;
            this.DGV_Details.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Details.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Details.Location = new System.Drawing.Point(3, 123);
            this.DGV_Details.Name = "DGV_Details";
            this.DGV_Details.Size = new System.Drawing.Size(662, 679);
            this.DGV_Details.TabIndex = 0;
            // 
            // DGV_Recheck
            // 
            this.DGV_Recheck.AllowUserToAddRows = false;
            this.DGV_Recheck.AllowUserToDeleteRows = false;
            this.DGV_Recheck.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Recheck.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Recheck.Location = new System.Drawing.Point(671, 123);
            this.DGV_Recheck.Name = "DGV_Recheck";
            this.DGV_Recheck.Size = new System.Drawing.Size(662, 679);
            this.DGV_Recheck.TabIndex = 1;
            // 
            // txtLogfile
            // 
            this.txtLogfile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtLogfile.Location = new System.Drawing.Point(3, 3);
            this.txtLogfile.Multiline = true;
            this.txtLogfile.Name = "txtLogfile";
            this.txtLogfile.Size = new System.Drawing.Size(662, 74);
            this.txtLogfile.TabIndex = 2;
            this.txtLogfile.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLogfile_KeyDown);
            // 
            // numPCS
            // 
            this.numPCS.Location = new System.Drawing.Point(93, 47);
            this.numPCS.Maximum = new decimal(new int[] {
            300,
            0,
            0,
            0});
            this.numPCS.Name = "numPCS";
            this.numPCS.Size = new System.Drawing.Size(55, 20);
            this.numPCS.TabIndex = 22;
            this.numPCS.Value = new decimal(new int[] {
            54,
            0,
            0,
            0});
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 49);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 13);
            this.label3.TabIndex = 21;
            this.label3.Text = "PCS num";
            // 
            // cbMachine
            // 
            this.cbMachine.FormattingEnabled = true;
            this.cbMachine.Items.AddRange(new object[] {
            "YAMAHA",
            "TAIYO"});
            this.cbMachine.Location = new System.Drawing.Point(62, 9);
            this.cbMachine.Name = "cbMachine";
            this.cbMachine.Size = new System.Drawing.Size(86, 21);
            this.cbMachine.TabIndex = 20;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(48, 13);
            this.label2.TabIndex = 19;
            this.label2.Text = "Machine";
            // 
            // btnReplace
            // 
            this.btnReplace.Location = new System.Drawing.Point(199, 9);
            this.btnReplace.Name = "btnReplace";
            this.btnReplace.Size = new System.Drawing.Size(97, 37);
            this.btnReplace.TabIndex = 23;
            this.btnReplace.Text = "Replace Data";
            this.btnReplace.UseVisualStyleBackColor = true;
            this.btnReplace.Click += new System.EventHandler(this.btnReplace_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(321, 9);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(97, 37);
            this.btnUpdate.TabIndex = 24;
            this.btnUpdate.Text = "Update Database";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(443, 9);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(97, 37);
            this.btnExport.TabIndex = 25;
            this.btnExport.Text = "Export to Excel";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // tblMain
            // 
            this.tblMain.ColumnCount = 2;
            this.tblMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tblMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tblMain.Controls.Add(this.DGV_Details, 0, 2);
            this.tblMain.Controls.Add(this.DGV_Recheck, 1, 2);
            this.tblMain.Controls.Add(this.txtLogfile, 0, 0);
            this.tblMain.Controls.Add(this.panelControl, 1, 0);
            this.tblMain.Controls.Add(this.label1, 0, 1);
            this.tblMain.Controls.Add(this.label4, 1, 1);
            this.tblMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblMain.Location = new System.Drawing.Point(0, 0);
            this.tblMain.Name = "tblMain";
            this.tblMain.RowCount = 3;
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tblMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 85F));
            this.tblMain.Size = new System.Drawing.Size(1336, 805);
            this.tblMain.TabIndex = 26;
            // 
            // panelControl
            // 
            this.panelControl.Controls.Add(this.cbMachine);
            this.panelControl.Controls.Add(this.btnExport);
            this.panelControl.Controls.Add(this.label2);
            this.panelControl.Controls.Add(this.btnUpdate);
            this.panelControl.Controls.Add(this.label3);
            this.panelControl.Controls.Add(this.btnReplace);
            this.panelControl.Controls.Add(this.numPCS);
            this.panelControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelControl.Location = new System.Drawing.Point(671, 3);
            this.panelControl.Name = "panelControl";
            this.panelControl.Size = new System.Drawing.Size(662, 74);
            this.panelControl.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Arial", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 80);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(662, 40);
            this.label1.TabIndex = 4;
            this.label1.Text = "Dữ liệu NG";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Font = new System.Drawing.Font("Arial", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(671, 80);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(662, 40);
            this.label4.TabIndex = 5;
            this.label4.Text = "Dữ liệu Logfile(Retry)";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // ReCheck_Data
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1336, 805);
            this.Controls.Add(this.tblMain);
            this.Name = "ReCheck_Data";
            this.Text = "ReCheck_Data";
            this.Load += new System.EventHandler(this.ReCheck_Data_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Details)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Recheck)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPCS)).EndInit();
            this.tblMain.ResumeLayout(false);
            this.tblMain.PerformLayout();
            this.panelControl.ResumeLayout(false);
            this.panelControl.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView DGV_Details;
        private System.Windows.Forms.DataGridView DGV_Recheck;
        private System.Windows.Forms.TextBox txtLogfile;
        private System.Windows.Forms.NumericUpDown numPCS;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbMachine;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnReplace;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.TableLayoutPanel tblMain;
        private System.Windows.Forms.Panel panelControl;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
    }
}