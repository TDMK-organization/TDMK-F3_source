
namespace Bending_Items
{
    partial class Form1
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
            this.txtLogfile_before = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.DGV_Data = new System.Windows.Forms.DataGridView();
            this.btnLoad = new System.Windows.Forms.Button();
            this.DGV_Data2 = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.txtlogfile_After = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtFormatFile = new System.Windows.Forms.TextBox();
            this.lstData_Addr = new System.Windows.Forms.ListBox();
            this.lstNet = new System.Windows.Forms.ListBox();
            this.txtUUT_No = new System.Windows.Forms.TextBox();
            this.txtNet_No = new System.Windows.Forms.TextBox();
            this.btnTestSum = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data2)).BeginInit();
            this.SuspendLayout();
            // 
            // txtLogfile_before
            // 
            this.txtLogfile_before.Location = new System.Drawing.Point(47, 20);
            this.txtLogfile_before.Name = "txtLogfile_before";
            this.txtLogfile_before.Size = new System.Drawing.Size(615, 20);
            this.txtLogfile_before.TabIndex = 0;
            this.txtLogfile_before.Text = "D:\\Customer Projects\\SEEV\\SMT Project\\SEEV Data\\F3 Data\\DATA\\Type4\\flex-717985-03" +
    "\\20220515-bf1.DAT";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Logfile";
            // 
            // DGV_Data
            // 
            this.DGV_Data.AllowUserToAddRows = false;
            this.DGV_Data.AllowUserToDeleteRows = false;
            this.DGV_Data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Data.Location = new System.Drawing.Point(9, 145);
            this.DGV_Data.Name = "DGV_Data";
            this.DGV_Data.Size = new System.Drawing.Size(538, 242);
            this.DGV_Data.TabIndex = 2;
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(1347, 11);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(88, 37);
            this.btnLoad.TabIndex = 3;
            this.btnLoad.Text = "Load Data";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // DGV_Data2
            // 
            this.DGV_Data2.AllowUserToAddRows = false;
            this.DGV_Data2.AllowUserToDeleteRows = false;
            this.DGV_Data2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Data2.Location = new System.Drawing.Point(12, 393);
            this.DGV_Data2.Name = "DGV_Data2";
            this.DGV_Data2.Size = new System.Drawing.Size(535, 274);
            this.DGV_Data2.TabIndex = 4;
            this.DGV_Data2.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGV_Data2_CellContentClick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(685, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Logfile";
            // 
            // txtlogfile_After
            // 
            this.txtlogfile_After.Location = new System.Drawing.Point(726, 20);
            this.txtlogfile_After.Name = "txtlogfile_After";
            this.txtlogfile_After.Size = new System.Drawing.Size(615, 20);
            this.txtlogfile_After.TabIndex = 5;
            this.txtlogfile_After.Text = "E:\\Git_Repo\\TestAreas\\DATA\\F3\\Bending Items\\719454-00003\\20240118-9CSPD-719454-03" +
    "-503827-BF-3L.DAT";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 79);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Format File";
            // 
            // txtFormatFile
            // 
            this.txtFormatFile.Location = new System.Drawing.Point(70, 76);
            this.txtFormatFile.Name = "txtFormatFile";
            this.txtFormatFile.Size = new System.Drawing.Size(615, 20);
            this.txtFormatFile.TabIndex = 7;
            this.txtFormatFile.Text = "D:\\Customer Projects\\SEEV\\SMT Project\\SEEV Data\\F3 Data\\Issues\\ORT\\Format LQ.xlsx" +
    "";
            // 
            // lstData_Addr
            // 
            this.lstData_Addr.FormattingEnabled = true;
            this.lstData_Addr.Location = new System.Drawing.Point(553, 145);
            this.lstData_Addr.Name = "lstData_Addr";
            this.lstData_Addr.Size = new System.Drawing.Size(882, 524);
            this.lstData_Addr.TabIndex = 9;
            // 
            // lstNet
            // 
            this.lstNet.FormattingEnabled = true;
            this.lstNet.Location = new System.Drawing.Point(249, 694);
            this.lstNet.Name = "lstNet";
            this.lstNet.Size = new System.Drawing.Size(99, 82);
            this.lstNet.TabIndex = 10;
            // 
            // txtUUT_No
            // 
            this.txtUUT_No.Location = new System.Drawing.Point(12, 694);
            this.txtUUT_No.Name = "txtUUT_No";
            this.txtUUT_No.Size = new System.Drawing.Size(97, 20);
            this.txtUUT_No.TabIndex = 11;
            // 
            // txtNet_No
            // 
            this.txtNet_No.Location = new System.Drawing.Point(131, 694);
            this.txtNet_No.Name = "txtNet_No";
            this.txtNet_No.Size = new System.Drawing.Size(97, 20);
            this.txtNet_No.TabIndex = 12;
            // 
            // btnTestSum
            // 
            this.btnTestSum.Location = new System.Drawing.Point(726, 73);
            this.btnTestSum.Name = "btnTestSum";
            this.btnTestSum.Size = new System.Drawing.Size(75, 36);
            this.btnTestSum.TabIndex = 13;
            this.btnTestSum.Text = "Test Sum";
            this.btnTestSum.UseVisualStyleBackColor = true;
            this.btnTestSum.Click += new System.EventHandler(this.btnTestSum_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1459, 794);
            this.Controls.Add(this.btnTestSum);
            this.Controls.Add(this.txtNet_No);
            this.Controls.Add(this.txtUUT_No);
            this.Controls.Add(this.lstNet);
            this.Controls.Add(this.lstData_Addr);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtFormatFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtlogfile_After);
            this.Controls.Add(this.DGV_Data2);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.DGV_Data);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtLogfile_before);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtLogfile_before;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView DGV_Data;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.DataGridView DGV_Data2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtlogfile_After;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtFormatFile;
        private System.Windows.Forms.ListBox lstData_Addr;
        private System.Windows.Forms.ListBox lstNet;
        private System.Windows.Forms.TextBox txtUUT_No;
        private System.Windows.Forms.TextBox txtNet_No;
        private System.Windows.Forms.Button btnTestSum;
    }
}

