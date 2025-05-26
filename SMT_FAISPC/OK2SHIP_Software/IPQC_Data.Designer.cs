namespace OK2SHIP_Software
{
    partial class IPQC_Data
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.GBIPQC_Items = new System.Windows.Forms.GroupBox();
            this.btnLoadSpec = new System.Windows.Forms.Button();
            this.RBAU_NI = new System.Windows.Forms.RadioButton();
            this.RBRoughness = new System.Windows.Forms.RadioButton();
            this.RB_CoverLay_Process = new System.Windows.Forms.RadioButton();
            this.RB_UV_Process = new System.Windows.Forms.RadioButton();
            this.RB_Printing_Process = new System.Windows.Forms.RadioButton();
            this.RB_CopperPlating = new System.Windows.Forms.RadioButton();
            this.RB_Etching_Process = new System.Windows.Forms.RadioButton();
            this.lblIPQC = new System.Windows.Forms.Label();
            this.DGV_DataView = new System.Windows.Forms.DataGridView();
            this.lblSpec = new System.Windows.Forms.Label();
            this.DGV_SpecView = new System.Windows.Forms.DataGridView();
            this.tblLayoutMain = new System.Windows.Forms.TableLayoutPanel();
            this.GBIPQC_Items.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_DataView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_SpecView)).BeginInit();
            this.tblLayoutMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // GBIPQC_Items
            // 
            this.GBIPQC_Items.Controls.Add(this.btnLoadSpec);
            this.GBIPQC_Items.Controls.Add(this.RBAU_NI);
            this.GBIPQC_Items.Controls.Add(this.RBRoughness);
            this.GBIPQC_Items.Controls.Add(this.RB_CoverLay_Process);
            this.GBIPQC_Items.Controls.Add(this.RB_UV_Process);
            this.GBIPQC_Items.Controls.Add(this.RB_Printing_Process);
            this.GBIPQC_Items.Controls.Add(this.RB_CopperPlating);
            this.GBIPQC_Items.Controls.Add(this.RB_Etching_Process);
            this.GBIPQC_Items.Dock = System.Windows.Forms.DockStyle.Fill;
            this.GBIPQC_Items.Location = new System.Drawing.Point(3, 3);
            this.GBIPQC_Items.Name = "GBIPQC_Items";
            this.GBIPQC_Items.Size = new System.Drawing.Size(1178, 70);
            this.GBIPQC_Items.TabIndex = 8;
            this.GBIPQC_Items.TabStop = false;
            this.GBIPQC_Items.Text = "Select Process";
            // 
            // btnLoadSpec
            // 
            this.btnLoadSpec.Location = new System.Drawing.Point(1026, 21);
            this.btnLoadSpec.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnLoadSpec.Name = "btnLoadSpec";
            this.btnLoadSpec.Size = new System.Drawing.Size(141, 30);
            this.btnLoadSpec.TabIndex = 13;
            this.btnLoadSpec.Text = "Load item Check";
            this.btnLoadSpec.UseVisualStyleBackColor = true;
            this.btnLoadSpec.Click += new System.EventHandler(this.btnLoadSpec_Click);
            // 
            // RBAU_NI
            // 
            this.RBAU_NI.AutoSize = true;
            this.RBAU_NI.Location = new System.Drawing.Point(876, 28);
            this.RBAU_NI.Name = "RBAU_NI";
            this.RBAU_NI.Size = new System.Drawing.Size(112, 17);
            this.RBAU_NI.TabIndex = 8;
            this.RBAU_NI.TabStop = true;
            this.RBAU_NI.Text = "AU_NI_Thickness";
            this.RBAU_NI.UseVisualStyleBackColor = true;
            this.RBAU_NI.CheckedChanged += new System.EventHandler(this.RBAU_NI_CheckedChanged);
            // 
            // RBRoughness
            // 
            this.RBRoughness.AutoSize = true;
            this.RBRoughness.Location = new System.Drawing.Point(759, 28);
            this.RBRoughness.Name = "RBRoughness";
            this.RBRoughness.Size = new System.Drawing.Size(79, 17);
            this.RBRoughness.TabIndex = 7;
            this.RBRoughness.TabStop = true;
            this.RBRoughness.Text = "Roughness";
            this.RBRoughness.UseVisualStyleBackColor = true;
            this.RBRoughness.CheckedChanged += new System.EventHandler(this.RBRoughness_CheckedChanged);
            // 
            // RB_CoverLay_Process
            // 
            this.RB_CoverLay_Process.AutoSize = true;
            this.RB_CoverLay_Process.Location = new System.Drawing.Point(601, 28);
            this.RB_CoverLay_Process.Name = "RB_CoverLay_Process";
            this.RB_CoverLay_Process.Size = new System.Drawing.Size(120, 17);
            this.RB_CoverLay_Process.TabIndex = 6;
            this.RB_CoverLay_Process.TabStop = true;
            this.RB_CoverLay_Process.Text = "Cover_Lay_Process";
            this.RB_CoverLay_Process.UseVisualStyleBackColor = true;
            this.RB_CoverLay_Process.CheckedChanged += new System.EventHandler(this.RB_CoverLay_Process_CheckedChanged);
            // 
            // RB_UV_Process
            // 
            this.RB_UV_Process.AutoSize = true;
            this.RB_UV_Process.Location = new System.Drawing.Point(479, 28);
            this.RB_UV_Process.Name = "RB_UV_Process";
            this.RB_UV_Process.Size = new System.Drawing.Size(84, 17);
            this.RB_UV_Process.TabIndex = 4;
            this.RB_UV_Process.TabStop = true;
            this.RB_UV_Process.Text = "UV_Process";
            this.RB_UV_Process.UseVisualStyleBackColor = true;
            this.RB_UV_Process.CheckedChanged += new System.EventHandler(this.RB_UV_Process_CheckedChanged);
            // 
            // RB_Printing_Process
            // 
            this.RB_Printing_Process.AutoSize = true;
            this.RB_Printing_Process.Location = new System.Drawing.Point(337, 28);
            this.RB_Printing_Process.Name = "RB_Printing_Process";
            this.RB_Printing_Process.Size = new System.Drawing.Size(104, 17);
            this.RB_Printing_Process.TabIndex = 3;
            this.RB_Printing_Process.TabStop = true;
            this.RB_Printing_Process.Text = "Printing_Process";
            this.RB_Printing_Process.UseVisualStyleBackColor = true;
            this.RB_Printing_Process.CheckedChanged += new System.EventHandler(this.RB_Printing_Process_CheckedChanged);
            // 
            // RB_CopperPlating
            // 
            this.RB_CopperPlating.AutoSize = true;
            this.RB_CopperPlating.Location = new System.Drawing.Point(158, 28);
            this.RB_CopperPlating.Name = "RB_CopperPlating";
            this.RB_CopperPlating.Size = new System.Drawing.Size(141, 17);
            this.RB_CopperPlating.TabIndex = 1;
            this.RB_CopperPlating.TabStop = true;
            this.RB_CopperPlating.Text = "Copper_Plating_Process";
            this.RB_CopperPlating.UseVisualStyleBackColor = true;
            this.RB_CopperPlating.CheckedChanged += new System.EventHandler(this.RB_CopperPlating_CheckedChanged);
            // 
            // RB_Etching_Process
            // 
            this.RB_Etching_Process.AutoSize = true;
            this.RB_Etching_Process.Location = new System.Drawing.Point(15, 28);
            this.RB_Etching_Process.Name = "RB_Etching_Process";
            this.RB_Etching_Process.Size = new System.Drawing.Size(105, 17);
            this.RB_Etching_Process.TabIndex = 0;
            this.RB_Etching_Process.TabStop = true;
            this.RB_Etching_Process.Text = "Etching_Process";
            this.RB_Etching_Process.UseVisualStyleBackColor = true;
            this.RB_Etching_Process.CheckedChanged += new System.EventHandler(this.RB_Etching_Process_CheckedChanged);
            // 
            // lblIPQC
            // 
            this.lblIPQC.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.lblIPQC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblIPQC.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblIPQC.Font = new System.Drawing.Font("Arial", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblIPQC.Location = new System.Drawing.Point(3, 304);
            this.lblIPQC.Name = "lblIPQC";
            this.lblIPQC.Size = new System.Drawing.Size(1178, 38);
            this.lblIPQC.TabIndex = 9;
            this.lblIPQC.Text = "IPQC Data";
            this.lblIPQC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DGV_DataView
            // 
            this.DGV_DataView.AllowUserToAddRows = false;
            this.DGV_DataView.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.DGV_DataView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.DGV_DataView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_DataView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_DataView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.DGV_DataView.Location = new System.Drawing.Point(3, 346);
            this.DGV_DataView.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.DGV_DataView.Name = "DGV_DataView";
            this.DGV_DataView.Size = new System.Drawing.Size(1178, 411);
            this.DGV_DataView.TabIndex = 10;
            this.DGV_DataView.ColumnHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_DataView_ColumnHeaderMouseDoubleClick);
            // 
            // lblSpec
            // 
            this.lblSpec.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.lblSpec.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSpec.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblSpec.Font = new System.Drawing.Font("Arial", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSpec.Location = new System.Drawing.Point(3, 76);
            this.lblSpec.Name = "lblSpec";
            this.lblSpec.Size = new System.Drawing.Size(1178, 38);
            this.lblSpec.TabIndex = 11;
            this.lblSpec.Text = "Specification";
            this.lblSpec.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DGV_SpecView
            // 
            this.DGV_SpecView.AllowUserToAddRows = false;
            this.DGV_SpecView.AllowUserToDeleteRows = false;
            this.DGV_SpecView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_SpecView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_SpecView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.DGV_SpecView.Location = new System.Drawing.Point(3, 118);
            this.DGV_SpecView.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.DGV_SpecView.Name = "DGV_SpecView";
            this.DGV_SpecView.Size = new System.Drawing.Size(1178, 182);
            this.DGV_SpecView.TabIndex = 12;
            // 
            // tblLayoutMain
            // 
            this.tblLayoutMain.ColumnCount = 1;
            this.tblLayoutMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tblLayoutMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tblLayoutMain.Controls.Add(this.DGV_DataView, 0, 4);
            this.tblLayoutMain.Controls.Add(this.GBIPQC_Items, 0, 0);
            this.tblLayoutMain.Controls.Add(this.lblSpec, 0, 1);
            this.tblLayoutMain.Controls.Add(this.DGV_SpecView, 0, 2);
            this.tblLayoutMain.Controls.Add(this.lblIPQC, 0, 3);
            this.tblLayoutMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tblLayoutMain.Location = new System.Drawing.Point(0, 0);
            this.tblLayoutMain.Name = "tblLayoutMain";
            this.tblLayoutMain.RowCount = 5;
            this.tblLayoutMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tblLayoutMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tblLayoutMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tblLayoutMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tblLayoutMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 55F));
            this.tblLayoutMain.Size = new System.Drawing.Size(1184, 761);
            this.tblLayoutMain.TabIndex = 13;
            // 
            // IPQC_Data
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1184, 761);
            this.Controls.Add(this.tblLayoutMain);
            this.Name = "IPQC_Data";
            this.Text = "IPQC_Data";
            this.Load += new System.EventHandler(this.IPQC_Data_Load);
            this.GBIPQC_Items.ResumeLayout(false);
            this.GBIPQC_Items.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_DataView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_SpecView)).EndInit();
            this.tblLayoutMain.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox GBIPQC_Items;
        private System.Windows.Forms.RadioButton RB_CoverLay_Process;
        private System.Windows.Forms.RadioButton RB_UV_Process;
        private System.Windows.Forms.RadioButton RB_Printing_Process;
        private System.Windows.Forms.RadioButton RB_CopperPlating;
        private System.Windows.Forms.RadioButton RB_Etching_Process;
        private System.Windows.Forms.Label lblIPQC;
        private System.Windows.Forms.DataGridView DGV_DataView;
        private System.Windows.Forms.Label lblSpec;
        private System.Windows.Forms.DataGridView DGV_SpecView;
        private System.Windows.Forms.Button btnLoadSpec;
        private System.Windows.Forms.RadioButton RBAU_NI;
        private System.Windows.Forms.RadioButton RBRoughness;
        private System.Windows.Forms.TableLayoutPanel tblLayoutMain;
    }
}