namespace OK2SHIP
{
    partial class ACF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ACF));
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.txtItemCode = new System.Windows.Forms.TextBox();
            this.mnuAction = new System.Windows.Forms.MenuStrip();
            this.tsmLoadData = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmExport = new System.Windows.Forms.ToolStripMenuItem();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.lblImage = new System.Windows.Forms.Label();
            this.DGV_Data = new System.Windows.Forms.DataGridView();
            this.lbltitle = new System.Windows.Forms.Label();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.picdetail = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tableLayoutPanel5 = new System.Windows.Forms.TableLayoutPanel();
            this.DGV_Roughness_spec = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.lbl_title_spec = new System.Windows.Forms.Label();
            this.btn_load_spec = new System.Windows.Forms.Button();
            this.cb_data_for = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtLotNo = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.txtOperator = new System.Windows.Forms.TextBox();
            this.mnuAction.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).BeginInit();
            this.tableLayoutPanel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picdetail)).BeginInit();
            this.panel1.SuspendLayout();
            this.tableLayoutPanel5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Roughness_spec)).BeginInit();
            this.tableLayoutPanel4.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(455, 7);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(45, 16);
            this.label11.TabIndex = 38;
            this.label11.Text = "Lot No";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(284, 7);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(62, 16);
            this.label10.TabIndex = 34;
            this.label10.Text = "ItemCode";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtItemCode
            // 
            this.txtItemCode.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtItemCode.Location = new System.Drawing.Point(352, 4);
            this.txtItemCode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtItemCode.Name = "txtItemCode";
            this.txtItemCode.Size = new System.Drawing.Size(100, 22);
            this.txtItemCode.TabIndex = 35;
            this.txtItemCode.TextChanged += new System.EventHandler(this.txtItemCode_TextChanged);
            // 
            // mnuAction
            // 
            this.mnuAction.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mnuAction.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmLoadData,
            this.tsmExport});
            this.mnuAction.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow;
            this.mnuAction.Location = new System.Drawing.Point(0, 0);
            this.mnuAction.Name = "mnuAction";
            this.mnuAction.Size = new System.Drawing.Size(1266, 26);
            this.mnuAction.TabIndex = 33;
            this.mnuAction.Text = "menuStrip1";
            this.mnuAction.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.mnuAction_ItemClicked);
            // 
            // tsmLoadData
            // 
            this.tsmLoadData.Name = "tsmLoadData";
            this.tsmLoadData.Size = new System.Drawing.Size(94, 22);
            this.tsmLoadData.Text = "Load Data";
            this.tsmLoadData.Click += new System.EventHandler(this.tsmLoadData_Click);
            // 
            // tsmExport
            // 
            this.tsmExport.Name = "tsmExport";
            this.tsmExport.Size = new System.Drawing.Size(165, 22);
            this.tsmExport.Text = "Export to checksheet";
            this.tsmExport.Click += new System.EventHandler(this.tsmExport_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 61.38498F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 38.61502F));
            this.tableLayoutPanel1.Controls.Add(this.lblImage, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.DGV_Data, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.lbltitle, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel3, 1, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 43);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1260, 699);
            this.tableLayoutPanel1.TabIndex = 42;
            // 
            // lblImage
            // 
            this.lblImage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblImage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblImage.Font = new System.Drawing.Font("Arial", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblImage.Location = new System.Drawing.Point(776, 0);
            this.lblImage.Name = "lblImage";
            this.lblImage.Size = new System.Drawing.Size(481, 40);
            this.lblImage.TabIndex = 6;
            this.lblImage.Text = "Image / Graph Details";
            this.lblImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DGV_Data
            // 
            this.DGV_Data.AllowUserToAddRows = false;
            this.DGV_Data.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DGV_Data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Data.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Data.Location = new System.Drawing.Point(3, 43);
            this.DGV_Data.Name = "DGV_Data";
            this.DGV_Data.RowHeadersVisible = false;
            this.DGV_Data.Size = new System.Drawing.Size(767, 653);
            this.DGV_Data.TabIndex = 6;
            this.DGV_Data.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGV_Data_CellContentClick);
            // 
            // lbltitle
            // 
            this.lbltitle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lbltitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbltitle.Font = new System.Drawing.Font("Arial", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbltitle.Location = new System.Drawing.Point(3, 0);
            this.lbltitle.Name = "lbltitle";
            this.lbltitle.Size = new System.Drawing.Size(767, 40);
            this.lbltitle.TabIndex = 5;
            this.lbltitle.Text = "Data For";
            this.lbltitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 1;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Controls.Add(this.picdetail, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.panel1, 0, 1);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(776, 43);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 2;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 200F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(481, 653);
            this.tableLayoutPanel3.TabIndex = 44;
            // 
            // picdetail
            // 
            this.picdetail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picdetail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.picdetail.Location = new System.Drawing.Point(3, 3);
            this.picdetail.Name = "picdetail";
            this.picdetail.Size = new System.Drawing.Size(475, 447);
            this.picdetail.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picdetail.TabIndex = 7;
            this.picdetail.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.tableLayoutPanel5);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 456);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(475, 194);
            this.panel1.TabIndex = 11;
            // 
            // tableLayoutPanel5
            // 
            this.tableLayoutPanel5.ColumnCount = 1;
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel5.Controls.Add(this.DGV_Roughness_spec, 0, 1);
            this.tableLayoutPanel5.Controls.Add(this.tableLayoutPanel4, 0, 0);
            this.tableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel5.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel5.Name = "tableLayoutPanel5";
            this.tableLayoutPanel5.RowCount = 2;
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel5.Size = new System.Drawing.Size(473, 192);
            this.tableLayoutPanel5.TabIndex = 44;
            // 
            // DGV_Roughness_spec
            // 
            this.DGV_Roughness_spec.AllowUserToAddRows = false;
            this.DGV_Roughness_spec.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Roughness_spec.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DGV_Roughness_spec.Location = new System.Drawing.Point(3, 43);
            this.DGV_Roughness_spec.Name = "DGV_Roughness_spec";
            this.DGV_Roughness_spec.Size = new System.Drawing.Size(467, 146);
            this.DGV_Roughness_spec.TabIndex = 9;
            this.DGV_Roughness_spec.Visible = false;
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.ColumnCount = 2;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 79.72603F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.27397F));
            this.tableLayoutPanel4.Controls.Add(this.lbl_title_spec, 0, 0);
            this.tableLayoutPanel4.Controls.Add(this.btn_load_spec, 1, 0);
            this.tableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel4.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 1;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(467, 34);
            this.tableLayoutPanel4.TabIndex = 10;
            // 
            // lbl_title_spec
            // 
            this.lbl_title_spec.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lbl_title_spec.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_title_spec.Font = new System.Drawing.Font("Arial", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_title_spec.Location = new System.Drawing.Point(3, 0);
            this.lbl_title_spec.Name = "lbl_title_spec";
            this.lbl_title_spec.Size = new System.Drawing.Size(366, 34);
            this.lbl_title_spec.TabIndex = 8;
            this.lbl_title_spec.Text = "Specificion";
            this.lbl_title_spec.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lbl_title_spec.Visible = false;
            // 
            // btn_load_spec
            // 
            this.btn_load_spec.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_load_spec.Location = new System.Drawing.Point(375, 3);
            this.btn_load_spec.Name = "btn_load_spec";
            this.btn_load_spec.Size = new System.Drawing.Size(89, 28);
            this.btn_load_spec.TabIndex = 9;
            this.btn_load_spec.Text = "LOAD SPEC";
            this.btn_load_spec.UseVisualStyleBackColor = true;
            this.btn_load_spec.Visible = false;
            this.btn_load_spec.Click += new System.EventHandler(this.btn_load_spec_Click);
            // 
            // cb_data_for
            // 
            this.cb_data_for.FormattingEnabled = true;
            this.cb_data_for.Items.AddRange(new object[] {
            "ALL",
            "ACF_CLEANING_IMAGE",
            "ACF_FLATNESS",
            "ACF_BEFORE_TAPE_TEST_IMAGE",
            "ACF_AFTER_TAPE_TEST_IMAGE",
            "ACF_GLASS_COUPON_IMAGE",
            "ACF_GRAPH_FORCE_IMAGE",
            "ACF_ROUGHNESS"});
            this.cb_data_for.Location = new System.Drawing.Point(870, 5);
            this.cb_data_for.Name = "cb_data_for";
            this.cb_data_for.Size = new System.Drawing.Size(156, 21);
            this.cb_data_for.TabIndex = 41;
            this.cb_data_for.SelectedIndexChanged += new System.EventHandler(this.cb_data_for_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(808, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 16);
            this.label1.TabIndex = 40;
            this.label1.Text = "Data for ";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtLotNo
            // 
            this.txtLotNo.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLotNo.Location = new System.Drawing.Point(503, 4);
            this.txtLotNo.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtLotNo.Name = "txtLotNo";
            this.txtLotNo.Size = new System.Drawing.Size(100, 22);
            this.txtLotNo.TabIndex = 36;
            this.txtLotNo.TextChanged += new System.EventHandler(this.txtLotNo_TextChanged);
            this.txtLotNo.Validated += new System.EventHandler(this.txtLotNo_Validated);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel1, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.mnuAction, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 2;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 99.99999F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(1266, 745);
            this.tableLayoutPanel2.TabIndex = 43;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(614, 6);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 16);
            this.label2.TabIndex = 44;
            this.label2.Text = "Operator";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtOperator
            // 
            this.txtOperator.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOperator.Location = new System.Drawing.Point(682, 3);
            this.txtOperator.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtOperator.Name = "txtOperator";
            this.txtOperator.Size = new System.Drawing.Size(100, 22);
            this.txtOperator.TabIndex = 45;
            // 
            // ACF
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1266, 745);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtOperator);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.txtItemCode);
            this.Controls.Add(this.cb_data_for);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtLotNo);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ACF";
            this.Text = "ACF Process";
            this.Load += new System.EventHandler(this.ACF_Load);
            this.mnuAction.ResumeLayout(false);
            this.mnuAction.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).EndInit();
            this.tableLayoutPanel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picdetail)).EndInit();
            this.panel1.ResumeLayout(false);
            this.tableLayoutPanel5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Roughness_spec)).EndInit();
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtItemCode;
        private System.Windows.Forms.MenuStrip mnuAction;
        private System.Windows.Forms.ToolStripMenuItem tsmLoadData;
        private System.Windows.Forms.ToolStripMenuItem tsmExport;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label lbltitle;
        private System.Windows.Forms.DataGridView DGV_Data;
        private System.Windows.Forms.ComboBox cb_data_for;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtLotNo;
        private System.Windows.Forms.PictureBox picdetail;
        private System.Windows.Forms.Label lblImage;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.DataGridView DGV_Roughness_spec;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel5;
        private System.Windows.Forms.Label lbl_title_spec;
        private System.Windows.Forms.Button btn_load_spec;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtOperator;
    }
}