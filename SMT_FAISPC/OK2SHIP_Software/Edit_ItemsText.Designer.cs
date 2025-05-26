namespace OK2SHIP_Software
{
    partial class Edit_ItemsText
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
            this.GBIPQC_Items = new System.Windows.Forms.GroupBox();
            this.RB_CoverLay_Process = new System.Windows.Forms.RadioButton();
            this.RB_Printing_Process = new System.Windows.Forms.RadioButton();
            this.RB_CopperPlating = new System.Windows.Forms.RadioButton();
            this.RB_Etching_Process = new System.Windows.Forms.RadioButton();
            this.btnExit = new System.Windows.Forms.Button();
            this.DGV_Data = new System.Windows.Forms.DataGridView();
            this.RB_Depart = new System.Windows.Forms.GroupBox();
            this.RB_Product = new System.Windows.Forms.RadioButton();
            this.RB_UV_Process = new System.Windows.Forms.RadioButton();
            this.RB_Lab = new System.Windows.Forms.RadioButton();
            this.btnLoad = new System.Windows.Forms.Button();
            this.GBIPQC_Items.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).BeginInit();
            this.RB_Depart.SuspendLayout();
            this.SuspendLayout();
            // 
            // GBIPQC_Items
            // 
            this.GBIPQC_Items.Controls.Add(this.RB_CoverLay_Process);
            this.GBIPQC_Items.Controls.Add(this.RB_Printing_Process);
            this.GBIPQC_Items.Controls.Add(this.RB_CopperPlating);
            this.GBIPQC_Items.Controls.Add(this.RB_Etching_Process);
            this.GBIPQC_Items.Location = new System.Drawing.Point(12, 3);
            this.GBIPQC_Items.Name = "GBIPQC_Items";
            this.GBIPQC_Items.Size = new System.Drawing.Size(547, 58);
            this.GBIPQC_Items.TabIndex = 15;
            this.GBIPQC_Items.TabStop = false;
            this.GBIPQC_Items.Text = "Select Process";
            // 
            // RB_CoverLay_Process
            // 
            this.RB_CoverLay_Process.AutoSize = true;
            this.RB_CoverLay_Process.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RB_CoverLay_Process.Location = new System.Drawing.Point(408, 23);
            this.RB_CoverLay_Process.Name = "RB_CoverLay_Process";
            this.RB_CoverLay_Process.Size = new System.Drawing.Size(124, 18);
            this.RB_CoverLay_Process.TabIndex = 6;
            this.RB_CoverLay_Process.TabStop = true;
            this.RB_CoverLay_Process.Text = "Cover_Lay_Process";
            this.RB_CoverLay_Process.UseVisualStyleBackColor = true;
            this.RB_CoverLay_Process.CheckedChanged += new System.EventHandler(this.RB_CoverLay_Process_CheckedChanged);
            // 
            // RB_Printing_Process
            // 
            this.RB_Printing_Process.AutoSize = true;
            this.RB_Printing_Process.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RB_Printing_Process.Location = new System.Drawing.Point(288, 23);
            this.RB_Printing_Process.Name = "RB_Printing_Process";
            this.RB_Printing_Process.Size = new System.Drawing.Size(106, 18);
            this.RB_Printing_Process.TabIndex = 3;
            this.RB_Printing_Process.TabStop = true;
            this.RB_Printing_Process.Text = "Printing_Process";
            this.RB_Printing_Process.UseVisualStyleBackColor = true;
            this.RB_Printing_Process.CheckedChanged += new System.EventHandler(this.RB_Printing_Process_CheckedChanged);
            // 
            // RB_CopperPlating
            // 
            this.RB_CopperPlating.AutoSize = true;
            this.RB_CopperPlating.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RB_CopperPlating.Location = new System.Drawing.Point(131, 23);
            this.RB_CopperPlating.Name = "RB_CopperPlating";
            this.RB_CopperPlating.Size = new System.Drawing.Size(143, 18);
            this.RB_CopperPlating.TabIndex = 1;
            this.RB_CopperPlating.TabStop = true;
            this.RB_CopperPlating.Text = "Copper_Plating_Process";
            this.RB_CopperPlating.UseVisualStyleBackColor = true;
            this.RB_CopperPlating.CheckedChanged += new System.EventHandler(this.RB_CopperPlating_CheckedChanged);
            // 
            // RB_Etching_Process
            // 
            this.RB_Etching_Process.AutoSize = true;
            this.RB_Etching_Process.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RB_Etching_Process.Location = new System.Drawing.Point(11, 23);
            this.RB_Etching_Process.Name = "RB_Etching_Process";
            this.RB_Etching_Process.Size = new System.Drawing.Size(106, 18);
            this.RB_Etching_Process.TabIndex = 0;
            this.RB_Etching_Process.TabStop = true;
            this.RB_Etching_Process.Text = "Etching_Process";
            this.RB_Etching_Process.UseVisualStyleBackColor = true;
            this.RB_Etching_Process.CheckedChanged += new System.EventHandler(this.RB_Etching_Process_CheckedChanged);
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(468, 520);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(124, 57);
            this.btnExit.TabIndex = 14;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // DGV_Data
            // 
            this.DGV_Data.AllowUserToAddRows = false;
            this.DGV_Data.AllowUserToDeleteRows = false;
            this.DGV_Data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Data.Location = new System.Drawing.Point(12, 67);
            this.DGV_Data.Name = "DGV_Data";
            this.DGV_Data.Size = new System.Drawing.Size(803, 447);
            this.DGV_Data.TabIndex = 13;
            this.DGV_Data.ColumnHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_Data_ColumnHeaderMouseDoubleClick);
            this.DGV_Data.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DGV_Data_RowHeaderMouseClick);
            // 
            // RB_Depart
            // 
            this.RB_Depart.Controls.Add(this.RB_Product);
            this.RB_Depart.Controls.Add(this.RB_UV_Process);
            this.RB_Depart.Controls.Add(this.RB_Lab);
            this.RB_Depart.Location = new System.Drawing.Point(565, 3);
            this.RB_Depart.Name = "RB_Depart";
            this.RB_Depart.Size = new System.Drawing.Size(250, 58);
            this.RB_Depart.TabIndex = 16;
            this.RB_Depart.TabStop = false;
            this.RB_Depart.Text = "Depart";
            // 
            // RB_Product
            // 
            this.RB_Product.AutoSize = true;
            this.RB_Product.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RB_Product.Location = new System.Drawing.Point(70, 23);
            this.RB_Product.Name = "RB_Product";
            this.RB_Product.Size = new System.Drawing.Size(76, 18);
            this.RB_Product.TabIndex = 1;
            this.RB_Product.TabStop = true;
            this.RB_Product.Text = "Production";
            this.RB_Product.UseVisualStyleBackColor = true;
            this.RB_Product.CheckedChanged += new System.EventHandler(this.RB_Etching_CheckedChanged);
            // 
            // RB_UV_Process
            // 
            this.RB_UV_Process.AutoSize = true;
            this.RB_UV_Process.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RB_UV_Process.Location = new System.Drawing.Point(161, 23);
            this.RB_UV_Process.Name = "RB_UV_Process";
            this.RB_UV_Process.Size = new System.Drawing.Size(86, 18);
            this.RB_UV_Process.TabIndex = 4;
            this.RB_UV_Process.TabStop = true;
            this.RB_UV_Process.Text = "UV_Process";
            this.RB_UV_Process.UseVisualStyleBackColor = true;
            this.RB_UV_Process.CheckedChanged += new System.EventHandler(this.RB_UV_Process_CheckedChanged);
            // 
            // RB_Lab
            // 
            this.RB_Lab.AutoSize = true;
            this.RB_Lab.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RB_Lab.Location = new System.Drawing.Point(4, 23);
            this.RB_Lab.Name = "RB_Lab";
            this.RB_Lab.Size = new System.Drawing.Size(46, 18);
            this.RB_Lab.TabIndex = 0;
            this.RB_Lab.TabStop = true;
            this.RB_Lab.Text = "LAB";
            this.RB_Lab.UseVisualStyleBackColor = true;
            this.RB_Lab.CheckedChanged += new System.EventHandler(this.RB_Lab_CheckedChanged);
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(199, 520);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(124, 57);
            this.btnLoad.TabIndex = 17;
            this.btnLoad.Text = "Load Data";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // Edit_ItemsText
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(833, 588);
            this.Controls.Add(this.GBIPQC_Items);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.DGV_Data);
            this.Controls.Add(this.RB_Depart);
            this.Controls.Add(this.btnLoad);
            this.Name = "Edit_ItemsText";
            this.Text = "Edit_ItemsText";
            this.Load += new System.EventHandler(this.Edit_ItemsText_Load);
            this.GBIPQC_Items.ResumeLayout(false);
            this.GBIPQC_Items.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Data)).EndInit();
            this.RB_Depart.ResumeLayout(false);
            this.RB_Depart.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox GBIPQC_Items;
        private System.Windows.Forms.RadioButton RB_CoverLay_Process;
        private System.Windows.Forms.RadioButton RB_Printing_Process;
        private System.Windows.Forms.RadioButton RB_CopperPlating;
        private System.Windows.Forms.RadioButton RB_Etching_Process;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.DataGridView DGV_Data;
        private System.Windows.Forms.GroupBox RB_Depart;
        private System.Windows.Forms.RadioButton RB_Product;
        private System.Windows.Forms.RadioButton RB_UV_Process;
        private System.Windows.Forms.RadioButton RB_Lab;
        private System.Windows.Forms.Button btnLoad;
    }
}