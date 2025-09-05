namespace OK2SHIP_SMT.Views
{
    partial class ReferForm
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.tdmK_Label2 = new OK2SHIP_SMT.ToolBoxs.TDMK_Label();
            this.tdmK_Label1 = new OK2SHIP_SMT.ToolBoxs.TDMK_Label();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.btn_refer = new OK2SHIP_SMT.ToolBoxs.TDMK_Button();
            this.btn_cancel = new OK2SHIP_SMT.ToolBoxs.TDMK_Button();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.lb_namerefer = new OK2SHIP_SMT.ToolBoxs.TDMK_Label();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 48.69685F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 51.30315F));
            this.tableLayoutPanel1.Controls.Add(this.listBox2, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.tdmK_Label2, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.tdmK_Label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.listBox1, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 52);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.881423F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 90.11858F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(729, 388);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // listBox2
            // 
            this.listBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox2.FormattingEnabled = true;
            this.listBox2.ItemHeight = 33;
            this.listBox2.Location = new System.Drawing.Point(358, 41);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(368, 344);
            this.listBox2.TabIndex = 6;
            // 
            // tdmK_Label2
            // 
            this.tdmK_Label2.AutoSize = true;
            this.tdmK_Label2.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.tdmK_Label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tdmK_Label2.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold);
            this.tdmK_Label2.Location = new System.Drawing.Point(358, 0);
            this.tdmK_Label2.Name = "tdmK_Label2";
            this.tdmK_Label2.Size = new System.Drawing.Size(368, 38);
            this.tdmK_Label2.TabIndex = 5;
            this.tdmK_Label2.Text = "List of referred codes";
            this.tdmK_Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tdmK_Label1
            // 
            this.tdmK_Label1.AutoSize = true;
            this.tdmK_Label1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.tdmK_Label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tdmK_Label1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold);
            this.tdmK_Label1.Location = new System.Drawing.Point(3, 0);
            this.tdmK_Label1.Name = "tdmK_Label1";
            this.tdmK_Label1.Size = new System.Drawing.Size(349, 38);
            this.tdmK_Label1.TabIndex = 1;
            this.tdmK_Label1.Text = "List ItemCode";
            this.tdmK_Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // listBox1
            // 
            this.listBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 33;
            this.listBox1.Location = new System.Drawing.Point(3, 41);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(349, 344);
            this.listBox1.TabIndex = 2;
            // 
            // btn_refer
            // 
            this.btn_refer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_refer.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold);
            this.btn_refer.Location = new System.Drawing.Point(3, 3);
            this.btn_refer.Name = "btn_refer";
            this.btn_refer.Size = new System.Drawing.Size(358, 51);
            this.btn_refer.TabIndex = 3;
            this.btn_refer.Text = "Add Refer";
            this.btn_refer.UseVisualStyleBackColor = true;
            this.btn_refer.Click += new System.EventHandler(this.btn_refer_Click);
            // 
            // btn_cancel
            // 
            this.btn_cancel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold);
            this.btn_cancel.Location = new System.Drawing.Point(367, 3);
            this.btn_cancel.Name = "btn_cancel";
            this.btn_cancel.Size = new System.Drawing.Size(359, 51);
            this.btn_cancel.TabIndex = 4;
            this.btn_cancel.Text = "Remove";
            this.btn_cancel.UseVisualStyleBackColor = true;
            this.btn_cancel.Click += new System.EventHandler(this.btn_cancel_Click);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel1, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.lb_namerefer, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel4, 0, 2);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 3;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.06719F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 88.93281F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 62F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(735, 506);
            this.tableLayoutPanel2.TabIndex = 1;
            // 
            // lb_namerefer
            // 
            this.lb_namerefer.AutoSize = true;
            this.lb_namerefer.BackColor = System.Drawing.Color.DarkOrange;
            this.lb_namerefer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lb_namerefer.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold);
            this.lb_namerefer.Location = new System.Drawing.Point(3, 0);
            this.lb_namerefer.Name = "lb_namerefer";
            this.lb_namerefer.Size = new System.Drawing.Size(729, 49);
            this.lb_namerefer.TabIndex = 1;
            this.lb_namerefer.Text = "Error";
            this.lb_namerefer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.ColumnCount = 2;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel4.Controls.Add(this.btn_refer, 0, 0);
            this.tableLayoutPanel4.Controls.Add(this.btn_cancel, 1, 0);
            this.tableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel4.Location = new System.Drawing.Point(3, 446);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 1;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(729, 57);
            this.tableLayoutPanel4.TabIndex = 3;
            // 
            // ReferForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(735, 506);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Name = "ReferForm";
            this.Text = "ReferForm";
            this.Load += new System.EventHandler(this.ReferForm_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.tableLayoutPanel4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private ToolBoxs.TDMK_Label lb_namerefer;
        private ToolBoxs.TDMK_Label tdmK_Label1;
        private System.Windows.Forms.ListBox listBox1;
        private ToolBoxs.TDMK_Button btn_refer;
        private ToolBoxs.TDMK_Button btn_cancel;
        private System.Windows.Forms.ListBox listBox2;
        private ToolBoxs.TDMK_Label tdmK_Label2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
    }
}