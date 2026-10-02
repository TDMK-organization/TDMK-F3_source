using System.ComponentModel;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    partial class RunProgess
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.dgv = new AntdUI.Table();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.btn_run = new AntdUI.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 74.32258F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.67742F));
            this.tableLayoutPanel1.Controls.Add(this.dgv, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(775, 418);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // dgv
            // 
            this.dgv.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv.Gap = 12;
            this.dgv.Location = new System.Drawing.Point(3, 3);
            this.dgv.Name = "dgv";
            this.dgv.Size = new System.Drawing.Size(570, 412);
            this.dgv.TabIndex = 0;
            this.dgv.Text = "table1";
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Controls.Add(this.btn_run, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(579, 3);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 2;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 91F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(193, 412);
            this.tableLayoutPanel2.TabIndex = 1;
            // 
            // btn_run
            // 
            this.btn_run.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_run.Location = new System.Drawing.Point(3, 3);
            this.btn_run.Name = "btn_run";
            this.btn_run.Size = new System.Drawing.Size(187, 85);
            this.btn_run.TabIndex = 0;
            this.btn_run.Text = "Run Progess";
            this.btn_run.Click += new System.EventHandler(this.btn_run_Click);
            // 
            // RunProgess
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "RunProgess";
            this.Size = new System.Drawing.Size(775, 418);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private AntdUI.Button btn_run;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;

        private AntdUI.Table dgv;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;

        #endregion
    }
}