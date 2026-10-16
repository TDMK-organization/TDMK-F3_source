using System.ComponentModel;

namespace OK2SHIP_SMT.UserControls.Bending
{
    partial class CustomSpec
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CustomSpec));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.btn_cancel = new AntdUI.Button();
            this.btn_cf = new AntdUI.Button();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.label1 = new AntdUI.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btn_add = new AntdUI.Button();
            this.tb_logfile = new AntdUI.Input();
            this.select = new AntdUI.SelectMultiple();
            this.dgv = new AntdUI.Table();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 24.32432F));
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel3, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.dgv, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.4893F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 250F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 49F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(622, 367);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 3;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel2.Controls.Add(this.btn_cancel, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.btn_cf, 2, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 321);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(616, 43);
            this.tableLayoutPanel2.TabIndex = 0;
            // 
            // btn_cancel
            // 
            this.btn_cancel.BorderWidth = 2F;
            this.btn_cancel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_cancel.Ghost = true;
            this.btn_cancel.Location = new System.Drawing.Point(3, 3);
            this.btn_cancel.Name = "btn_cancel";
            this.btn_cancel.Size = new System.Drawing.Size(199, 37);
            this.btn_cancel.TabIndex = 0;
            this.btn_cancel.Text = "Cancel";
            this.btn_cancel.Type = AntdUI.TTypeMini.Error;
            this.btn_cancel.Click += new System.EventHandler(this.btn_cancel_Click);
            // 
            // btn_cf
            // 
            this.btn_cf.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_cf.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_cf.Location = new System.Drawing.Point(413, 3);
            this.btn_cf.Name = "btn_cf";
            this.btn_cf.Size = new System.Drawing.Size(200, 37);
            this.btn_cf.TabIndex = 1;
            this.btn_cf.Text = "Confirm";
            this.btn_cf.Type = AntdUI.TTypeMini.Primary;
            this.btn_cf.Click += new System.EventHandler(this.btn_cf_Click);
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 3;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 23.37662F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 54.54546F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 22.24026F));
            this.tableLayoutPanel3.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.panel1, 2, 0);
            this.tableLayoutPanel3.Controls.Add(this.select, 1, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 1;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(616, 62);
            this.tableLayoutPanel3.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(137, 56);
            this.label1.TabIndex = 1;
            this.label1.Text = "Cycle";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btn_add);
            this.panel1.Controls.Add(this.tb_logfile);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(481, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(132, 56);
            this.panel1.TabIndex = 7;
            // 
            // btn_add
            // 
            this.btn_add.BorderWidth = 1F;
            this.btn_add.Dock = System.Windows.Forms.DockStyle.Right;
            this.btn_add.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_add.IconGap = 0F;
            this.btn_add.IconHoverSvg = "";
            this.btn_add.IconSvg = resources.GetString("btn_add.IconSvg");
            this.btn_add.JoinLeft = true;
            this.btn_add.Location = new System.Drawing.Point(78, 0);
            this.btn_add.Margin = new System.Windows.Forms.Padding(2);
            this.btn_add.Name = "btn_add";
            this.btn_add.Size = new System.Drawing.Size(54, 56);
            this.btn_add.TabIndex = 6;
            this.btn_add.Type = AntdUI.TTypeMini.Primary;
            this.btn_add.WaveSize = 0;
            this.btn_add.Click += new System.EventHandler(this.btn_add_Click);
            // 
            // tb_logfile
            // 
            this.tb_logfile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tb_logfile.JoinRight = true;
            this.tb_logfile.Location = new System.Drawing.Point(0, 0);
            this.tb_logfile.Margin = new System.Windows.Forms.Padding(2);
            this.tb_logfile.Name = "tb_logfile";
            this.tb_logfile.PlaceholderText = "Value (%)";
            this.tb_logfile.Size = new System.Drawing.Size(132, 56);
            this.tb_logfile.TabIndex = 5;
            this.tb_logfile.WaveSize = 0;
            // 
            // select
            // 
            this.select.Dock = System.Windows.Forms.DockStyle.Fill;
            this.select.Location = new System.Drawing.Point(146, 3);
            this.select.Name = "select";
            this.select.PlaceholderText = "Select cycle";
            this.select.Size = new System.Drawing.Size(329, 56);
            this.select.TabIndex = 0;
            // 
            // dgv
            // 
            this.dgv.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv.Gap = 12;
            this.dgv.Location = new System.Drawing.Point(3, 71);
            this.dgv.Name = "dgv";
            this.dgv.Size = new System.Drawing.Size(616, 244);
            this.dgv.TabIndex = 2;
            this.dgv.Text = "table1";
            // 
            // CustomSpec
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(622, 367);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "CustomSpec";
            this.Text = "CustomSpec";
            this.Load += new System.EventHandler(this.CustomSpec_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.Panel panel1;

        private AntdUI.Button btn_add;

        private AntdUI.Input tb_logfile;

        private AntdUI.Label label1;

        private AntdUI.Table dgv;

        private AntdUI.SelectMultiple select;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;

        private AntdUI.Button btn_cancel;
        private AntdUI.Button btn_cf;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;

        #endregion
    }
}