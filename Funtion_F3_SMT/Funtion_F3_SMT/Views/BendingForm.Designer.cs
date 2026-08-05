using System.ComponentModel;

namespace OK2SHIP_SMT.Views
{
    partial class BendingForm
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.btn_new = new AntdUI.Button();
            this.btn_oldversion = new AntdUI.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.btn_new, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.btn_oldversion, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 13.17365F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(901, 167);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // btn_new
            // 
            this.btn_new.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_new.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_new.Location = new System.Drawing.Point(453, 3);
            this.btn_new.Name = "btn_new";
            this.btn_new.Size = new System.Drawing.Size(445, 161);
            this.btn_new.TabIndex = 2;
            this.btn_new.Text = "New Version";
            this.btn_new.Type = AntdUI.TTypeMini.Primary;
            this.btn_new.Click += new System.EventHandler(this.btn_new_Click_1);
            // 
            // btn_oldversion
            // 
            this.btn_oldversion.BorderWidth = 2F;
            this.btn_oldversion.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_oldversion.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_oldversion.Ghost = true;
            this.btn_oldversion.Location = new System.Drawing.Point(3, 3);
            this.btn_oldversion.Name = "btn_oldversion";
            this.btn_oldversion.Size = new System.Drawing.Size(444, 161);
            this.btn_oldversion.TabIndex = 1;
            this.btn_oldversion.Text = "Old Version";
            this.btn_oldversion.Type = AntdUI.TTypeMini.Primary;
            // 
            // BendingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(901, 167);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "BendingForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Option";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private AntdUI.Button btn_new;

        private AntdUI.Button btn_oldversion;

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;

        #endregion
    }
}