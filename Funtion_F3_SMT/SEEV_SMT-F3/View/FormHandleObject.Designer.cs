using SEEV_SMT_F3.TDMK_Tool;

namespace SEEV_SMT_F3
{
    partial class FormHandleObject
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
            menuStrip1 = new MenuStrip();
            toolStripMenuItem1 = new ToolStripMenuItem();
            toolStripMenuItem2 = new ToolStripMenuItem();
            toolStripSeparator1 = new ToolStripSeparator();
            exitToolStripMenuItem = new ToolStripMenuItem();
            toolStripComboBox1 = new DropDown();
            tableLayoutPanel = new TableDiv();
            userControlHandle1 = new SEEV_SMT_F3.UserControls.UserControlHandle();
            menuStrip1.SuspendLayout();
            tableLayoutPanel.SuspendLayout();
            SuspendLayout();
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, toolStripComboBox1 });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(1041, 33);
            menuStrip1.TabIndex = 0;
            menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.DropDownItems.AddRange(new ToolStripItem[] { toolStripMenuItem2, toolStripSeparator1, exitToolStripMenuItem });
            toolStripMenuItem1.Font = new Font("Segoe UI", 14.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(103, 29);
            toolStripMenuItem1.Text = "Function";
            // 
            // toolStripMenuItem2
            // 
            toolStripMenuItem2.Name = "toolStripMenuItem2";
            toolStripMenuItem2.Size = new Size(131, 30);
            toolStripMenuItem2.Text = "Reset";
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new Size(128, 6);
            // 
            // exitToolStripMenuItem
            // 
            exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            exitToolStripMenuItem.Size = new Size(131, 30);
            exitToolStripMenuItem.Text = "Exit";
            // 
            // toolStripComboBox1
            // 
            toolStripComboBox1.ForeColor = Color.Gray;
            toolStripComboBox1.Items.AddRange(new object[] { " CROSS_SECTION", " GAP_CONNECTOR", " PEEL_TEST", " MATING_PULL_TEST", " SHEAR_TEST", " IQC_UNMATING_PULL_TEST", " IQC_LINER_PEELING_COUPON", " IQC_PSA_PEELING_COUPON", " LINER_PEEL_TEST_ON_PRODUCT", " PSA_PEEL_TEST_ON_PRODUCT", " ACF" });
            toolStripComboBox1.Name = "toolStripComboBox1";
            toolStripComboBox1.Size = new Size(200, 29);
            toolStripComboBox1.Text = "Chosse Your Process";
            toolStripComboBox1.SelectedIndexChanged += toolStripComboBox1_SelectedIndexChanged;
            // 
            // tableLayoutPanel
            // 
            tableLayoutPanel.BackColor = SystemColors.Window;
            tableLayoutPanel.ColumnCount = 1;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel.Controls.Add(userControlHandle1, 0, 1);
            tableLayoutPanel.Dock = DockStyle.Fill;
            tableLayoutPanel.Location = new Point(0, 33);
            tableLayoutPanel.Name = "tableLayoutPanel";
            tableLayoutPanel.RowCount = 2;
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 486F));
            tableLayoutPanel.Size = new Size(1041, 537);
            tableLayoutPanel.TabIndex = 2;
            // 
            // userControlHandle1
            // 
            userControlHandle1.BackColor = SystemColors.ActiveCaption;
            userControlHandle1.Dock = DockStyle.Fill;
            userControlHandle1.Location = new Point(3, 54);
            userControlHandle1.Name = "userControlHandle1";
            userControlHandle1.Size = new Size(1035, 480);
            userControlHandle1.TabIndex = 0;
            // 
            // FormHandleObject
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1041, 570);
            Controls.Add(tableLayoutPanel);
            Controls.Add(menuStrip1);
            MainMenuStrip = menuStrip1;
            Name = "FormHandleObject";
            Text = "FormHandleObject";
            FormClosed += FormHandleObject_FormClosed;
            Load += FormHandleObject_Load;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            tableLayoutPanel.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private MenuStrip menuStrip1;
        private DropDown toolStripComboBox1;
        private TableDiv tableLayoutPanel;
        private ToolStripMenuItem toolStripMenuItem1;
        private ToolStripMenuItem exitToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripMenuItem toolStripMenuItem2;
        private UserControls.UserControlHandle userControlHandle1;
    }
}