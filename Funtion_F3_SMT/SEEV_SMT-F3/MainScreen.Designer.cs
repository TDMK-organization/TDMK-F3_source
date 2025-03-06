namespace SEEV_SMT_F3
{
    partial class MainScreen
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainScreen));
            btn_setup = new Button();
            btn_InputData = new Button();
            Tablelayout = new TableLayoutPanel();
            Tablelayout.SuspendLayout();
            SuspendLayout();
            // 
            // btn_setup
            // 
            resources.ApplyResources(btn_setup, "btn_setup");
            btn_setup.Name = "btn_setup";
            btn_setup.UseVisualStyleBackColor = true;
            // 
            // btn_InputData
            // 
            resources.ApplyResources(btn_InputData, "btn_InputData");
            btn_InputData.Name = "btn_InputData";
            btn_InputData.UseVisualStyleBackColor = true;
            btn_InputData.Click += btn_InputData_Click;
            // 
            // Tablelayout
            // 
            resources.ApplyResources(Tablelayout, "Tablelayout");
            Tablelayout.Controls.Add(btn_InputData, 1, 0);
            Tablelayout.Controls.Add(btn_setup, 0, 0);
            Tablelayout.Name = "Tablelayout";
            // 
            // MainScreen
            // 
            resources.ApplyResources(this, "$this");
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(Tablelayout);
            MaximizeBox = false;
            Name = "MainScreen";
            FormClosing += MainScreen_FormClosing;
            Tablelayout.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Button btn_setup;
        private Button btn_InputData;
        private TableLayoutPanel Tablelayout;
    }
}