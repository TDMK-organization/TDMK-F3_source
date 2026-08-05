namespace OK2SHIP_SMT.Views
{
    partial class BendingWindow
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
            AntdUI.Tabs.StyleCard styleCard1 = new AntdUI.Tabs.StyleCard();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BendingWindow));
            this.titlebar = new AntdUI.PageHeader();
            this.button_collapse = new AntdUI.Button();
            this.tabs = new AntdUI.Tabs();
            this.tabPage = new AntdUI.TabPage();
            this.panel_content = new AntdUI.Panel();
            this.menu = new AntdUI.Menu();
            this.panel_left = new AntdUI.Panel();
            this.panel1 = new AntdUI.Panel();
            this.tabs.SuspendLayout();
            this.tabPage.SuspendLayout();
            this.panel_left.SuspendLayout();
            this.SuspendLayout();
            // 
            // titlebar
            // 
            this.titlebar.DividerShow = true;
            this.titlebar.Dock = System.Windows.Forms.DockStyle.Top;
            this.titlebar.Location = new System.Drawing.Point(0, 0);
            this.titlebar.Name = "titlebar";
            this.titlebar.ShowButton = true;
            this.titlebar.ShowIcon = true;
            this.titlebar.Size = new System.Drawing.Size(1225, 36);
            this.titlebar.SubText = "Demo";
            this.titlebar.TabIndex = 0;
            this.titlebar.Text = "AntdUI";
            // 
            // button_collapse
            // 
            this.button_collapse.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.button_collapse.Ghost = true;
            this.button_collapse.IconRatio = 1F;
            this.button_collapse.IconSvg = "MenuUnfoldOutlined";
            this.button_collapse.Location = new System.Drawing.Point(0, 564);
            this.button_collapse.Name = "button_collapse";
            this.button_collapse.Radius = 0;
            this.button_collapse.Size = new System.Drawing.Size(58, 40);
            this.button_collapse.TabIndex = 1;
            this.button_collapse.ToggleIconSvg = "MenuFoldOutlined";
            this.button_collapse.WaveSize = 0;
            // 
            // tabs
            // 
            this.tabs.CloseDisposePage = true;
            this.tabs.Controls.Add(this.tabPage);
            this.tabs.Cursor = System.Windows.Forms.Cursors.Default;
            this.tabs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabs.Gap = 16;
            this.tabs.Location = new System.Drawing.Point(58, 36);
            this.tabs.Margin = new System.Windows.Forms.Padding(4, 4, 4, 0);
            this.tabs.Name = "tabs";
            this.tabs.Pages.Add(this.tabPage);
            this.tabs.Size = new System.Drawing.Size(1167, 604);
            styleCard1.Closable = true;
            this.tabs.Style = styleCard1;
            this.tabs.TabIndex = 2;
            this.tabs.Type = AntdUI.TabType.Card;
            // 
            // tabPage
            // 
            this.tabPage.Controls.Add(this.panel_content);
            this.tabPage.IconSvg = "HomeOutlined";
            this.tabPage.Location = new System.Drawing.Point(0, 37);
            this.tabPage.Name = "tabPage";
            this.tabPage.ReadOnly = true;
            this.tabPage.Size = new System.Drawing.Size(1167, 567);
            this.tabPage.TabIndex = 1;
            this.tabPage.Text = "Home page";
            // 
            // panel_content
            // 
            this.panel_content.Back = System.Drawing.Color.Transparent;
            this.panel_content.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel_content.Location = new System.Drawing.Point(0, 0);
            this.panel_content.Name = "panel_content";
            this.panel_content.Radius = 0;
            this.panel_content.Size = new System.Drawing.Size(1167, 567);
            this.panel_content.TabIndex = 0;
            // 
            // menu
            // 
            this.menu.Collapsed = true;
            this.menu.Dock = System.Windows.Forms.DockStyle.Fill;
            this.menu.IconRatio = 1F;
            this.menu.Indent = true;
            this.menu.Location = new System.Drawing.Point(0, 0);
            this.menu.Name = "menu";
            this.menu.Padding = new System.Windows.Forms.Padding(4);
            this.menu.Size = new System.Drawing.Size(58, 564);
            this.menu.TabIndex = 0;
            this.menu.Unique = true;
            // 
            // panel_left
            // 
            this.panel_left.Back = System.Drawing.Color.Transparent;
            this.panel_left.Controls.Add(this.menu);
            this.panel_left.Controls.Add(this.button_collapse);
            this.panel_left.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel_left.Location = new System.Drawing.Point(0, 36);
            this.panel_left.Name = "panel_left";
            this.panel_left.Radius = 0;
            this.panel_left.Size = new System.Drawing.Size(58, 604);
            this.panel_left.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.Back = System.Drawing.Color.Transparent;
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Radius = 0;
            this.panel1.Size = new System.Drawing.Size(966, 567);
            this.panel1.TabIndex = 0;
            // 
            // BendingWindow
            // 
            this.ClientSize = new System.Drawing.Size(1225, 640);
            this.ControlBox = false;
            this.Controls.Add(this.tabs);
            this.Controls.Add(this.panel_left);
            this.Controls.Add(this.titlebar);
            this.Font = new System.Drawing.Font("Microsoft YaHei UI", 10F);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "BendingWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AntdUI Demo";
            this.tabs.ResumeLayout(false);
            this.tabPage.ResumeLayout(false);
            this.panel_left.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private AntdUI.TabPage tabPage;
        private AntdUI.Panel panel1;

        #endregion

        private AntdUI.PageHeader titlebar;
        private AntdUI.Panel panel_left;
        private AntdUI.Button button_collapse;
        private AntdUI.Menu menu;
        private AntdUI.Tabs tabs;
        private AntdUI.Panel panel_content;
    }
}