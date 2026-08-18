using AntdUI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using OK2SHIP_SMT.UserControls;
using OK2SHIP_SMT.UserControls.Bending;

namespace OK2SHIP_SMT.Views
{
    public partial class BendingWindow : AntdUI.Window
    {
        private UserControl currControl;
        private bool isUpdatingTabs = false;
        public event EventHandler LanguageChanged;

        public BendingWindow()
        {
            InitializeComponent();
            Config.IsLight = true;
            titlebar.Text = "Bending " + titlebar.ProductVersion;
            LoadMenu();
            InitData();
            BindEventHandler();
        }

        private void LoadMenu(string filter = "")
        {
            menu.Items.Clear();

            
            Dictionary<string, string[]> menuItems = new Dictionary<string, string[]>
            {
                {
                    "bending_data", new[]
                    {
                        "BarChartOutlined", // Icon Svg
                        "Bending Data", // Text
                    }
                },
                {
                    "xray_pic", new[]
                    {
                        "LayoutOutlined",
                        "X-ray Picture",
                    }
                }, 
                {
                    "ers_spec", new[]
                    {
                        "AppstoreOutlined",
                        "Ers spec",
                    }
                },
            };
            foreach (KeyValuePair<string, string[]> rootItem in menuItems)
            {
                var rootKey = rootItem.Key.ToLower();
                var rootMenu = new AntdUI.MenuItem
                {
                    Tag = rootItem.Key,
                    Text = rootItem.Value[1],
                    IconSvg = menuItems.TryGetValue(rootItem.Key, out string[] icon) ? icon[0] : "UnorderedListOutlined",
                };
           
                if (rootKey.Contains(filter))
                {
                    menu.Items.Add(rootMenu);
                }
            }
        }


        private void InitData()
        {
            // Theme().Button(button_color);
            // Config.IsLight = ThemeHelper.IsLightMode();
            // dropdown_translate.SelectedValue = dropdown_translate.Items[0];
            // //初始化消息弹出位置
            // Config.ShowInWindow = true;
            //
            // UserControl control = new Wellcome(this) { Dock = DockStyle.Fill };
            // AutoDpi(control);
            // panel_content.Controls.Add(control);
            //
            // //global
            // tabs.Pages[0].Text = AntdUI.Localization.Get("home", "主页");
        }

        private void BindEventHandler()
        {
            menu.SelectChanged += Menu_SelectChanged;
            button_collapse.Click += Button_collapse_Click;
            tabs.Click += Tabs_Click;
            tabs.SelectedIndexChanged += Tabs_SelectedIndexChanged;
        }

        private void Tabs_Click(object sender, EventArgs e)
        {
            MouseEventArgs mouseEventArgs = e as MouseEventArgs;

            if (mouseEventArgs != null)
            {
                if (mouseEventArgs.Button == MouseButtons.Right)
                {
                    string closeall = AntdUI.Localization.Get("closeall", "关闭所有选项卡");
                    var menulist = new AntdUI.IContextMenuStripItem[]
                    {
                        new AntdUI.ContextMenuStripItem(closeall)
                        {
                            IconSvg = "CloseOutlined"
                        }
                    };

                    AntdUI.ContextMenuStrip.open(tabs, item =>
                    {
                        if (item.Text == closeall)
                        {
                            tabs.SelectedIndex = 0;
                            if (tabs.Pages.Count > 1)
                                tabs.Pages.RemoveRange(1, tabs.Pages.Count - 1); // 从索引1开始，移除后面的所有页面
                            menu.Select(null);
                            menu.Refresh();
                        }
                    }, menulist);
                }
            }
        }


        private void Tabs_SelectedIndexChanged(object sender, IntEventArgs e)
        {
            SelectMenu();
        }

        private void Button_collapse_Click(object sender, EventArgs e)
        {
            menu.Collapsed = button_collapse.Toggle = !menu.Collapsed;
            if (menu.Collapsed) panel_left.Width = (int)(50 * Config.Dpi);
            else panel_left.Width = (int)(250 * Config.Dpi);
        }



   

        private void SelectMenu()
        {
            if (isUpdatingTabs) return;
            var text = tabs.SelectedTab?.Text;
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            if (text == AntdUI.Localization.Get("home", "主页"))
            {
                return;
            }

            var rootIndex = 0;
            var subIndex = 0;
            var menuItemsCopy = menu.Items.ToList(); // 创建副本
            for (int i = 0; i < menuItemsCopy.Count; i++)
            {
                for (int j = 0; j < menuItemsCopy[i].Sub.Count; j++)
                {
                    if (menuItemsCopy[i].Sub[j].Tag.ToString() == text)
                    {
                        rootIndex = i;
                        subIndex = j;
                        break;
                    }
                }
            }

            menu.SelectIndex(rootIndex, subIndex, true);
        }

        private void Menu_SelectChanged(object sender, MenuSelectEventArgs e)
        {
            string name = (string)e.Value.Tag;

            foreach (var tab in tabs.Pages)
            {
                if (tab is AntdUI.TabPage existingTab && existingTab.Text == name)
                {
                    isUpdatingTabs = true;
                    tabs.SelectedTab = existingTab; 
                    isUpdatingTabs = false;
                    currControl = existingTab.Controls.Count > 0 ? existingTab.Controls[0] as UserControl : null;
                    return;
                }
            }

            UserControl control = null;
            switch (name)
            {
                case "bending_data":
                    control = new UC_BendingData(this);
                    break;
                case "xray_pic":
                    control = new SEM("X-Ray picture") { Dock = DockStyle.Fill };
                    break;
                case "ers_spec":
                    control = new UC_ERS_NETSpec(this);
                    break;
                default:
                    break;
            }

            if (control != null)
            {
                control.Dock = DockStyle.Fill;
                AutoDpi(control);
                AntdUI.TabPage tabPage = new AntdUI.TabPage()
                {
                    Text = name,
                };
                tabPage.Controls.Add(control);
                isUpdatingTabs = true;
                tabs.AddTabSelect(tabPage);
                isUpdatingTabs = false;
                currControl = control;
            }
        }

      

        private void button_color_Click(object sender, EventArgs e)
        {
            Config.IsLight = !Config.IsLight;
        }

        private void titlebar_StyleChanged(object sender, EventArgs e)
        {
            Debugger.Break();
        }
    }
}