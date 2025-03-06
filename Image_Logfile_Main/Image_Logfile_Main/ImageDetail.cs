using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using VHX;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace OK2SHIP
{
    public partial class ImageDetail : Form
    {
        public string path_start = "";
        private string path_select = "";
        private bool _click = false;
        //public ImageDetail()
        //{
        //    InitializeComponent();
        //}
        public ImageDetail(Get_path sender)
        {

            InitializeComponent();
            this.getpath = sender;
        }
        public delegate void Get_path(string path);


        //   public DelSendpath SendMsg;
        public Get_path getpath;

       


        private void button1_Click(object sender, EventArgs e)
        {
           Load_full_node(txtPath.Text);
           path_start = txtPath.Text;

        }

        public void Load_full_node(string src_path)
        {
            DirectoryInfo tem_dir = new DirectoryInfo(src_path);
            string folder_name = tem_dir.Name;
            TreeNode MainNode = new TreeNode();
            MainNode.Text = folder_name;

            treeView1.Nodes.Add(MainNode);

            DirectoryInfo di = new DirectoryInfo(src_path);           
            Get_sub_node(di, MainNode);
           
        }

        public void Get_sub_node(DirectoryInfo di, TreeNode main_node)
        {
            try
            {
                DirectoryInfo[] diArr = di.GetDirectories();
                foreach (DirectoryInfo dri in diArr)
                {
                    TreeNode Subnode = new TreeNode();
                    Subnode.Text = dri.Name;
                    main_node.Nodes.Add(Subnode);
                    Get_sub_node(dri, Subnode);
                }
            }
            catch
            {

            }
          
        
   
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            //txtPath.Text = Get_fullpath_after_select(txtPath.Text, treeView1.SelectedNode);
            //txtPath.Text =path_start.Replace(path_start.Split('\\').Last(), string.Empty) + Get_node_path(treeView1.SelectedNode);
           
            txtPath.Text = Directory.GetParent(path_start).ToString() + "\\" + Get_node_path(treeView1.SelectedNode);
        }
        public string Get_node_path(TreeNode node)
        {
            string str = "";
            if (node.Parent != null)
            {
                str += Get_node_path(node.Parent) + "\\" + node.Text;
            }
            else
            {
                str += node.Text;
            }
            return str;
        }
        public string startpath
        {
            get { return path_start; }
            set { path_start = value; }
        }
        public string selectpath
        {
            get { return path_select; }
            set { path_select = value; }
        }

        public bool click
        {
            get { return _click; }
            set { _click = value; }
        }

        private void treeView1_MouseClick(object sender, MouseEventArgs e)
        {           
        }

        private void treeView1_NodeMouseClick_1(object sender, TreeNodeMouseClickEventArgs e)
        {
           
        }

        private void txtPath_KeyDown(object sender, KeyEventArgs e)
        {
           
        }

        private void btn_selectpath_Click(object sender, EventArgs e)
        {
            if (txtPath.Text != "")
            {
                treeView1.Nodes.Clear();
                Load_full_node(txtPath.Text);
                path_start = txtPath.Text;
            }
        }

        private void ImageDetail_Load(object sender, EventArgs e)
        {
            txtPath.Text = path_start;
        }
        public void Get_Image_Multi(string in_src, ref List<Byte[]> lst_result)
        {
            try
            {
                DirectoryInfo tar_d = new DirectoryInfo(in_src);
                FileInfo[] temp_lst = tar_d.GetFiles("*.jpg");
                if (temp_lst.Length > 0)
                {
                    for (int i = 0; i < temp_lst.Length; i++)
                    {
                        var sel_img = Bitmap.FromFile(temp_lst[i].FullName);
                        ImageConverter imgcon = new ImageConverter();
                        byte[] img_data = (byte[])imgcon.ConvertTo(sel_img, typeof(byte[]));
                        lst_result.Add(img_data);

                    }
                }
            }
            catch 
            {
                MessageBox.Show("Could not find a part of the path " + in_src, "Warning");
            }
        
            //else
            //{
            //    DirectoryInfo[] inter_type_lst = tar_d.GetDirectories();
            //    if (inter_type_lst.Length > 0)
            //    {
            //        foreach (var inter_lst in inter_type_lst)
            //        {
            //            Get_Image_Multi(inter_lst.FullName, ref lst_result);
            //        }
            //    }
            //}
        }

        public void View_All_Image(string in_src)
        {
          //  DGV_Image.Columns.Clear();
            DGV_Image.Rows.Clear();
            DataTable dt_image = new DataTable();
          //  dt_image = (DataTable)DGV_Image.DataSource;
            List<Byte[]> lst_image = new List<Byte[]>();
            Get_Image_Multi(in_src, ref lst_image);
            int count_row = (lst_image.Count) / 8;
          
            if ( (lst_image.Count) % 8 != 0)
                count_row++;
            for(int i = 0;  i < count_row ; i++)
            {
                
                DGV_Image.Rows.Add();
                
                DataGridViewRow row = (DataGridViewRow)DGV_Image.Rows[i];
                for (int j = 0; j < 8; j ++)
                {
                    if(j < lst_image.Count - 8*i)
                    row.Cells[j].Value = lst_image[8 * i + j];
                }
            
            }
            for(int i =0; i < DGV_Image.Columns.Count ;i++)
            {
                ((DataGridViewImageColumn)DGV_Image.Columns[i]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                ((DataGridViewImageColumn)DGV_Image.Columns[i]).Width = 100;
            }
            
            foreach (DataGridViewRow dr in DGV_Image.Rows)
            {
                dr.Height = 70;
            }
            
        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            View_All_Image(txtPath.Text);
        }

        private void btn_accept_Click(object sender, EventArgs e)
        {
            //path_select = txtPath.Text;
            //_click = true;

            this.getpath(this.txtPath.Text);

            this.Close();
        }

        private void DGV_Image_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int col_inx = e.ColumnIndex;
            int r_inx = e.RowIndex;
            DataGridViewCell cur_cell = DGV_Image.CurrentCell;
            //if (DGV_Image.Columns[col_inx].Name.Contains("Image"))
            //{
                byte[] data = (byte[])cur_cell.Value;
            if(data != null)
            {
                using (MemoryStream ms = new MemoryStream(data))
                {
                    pic_detail.Image = Image.FromStream(ms);
                }
            }
               
           // }
        }
    }
}
