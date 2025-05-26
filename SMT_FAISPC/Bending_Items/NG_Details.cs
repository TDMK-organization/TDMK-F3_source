using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Resize_Lib;

namespace Bending_Items
{
    public partial class NG_Details : UserControl
    {
        Resize rs = new Resize();
        ECheck_Process proc_data = new ECheck_Process();
        public List<ECheck_Process.NG_list2> myNGlst2 = new List<ECheck_Process.NG_list2>();
        string lblItem_title = "NG NET List of Item: ";
        public DataTable Net_Spec_tbl { get; set; }
        public List<int> NET_index_lst { get; set; }
        public List<string> item_NG_lst { get; set; }
        public NG_Details()
        {
            InitializeComponent();
        }

        private void NG_Details_Load(object sender, EventArgs e)
        {
            form_load();
        }
        public void fill_data(DataTable src_data, DataTable src_Spec_tbl)
        {
            DataTable tbl_NG_NET = new DataTable();
            List<int> _NET_index_lst = new List<int>();
            List<string> _item_NG_lst = new List<string>();
            Net_Spec_tbl = src_Spec_tbl;
            proc_data.Submit_Data_2(src_Spec_tbl, src_data, ref tbl_NG_NET, ref _NET_index_lst, DGV_NET_NG, ref _item_NG_lst, ref myNGlst2);
            lstNG.DataSource = _item_NG_lst;
            txtTotal_NET.Text = _NET_index_lst.Count.ToString();
            txtTotalItem.Text = _item_NG_lst.Count.ToString();
            NET_index_lst = _NET_index_lst;
            item_NG_lst = _item_NG_lst;
        }

        private void lstNG_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstNG.SelectedIndex > -1)
            {
                int inx = lstNG.SelectedIndex;
                lblItemNG.Text = lblItem_title + lstNG.Items[inx];
                DGV_NG_detail.DataSource = proc_data.NG_NET_detail(Net_Spec_tbl, myNGlst2[inx]);
                int r_inx = 0;
                foreach (int index in myNGlst2[inx].lst_row_inx)
                {
                    DGV_NG_detail.Rows[r_inx].HeaderCell.Value = (index + 1).ToString();
                    r_inx++;
                }
                DGV_NG_detail.AutoResizeColumns();
                DGV_NG_detail.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            }
            else
            {
                lblItemNG.Text = lblItem_title;
            }
        }
        public void reset_all()
        {
            DGV_NET_NG.DataSource = null;
            DGV_NG_detail.DataSource = null;
            lstNG.DataSource = null;
            txtTotalItem.Clear();
            txtTotal_NET.Clear();
        }

        private void NG_Details_Resize(object sender, EventArgs e)
        {
            
        }
        public void resize()
        {
            rs.ResizeAllControls(this);
        }
        public void form_load()
        {
            rs.FindAllControls(this);
        }
    }
}
