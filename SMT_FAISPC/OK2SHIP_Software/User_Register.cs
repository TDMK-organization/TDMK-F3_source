using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TDMK_SQL;
using OK2SHIP_Measurements;

namespace OK2SHIP_Software
{
    public partial class User_Register : Form
    {
        TDMK_SQL_Lib TDMK_Code = new TDMK_SQL_Lib();
        public string depart = "QA";
        public User_Register()
        {
            InitializeComponent();
        }
        private void User_Register_Load(object sender, EventArgs e)
        {
            TDMK_Code.fill_dataset_DGV("UserInfo", TDMK_Code.SQL_CMD("UserInfo"),DGV_User, myVar.sqlcon_OK2SHIP);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if( check_empty())
            {
                MessageBox.Show("Please, fill data", "Warning");
            }
            else
            {
                if (txtPassword.Text != txtReTypePass.Text)
                {
                    MessageBox.Show("Please, enter password again!", "Warning");
                }
                else
                {
                    string[] item_arr = new string[5];
                    string[] item_arr_val = new string[5];
                    AutoCompleteStringCollection userID_lst = new AutoCompleteStringCollection();
                    item_arr[0] = "ID";
                    item_arr[1] = "User_Name";
                    item_arr[2] = "User_ID";
                    item_arr[3] = "User_Password";
                    item_arr[4] = "Department";
                    item_arr_val[1] = txtUsername.Text;
                    item_arr_val[2] = txtUserID.Text;
                    item_arr_val[3] = txtPassword.Text;
                    item_arr_val[4] = depart;
                    userID_lst = TDMK_Code.Load_Item_Names("UserInfo", "User_ID", myVar.sqlcon_OK2SHIP);
                    if(!TDMK_Code.check_exist(txtUserID.Text,userID_lst))
                    {
                        item_arr_val[0] = (TDMK_Code.SQL_MAX("UserInfo", "ID", myVar.sqlcon_OK2SHIP) + 1).ToString();
                        TDMK_Code.insert_val_arr("UserInfo", myVar.sqlcon_OK2SHIP, item_arr, item_arr_val);
                    }
                    else
                    {
                        if(MessageBox.Show("Mã nhân viên đã được nhập" + "\r\n" + "Bạn muốn cập nhật lại mật khẩu và tên nhân viên ?","Warning", MessageBoxButtons.YesNo)== DialogResult.Yes)
                        {
                            string tar_ID;
                            tar_ID = TDMK_Code.Get_item_val("UserInfo", "User_ID = '" + txtUserID.Text + "'", "ID", myVar.sqlcon_OK2SHIP);
                            TDMK_Code.updatebyID_val_arr("UserInfo", myVar.sqlcon_OK2SHIP, Convert.ToInt32(tar_ID), item_arr, item_arr_val);
                        }
                    }
                    TDMK_Code.fill_dataset_DGV("UserInfo", TDMK_Code.SQL_CMD("UserInfo"), DGV_User, myVar.sqlcon_OK2SHIP);

                }
            }
        }
        public bool check_empty()
        {
            bool _result = false;
            foreach (Control c in GBUserInfo.Controls)
            {
                if (c is TextBox )
                {
                    if(c.Text=="")
                    {
                        c.BackColor = Color.Red;
                        _result = true;
                    }
                }
            }
            return _result;
        }
        public void check_txt_null (TextBox src_txt)
        {
            if(src_txt.Text=="")
            {
                src_txt.BackColor = Color.Red;
            }
            else
            {
                src_txt.BackColor = Color.White;
            }
        }

        private void txtUsername_TextChanged(object sender, EventArgs e)
        {
            check_txt_null(txtUsername);
        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {
            check_txt_null(txtPassword);
        }

        private void txtUserID_TextChanged(object sender, EventArgs e)
        {
            check_txt_null(txtUserID);
        }

        private void txtReTypePass_TextChanged(object sender, EventArgs e)
        {
            check_txt_null(txtReTypePass);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if(txtUserID.Text!="")
            {
                if(Convert.ToInt32(TDMK_Code.SQL_Operator("UserInfo", "User_ID", myVar.sqlcon_OK2SHIP, "COUNT", "User_ID = '" + txtUserID.Text + "'")) > 0)
                {
                    TDMK_Code.Delelte_FilteredItem_arr("UserInfo", myVar.sqlcon_OK2SHIP, "User_ID = '" + txtUserID.Text + "'");
                    MessageBox.Show("Complete delete User ID: " + txtUserID.Text);
                    TDMK_Code.fill_dataset_DGV("UserInfo", TDMK_Code.SQL_CMD("UserInfo"), DGV_User, myVar.sqlcon_OK2SHIP);
                }
            }  
        }
        private void DGV_User_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if(DGV_User.Rows.Count>0)
            {
                int row_inx=DGV_User.SelectedRows[0].Index;
                string sel_ID = DGV_User.Rows[row_inx].Cells["User_ID"].Value.ToString();
                txtUsername.Text = DGV_User.Rows[row_inx].Cells["User_Name"].Value.ToString();
                txtUserID.Text = sel_ID;
                txtPassword.Text = TDMK_Code.Get_item_val("UserInfo", "User_ID = '" + sel_ID + "'", "User_Password", myVar.sqlcon_OK2SHIP);
            }
        }
        private void rbQA_CheckedChanged(object sender, EventArgs e)
        {
            if(rbQA.Checked)
            {
                depart = rbQA.Text;
            }
        }

        private void rbDE_CheckedChanged(object sender, EventArgs e)
        {
            if (rbDE.Checked)
            {
                depart = rbDE.Text;
            }
        }

        private void rbNPI_CheckedChanged(object sender, EventArgs e)
        {
            if (rbNPI.Checked)
            {
                depart = rbNPI.Text;
            }
        }

        private void rbPro_CheckedChanged(object sender, EventArgs e)
        {
            if (rbPro.Checked)
            {
                depart = rbPro.Text;
            }
        }

        private void rbPE_CheckedChanged(object sender, EventArgs e)
        {
            if (rbPE.Checked)
            {
                depart = rbPE.Text;
            }
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
