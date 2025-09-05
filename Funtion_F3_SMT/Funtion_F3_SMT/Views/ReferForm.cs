using OK2SHIP_SMT.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OK2SHIP_SMT.Views
{
    public partial class ReferForm : Form
    {
        private List<KeyValuePair<string, string>> _referList = new List<KeyValuePair<string, string>>();
        private List<KeyValuePair<string, string>> _listRefered = new List<KeyValuePair<string, string>>();
        private AirBubbleService _airBubbleService = new AirBubbleService();
        public ReferForm(AirBubbleService airBubbleService)
        {
            _airBubbleService = airBubbleService;
            InitializeComponent();
            string itemCode = _airBubbleService._itemCode;
            string lotNo = _airBubbleService._lotNo;
            lb_namerefer.Text = $"{itemCode} - {lotNo}";
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {

            int i = listBox2.SelectedIndex;
          
            _airBubbleService.RemoveRefer(_listRefered[i].Key.Trim(), _listRefered[i].Value.Trim());
            fillForm();
        }

        private void btn_refer_Click(object sender, EventArgs e)
        {
            int i = listBox1.SelectedIndex;
            _airBubbleService.AddRefer(_referList[i]);

            fillForm();
        }
        private void ReferForm_Load(object sender, EventArgs e)
        {
            fillForm();
        }
        private void fillForm()
        {
            try
            {
                _referList = _airBubbleService.LoadRefer();
                _listRefered = _airBubbleService.LoadRefered();
                FillData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void FillData()
        {
            listBox1.Items.Clear();
            foreach (KeyValuePair<string, string> item in _referList)
            {
                listBox1.Items.Add($"{item.Key} - {item.Value}");
            }
            listBox2.Items.Clear();
            foreach (KeyValuePair<string, string> item in _listRefered)
            {
                listBox2.Items.Add($"{item.Key} - {item.Value}");
            }
        }
    }
}
