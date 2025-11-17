using Export_FPCA_OK2ship_Auto_System.Repositories;
using Export_FPCA_OK2ship_Auto_System.Services;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace Export_FPCA_OK2ship_Auto_System.Views
{
    public partial class ExportManager : Form
    {
        ExportService _service = new ExportService();
        public ExportManager()
        {

            InitializeComponent();

            setupView();

        }
        private void setupView()
        {
            this.WindowState = FormWindowState.Maximized;
            DGV_Main.DataSource = _service.setupDGV();
        }
       
        private void UpdateStatus(string category, string type, string status)
        {
            switch (type)
            {
                case "StatusDB":
                    DGV_Main.Rows[((int)_service._dicCategory[category]) - 1].Cells[type].Value = status.ToString();
                    break;
            }
        }
        private void btn_CHECKDB_Click(object sender, System.EventArgs e)
        {

        }
    }
}
