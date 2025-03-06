using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SEEV_SMT_F3.TDMK_Tool
{
    public class TableDiv : TableLayoutPanel
    {
        public IEnumerable<string> Column_Ratio { get; set; }
        public IEnumerable<string> Row_Ratio { get; set; }
        public TableDiv()
        {
            this.Resize += TableDiv_Resize;
        }

        private void TableDiv_Resize(object? sender, EventArgs e)
        {
            if (this.Height != 0 && Column_Ratio.Count() > 0)
            {
               foreach(var item in Column_Ratio)
                {
                }
            }
        }
    }
}

