using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Funtion_F3_SMT
{
    public class Funtion_SMT
    {
        public struct Peeltest_data
        {
            public string data_val { get; set; }
            public byte[] grap_data { get; set; }

            public byte[] image_data { get; set; }
            public Peeltest_data(string in_data, byte[] in_graph, byte[] in_image)
            {
                data_val = in_data;
                grap_data = in_graph;
                image_data = in_image;
            }
        }
    }

   



}
