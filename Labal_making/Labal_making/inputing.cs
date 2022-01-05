using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Labal_making
{
    internal partial class inputing
    {
        public string Got_now_url()
        {
            
            string path = Path.GetDirectoryName(typeof(inputing).Assembly.Location).ToString();
            return path;
        }

        public List<projn> got_list(string path_of_file)
        {

        }
       
    }
}
