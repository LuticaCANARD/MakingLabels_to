using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Labal_making
{
    internal partial class projn
    {
        public string name; public string prof_n; public int year; public string agency; public int many_Lazy; public int many_small;
        public projn(string _name, string _prof_n, int _year, string _agency, int _many_Lazy, int _many_small)
        {
            name = _name;
            prof_n = _prof_n;
            year = _year;
            agency = _agency;
            many_Lazy = _many_Lazy;
            many_small = _many_small;
        }
        public partial class projn_non
        {
            public string name; public string prof_n; int year; public string agency;

            public projn_non(string _name, string _prof_n, int _year, string _agency)
            {
                name = _name; prof_n = _prof_n; year = _year; agency = _agency;
            }
        } 
            
            
    }

    }

