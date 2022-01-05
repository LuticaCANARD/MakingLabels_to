using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Labal_making
{
    internal class Largeset
    {
        public List<projn.projn_non> Get_Anomimuslist(List<projn> alllist,bool smallmode)
        {
            List < projn.projn_non > output = new List<projn.projn_non>();
            for (int i = 0; i < alllist.Count; i++)
            {
                var checkpoint = alllist[i].many_Lazy;
                if (smallmode==true)
                {
                    checkpoint = alllist[i].many_small;
                }
                if (checkpoint == 0)
                {
                    continue;
                }
                else
                {
                    projn.projn_non jx = new projn.projn_non(
                        alllist[i].name, alllist[i].prof_n, alllist[i].year, alllist[i].agency
                        );
                    output.Add(jx);
                }
            }
            return output;
        }
    }
}
