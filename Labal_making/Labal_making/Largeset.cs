using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Labal_making
{
    internal partial class Largeset
    {
        public static List<projn.projn_non> Get_Anomimuslist(List<projn> alllist,bool smallmode)
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
                    for (int j = 0; j < checkpoint; j++)
                    {
                        projn.projn_non jx = new projn.projn_non(alllist[i].name, alllist[i].prof_n, alllist[i].year, alllist[i].agency);
                        output.Add(jx);
                    }
                }
            }
            return output;
        }

        public static List<Layout> Make_Layout(List<projn.projn_non> projn_s)
        {
            List<Layout> output = new List<Layout>();
            if (projn_s.Count % 7 == 0)
            {
                for (int i = 0;i < projn_s.Count; i +=7)
                {
                    Layout L_o = new Layout(projn_s[i], projn_s[i + 1], projn_s[i + 2], projn_s[i + 3], projn_s[i + 4], projn_s[i + 5], projn_s[i + 6]);
                    output.Add(L_o);
                }
            }else
            {
                int counter = projn_s.Count;
                List<projn.projn_non> ocut = new List<projn.projn_non>();
                while ( (counter !=0))
                {
                    ocut.Add(projn_s[counter-1]);
                    if(ocut.Count == 7)
                    {
                        Layout layout = new Layout(ocut[0], ocut[1], ocut[2], ocut[3], ocut[4], ocut[5], ocut[6]);
                        output.Add(layout);
                        ocut.RemoveRange(0,ocut.Count);
                    }
                    counter--;
                }
                for(int i = ocut.Count; i < 7; i++)
                {
                    projn.projn_non NON = new projn.projn_non("","",0,"");
                    ocut.Add(NON);
                }
                Layout layout2 = new Layout(ocut[0], ocut[1], ocut[2], ocut[3], ocut[4], ocut[5], ocut[6]);
                output.Add(layout2);

            }
            return output ;
        }
    }
}
