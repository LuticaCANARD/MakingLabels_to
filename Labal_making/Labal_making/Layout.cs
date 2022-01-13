using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Labal_making
{
    internal partial class Layout
    {
        public string title1; public string title2; public string title3; public string title4; public string title5; public string title6; public string title7;
        public string profn1; public string profn2; public string profn3; public string profn4; public string profn5; public string profn6; public string profn7;
        public string title_su1; public string title_su2; public string title_su3; public string title_su4; public string title_su5; public string title_su6; public string title_su7;
        public int year1; public int year2; public int year3; public int year4; public int year5; public int year6; public int year7;
        public string agent1; public string agent2; public string agent3; public string agent4; public string agent5; public string agent6; public string agent7;

        public Layout(projn.projn_non a1, projn.projn_non a2, projn.projn_non a3, projn.projn_non a4, projn.projn_non a5, projn.projn_non a6, projn.projn_non a7)
        {
            title1 = "[ " + a1.prof_n + " - " + a1.agency + " ] " + "\n" + a1.name;
            title2 = "[ " + a2.prof_n + " - " + a2.agency + " ] " + "\n" + a2.name;
            title3 = "[ " + a3.prof_n + " - " + a3.agency + " ] " + "\n" + a3.name;
            title4 = "[ " + a4.prof_n + " - " + a4.agency + " ] " + "\n" + a4.name;
            title5 = "[ " + a5.prof_n + " - " + a5.agency + " ] " + "\n" + a5.name;
            title6 = "[ " + a6.prof_n + " - " + a6.agency + " ] " + "\n" + a6.name;
            title7 = "[ " + a7.prof_n + " - " + a7.agency + " ] " + "\n" + a7.name;

            profn1 = a1.prof_n; year1 = a1.year; title_su1 = a1.name; agent1 = a1.agency;
            profn2 = a2.prof_n; year2 = a2.year; title_su2 = a2.name; agent2 = a2.agency;
            profn3 = a3.prof_n; year3 = a3.year; title_su3 = a3.name; agent3 = a3.agency;
            profn4 = a4.prof_n; year4 = a4.year; title_su4 = a4.name; agent4 = a4.agency;
            profn5 = a5.prof_n; year5 = a5.year; title_su5 = a5.name; agent5 = a5.agency;
            profn6 = a6.prof_n; year6 = a6.year; title_su6 = a6.name; agent6 = a6.agency;
            profn7 = a7.prof_n; year7 = a7.year; title_su7 = a7.name; agent7 = a7.agency;


        }


    }
}
