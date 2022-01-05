using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Labal_making
{
    internal partial class Layout
    {
        protected string title1; protected string title2; protected string title3; protected string title4; protected string title5; protected string title6; protected string title7;
        protected string profn1; protected string profn2; protected string profn3; protected string profn4; protected string profn5; protected string profn6; protected string profn7;
        protected string title_su1; protected string title_su2; protected string title_su3; protected string title_su4; protected string title_su5; protected string title_su6; protected string title_su7;
        protected int year1; protected int year2; protected int year3; protected int year4; protected int year5; protected int year6; protected int year7;

        public Layout(projn.projn_non a1, projn.projn_non a2, projn.projn_non a3, projn.projn_non a4, projn.projn_non a5, projn.projn_non a6, projn.projn_non a7)
        {
            title1 = "[ " + a1.prof_n + " - " + a1.agency + " ] " + "\n" + a1.name;
            title2 = "[ " + a2.prof_n + " - " + a2.agency + " ] " + "\n" + a2.name;
            title3 = "[ " + a3.prof_n + " - " + a3.agency + " ] " + "\n" + a3.name;
            title4 = "[ " + a4.prof_n + " - " + a4.agency + " ] " + "\n" + a4.name;
            title5 = "[ " + a5.prof_n + " - " + a5.agency + " ] " + "\n" + a5.name;
            title6 = "[ " + a6.prof_n + " - " + a6.agency + " ] " + "\n" + a6.name;
            title7 = "[ " + a7.prof_n + " - " + a7.agency + " ] " + "\n" + a7.name;

            profn1 = a1.prof_n; year1 = a1.year; title_su1 = a1.name;
            profn2 = a2.prof_n; year2 = a2.year; title_su2 = a2.name;
            profn3 = a3.prof_n; year3 = a3.year; title_su3 = a3.name;
            profn4 = a4.prof_n; year4 = a4.year; title_su4 = a4.name;
            profn5 = a5.prof_n; year5 = a5.year; title_su5 = a5.name;
            profn6 = a6.prof_n; year6 = a6.year; title_su6 = a6.name;
            profn7 = a7.prof_n; year7 = a7.year; title_su7 = a7.name;


        }


    }
}
