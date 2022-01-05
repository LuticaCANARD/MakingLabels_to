using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Labal_making
{
    internal partial class output
    {
        public void outputint (List<Layout> Layout,string path_file, bool Large)
        {
            Excel.Application Excellapp = null;
            Excel.Workbook wb = null;
            try
            {
                int page = 2;
                if (Large)
                {
                    page = 1;
                }
                Excellapp = new Excel.Application();
                wb = Excellapp.Workbooks.Open(path_file);
                Excel.Worksheet ws = wb.Worksheets.get_Item(page);
                for (int i = 0; i < Layout.Count; i++)
                {

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("에러 : " + ex.Message);
            }
            finally
            {
                Console.WriteLine("파싱완료");
                wb.Close();
            }
        }
    }
}
