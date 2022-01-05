using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

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
            List<projn> output = new List<projn>();
            Excel.Application Excellapp = null;
            Excel.Workbook wb = null;
            

            try
            {
                Excellapp = new Excel.Application();
                wb = Excellapp.Workbooks.Open(path_of_file);
                Excel.Worksheet ws = wb.Worksheets.get_Item(1);
               int readingpoint = 2;
                while (!(ws.Cells[readingpoint, 1] != ""&& readingpoint<5000))
                {
                    projn pix = new projn(
                        ws.Cells[readingpoint, 1],
                        ws.Cells[readingpoint, 2],
                        ws.Cells[readingpoint, 4],
                        ws.Cells[readingpoint, 3],
                        ws.Cells[readingpoint, 5],
                        ws.Cells[readingpoint, 6]);
                    output.Add(pix);
                    readingpoint++;
                }

                
            }catch (Exception ex)
            {
                Console.WriteLine("에러 : "+ex.Message);
            }
            finally
            {
                Console.WriteLine("파싱완료");
                wb.Close();
            }
            return output;

            


        }
       
    }
}
