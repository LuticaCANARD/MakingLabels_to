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

        public static List<projn> got_list(string path_of_file)
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
                Console.WriteLine(ws.Cells[readingpoint, 1].Value());
                while ((ws.Cells[readingpoint, 1].Value() != ""&& readingpoint<5000))
                {
                    int kimio = 0;
                    if(ws.Cells[readingpoint, 6].Value() != null)
                    {
                        kimio = ws.Cells[readingpoint, 6].Value();
                    }
                    projn pix = new projn(
                        ws.Cells[readingpoint, 1].Value(),
                        ws.Cells[readingpoint, 2].Value(),
                        (int)ws.Cells[readingpoint, 4].Value(),
                        ws.Cells[readingpoint, 3].Value(),
                        (int)ws.Cells[readingpoint, 5].Value(),
                        (int)kimio);
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
