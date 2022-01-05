using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
using System.Runtime.InteropServices;
using System.Drawing;

namespace Labal_making
{
    internal partial class output
    {
        public void outputint (List<Layout> Layout,string path_file, bool Large)
        {
            Excel.Application Excellapp = null;
            Excel.Workbook wb = null;
            string buseo = "한국교통대학교 산학협력단";
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
                    //측면부 제목부 초기화및 입력
                    int xp = i;
                    int sub_title1_y_n = 11 + xp; int sub_title2_y_n = 34 + xp;
                    int sub_title1_x = 2; int sub_title2_x = 4;int sub_title3_x = 6; int sub_title4_x = 8;int sub_title5_x = 10;int sub_title6_x = 12;int sub_title7_x = 14;
                    ws.Range[ws.Cells[sub_title1_y_n, sub_title1_x], ws.Cells[sub_title1_y_n, sub_title1_x]].Merge();
                    ws.Range[ws.Cells[sub_title1_y_n, sub_title2_x], ws.Cells[sub_title1_y_n, sub_title2_x]].Merge();
                    ws.Range[ws.Cells[sub_title1_y_n, sub_title3_x], ws.Cells[sub_title1_y_n, sub_title3_x]].Merge();
                    ws.Range[ws.Cells[sub_title1_y_n, sub_title4_x], ws.Cells[sub_title1_y_n, sub_title4_x]].Merge();
                    ws.Range[ws.Cells[sub_title1_y_n, sub_title5_x], ws.Cells[sub_title1_y_n, sub_title5_x]].Merge();
                    ws.Range[ws.Cells[sub_title1_y_n, sub_title6_x], ws.Cells[sub_title1_y_n, sub_title6_x]].Merge();
                    ws.Range[ws.Cells[sub_title1_y_n, sub_title7_x], ws.Cells[sub_title1_y_n, sub_title7_x]].Merge();

                    ws.Cells[sub_title1_y_n, sub_title1_x] = Layout[i].title_su1;
                    ws.Cells[sub_title1_y_n, sub_title2_x] = Layout[i].title_su2;
                    ws.Cells[sub_title1_y_n, sub_title3_x] = Layout[i].title_su3;
                    ws.Cells[sub_title1_y_n, sub_title4_x] = Layout[i].title_su4;
                    ws.Cells[sub_title1_y_n, sub_title5_x] = Layout[i].title_su5;
                    ws.Cells[sub_title1_y_n, sub_title6_x] = Layout[i].title_su6;
                    ws.Cells[sub_title1_y_n, sub_title7_x] = Layout[i].title_su7;

                    // 측면부 부서부 초기화 및 입력
                    int sanhak_a = 36 + xp; int sanhak_b = 40 + xp;
                    ws.Range[ws.Cells[sanhak_a, sub_title1_x], ws.Cells[sanhak_b, sub_title1_x]].Merge();
                    ws.Range[ws.Cells[sanhak_a, sub_title2_x], ws.Cells[sanhak_b, sub_title2_x]].Merge();
                    ws.Range[ws.Cells[sanhak_a, sub_title3_x], ws.Cells[sanhak_b, sub_title3_x]].Merge();
                    ws.Range[ws.Cells[sanhak_a, sub_title4_x], ws.Cells[sanhak_b, sub_title4_x]].Merge();
                    ws.Range[ws.Cells[sanhak_a, sub_title5_x], ws.Cells[sanhak_b, sub_title5_x]].Merge();
                    ws.Range[ws.Cells[sanhak_a, sub_title6_x], ws.Cells[sanhak_b, sub_title6_x]].Merge();
                    ws.Range[ws.Cells[sanhak_a, sub_title7_x], ws.Cells[sanhak_b, sub_title7_x]].Merge();

                    ws.Cells[sanhak_a, sub_title1_x] = buseo;
                    ws.Cells[sanhak_a, sub_title2_x] = buseo;
                    ws.Cells[sanhak_a, sub_title3_x] = buseo;
                    ws.Cells[sanhak_a, sub_title4_x] = buseo;
                    ws.Cells[sanhak_a, sub_title5_x] = buseo;
                    ws.Cells[sanhak_a, sub_title6_x] = buseo;
                    ws.Cells[sanhak_a, sub_title7_x] = buseo;


                    // 측면부 제목 상단부 초기화
                    int name_y = 3 + xp;
                    int year_y = 5 + xp;
                    int kk_y = 7 + xp;
                    int num_y = 9 + xp;
                    int vio = 30;
                    ws.Range[ws.Cells[name_y,1], ws.Cells[name_y, 4]].RowHeight = 30;
                    ws.Range[ws.Cells[year_y, 1], ws.Cells[year_y, 4]].RowHeight = 30;
                    ws.Range[ws.Cells[kk_y, 1], ws.Cells[kk_y, 4]].RowHeight = 30;
                    ws.Range[ws.Cells[num_y, 1], ws.Cells[num_y, 4]].RowHeight = 30;
                    for( int j = 0; j< 12; i++)
                    {
                        int kipa = j + 2;

                    }




                    // 정면부 제목부 및 입력
                    int t1a_y = 2 + xp;  int t1b_y = 5 + xp; int ta_x = 16 + xp; int tb_x = 21 + xp;
                    int t2a_y = 6 + xp; int t2b_y = 10 + xp;
                    int t3a_y = 11 + xp; int t3b_y = 17 + xp;
                    int t4a_y = 18 + xp; int t4b_y = 24 + xp;
                    int t5a_y = 25 + xp; int t5b_y = 30 + xp;
                    int t6a_y = 31 + xp; int t6b_y = 35 + xp;
                    int t7a_y = 36 + xp; int t7b_y = 40 + xp;
                    ws.Range[ws.Cells[t1a_y, ta_x], ws.Cells[t1b_y, tb_x]].Merge();
                    ws.Range[ws.Cells[t2a_y, ta_x], ws.Cells[t2b_y, tb_x]].Merge();
                    ws.Range[ws.Cells[t3a_y, ta_x], ws.Cells[t3b_y, tb_x]].Merge();
                    ws.Range[ws.Cells[t4a_y, ta_x], ws.Cells[t4b_y, tb_x]].Merge();
                    ws.Range[ws.Cells[t5a_y, ta_x], ws.Cells[t5b_y, tb_x]].Merge();
                    ws.Range[ws.Cells[t6a_y, ta_x], ws.Cells[t6b_y, tb_x]].Merge();
                    ws.Range[ws.Cells[t7a_y, ta_x], ws.Cells[t7b_y, tb_x]].Merge();

                    ws.Cells[t1a_y, ta_x] = Layout[i].title1;
                    ws.Cells[t2a_y, ta_x] = Layout[i].title2;
                    ws.Cells[t3a_y, ta_x] = Layout[i].title3;
                    ws.Cells[t4a_y, ta_x] = Layout[i].title4;
                    ws.Cells[t5a_y, ta_x] = Layout[i].title5;
                    ws.Cells[t6a_y, ta_x] = Layout[i].title6;
                    ws.Cells[t7a_y, ta_x] = Layout[i].title7;


                    //정면부 부서부
                    int x_a = 25; int x_b = 28; int x_c = 30; int x_d = 33;
                    int y_1a = 3 + xp; int y_1b = 5 + xp;
                    int y_2a = 7 + xp; int y_2b = 9 + xp;
                    int y_3a = 11 + xp; int y_3b = 15 + xp;
                    int y_4a = 17 + xp; int y_4b = 20 + xp;

                    ws.Range[ws.Cells[y_1a, x_a], ws.Cells[y_1b, x_b]].Merge();
                    ws.Range[ws.Cells[y_2a, x_a], ws.Cells[y_2b, x_b]].Merge();
                    ws.Range[ws.Cells[y_3a, x_a], ws.Cells[y_3b, x_b]].Merge();
                    ws.Range[ws.Cells[y_4a, x_a], ws.Cells[y_4b, x_b]].Merge();
                    ws.Range[ws.Cells[y_1a, x_c], ws.Cells[y_1b, x_d]].Merge();
                    ws.Range[ws.Cells[y_2a, x_c], ws.Cells[y_3b, x_d]].Merge();
                    ws.Range[ws.Cells[y_3a, x_c], ws.Cells[y_3b, x_d]].Merge();

                    ws.Cells[y_3a, x_c] = buseo;
                    ws.Cells[y_2a, x_c] = buseo;
                    ws.Cells[y_1a, x_c] = buseo;
                    ws.Cells[y_1a, x_a] = buseo;
                    ws.Cells[y_2a, x_a] = buseo;
                    ws.Cells[y_3a, x_a] = buseo;
                    ws.Cells[y_4a, x_a] = buseo;


                    //정면부 연도부
                    int yz1 = 23 + xp; int xpz1 = 25; int xpz2 = 26; int xppp1 = 28; int xppp2 = 29; int xppp3 = 31; int xppp4 = 32;
                    int yv1 = 25 + xp;
                    int yz2 = 26 + xp;
                    int yv2 = 28 + xp;
                    int yz3 = 30 + xp;
                    int yv3 = 32 + xp;

                    ws.Range[ws.Cells[yz1, xpz1],ws.Cells[yv1, xpz2]].Merge();
                    ws.Range[ws.Cells[yz2, xpz1], ws.Cells[yv2, xpz2]].Merge();
                    ws.Range[ws.Cells[yz3, xpz1], ws.Cells[yv3, xpz2]].Merge();
                    ws.Range[ws.Cells[yz1, xppp1], ws.Cells[yv1, xppp1]].Merge();
                    ws.Range[ws.Cells[yz2, xppp1], ws.Cells[yv2, xppp1]].Merge();
                    ws.Range[ws.Cells[yz1, xppp3], ws.Cells[yv1, xppp4]].Merge();
                    ws.Range[ws.Cells[yz2, xppp3], ws.Cells[yv2, xppp4]].Merge();


                    ws.Cells[yz1, xpz1] = Layout[i].year1;
                    ws.Cells[yz2, xpz1] = Layout[i].year2;
                    ws.Cells[yz3, xpz1] = Layout[i].year3;
                    ws.Cells[yz1, xppp1] = Layout[i].year4;
                    ws.Cells[yz2, xppp1] = Layout[i].year5;
                    ws.Cells[yz1, xppp3] = Layout[i].year6;
                    ws.Cells[yz2, xppp3] = Layout[i].year7;




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
