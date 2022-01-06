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
        public static void outputint (List<Layout> Layout,string path_file, bool Large)
        {
            Excel.Application Excellapp = null;
            Excel.Workbook wb = null;
            string buseo = "한국교통대학교 산학협력단";
            try
            {
                int page = 2;
                if (Large==true)
                {
                    page = 1;
                }
                
                Excellapp = new Excel.Application();
                wb = Excellapp.Workbooks.Open(path_file);
                Excel.Worksheet ws = wb.Worksheets.get_Item(page);
                for (int i = 0; i < Layout.Count; i++)
                {
                    int xp = 40*i;

                    //측면부 제목부 초기화및 입력
                    
                    int sub_title1_y_n = 11 + xp; int sub_title2_y_n = 34 + xp;
                    int sub_title1_x = 2; int sub_title2_x = 4;int sub_title3_x = 6; int sub_title4_x = 8;int sub_title5_x = 10;int sub_title6_x = 12;int sub_title7_x = 14;

                    
                    
                    
                    ws.Range[ws.Cells[11+xp, 2], ws.Cells[34+xp, 2]].Merge();
                    ws.Range[ws.Cells[11+xp, 4], ws.Cells[34+xp, 4]].Merge();
                    ws.Range[ws.Cells[11+xp, 6], ws.Cells[34+xp, 6]].Merge();
                    ws.Range[ws.Cells[11 + xp, 8], ws.Cells[34 + xp, 8]].Merge();
                    ws.Range[ws.Cells[11 + xp, 10], ws.Cells[34 + xp, 10]].Merge();
                    ws.Range[ws.Cells[11 + xp, 12], ws.Cells[34 + xp, 12]].Merge();
                    ws.Range[ws.Cells[11 + xp, 14], ws.Cells[34 + xp, 14]].Merge();

                    string ik1 = Layout[i].title_su1;

                    ws.Cells[sub_title1_y_n, sub_title1_x] = Layout[i].title_su1;
                    ws.Cells[sub_title1_y_n, sub_title2_x] = Layout[i].title_su2;
                    ws.Cells[sub_title1_y_n, sub_title3_x] = Layout[i].title_su3;
                    ws.Cells[sub_title1_y_n, sub_title4_x] = Layout[i].title_su4;
                    ws.Cells[sub_title1_y_n, sub_title5_x] = Layout[i].title_su5;
                    ws.Cells[sub_title1_y_n, sub_title6_x] = Layout[i].title_su6;
                    ws.Cells[sub_title1_y_n, sub_title7_x] = Layout[i].title_su7;
                    Console.WriteLine("제목부 머지 완료");



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
                    Console.WriteLine("부서부 머지 완료");

                    // 측면부 제목 상단부 초기화
                    int name_y = 3 + xp;
                    int year_y = 5 + xp;
                    int kk_y = 7 + xp;
                    int num_y = 9 + xp;
                    int vio = 30;
                    int[] kise = { Layout[i].year1, Layout[i].year2, Layout[i].year3, Layout[i].year4, Layout[i].year5, Layout[i].year6, Layout[i].year7 };
                    int lk = 0 ;
                    for (int j = 0; j < 13; j += 2)
                    {
                        Console.WriteLine("관리 장부 진입");
                        int kipa = j + 2;
                        ws.Cells[name_y - 1, kipa] = "관리번호";
                        ws.Cells[year_y - 1, kipa] = "생산연도";
                        ws.Cells[kk_y - 1, kipa] = "보존기한";
                        ws.Cells[year_y, kipa] = kise[lk];
                        lk++;
                        ws.Cells[num_y - 1, kipa] = "분류번호";
                        ws.Cells[10 + xp, kipa] = "제목";
                        ws.Cells[35 + xp, kipa] = "부서명";
                    }


                    ws.Range[ws.Cells[name_y,1], ws.Cells[name_y, 4]].RowHeight = 30;
                    ws.Range[ws.Cells[year_y, 1], ws.Cells[year_y, 4]].RowHeight = 30;
                    ws.Range[ws.Cells[kk_y, 1], ws.Cells[kk_y, 4]].RowHeight = 30;
                    ws.Range[ws.Cells[num_y, 1], ws.Cells[num_y, 4]].RowHeight = 30;
                    Console.WriteLine("측면부 설정 완료");

                    // 이름부 입력
                    ws.Cells[3 + xp, 2] = Layout[i].profn1;
                    ws.Cells[3 + xp, 4] = Layout[i].profn2;
                    ws.Cells[3 + xp, 6] = Layout[i].profn3;
                    ws.Cells[3 + xp, 8] = Layout[i].profn4;
                    ws.Cells[3 + xp, 10] = Layout[i].profn5;
                    ws.Cells[3 + xp, 12] = Layout[i].profn6;
                    ws.Cells[3 + xp, 14] = Layout[i].profn7;


                    // 정면부 제목부 및 입력
                    int t1a_y = 2 + xp;  int t1b_y = 5 + xp; int ta_x = 16 ; int tb_x = 21 ;
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
                    Console.WriteLine("정면부 머지 완료");


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
                    ws.Range[ws.Cells[y_2a, x_c], ws.Cells[y_2b, x_d]].Merge();
                    ws.Range[ws.Cells[y_3a, x_c], ws.Cells[y_3b, x_d]].Merge();

                    ws.Cells[y_3a, x_c] = buseo;
                    ws.Cells[y_2a, x_c] = buseo;
                    ws.Cells[y_1a, x_c] = buseo;
                    ws.Cells[y_1a, x_a] = buseo;
                    ws.Cells[y_2a, x_a] = buseo;
                    ws.Cells[y_3a, x_a] = buseo;
                    ws.Cells[y_4a, x_a] = buseo;
                    Console.WriteLine("정면부 부서 완료");


                    //정면부 연도부
                    int yz1 = 22 + xp; int xpz1 = 25; int xpz2 = 26; int xppp1 = 28; int xppp2 = 29; int xppp3 = 31; int xppp4 = 32;
                    int yv1 = 24 + xp;
                    int yz2 = 26 + xp;
                    int yv2 = 28 + xp;
                    int yz3 = 30 + xp;
                    int yv3 = 32 + xp;

                    ws.Range[ws.Cells[yz1, xpz1],ws.Cells[yv1, xpz2]].Merge();
                    ws.Range[ws.Cells[yz2, xpz1], ws.Cells[yv2, xpz2]].Merge();
                    ws.Range[ws.Cells[yz3, xpz1], ws.Cells[yv3, xpz2]].Merge();
                    ws.Range[ws.Cells[yz1, xppp1], ws.Cells[yv1, xppp2]].Merge();
                    ws.Range[ws.Cells[yz2, xppp1], ws.Cells[yv2, xppp2]].Merge();
                    ws.Range[ws.Cells[yz1, xppp3], ws.Cells[yv1, xppp4]].Merge();
                    ws.Range[ws.Cells[yz2, xppp3], ws.Cells[yv2, xppp4]].Merge();


                    ws.Cells[yz1, xpz1] = Layout[i].year1;
                    ws.Cells[yz2, xpz1] = Layout[i].year2;
                    ws.Cells[yz3, xpz1] = Layout[i].year3;
                    ws.Cells[yz1, xppp1] = Layout[i].year4;
                    ws.Cells[yz2, xppp1] = Layout[i].year5;
                    ws.Cells[yz1, xppp3] = Layout[i].year6;
                    ws.Cells[yz2, xppp3] = Layout[i].year7;
                    Console.WriteLine("정면부 연도 완료");

                    Console.WriteLine("한바퀴 순회 완료");

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("에러 : " + ex.Message);
            }
            finally
            {
                Console.WriteLine("파일작업 완료");
                if (Large == false) { wb.Close(); }
              
            }
        }
    }
}
