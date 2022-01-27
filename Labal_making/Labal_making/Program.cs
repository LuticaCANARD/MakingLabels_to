// See https://aka.ms/new-console-template for more information
using System.Reflection;

var url = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
Console.WriteLine("made by : Lutica (presan100@gmail.com)");
Console.Write("for more information : https://github.com/LuticaCANARD/MakingLabels_to");

Console.WriteLine(url);
List<Labal_making.projn> pio = new List<Labal_making.projn>();
pio = Labal_making.inputing.got_list(url + "\\입력.xlsx");
List<Labal_making.projn.projn_non>big_list = new List<Labal_making.projn.projn_non>();
List<Labal_making.projn.projn_non> samll_list = new List<Labal_making.projn.projn_non>();
big_list = Labal_making.Largeset.Get_Anomimuslist(pio, false);
samll_list = Labal_making.Largeset.Get_Anomimuslist(pio, true);

//big
List<Labal_making.Layout>L_big = new List<Labal_making.Layout>();
L_big = Labal_making.Largeset.Make_Layout(big_list);
Labal_making.output.outputint(L_big, url + "\\결과.xlsx", 1);

Console.WriteLine(" 완료");