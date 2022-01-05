// See https://aka.ms/new-console-template for more information
using System.Reflection;

var url = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
Console.WriteLine("Hello, World!");

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
Labal_making.output.outputint(L_big, url + "\\서식화 본.xlsx", true);
