using Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReadExcelBase;
using Operation;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {

            //ReadExcelBase.ReadExcelBase reb = new ReadExcelBase.ReadExcelBase(@"C:\Users\dell\Desktop\all_0.xlsx");
            //reb.ReadExcelToList(new DefalutModel());

            ////IList<DefalutModel> list = reb.List as IList<DefalutModel>;
            //string[] tittle = new string[9] { "交易编号", "城市", "区域", "片区", "店铺", "经纪人", "经纪人号码", "成交类型", "可分配业绩" };
            //WriteExcelBase.WriteExcelBase web = new WriteExcelBase.WriteExcelBase(@"C:\Users\dell\Desktop\all_1.xlsx", "生成类型1", tittle);

            //IList<ModelBase> listmb = reb.List;
            ////对集合进行一些列操作
            //OperatDefault od = new OperatDefault();

            //var mb = od.FilterByColumn(listmb);
            //mb = od.RepairByColumn(mb.ToList());
            //mb = od.OrederOperation(mb.ToList());

            //int i = web.ListToExcel(mb.ToList());


            DateTime dt1 = new DateTime(2016, 12, 6);
            DateTime dt2 = new DateTime(2017, 1, 6);
           
            int j = dt1.Month;
        }
    }
}
