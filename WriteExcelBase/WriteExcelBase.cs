using Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Operation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteExcelBase
{
    public class WriteExcelBase
    {
        public WriteExcelBase(string filePath, string[] tittle, string[] coltittle)
        {
            FilePath = filePath;
            Tittle = tittle;
            ColTittle = coltittle;
        }

        public string[] Tittle { get; set; }

        public string[] ColTittle { get; set; }
        public string FilePath { get; set; }

        //private static ICellStyle style { get; set; }

        private IWorkbook workbook = null;

        public int ListToExcel(IList<ModelBase> list)
        {
            Dictionary<string, List<ModelBase>> GroupList = new Dictionary<string, List<ModelBase>>();
            //1.按城市分类，有哪些城市 2.将城市名做为sheet，循环写入Excel
            IEnumerable<IGrouping<string, ModelBase>> grouplist = GroupByCity(list).ToList();
            List<ModelBase> li = new List<ModelBase>();
            foreach (var group in grouplist)
            {
                li = OperatDefault.GetGroupList(list, group.Key).ToList();
                GroupList.Add(group.Key, li);
            }

            int count = 0;
            ISheet sheet = null;
            using (FileStream fs = new FileStream(FilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                if (FilePath.EndsWith(".xlsx"))
                    workbook = new XSSFWorkbook();
                else if (FilePath.EndsWith(".xls"))
                    workbook = new HSSFWorkbook();
                try
                {
                    //导出各城市数据
                    foreach (var dic in GroupList)
                    {
                        if (workbook != null)
                        {
                            //style = workbook.CreateCellStyle();
                            //style.DataFormat = HSSFDataFormat.GetBuiltinFormat("yyyy-MM-dd");
                            sheet = workbook.CreateSheet(dic.Key);
                        }
                        else
                        {
                            return -1;
                        }

                        if (dic.Value != null)
                        {
                            //开始将list写入Excel中 1.写标题行 2.写正文
                            IRow tittleRow = sheet.CreateRow(0);

                            //标题
                            for (int i = 0; i < Tittle.Length; i++)
                            {
                                tittleRow.CreateCell(i).SetCellValue(Tittle[i]);
                            }
                            //正文
                            for (int i = 0; i < dic.Value.Count; i++)
                            {
                                DefalutModelExport(dic.Value[i], i + 1, sheet);
                                count++;
                            }
                        }
                    }
                    //导出汇总数据
                    //Colloge model = GetCollegeList(list, key);

                    sheet = workbook.CreateSheet("成交数据汇总");
                    IRow ColtittleRow = sheet.CreateRow(0);
                    for (int i = 0; i < ColTittle.Length; i++)
                    {
                        ColtittleRow.CreateCell(i).SetCellValue(ColTittle[i]);
                    }

                    GroupList.Add("总计", null);
                    int a = 1;
                    foreach (var key in GroupList.Keys)
                    {
                        Colloge model = GetCollegeList(list, key);
                        IRow row = sheet.CreateRow(a);
                        row.CreateCell(0).SetCellValue(model.Key);
                        row.CreateCell(1).SetCellValue(model.SaleCount);
                        row.CreateCell(2).SetCellValue(model.SaleSum);
                        row.CreateCell(3).SetCellValue(model.RentCount);
                        row.CreateCell(4).SetCellValue(model.RentSum);
                        row.CreateCell(5).SetCellValue(model.SumCount);
                        row.CreateCell(6).SetCellValue(model.SumAchievemen);
                        a++;
                    }

                    workbook.Write(fs);
                    return count;
                }
                catch (Exception ex)
                {
                    return -1;
                    throw new Exception("Exception: " + ex.Message);
                }
            }
        }



        public static IRow DefalutModelExport(ModelBase mb, int i, ISheet sheet)
        {
            DefalutModel dm = mb as DefalutModel;
            IRow row = sheet.CreateRow(i);
            row.CreateCell(0).SetCellValue(dm.City);
            row.CreateCell(1).SetCellValue(dm.Area);
            row.CreateCell(2).SetCellValue(dm.Zone);
            row.CreateCell(3).SetCellValue(dm.Store);
            row.CreateCell(4).SetCellValue(dm.Agent);
            row.CreateCell(5).SetCellValue(dm.AgentTel);
            row.CreateCell(6).SetCellValue(TradeTypeToString(dm.TradeType));
            string achieve = GetAchievement(dm.DistributableAchievement);
            row.CreateCell(7).SetCellValue(achieve);
            ICell dateCell = row.CreateCell(8);
            dateCell.SetCellValue(dm.TradeDate.ToString("yyyy-MM-dd"));
            //dateCell.SetCellValue(dm.TradeDate.ToString("yyyy-MM-dd")); 时间格式不便于筛选
            //dateCell.CellStyle = style;
            row.CreateCell(9).SetCellValue(dm.No);
            row.CreateCell(10).SetCellValue(dm.IsCoilInTimeRight);
            string description = string.Empty;
            row.CreateCell(11).SetCellValue(CustomerSource(dm, achieve, out description));
            row.CreateCell(12).SetCellValue(description);
            return row;
        }

        /// <summary>
        /// 获得 客户来源 字段
        /// </summary>
        /// <param name="dm"></param>
        /// <returns></returns>
        public static string CustomerSource(DefalutModel dm, string achieve, out string description)
        {
            string customerSource = string.Empty;
            if (dm.IsCoilInTimeRight.ToUpper() == "YES")
            {
                customerSource = "进线";
            }
            if (dm.IsOrderTimeRight.ToUpper() == "YES")
            {
                if (string.IsNullOrEmpty(customerSource))
                {
                    customerSource = "预约";
                }
            }
            if (dm.IsQQTalkTimeRight.ToUpper() == "YES")
            {
                if (string.IsNullOrEmpty(customerSource))
                {
                    customerSource = "Q聊";
                }
            }
            if (!string.IsNullOrEmpty(customerSource))
            {
                description = dm.Store + dm.Agent + achieve;
            }
            else
            {
                description = string.Empty;
            }
            return customerSource;
        }

        /// <summary>
        /// 获取描述信息和成交方式
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static string GetAchievement(double d)
        {
            string achievement = string.Empty;
            if (d >= 10000)
            {
                achievement = Math.Round(d / 10000, 2).ToString() + "万元";
            }
            else if (d > 5000 && d < 10000)
            {
                achievement = "近万元";
            }
            else if (d == 5000)
            {
                achievement = "半万元";
            }
            else if (d < 5000)
            {
                achievement = "NNN万元";
            }
            return achievement;
        }


        public static string TradeTypeToString(TradeTypeEnum tte)
        {
            string tradeType = string.Empty;
            if (tte == TradeTypeEnum.RENT)
            {
                tradeType = "租";
            }
            else if (tte == TradeTypeEnum.SALE)
            {
                tradeType = "售";
            }
            else
            {
                throw new Exception("成交类型出错~");
            }
            return tradeType;
        }

        /// <summary>
        /// 1.按城市进行分组，并获取城市的数组集合
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static IEnumerable<IGrouping<string, ModelBase>> GroupByCity(IList<ModelBase> list)
        {
            IList<DefalutModel> listdm = list.Select(x => (DefalutModel)x).ToList();
            return listdm.GroupBy(x => x.City);
        }


        /// <summary>
        ///获取汇总数据源 
        /// </summary>
        /// <param name="list"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static Colloge GetCollegeList(IList<ModelBase> list, string key)
        {
            IList<DefalutModel> listdm = list.Select(x => (DefalutModel)x).ToList();
            IEnumerable<DefalutModel> saleList = listdm.Where(x => x.City == key && x.TradeType == TradeTypeEnum.SALE); //城市售单集合
            int saleCount = saleList.Count(); //售单
            double saleSum = saleList.Sum(x => x.DistributableAchievement); //售金额

            IEnumerable<DefalutModel> rentList = listdm.Where(x => x.City == key && x.TradeType == TradeTypeEnum.RENT);//城市租单集合
            int rentCount = rentList.Count(); //租单
            double rentSum = rentList.Sum(x => x.DistributableAchievement); //租金额

            int sumCount = saleCount + rentCount;  //总单数 

            double sumAchievement = saleSum + rentSum; //总金额

            if (key == "总计")
            {
                return new Colloge
                {
                    Key = key,
                    SaleCount = listdm.Count(x => x.TradeType == TradeTypeEnum.SALE),
                    SaleSum = listdm.Where(x => x.TradeType == TradeTypeEnum.SALE).Sum(x => x.DistributableAchievement),
                    RentCount = listdm.Count(x => x.TradeType == TradeTypeEnum.RENT),
                    RentSum = listdm.Where(x => x.TradeType == TradeTypeEnum.RENT).Sum(x => x.DistributableAchievement),
                    SumCount = listdm.Count(),
                    SumAchievemen = listdm.Sum(x => x.DistributableAchievement),
                };
            }

            return new Colloge
            {
                Key = key,
                SaleCount = saleCount,
                SaleSum = saleSum,
                RentCount = rentCount,
                RentSum = rentSum,
                SumCount = sumCount,
                SumAchievemen = sumAchievement
            };
        }
    }
}
