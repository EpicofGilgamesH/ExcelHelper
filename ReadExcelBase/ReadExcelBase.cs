using Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelBase
{
    public class ReadExcelBase
    {
        public ReadExcelBase(string filePath)
        {
            _filePath = filePath;
        }
        private string _filePath = string.Empty;
        private IWorkbook _workbook = null;
        public string[] Tittle { get; set; }
        public string[] SheetName { get; set; }
        public IList<ModelBase> List { get; set; }

        public IList<ModelBase> ReadExcelToList(ModelBase t)
        {
            ISheet sheet = null;
            IList<ModelBase> list = new List<ModelBase>();
            using (FileStream fs = new FileStream(_filePath, FileMode.OpenOrCreate))
            {
                if (_filePath.EndsWith(".xlsx"))//2007以上版本
                {
                    _workbook = new XSSFWorkbook(fs);
                }
                else if (_filePath.EndsWith(".xls")) //2003版本
                {
                    _workbook = new HSSFWorkbook(fs);
                }

                if (_workbook != null)
                {
                    //获取excel的第一个sheet
                    sheet = _workbook.GetSheetAt(0);
                    //总行数
                    int rowCount = sheet.LastRowNum;
                    IRow headRow = sheet.GetRow(0);
                    //总列数
                    int cellCounnt = headRow.Count();
                    //标题得单独读取出来
                    string[] tittle = new string[cellCounnt];
                    IRow tittleRow = sheet.GetRow(0);
                    for (int i = 0; i < cellCounnt; i++)
                    {
                        tittle[i] = tittleRow.GetCell(i).StringCellValue;
                    }
                    Tittle = tittle;

                    for (int i = 1; i < rowCount + 1; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        //将一行中的每个元素读取出来  
                        ModelBase mb;
                        mb = GetType(t, row);
                        //将改行的所有元素填充到List集合中
                        list.Add(mb);
                    }
                }
                List = list;

                return list;
            }
        }

        /// <summary>
        /// 判断导入的Excel模板对应的Model 并装载Model
        /// </summary>
        /// <param name="t"></param>
        /// <param name="errorMsg"></param>
        /// <returns></returns>
        public static ModelBase GetType(ModelBase t, IRow row)
        {
            ModelBase mb = null;
            switch (t.GetType().Name)
            {
                case "DefalutModel":
                    mb = FillDefalutModel(row);
                    break;
                default:
                    throw new Exception("该Excel模板还不被支持~");
            }
            return mb;
        }


        public static DefalutModel FillDefalutModel(IRow row)
        {
            DefalutModel dm = new DefalutModel();
            dm.No = row.GetCell(0).StringCellValue;
            dm.TradeDate = row.GetCell(1).DateCellValue;
            dm.TradeType = IsRentOrSale(row.GetCell(2).StringCellValue);
            dm.Agent = row.GetCell(3).StringCellValue;
            dm.AgentTel = GetDistributableAchievement(row.GetCell(4));
            dm.Store = row.GetCell(5).StringCellValue;
            dm.Zone = row.GetCell(6).StringCellValue;
            dm.Area = row.GetCell(7).StringCellValue;
            dm.DistributableAchievement = GetDistributableAchievement(row.GetCell(8));
            dm.CustomerInDate = row.GetCell(9).DateCellValue;
            dm.City = row.GetCell(10).StringCellValue;
            return dm;
        }

        /// <summary>
        /// 可分配业绩兼容文本类型和值类型的方法   [这个try catch 不可取，花的时间太长了]
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static double GetDistributableAchievement(ICell cell)
        {

            if (cell.CellType == CellType.Numeric)  //数字类型
                return cell.NumericCellValue;
            else if (cell.CellType == CellType.String) //文本类型
                return Convert.ToDouble(cell.StringCellValue);
            else
            {
                throw new Exception("[可分配业绩]类型出错，请检查~");
            }
        }


        /// <summary>
        /// 判断成交类型为租还是售
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static TradeTypeEnum IsRentOrSale(string str)
        {
            TradeTypeEnum tte;
            string s = str.ToUpper();
            //如果包含租且不包含售
            if (s.Contains("RENT") && !s.Contains("SALE"))
            {
                tte = TradeTypeEnum.RENT;
            }
            else if (s.Contains("SALE") && !s.Contains("RENT"))
            {
                tte = TradeTypeEnum.SALE;
            }
            else
            {
                throw new Exception("成交类型出错，请检查~");
            }
            return tte;
        }
    }
}
