using Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Operation
{
    /// <summary>
    /// 默认操作类   考虑代码量>性能 的因素，暂不屏蔽频繁的拆装箱操作
    /// </summary>
    public class OperatDefault : OrderByBase, IFilterOperation
    {
        /// <summary>
        /// 去除多余字符
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public IEnumerable<ModelBase> RepairByColumn(IList<ModelBase> list)
        {
            //将List<ModelBase>转换为List<DefalutModel>
            IList<DefalutModel> listdm = list.Select(x => (DefalutModel)x).ToList();
            return listdm.Select(RepairCity);
        }

        /// <summary>
        /// 根据条件筛选
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public IEnumerable<ModelBase> FilterByColumn(IList<ModelBase> list)
        {
            //将List<ModelBase>转换为List<DefalutModel>
            IList<DefalutModel> fbclist = new List<DefalutModel>();
            IList<DefalutModel> listdm = list.Select(x => (DefalutModel)x).ToList();
            foreach (DefalutModel dm in listdm)
            {
                if (dm.TradeDate.Date > dm.CustomerInDate.Date) //进线>成交的去掉
                    fbclist.Add(dm);
                else if (dm.TradeDate.Date == dm.CustomerInDate.Date && dm.TradeType == TradeTypeEnum.RENT) //进线=成交 只保留租
                    fbclist.Add(dm);
            }
            return fbclist;
        }

        /// <summary>
        /// 根据月份 进行筛选
        /// </summary>
        /// <param name="list"></param>
        /// <param name="isThisMonth"></param>
        /// <returns></returns>
        public IEnumerable<ModelBase> FilterByMonth(List<ModelBase> list, bool isThisMonth = true)
        {
            IList<DefalutModel> listdm = list.Select(x => (DefalutModel)x).ToList();
            IEnumerable<DefalutModel> li;
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            if (isThisMonth)
            {
                int y = listdm[0].TradeDate.Year;
                int m = listdm[0].TradeDate.Month;
                li = listdm.Where(x => x.TradeDate.Year == year && x.TradeDate.Month == month);
            }
            else
            {
                li = month == 1 ? listdm.Where(x => x.TradeDate.Year == year - 1 && x.TradeDate.Month == 12) :
                    listdm.Where(x => x.TradeDate.Year == DateTime.Now.Year && x.TradeDate.Month == DateTime.Now.Month - 1);
            }
            return li;
        }

        /// <summary>
        /// 排序
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public override IEnumerable<ModelBase> OrederOperation(IList<ModelBase> list)
        {
            //将List<ModelBase>转换为List<DefalutModel>
            IList<DefalutModel> listdm = list.Select(x => (DefalutModel)x).ToList();
            listdm = Dictinct(listdm).ToList();
            return listdm = listdm.OrderBy(x => x.City).ThenBy(x => x.Area).ThenBy(x => x.Zone).ToList();
        }

        /// <summary>
        /// 除去城市字段中的多余字符  
        /// </summary>
        /// <param name="str"></param>
        public static DefalutModel RepairCity(DefalutModel dm)
        {
            dm.City = dm.City.Replace("世华", "").Replace("云房源", "").Replace("云房", "").Replace("大", "").Replace("区域", "");
            return dm;
        }

        /// <summary>
        /// 根据城市 返回数据
        /// </summary>
        /// <param name="list"></param>
        /// <param name="city"></param>
        /// <returns></returns>
        public static IEnumerable<ModelBase> GetGroupList(IList<ModelBase> list, string city)
        {
            IList<DefalutModel> listdm = list.Select(x => (DefalutModel)x).ToList();
            var li = listdm.Where(x => x.City == city).ToList();
            return li;
        }

        /// <summary>
        /// List<集合>去重操作   弃用
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        //public static IEnumerable<DefalutModel> Dictinct(IList<DefalutModel> list)
        //{
        //    for (int i = 0; i < list.Count; i++)
        //    {
        //        for (int j = i + 1; j < list.Count; j++)  //如果内层循环是从最后一位 循环  则不会出现序列错误的情况
        //        {
        //            if (list[i].No == list[j].No && list[i].Agent == list[j].Agent)
        //            {
        //                list.RemoveAt(j);
        //                j--;
        //            }
        //        }
        //    }
        //    return list;
        //}

        public static IEnumerable<DefalutModel> Dictinct(IList<DefalutModel> list)
        {
            for (int i = 0; i < list.Count; i++)
            {
                for (int j = list.Count - 1; j > i; j--)  //内层循环使用从后到前的方式，不会出现序列错误的问题
                {
                    if (list[i].No == list[j].No && list[i].Agent == list[j].Agent)
                    {
                        list.RemoveAt(j);
                    }
                }
            }
            return list;
        }
    }

    public class DefalutModelComparer : IEqualityComparer<DefalutModel>
    {
        public bool Equals(DefalutModel x, DefalutModel y)
        {
            if (x == null)
                return y == null;
            return (x.No == y.No && x.Store == y.Store && x.TradeDate == y.TradeDate && x.TradeType
                == y.TradeType && x.Zone == y.Zone && x.Agent == y.Agent && x.AgentTel ==
                y.AgentTel && x.Area == y.Area && x.City == y.City && x.CustomerInDate ==
                y.CustomerInDate && x.DistributableAchievement == y.DistributableAchievement);
        }

        public int GetHashCode(DefalutModel obj)
        {
            if (obj == null)
                return 0;
            return obj.GetHashCode();
        }
    }
}
