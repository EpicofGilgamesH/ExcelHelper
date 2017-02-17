using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    //default 默认的excel表字段
    public class DefalutModel : ModelBase
    {
        //交易编号
        public string No { get; set; }

        //成交日期
        public DateTime TradeDate { get; set; }

        /// <summary>
        /// 成交类型
        /// </summary>
        public TradeTypeEnum TradeType { get; set; }

        /// <summary>
        /// 经纪人
        /// </summary>
        public string Agent { get; set; }

        /// <summary>
        /// 经纪人号码
        /// </summary>
        public double AgentTel { get; set; }

        /// <summary>
        /// 店铺
        /// </summary>
        public string Store { get; set; }

        /// <summary>
        /// 片区
        /// </summary>
        public string Zone { get; set; }

        /// <summary>
        /// 区域
        /// </summary>
        public string Area { get; set; }

        /// <summary>
        /// 可分配业绩
        /// </summary>
        public double DistributableAchievement { get; set; }

        /// <summary>
        /// 客户进线时间
        /// </summary>
        public DateTime CustomerInDate { get; set; }

        /// <summary>
        /// 城市
        /// </summary>
        public string City { get; set; }
    }
}
