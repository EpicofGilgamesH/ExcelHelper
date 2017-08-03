using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// 输出模型
    /// </summary>
    public class DefaultModelTransfor : ModelBase
    {
        //交易编号
        public string No { get; set; }

        /// <summary>
        /// 城市
        /// </summary>
        public string City { get; set; }

        /// <summary>
        /// 区域
        /// </summary>
        public string Area { get; set; }

        /// <summary>
        /// 片区
        /// </summary>
        public string Zone { get; set; }

        /// <summary>
        /// 店铺
        /// </summary>
        public string Store { get; set; }

        //成交日期
        public DateTime TradeDate { get; set; }

        /// <summary>
        /// 经纪人
        /// </summary>
        public string Agent { get; set; }

        /// <summary>
        /// 经纪人号码
        /// </summary>
        public string AgentTel { get; set; }

        /// <summary>
        /// 成交类型
        /// </summary>
        public TradeTypeEnum TradeType { get; set; }

        /// <summary>
        /// 可分配业绩
        /// </summary>
        public double DistributableAchievement { get; set; }
    }
}
