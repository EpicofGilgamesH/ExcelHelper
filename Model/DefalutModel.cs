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
        /// <summary>
        /// 交易编号
        /// </summary>
        public string No { get; set; }

        /// <summary>
        /// 成交日期
        /// </summary>
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
        public string AgentTel { get; set; }

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

        //2.0新增字段
        /// <summary>
        /// 进线精准匹配
        /// </summary>
        public string IsCoilInTimeRight { get; set; }

        /// <summary>
        /// 预约时间
        /// </summary>
        public DateTime OrderTime { get; set; }

        /// <summary>
        /// 预约精准匹配
        /// </summary>
        public string IsOrderTimeRight { get; set; }

        /// <summary>
        /// Q聊时间
        /// </summary>
        public DateTime QQTalkTime { get; set; }

        /// <summary>
        /// Q聊时间精确匹配
        /// </summary>
        public string IsQQTalkTimeRight { get; set; }

        //3.0增加文本描述字段
        /// <summary>
        /// 成交方式
        /// </summary>
        public string ClosingMehtod { get; set; }

        /// <summary>
        /// 文本描述
        /// </summary>
        public string Description { get; set; }

    }
}
