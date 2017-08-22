using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    public enum TittleEnum
    {
        [Description("交易编号")]
        No = 0,
        [Description("成交日期")]
        TradeDate = 1,
        [Description("成交类型")]
        TradeType = 2,
        [Description("经纪人")]
        Agent = 3,
        [Description("经纪人号码")]
        AgentTel = 4,
        [Description("店铺")]
        Store = 5,
        [Description("片区")]
        Zone = 6,
        [Description("区域")]
        Area = 7,
        [Description("可分配业绩")]
        DistributableAchievement = 8,
        [Description("进线时间")]
        CustomerInDate = 9,
        [Description("城市")]
        City = 10,
        [Description("进线精准匹配")]
        IsCoilInTimeRight = 11,
        [Description("预约时间")]
        OrderTime = 12,
        [Description("预约精准匹配")]
        IsOrderTimeRight = 13,
        [Description("Q聊时间")]
        QQTalkTime = 14,
        [Description("Q聊精准匹配")]
        IsQQTalkTimeRight = 15
    }
}
