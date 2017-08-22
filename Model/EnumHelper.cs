using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    public static class EnumHelper
    {
        /// <summary>
        /// 获取枚举的描述
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="t"></param>
        /// <returns></returns>
        public static string GetEnumDescription(Enum e)
        {
            //获取类型
            Type type = e.GetType();
            //获取成员
            MemberInfo[] mis = type.GetMember(e.ToString());
            if (mis != null && mis.Length > 0)
            {
                DescriptionAttribute[] descriptions = mis[0].GetCustomAttributes(typeof(DescriptionAttribute), false) as DescriptionAttribute[];
                if (descriptions != null && descriptions.Length > 0)
                {
                    return descriptions[0].Description;
                }
            }
            return e.ToString();
        }
    }
}
