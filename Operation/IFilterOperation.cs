using Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Operation
{
    //操作类：
    //1.筛选 2.条件（对比后过滤）3.排序（排列与统计[字段]）

    /// <summary>
    /// 筛选类  根据什么字段进行筛选
    /// </summary>
    public interface IFilterOperation
    {
        /// <summary>
        /// 对列进行字段的修复（比如将深圳总部替换为深圳）
        /// </summary>
        /// <returns></returns>
        IEnumerable<ModelBase> RepairByColumn(IList<ModelBase> list);

        /// <summary>
        /// 根据列进行筛选
        /// </summary>
        /// <returns>返回筛选后集合</returns>
        IEnumerable<ModelBase> FilterByColumn(IList<ModelBase> list);

    }
}
