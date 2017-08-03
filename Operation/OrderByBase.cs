using Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Operation
{
    /// <summary>
    /// 排序抽象类
    /// </summary>
    public abstract class OrderByBase
    {
        public abstract IEnumerable<ModelBase> OrederOperation(IList<ModelBase> list);
    }
}
