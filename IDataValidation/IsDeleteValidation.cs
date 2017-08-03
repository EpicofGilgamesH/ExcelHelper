using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IDataValidation
{
    public interface IsDeleteValidation
    {
        /// <summary>
        /// 删除检验（根据列值直接删除）
        /// </summary>
        /// <returns></returns>
        bool IsDelete();

    }
}
