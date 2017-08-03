using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IDataValidation
{
   public interface IFormatValidation
    {
        /// <summary>
        /// 数据格式检验
        /// </summary>
        /// <returns></returns>
        bool IsRightFormat();
    }
}
