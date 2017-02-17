using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IDataValidation
{
    public interface INullValidation
    {
        /// <summary>
        /// 非空检验
        /// </summary>
        /// <returns></returns>
        bool IsNotNull();

    }
}
