using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel数据导入.ExcelImport
{
    public enum EnumDefaultValue
    {
        /// <summary>
        /// 无默认值
        /// </summary>
        NONE = 0,

        /// <summary>
        /// 唯一ID GUID随机生成
        /// </summary>
        ID = 1,

        /// <summary>
        /// 当前时间 获取系统当前时间
        /// </summary>
        CURRENT_TIME = 2
    }
}
