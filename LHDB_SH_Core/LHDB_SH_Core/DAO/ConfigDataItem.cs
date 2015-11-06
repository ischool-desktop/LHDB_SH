using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LHDB_SH_Core.DAO
{
    public class ConfigDataItem
    {
        /// <summary>
        /// 名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 對照明稱
        /// </summary>
        public string TargetName { get; set; }

        /// <summary>
        /// 值
        /// </summary>
        public string Value { get; set; }
    }
}
