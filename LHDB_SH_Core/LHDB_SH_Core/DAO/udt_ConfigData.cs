using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace LHDB_SH_Core.DAO
{
    [TableName("lhdb_sh_config_data")]
    public class udt_ConfigData:ActiveRecord
    {
        ///<summary>
        /// 設定名稱
        ///</summary>
        [Field(Field = "config_name", Indexed = false)]
        public string ConfigName { get; set; }

        ///<summary>
        /// 內容
        ///</summary>
        [Field(Field = "content", Indexed = false)]
        public string Content { get; set; }

    }
}
