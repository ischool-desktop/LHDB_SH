using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using FISCA.UDT;

namespace LHDB_SH_Core.DAO
{
    public class ConfigData
    {

        public enum ConfigType {部別代碼,班級代碼}
        /// <summary>
        /// 儲存代碼對照
        /// </summary>
        /// <param name="items"></param>
        public void SetCode(List<ConfigDataItem> items,ConfigType type)
        { 
            //取得udt資料
            AccessHelper accHelper = new AccessHelper();
            udt_ConfigData data = null;

            if (data == null)
                data = new udt_ConfigData();

            data.ConfigName = type.ToString();
            

        }

        /// <summary>
        /// 取得代碼對照
        /// </summary>
        /// <returns></returns>
        public List<ConfigDataItem> GetDepCode(ConfigType type)
        {
            List<ConfigDataItem> Value = new List<ConfigDataItem>();

            return Value;
        }

        private XElement ConvertItemToXML(List<ConfigDataItem> items)
        {
            XElement elmRoot = new XElement("Content");
            return elmRoot;
        }

    }
}
