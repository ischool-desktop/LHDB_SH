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
        /// <summary>
        /// 儲存代碼對照
        /// </summary>
        /// <param name="items"></param>
        /// <param name="type"></param>
        public void SetConfigDataItem(List<ConfigDataItem> items,string ConfigName)
        { 
            //取得udt資料
            AccessHelper accHelper = new AccessHelper();
            udt_ConfigData data = null;
            string strQuery = "config_name='" + ConfigName+"'";
            List<udt_ConfigData> datas = accHelper.Select<udt_ConfigData>(strQuery);

            if (datas.Count > 0)
                data = datas[0];

            if (data == null)
                data = new udt_ConfigData();

            data.ConfigName = ConfigName;
            data.Content = ConvertItemToXML(items).ToString();
            data.Save();

        }

        /// <summary>
        /// 取得所有代碼對照
        /// </summary>
        /// <returns></returns>
        public Dictionary<string,List<ConfigDataItem>> GetConfigDataItemDict()
        {
            Dictionary<string, List<ConfigDataItem>> Value = new Dictionary<string, List<ConfigDataItem>>();
            AccessHelper accHelper = new AccessHelper();
            List<udt_ConfigData> datas = accHelper.Select<udt_ConfigData>();
            foreach(udt_ConfigData data in datas)
            {
                if(!string.IsNullOrEmpty(data.ConfigName ))
                    Value.Add(data.ConfigName,ConvertXMLStrToItem(data.Content));
            }

            return Value;
        }

        private XElement ConvertItemToXML(List<ConfigDataItem> items)
        {
            XElement elmRoot = new XElement("Content");
            foreach (ConfigDataItem item in items)
            {
                XElement elm = new XElement("Item");
                elm.SetAttributeValue("Name", item.Name);
                elm.SetAttributeValue("Value", item.Value);
                elm.SetAttributeValue("TargetName", item.TargetName);
                elmRoot.Add(elm);
            }
            return elmRoot;
        }

        private List<ConfigDataItem> ConvertXMLStrToItem(string content)
        {
            List<ConfigDataItem> value = new List<ConfigDataItem>();
            try
            {
                XElement elmRoot = XElement.Parse(content);
                foreach(XElement elm in elmRoot.Elements("Item"))
                {
                    ConfigDataItem item = new ConfigDataItem();
                    item.Name = GetXmlAttributeString(elm, "Name");
                    item.Value = GetXmlAttributeString(elm, "Value");
                    item.TargetName = GetXmlAttributeString(elm, "TargetName");                    
                    value.Add(item);
                }
            }
            catch (Exception ex)
            { 
            }
            return value;
        }

        private string GetXmlAttributeString(XElement elm,string name)
        {
            string value = "";
            if (elm.Attribute(name) != null)
                value = elm.Attribute(name).Value;
            return value;
        }

        private Dictionary<string, string> KeyValueItemDict = new Dictionary<string, string>();

        public void ClearKeyValueItem()
        {
            KeyValueItemDict.Clear();
        }

        public void AddKeyValueItem(string key,string Value)
        {
            if (!KeyValueItemDict.ContainsKey(key))
                KeyValueItemDict.Add(key, Value);
        }

        public void SaveKeyValueItem(string ConfigName)
        {
            List<ConfigDataItem> its = new List<ConfigDataItem>();
            foreach(string key in KeyValueItemDict.Keys)
            {
                ConfigDataItem cdi = new ConfigDataItem();
                cdi.Name = key;
                cdi.Value = KeyValueItemDict[key];
                its.Add(cdi);
            }
            SetConfigDataItem(its, ConfigName);
        }

        public Dictionary<string,string> GetKeyValueItem(string ConfigName)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();

            Dictionary<string, List<ConfigDataItem>> datas = GetConfigDataItemDict();
            if(datas.ContainsKey(ConfigName))
            {
                foreach(ConfigDataItem cdi in  datas[ConfigName])
                {
                    if (!value.ContainsKey(cdi.Name))
                        value.Add(cdi.Name, cdi.Value);
                }
            }

            return value;
        }

    }
}
