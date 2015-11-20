using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using LHDB_SH_Core.DAO;
using SHSchool.Data;

namespace LHDB_SH_Core.Config
{
    public partial class DepConfigForm : BaseForm
    {
        ConfigData _cd;

        string _ConfigDepName = "部別代碼";
        string _ConfigClsName = "班別代碼";       

        public DepConfigForm()
        {
            InitializeComponent();
            _cd = new ConfigData();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // 檢查

                List<ConfigDataItem> datas = new List<ConfigDataItem>();
                // 儲存 部別
                foreach (DataGridViewRow dgvr in dgDepData.Rows)
                {
                    if (dgvr.IsNewRow)
                        continue;

                    ConfigDataItem cdi = new ConfigDataItem();
                    cdi.Name = Utility.GetDgCellValue(dgvr.Cells[colDepName.Index]);
                    cdi.Value = Utility.GetDgCellValue(dgvr.Cells[colDepValue.Index]);
                    cdi.TargetName = Utility.GetDgCellValue(dgvr.Cells[colDepTagName.Index]);
                    datas.Add(cdi);
                }
                _cd.SetConfigDataItem(datas, _ConfigDepName);

                List<ConfigDataItem> datas1 = new List<ConfigDataItem>();
                foreach(DataGridViewRow dgvr in dgClsData.Rows)
                {
                    if (dgvr.IsNewRow)
                        continue;
                    ConfigDataItem cdi = new ConfigDataItem();
                    cdi.Name = Utility.GetDgCellValue(dgvr.Cells[colClsName.Index]);
                    cdi.Value = Utility.GetDgCellValue(dgvr.Cells[colClsValue.Index]);
                    cdi.TargetName = Utility.GetDgCellValue(dgvr.Cells[colDepTagName.Index]);
                    datas1.Add(cdi);
                }
                _cd.SetConfigDataItem(datas1, _ConfigClsName);

                MsgBox.Show("儲存完成");
                this.Close();
            }
            catch(Exception ex)
            {
                MsgBox.Show("儲存過程發生錯誤," + ex.Message);
            }
        }

       
        private void DepConfigForm_Load(object sender, EventArgs e)
        {                // 
            List<SHTagConfigRecord> tagRecList = SHTagConfig.SelectAll();
            List<string> TagNameList = new List<string>();
            foreach (SHTagConfigRecord rec in tagRecList)
                TagNameList.Add(rec.FullName);

            TagNameList.Sort();

            colDepTagName.Items.AddRange(TagNameList.ToArray());
            colDepTagName.DropDownStyle = ComboBoxStyle.DropDownList;
            colClsTagName.Items.AddRange(TagNameList.ToArray());
            colClsTagName.DropDownStyle = ComboBoxStyle.DropDownList;

            Dictionary<string, List<ConfigDataItem>> dataDict = _cd.GetConfigDataItemDict();
            // 部別
            if(dataDict.ContainsKey(_ConfigDepName))
            {
                foreach(ConfigDataItem cdi in dataDict[_ConfigDepName])
                {
                    int rowIdx = dgDepData.Rows.Add();
                    dgDepData.Rows[rowIdx].Cells[colDepName.Index].Value = cdi.Name;
                    dgDepData.Rows[rowIdx].Cells[colDepValue.Index].Value = cdi.Value;
                    dgDepData.Rows[rowIdx].Cells[colDepTagName.Index].Value = cdi.TargetName;
                }
            }else
            {
                Dictionary<string,string> dDict=GetDepDefault();
                foreach(string key in dDict.Keys)
                {
                    int rowIdx = dgDepData.Rows.Add();
                    dgDepData.Rows[rowIdx].Cells[colDepName.Index].Value = key;
                    dgDepData.Rows[rowIdx].Cells[colDepValue.Index].Value = dDict[key];                    
                }
            }
            

            // 班別
            if(dataDict.ContainsKey(_ConfigClsName))
            {
                foreach(ConfigDataItem cdi in dataDict[_ConfigClsName])
                {
                    int rowIdx = dgClsData.Rows.Add();
                    dgClsData.Rows[rowIdx].Cells[colClsName.Index].Value = cdi.Name;
                    dgClsData.Rows[rowIdx].Cells[colClsValue.Index].Value = cdi.Value;
                    dgClsData.Rows[rowIdx].Cells[colClsTagName.Index].Value = cdi.TargetName;
                }
            }else
            {
                Dictionary<string, string> dDict = GetClsDefault();
                foreach(string key in dDict.Keys)
                {
                    int rowIdx = dgClsData.Rows.Add();
                    dgClsData.Rows[rowIdx].Cells[colClsName.Index].Value = key;
                    dgClsData.Rows[rowIdx].Cells[colClsValue.Index].Value = dDict[key];                    
                }
            }

        }

        /// <summary>
        /// 取得部別預設
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, string> GetDepDefault()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("日間部", "1");
            value.Add("夜間部", "2");
            value.Add("實用技能學程(日)", "3");
            value.Add("實用技能學程(夜)", "4");
            value.Add("進修部(學校)", "5");
            return value;
        }

        /// <summary>
        /// 取得班別預設
        /// </summary>
        /// <returns></returns>
        private Dictionary<string,string> GetClsDefault()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("編制班", "1");
            value.Add("建教班", "2");
            value.Add("重點產業班", "3");
            value.Add("雙軌旗艦訓練計劃專班", "4");
            value.Add("員工進修班", "5");
            return value;
        }
    }
}
