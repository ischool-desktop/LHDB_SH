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

namespace LHDB_SH_Core.Config
{
    public partial class DepConfigForm : BaseForm
    {
        ConfigData _cd;

        string _ConfigName = "部別代碼";

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
                // 儲存
                foreach (DataGridViewRow dgvr in dgData.Rows)
                {
                    if (dgvr.IsNewRow)
                        continue;

                    ConfigDataItem cdi = new ConfigDataItem();
                    cdi.Name = GetCellValue(dgvr, colName.Index);
                    cdi.Value = GetCellValue(dgvr, colValue.Index);
                    cdi.TargetName = GetCellValue(dgvr, colTagName.Index);
                    datas.Add(cdi);
                }
                _cd.SetConfigDataItem(datas, _ConfigName);
                MsgBox.Show("儲存完成");
                this.Close();
            }
            catch(Exception ex)
            {
                MsgBox.Show("儲存過程發生錯誤," + ex.Message);
            }
        }

        private string GetCellValue(DataGridViewRow dr, int colIndx)
        {
            string value = "";
            if (dr.Cells[colIndx].Value != null)
                value = dr.Cells[colIndx].Value.ToString();

            return value;
        }

        private void DepConfigForm_Load(object sender, EventArgs e)
        {
            Dictionary<string, List<ConfigDataItem>> dataDict = _cd.GetConfigDataItemDict();
            if(dataDict.ContainsKey(_ConfigName))
            {
                foreach(ConfigDataItem cdi in dataDict[_ConfigName])
                {
                    int rowIdx = dgData.Rows.Add();
                    dgData.Rows[rowIdx].Cells[colName.Index].Value = cdi.Name;
                    dgData.Rows[rowIdx].Cells[colValue.Index].Value = cdi.Value;
                    dgData.Rows[rowIdx].Cells[colTagName.Index].Value = cdi.TargetName;
                }
            }

        }
    }
}
