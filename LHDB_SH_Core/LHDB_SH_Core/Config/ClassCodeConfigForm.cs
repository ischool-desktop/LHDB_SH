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
    public partial class ClassCodeConfigForm : BaseForm
    {
        ConfigData _cd;

        string _ConfigClassName = "班級代碼";
        

        public ClassCodeConfigForm()
        {
            _cd = new ConfigData();
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                List<ConfigDataItem> datas = new List<ConfigDataItem>();
                foreach(DataGridViewRow dgvr in dgClassCode.Rows)
                {
                    if (dgvr.IsNewRow)
                        continue;

                    ConfigDataItem cdi = new ConfigDataItem();
                    cdi.Name = Utility.GetDgCellValue(dgvr.Cells[colClassName.Index]);
                    cdi.TargetName = Utility.GetDgCellValue(dgvr.Cells[colStClassCode.Index]);
                    cdi.Value = Utility.GetDgCellValue(dgvr.Cells[colLHClassCode.Index]);
                    datas.Add(cdi);
                }
                _cd.SetConfigDataItem(datas, _ConfigClassName);
                MsgBox.Show("儲存完成。");
                this.Close();
            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存設定發生錯誤," + ex.Message);
            }

        }

        private void ClassCodeConfigForm_Load(object sender, EventArgs e)
        {
            GetClassCodeToDG();
        }

        /// <summary>
        /// 取得班級代碼設定
        /// </summary>
        private void GetClassCodeToDG()
        {
            Dictionary<string, List<ConfigDataItem>> datas = _cd.GetConfigDataItemDict();
            if(datas.ContainsKey(_ConfigClassName))
            {
                foreach(ConfigDataItem cdi in datas[_ConfigClassName])
                {
                    int rowIdx = dgClassCode.Rows.Add();
                    dgClassCode.Rows[rowIdx].Cells[colClassName.Index].Value = cdi.Name;
                    dgClassCode.Rows[rowIdx].Cells[colStClassCode.Index].Value = cdi.TargetName;
                    dgClassCode.Rows[rowIdx].Cells[colLHClassCode.Index].Value = cdi.Value;
                } 
            }
            else
            {
                GetDefaultToDG();
            }
        }

        private void GetDefaultToDG()
        {
            dgClassCode.Rows.Clear();
            // 取得繁星班級代碼設定當預設
            Dictionary<string, string> tmpClassCodeDict = Utility.GetClassCodeDict();
            foreach (string key in tmpClassCodeDict.Keys)
            {
                int rowIdx = dgClassCode.Rows.Add();
                dgClassCode.Rows[rowIdx].Cells[colClassName.Index].Value = key;
                dgClassCode.Rows[rowIdx].Cells[colStClassCode.Index].Value = tmpClassCodeDict[key];
                dgClassCode.Rows[rowIdx].Cells[colLHClassCode.Index].Value = tmpClassCodeDict[key];
            }

            if(tmpClassCodeDict.Count==0)
            {
                MsgBox.Show("請先設定大學繁星班級代碼，功能位置：學生>大學繁星>設定班級代碼。");              
            }
        }
    }
}
