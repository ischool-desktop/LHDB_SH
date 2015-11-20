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
    public partial class SpecClassNameConfigForm : BaseForm
    {
        ConfigData _cd;

        string _ConfigClassName = "特色班實驗班名稱對照";

        public SpecClassNameConfigForm()
        {
            _cd = new ConfigData();
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SpecClassNameConfigForm_Load(object sender, EventArgs e)
        {
            GetClassCodeToDG();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            try
            {
                SaveDgData();
                MsgBox.Show("儲存完成。");
                this.Close();
            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存設定發生錯誤," + ex.Message);
            }
        }

        private void SaveDgData()
        {
            List<ConfigDataItem> datas = new List<ConfigDataItem>();
            foreach (DataGridViewRow dgvr in dgClassName.Rows)
            {
                if (dgvr.IsNewRow)
                    continue;

                ConfigDataItem cdi = new ConfigDataItem();
                cdi.Name = Utility.GetDgCellValue(dgvr.Cells[colClassName.Index]);                
                cdi.Value = Utility.GetDgCellValue(dgvr.Cells[colSpecClassName.Index]);
                datas.Add(cdi);
            }
            _cd.SetConfigDataItem(datas, _ConfigClassName);
        }

        private void GetClassCodeToDG()
        {
            Dictionary<string, List<ConfigDataItem>> datas = _cd.GetConfigDataItemDict();
            if (datas.ContainsKey(_ConfigClassName))
            {
                foreach (ConfigDataItem cdi in datas[_ConfigClassName])
                {
                    int rowIdx = dgClassName.Rows.Add();
                    dgClassName.Rows[rowIdx].Cells[colClassName.Index].Value = cdi.Name;                    
                    dgClassName.Rows[rowIdx].Cells[colSpecClassName.Index].Value = cdi.Value;
                }
            }
            else
            {
                // 完全沒資料
                GetDefaultToDG();
            }
        }

        private void GetDefaultToDG()
        {            
            dgClassName.Rows.Clear();

            List<SHClassRecord> ClassRecList = SHClass.SelectAll();
            ClassRecList = (from data in ClassRecList orderby data.Name select data).ToList();

            foreach (SHClassRecord rec in ClassRecList)
            {
                int rowIdx = dgClassName.Rows.Add();
                dgClassName.Rows[rowIdx].Cells[colClassName.Index].Value = rec.Name;
                dgClassName.Rows[rowIdx].Cells[colSpecClassName.Index].Value = "";
                
            }
        }
    }
}
