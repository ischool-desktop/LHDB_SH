using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using SHSchool.Data;
using K12.Data;
using LHDB_SH_Core.DAO;

namespace LHDB_SH_Core.Report
{
    public partial class AttendanceReport : BaseForm
    {
        ConfigData _cd;
        string _ConfigName = "缺勤代碼";
        BackgroundWorker _bgWorker;

        List<string> _StudentIDList;
        int _SchoolYear = 0, _Semester = 0;

        public AttendanceReport(List<string> StudentIDList)
        {
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.WorkerReportsProgress = true;
            _StudentIDList = StudentIDList;
            _cd = new ConfigData();

            InitializeComponent();
                    }

        void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            List<SHStudentRecord> StudentRecList= SHStudent.SelectByIDs(_StudentIDList);

            List<SHAttendanceRecord> AttendanceList = SHAttendance.SelectBySchoolYearAndSemester(StudentRecList, _SchoolYear, _Semester);
            Dictionary<string, StudAttendanceRec> AttendanceDict = new Dictionary<string, StudAttendanceRec>();
            foreach(SHAttendanceRecord rec in AttendanceList)
            {
                foreach(AttendancePeriod per in rec.PeriodDetail)
                {
                    string strDate=(per.OccurDate.Year-1911)+string.Format("{0:00}",per.OccurDate.Month)+string.Format("{0:00}",per.OccurDate.Day);
                    string key = rec.RefStudentID + "_" + strDate + "_" + per.AbsenceType;

                }                
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            // 儲存畫面資料
            _SchoolYear = iptSchoolYear.Value;
            _Semester = iptSemester.Value;

            _bgWorker.RunWorkerAsync();
        }

        private void AttendanceReport_Load(object sender, EventArgs e)
        {
            // 預設學年度、學期
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);

            // 載入代碼對照
            Dictionary<string, List<ConfigDataItem>> datas = _cd.GetConfigDataItemDict();
            Dictionary<string, string> dataDict = GetDefaultItemDict();

            // 有儲存過
            if(datas.ContainsKey(_ConfigName))
            {
                List<string> defList = new List<string>();
                int rowIdx = 0;
                foreach(ConfigDataItem cdi in datas[_ConfigName])
                {
                    defList.Add(cdi.Name);
                    rowIdx = dgData.Rows.Add();
                    dgData.Rows[rowIdx].Cells[colName.Index].Value = cdi.Name;
                    dgData.Rows[rowIdx].Cells[colValue.Index].Value = cdi.Value;
                    dgData.Rows[rowIdx].Cells[colType.Index].Value = cdi.TargetName;
                }
            
                // 檢查加入預設是否完整
                foreach(string key in dataDict.Keys)
                {
                    if(!defList.Contains(key))
                    {
                        rowIdx = dgData.Rows.Add();
                        dgData.Rows[rowIdx].Cells[colName.Index].Value = key;
                        dgData.Rows[rowIdx].Cells[colValue.Index].Value = dataDict[key];
                        dgData.Rows[rowIdx].Cells[colValue.Index].Value = "";
                    }
                }
            }
            else
            {
                // 完全沒使用過                
                foreach(string key in dataDict.Keys)
                {
                    int rowIdx = dgData.Rows.Add();
                    dgData.Rows[rowIdx].Cells[colName.Index].Value = key;
                    dgData.Rows[rowIdx].Cells[colValue.Index].Value = dataDict[key];
                    dgData.Rows[rowIdx].Cells[colValue.Index].Value = "";
                }
            }

            // 載入節次對照
            List<PeriodMappingInfo> PeriodMappingList = PeriodMapping.SelectAll();
            // 節次>類別            
            foreach (PeriodMappingInfo rec in PeriodMappingList)
            {
                string name =rec.Name+"_"+rec.Type;
                lvPeriod.Items.Add(name);
            }

            

        }

        /// <summary>
        /// 取得預設對照
        /// </summary>
        /// <returns></returns>
        private Dictionary<string,string> GetDefaultItemDict()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("公假", "01");
            value.Add("事假", "02");
            value.Add("病假", "03");
            value.Add("婚假", "04");
            value.Add("產前假", "05");
            value.Add("娩假", "06");
            value.Add("陪產假", "07");
            value.Add("流產假", "08");
            value.Add("育嬰假", "09");
            value.Add("生理假", "10");
            value.Add("喪假", "11");
            value.Add("曠課", "12");
            return value;
        }
    }
}
