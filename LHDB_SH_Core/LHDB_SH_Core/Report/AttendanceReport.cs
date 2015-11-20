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
using Aspose.Cells;
using System.IO;

namespace LHDB_SH_Core.Report
{
    public partial class AttendanceReport : BaseForm
    {
        ConfigData _cd;
        string _ConfigName1 = "缺勤紀錄名冊_假別設定";
        string _ConfigName2 = "缺勤紀錄名冊_節次設定";
        string _ConfigName3 = "缺勤紀錄名冊_畫面設定";
        BackgroundWorker _bgWorker;

        List<string> _StudentIDList;
        // 學校代碼、名冊別
        string _SchoolCode = "",_DocType="3";
        int _SchoolYear = 0, _Semester = 0;
        Workbook _wb = null;

        // 程式對照使用
        Dictionary<string, string> _AbsenceDict = new Dictionary<string, string>();
        List<string> _PeriodList = new List<string>();

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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("缺勤紀錄名冊產生中..", e.ProgressPercentage);
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("缺勤紀錄名冊產生完成.");
            btnExport.Enabled = true;
            if(_wb !=null)
            {
                try
                {
                    Utility.ExprotXls("缺勤紀錄名冊", _wb);
                }catch(Exception ex)
                {
                    MsgBox.Show("產生缺勤紀錄名冊匯入檔發生錯誤," + ex.Message);
                }
            }
            else
            { MsgBox.Show("無法產生缺勤紀錄名冊匯入檔"); }
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(10);
            _SchoolCode = K12.Data.School.Code;

            // 讀取缺勤名冊樣板
            _wb = new Workbook(new MemoryStream(Properties.Resources.缺勤紀錄名冊樣板));

            Worksheet wst1 = _wb.Worksheets["缺勤紀錄名冊封面"];
            Worksheet wst2 = _wb.Worksheets["缺勤紀錄名冊"];

            List<SHStudentRecord> StudentRecList= SHStudent.SelectByIDs(_StudentIDList);
            Dictionary<string, SHStudentRecord> StudentRecDict = new Dictionary<string, SHStudentRecord>();
            foreach (SHStudentRecord rec in StudentRecList)
                StudentRecDict.Add(rec.ID, rec);

            _bgWorker.ReportProgress(30);

            List<SHAttendanceRecord> AttendanceList = SHAttendance.SelectBySchoolYearAndSemester(StudentRecList, _SchoolYear, _Semester);
            Dictionary<string, StudAttendanceRec> AttendanceDict = new Dictionary<string, StudAttendanceRec>();
            foreach(SHAttendanceRecord rec in AttendanceList)
            {
                foreach(AttendancePeriod per in rec.PeriodDetail)
                {
                    if(_AbsenceDict.ContainsKey(per.AbsenceType) && _PeriodList.Contains(per.Period))
                    {
                        string strDate = Utility.ConvertChDateString(rec.OccurDate);
                        string key = rec.RefStudentID + "_" + strDate + "_" + per.AbsenceType;

                        if (!AttendanceDict.ContainsKey(key))
                        {
                            AttendanceDict.Add(key, new StudAttendanceRec());
                            AttendanceDict[key].StudentID = rec.RefStudentID;
                            AttendanceDict[key].BeginDate = strDate;
                            AttendanceDict[key].EndDate = strDate;
                            // 代碼轉換
                            if (_AbsenceDict.ContainsKey(per.AbsenceType))
                                AttendanceDict[key].AttendType = _AbsenceDict[per.AbsenceType];

                            AttendanceDict[key].AttendTypeCount = 0;
                        }
                        AttendanceDict[key].AttendTypeCount++;

                    }
                }                
            }

            _bgWorker.ReportProgress(60);
            // 讀取資料
            List<StudAttendanceRec> StudAttendanceRecList = AttendanceDict.Values.ToList();

            // 填入身分證、生日
            foreach(StudAttendanceRec rec in StudAttendanceRecList)
            {
                if (StudentRecDict.ContainsKey(rec.StudentID))
                {
                    rec.IDNumber = StudentRecDict[rec.StudentID].IDNumber.ToUpper();
                    rec.BirthDate = Utility.ConvertChDateString(StudentRecDict[rec.StudentID].Birthday);
                }
            }
            
            // 排序
            StudAttendanceRecList = (from data in StudAttendanceRecList orderby data.IDNumber, data.BeginDate,data.AttendType select data).ToList();

            _bgWorker.ReportProgress(80);
            
            // 寫入資料 名冊封面
            wst1.Cells[1, 0].PutValue(_SchoolCode);
            wst1.Cells[1, 1].PutValue(_SchoolYear);
            wst1.Cells[1, 2].PutValue(_Semester);
            wst1.Cells[1, 3].PutValue(_DocType);

            // 名冊內容
            // 身分證號,出生日期,缺勤種類代碼,缺勤節數,缺勤起始日期,缺勤結束日期
            int rowIdx=1;
            foreach(StudAttendanceRec rec in StudAttendanceRecList)
            {
                wst2.Cells[rowIdx, 0].PutValue(rec.IDNumber);
                wst2.Cells[rowIdx, 1].PutValue(rec.BirthDate);
                wst2.Cells[rowIdx, 2].PutValue(rec.AttendType);
                wst2.Cells[rowIdx, 3].PutValue(rec.AttendTypeCount);
                wst2.Cells[rowIdx, 4].PutValue(rec.BeginDate);
                wst2.Cells[rowIdx, 5].PutValue(rec.EndDate);
                rowIdx++;
            }
            _bgWorker.ReportProgress(100);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            btnExport.Enabled = false;

            // 儲存畫面資料
            _SchoolYear = iptSchoolYear.Value;
            _Semester = iptSemester.Value;

            // 儲存假別設定
            SaveAbsenceDg();

            // 儲存節次設定
            SavePeriodDg();
            
            // 儲存學年度學期
            _cd.ClearKeyValueItem();
            _cd.AddKeyValueItem("學年度",_SchoolYear.ToString());
            _cd.AddKeyValueItem("學期", _Semester.ToString());
            _cd.SaveKeyValueItem(_ConfigName3);
            
            _bgWorker.RunWorkerAsync();
        }

        private void SaveAbsenceDg()
        {
            _AbsenceDict.Clear();
            try
            {
                List<ConfigDataItem> dataList = new List<ConfigDataItem>();
                foreach (DataGridViewRow drv in dgData.Rows)
                {
                    if (drv.IsNewRow)
                        continue;

                    ConfigDataItem cdi = new ConfigDataItem();
                    cdi.Name = Utility.GetDgCellValue(drv.Cells[colName.Index]);
                    cdi.Value = Utility.GetDgCellValue(drv.Cells[colValue.Index]);
                    cdi.TargetName = Utility.GetDgCellValue(drv.Cells[colType.Index]);
                    dataList.Add(cdi);

                    if(!_AbsenceDict.ContainsKey(cdi.TargetName))
                        _AbsenceDict.Add(cdi.TargetName,cdi.Value);
                }
                _cd.SetConfigDataItem(dataList, _ConfigName1);
            }catch(Exception ex)
            {
                MsgBox.Show("儲存假別發生錯誤," + ex.Message);
            }
        }

        private void SavePeriodDg()
        {
            _PeriodList.Clear();
            try
            {
                List<ConfigDataItem> dataList = new List<ConfigDataItem>();
                foreach (DataGridViewRow drv in dgPerdata.Rows)
                {
                    if (drv.IsNewRow)
                        continue;

                    ConfigDataItem cdi = new ConfigDataItem();
                    cdi.Name = Utility.GetDgCellValue(drv.Cells[colPerName.Index]);                    
                    cdi.TargetName = Utility.GetDgCellValue(drv.Cells[colPerType.Index]);
                    dataList.Add(cdi);
                    _PeriodList.Add(cdi.TargetName);
                }
                _cd.SetConfigDataItem(dataList, _ConfigName2);
            }catch(Exception ex)
            {
                MsgBox.Show("儲存節次發生錯誤," + ex.Message);
            }
        }


        private void AttendanceReport_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            // 預設學年度、學期
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);
            colType.DropDownStyle = ComboBoxStyle.DropDownList;
            colPerType.DropDownStyle = ComboBoxStyle.DropDownList;

            // 載入代碼對照
            Dictionary<string, List<ConfigDataItem>> datas = _cd.GetConfigDataItemDict();
            Dictionary<string, string> dataDict = GetDefaultItemDict();
            List<string> dataList1 = GetDefaultPeridList();
            Dictionary<string, string> tmpDict1 = GetDefaultItemDict1();
            Dictionary<string, string> tmpDict2 = GetDefaultPeridDict();

            #region 讀取設定值
            // 有儲存過 假別
            if (datas.ContainsKey(_ConfigName1))
            {
                List<string> defList = new List<string>();
                int rowIdx = 0;
                foreach (ConfigDataItem cdi in datas[_ConfigName1])
                {
                    defList.Add(cdi.Name);
                    rowIdx = dgData.Rows.Add();
                    dgData.Rows[rowIdx].Cells[colName.Index].Value = cdi.Name;
                    dgData.Rows[rowIdx].Cells[colValue.Index].Value = cdi.Value;
                    dgData.Rows[rowIdx].Cells[colType.Index].Value = cdi.TargetName;
                }

                // 檢查加入預設是否完整
                foreach (string key in dataDict.Keys)
                {
                    if (!defList.Contains(key))
                    {
                        rowIdx = dgData.Rows.Add();
                        dgData.Rows[rowIdx].Cells[colName.Index].Value = key;
                        dgData.Rows[rowIdx].Cells[colValue.Index].Value = dataDict[key];                                                
                        if (tmpDict1.ContainsKey(key))
                            dgData.Rows[rowIdx].Cells[colType.Index].Value = tmpDict1[key];
                        else
                            dgData.Rows[rowIdx].Cells[colType.Index].Value = "";
                    }
                }
            }
            else
            {
                // 完全沒使用過                
                foreach (string key in dataDict.Keys)
                {
                    int rowIdx = dgData.Rows.Add();
                    dgData.Rows[rowIdx].Cells[colName.Index].Value = key;
                    dgData.Rows[rowIdx].Cells[colValue.Index].Value = dataDict[key];                    
                    if (tmpDict1.ContainsKey(key))
                        dgData.Rows[rowIdx].Cells[colType.Index].Value = tmpDict1[key];
                    else
                        dgData.Rows[rowIdx].Cells[colType.Index].Value = "";
                }
            }

            // 有儲存過 節次
            if (datas.ContainsKey(_ConfigName2))
            {
                List<string> defList = new List<string>();
                
                int rowIdx = 0;
                foreach (ConfigDataItem cdi in datas[_ConfigName2])
                {
                    defList.Add(cdi.Name);
                    rowIdx = dgPerdata.Rows.Add();
                    dgPerdata.Rows[rowIdx].Cells[colPerName.Index].Value = cdi.Name;
                    dgPerdata.Rows[rowIdx].Cells[colPerType.Index].Value = cdi.TargetName;
                }

                // 檢查加入預設是否完整
                foreach (string key in dataList1)
                {
                    if (!defList.Contains(key))
                    {
                        rowIdx = dgPerdata.Rows.Add();
                        dgPerdata.Rows[rowIdx].Cells[colPerName.Index].Value = key;                        
                        dgPerdata.Rows[rowIdx].Cells[colPerType.Index].Value ="" ;
                        if (tmpDict2.ContainsKey(key))
                            dgPerdata.Rows[rowIdx].Cells[colPerType.Index].Value = tmpDict2[key];
                    }
                }
            }
            else
            {
                // 完全沒使用過                
                foreach (string key in dataList1)
                {
                    int rowIdx = dgPerdata.Rows.Add();
                    dgPerdata.Rows[rowIdx].Cells[colPerName.Index].Value = key;
                    dgPerdata.Rows[rowIdx].Cells[colPerType.Index].Value = "";
                    if (tmpDict2.ContainsKey(key))
                        dgPerdata.Rows[rowIdx].Cells[colPerType.Index].Value = tmpDict2[key];
                }
            }

            Dictionary<string, string> ds = _cd.GetKeyValueItem(_ConfigName3);
             if (ds.ContainsKey("學年度"))
                    if (ds["學年度"]!="")
                        iptSchoolYear.Value = int.Parse(ds["學年度"]);
              if (ds.ContainsKey("學期"))
                  if (ds["學期"] != "")
                      iptSemester.Value = int.Parse(ds["學期"]);
            #endregion



            // 載入假別
            List<AbsenceMappingInfo> AbsenceMappingList = AbsenceMapping.SelectAll();
            List<string> aList = new List<string>();
            aList.Add("");
            foreach (AbsenceMappingInfo rec in AbsenceMappingList)
               aList.Add(rec.Name);
            
            colType.Items.AddRange(aList.ToArray());

            // 載入節次對照
            List<PeriodMappingInfo> PeriodMappingList = PeriodMapping.SelectAll();
            
            // 節次>類別
            List<string> pList = new List<string>();
            pList.Add("");
            foreach (PeriodMappingInfo rec in PeriodMappingList)
                pList.Add(rec.Name);            
            colPerType.Items.AddRange(pList.ToArray());

   

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

        /// <summary>
        /// 取得預設需要統計節次
        /// </summary>
        /// <returns></returns>
        private List<string> GetDefaultPeridList()
        {
            List<string> value = new List<string>();
            value.Add("1");
            value.Add("2");
            value.Add("3");
            value.Add("4");
            value.Add("5");
            value.Add("6");
            value.Add("7");
            return value;
        }

        private Dictionary<string,string> GetDefaultPeridDict()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("1", "一");
            value.Add("2", "二");
            value.Add("3", "三");
            value.Add("4", "四");
            value.Add("5", "五");
            value.Add("6", "六");
            value.Add("7", "七");
            return value;
        }

        private Dictionary<string, string> GetDefaultItemDict1()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("公假", "公假");
            value.Add("事假", "事假");
            value.Add("病假", "病假");
            value.Add("婚假", "婚假");
            value.Add("產前假", "產前假");
            value.Add("娩假", "娩假");
            value.Add("陪產假", "陪產假");
            value.Add("流產假", "流產假");
            value.Add("育嬰假", "育嬰假");
            value.Add("生理假", "生理假");
            value.Add("喪假", "喪假");
            value.Add("曠課", "曠課");
            return value;
        }

    }
}
