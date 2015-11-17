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
using Aspose.Cells;
using System.IO;

namespace LHDB_SH_Core.Report
{
    public partial class StudentSCETakeScoreReport : BaseForm
    {
        private int _SchoolYear = 0, _Semester = 0, _ExamNo=0;
        private string _SchoolCode = "";
        private string _ConfigName = "定期考查名冊_畫面設定";
        private ConfigData _cd;
        
        // 名冊別
        private string _DocType = "5";

        BackgroundWorker _bgWorker;

        Dictionary<string, string> _ExamNameDict;
        string _ExamID = "";

        Workbook _wb;
        List<string> _StudentIDList;

        public StudentSCETakeScoreReport(List<string> StudentIDList)
        {
            _cd = new ConfigData();
            _ExamNameDict = new Dictionary<string, string>();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.WorkerReportsProgress = true;
            InitializeComponent();
            _StudentIDList = StudentIDList;
        }

        void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("定期考查成績名冊產生中..",e.ProgressPercentage);
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("定期考查成績名冊產生完成");
            btnExport.Enabled = true;
            try
            {
                if(_wb!=null)
                {
                    Utility.ExprotXls("定期考查成績名冊", _wb);
                }
            }catch(Exception ex)
            { }
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            
            _bgWorker.ReportProgress(1);
            // 取得部別、班別對照
            ConfigData cd = new ConfigData();
            Dictionary<string, List<ConfigDataItem>> cdDict = cd.GetConfigDataItemDict();
            Dictionary<string, string> DepMappingDict = new Dictionary<string, string>();
            Dictionary<string, string> ClsMappingDict = new Dictionary<string, string>();
            Dictionary<string, string> DeptMappingDict = new Dictionary<string, string>();
            Dictionary<string, string> ClassNoMappingDict = new Dictionary<string, string>();
            Dictionary<string, string> ClassIDNameDict = new Dictionary<string, string>();
            Dictionary<string, List<string>> StudTagNameDict = new Dictionary<string, List<string>>();
            List<SHStudentTagRecord> SHStudentTagRecordList = SHStudentTag.SelectByStudentIDs(_StudentIDList);

            // 取得學生科別名稱
            Dictionary<string, string> StudeDeptNameDict = Utility.GetStudDeptNameDict(_StudentIDList);

            // 取得學生類別
            foreach (SHStudentTagRecord TRec in SHStudentTagRecordList)
            {
                if (!StudTagNameDict.ContainsKey(TRec.RefStudentID))
                    StudTagNameDict.Add(TRec.RefStudentID, new List<string>());

                StudTagNameDict[TRec.RefStudentID].Add(TRec.FullName);
            }

            _bgWorker.ReportProgress(20);

            foreach (SHClassRecord rec in SHClass.SelectAll())
                ClassIDNameDict.Add(rec.ID, rec.Name);

            // 班級代碼對照
            ClassNoMappingDict = Utility.GetClassCodeDict();

            // 班別對照
            if (cdDict.ContainsKey("班別代碼"))
            {
                foreach (ConfigDataItem cdi in cdDict["班別代碼"])
                {
                    if (!ClsMappingDict.ContainsKey(cdi.TargetName))
                        ClsMappingDict.Add(cdi.TargetName, cdi.Value);
                }
            }

            // 科別對照
            DeptMappingDict = Utility.GetDepartmetDict();
            
            // 取得學生基本資料
            List<SHStudentRecord> StudentRecList = SHStudent.SelectByIDs(_StudentIDList);
            _bgWorker.ReportProgress(40);


            // 取得學生定期成績資料
            Dictionary<string, List<StudentSCETakeRec>> StudSCETakeDict = Utility.GetStudentSCETakeDict(_StudentIDList, _ExamID,_SchoolYear,_Semester);

            // 填入對應值
            foreach(SHStudentRecord StudRec in StudentRecList)
            {
                if (StudSCETakeDict.ContainsKey(StudRec.ID))
                {
                    string IDNumber = StudRec.IDNumber.ToUpper();
                    string BirthDate = Utility.ConvertChDateString(StudRec.Birthday);

                    // 科/班/學程別代碼
                    string DCLCode="000";
                    if (StudeDeptNameDict.ContainsKey(StudRec.ID))
                    {
                        string name = StudeDeptNameDict[StudRec.ID];
                        if (DeptMappingDict.ContainsKey(name))
                            DCLCode = DeptMappingDict[name];
                    }


                    // 修課班別
                    string ClClassName="000";
                    if (StudTagNameDict.ContainsKey(StudRec.ID))
                    {
                        foreach (string str in StudTagNameDict[StudRec.ID])
                        {
                            if (ClsMappingDict.ContainsKey(str))
                                ClClassName = ClsMappingDict[str];
                        }
                    }

                    // 修課班級
                    string ClassCode = "000";
                    if (ClassIDNameDict.ContainsKey(StudRec.RefClassID))
                    {
                        string cName = ClassIDNameDict[StudRec.RefClassID];
                        if (ClassNoMappingDict.ContainsKey(cName))
                            ClassCode = ClassNoMappingDict[cName];
                    }

                    foreach(StudentSCETakeRec rec in StudSCETakeDict[StudRec.ID])
                    {
                        rec.IDNumber = IDNumber;
                        rec.BirthDate = BirthDate;
                        rec.ClassName = ClassCode;
                        rec.ClClassName = ClClassName;
                        rec.DCLCode = DCLCode;
                    }
                }
            }
            _bgWorker.ReportProgress(70);

            List<StudentSCETakeRec> StudentSCETakeRecList = new List<StudentSCETakeRec>();
            foreach(List<StudentSCETakeRec> recList in StudSCETakeDict.Values)
            {
                foreach (StudentSCETakeRec rec in recList)
                    StudentSCETakeRecList.Add(rec);
            }

            // 排序 依身分證,科目代碼
            StudentSCETakeRecList=(from data in StudentSCETakeRecList orderby data.IDNumber ascending,data.SubjectCode ascending select data).ToList();

            // 填入Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.定期考查成績名冊樣板));
            Worksheet wst1 = _wb.Worksheets["定期考查成績名冊封面"];
            Worksheet wst2 = _wb.Worksheets["定期考查成績名冊"];

            // 學校代碼 0,學年度 1,學期 2,段考別 3,名冊別 4
            wst1.Cells[1, 0].PutValue(_SchoolCode);
            wst1.Cells[1, 1].PutValue(_SchoolYear);
            wst1.Cells[1, 2].PutValue(_Semester);
            wst1.Cells[1, 3].PutValue(_ExamNo);
            wst1.Cells[1, 4].PutValue(_DocType);

            // 身分證號 0,出生日期 1,科目代碼 2,科目學分 3,修課科/班/學程別代碼 4,
            //修課班級 5,修課班別 6,考查成績 7,狀態代碼 8
            int rowIdx = 1;
            foreach(StudentSCETakeRec rec in StudentSCETakeRecList)
            {
                wst2.Cells[rowIdx, 0].PutValue(rec.IDNumber);
                wst2.Cells[rowIdx, 1].PutValue(rec.BirthDate);
                wst2.Cells[rowIdx, 2].PutValue(rec.SubjectCode);
                wst2.Cells[rowIdx, 3].PutValue(rec.SubjectCredit);
                wst2.Cells[rowIdx, 4].PutValue(rec.DCLCode);
                wst2.Cells[rowIdx, 5].PutValue(rec.ClassName);
                wst2.Cells[rowIdx, 6].PutValue(rec.ClClassName);
                wst2.Cells[rowIdx, 7].PutValue(rec.Score);
                wst2.Cells[rowIdx, 8].PutValue(rec.Status);
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
            try
            {
                if (_ExamNameDict.ContainsKey(cboExam.Text))
                    _ExamID = _ExamNameDict[cboExam.Text];
                _SchoolYear = iptSchoolYear.Value;
                _Semester = iptSemester.Value;
                _ExamNo = iptExamNo.Value;
                _SchoolCode = K12.Data.School.Code;

                _cd.ClearKeyValueItem();
                _cd.AddKeyValueItem("學年度", _SchoolYear.ToString());
                _cd.AddKeyValueItem("學期", _Semester.ToString());
                _cd.AddKeyValueItem("段考別", _ExamNo.ToString());
                _cd.AddKeyValueItem("試別", cboExam.Text);
                _cd.SaveKeyValueItem(_ConfigName);

                if (_ExamID != "")
                {
                    btnExport.Enabled = false;
                    _bgWorker.RunWorkerAsync();
                    
                }

            }
            catch(Exception ex)
            {

            }
        }

        private void StudentCourseScoreReport_Load(object sender, EventArgs e)
        {
            // 載入設值
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);
            cboExam.DropDownStyle = ComboBoxStyle.DropDownList;
            _ExamNameDict = Utility.GetExamNameDict();
            foreach (string name in _ExamNameDict.Keys)
                cboExam.Items.Add(name);

            // 讀取預設值
            Dictionary<string, string> ds = _cd.GetKeyValueItem(_ConfigName);
            if (ds.ContainsKey("學年度"))
                if (ds["學年度"] != "")
                    iptSchoolYear.Value = int.Parse(ds["學年度"]);
            if (ds.ContainsKey("學期"))
                if (ds["學期"] != "")
                    iptSemester.Value = int.Parse(ds["學期"]);

            if (ds.ContainsKey("段考別"))
                if (ds["段考別"] != "")
                    iptExamNo.Value = int.Parse(ds["段考別"]);
                        
            if (ds.ContainsKey("試別"))
                cboExam.Text = ds["試別"];
        }

    }
}
