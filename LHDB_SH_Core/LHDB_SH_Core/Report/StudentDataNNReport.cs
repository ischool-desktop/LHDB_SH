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
    public partial class StudentDataNNReport : BaseForm
    {
        private int _SchoolYear = 0, _Semester = 0;
        private string _SchoolCode = "";
        private string _ConfigName = "學生名冊非國教署_畫面設定";
        private string _DepCode = "";
        private string _ClassCode = "";
        private ConfigData _cd;
        // 名冊別
        private string _DocType = "1";

        BackgroundWorker _bgWorker;
        Workbook _wb;
        List<string> _StudentIDList;

        public StudentDataNNReport(List<string> StudentIDList)
        {
            _cd = new ConfigData();
            _bgWorker = new BackgroundWorker();
            _StudentIDList = StudentIDList;
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;
            InitializeComponent();
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學生資料名冊(非國教署主管學校)產生完成");
            btnExport.Enabled = true;
            if (_wb != null)
            {
                Utility.ExprotXls("學生資料名冊(非國教署主管學校)", _wb);
            }
        }

        void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學生資料名冊(非國教署主管學校)產生中..", e.ProgressPercentage);
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

            // 取得學期對照班級座號轉成大學繁星班代
            Dictionary<string, string> StudentSHClassSeatNoDict = Utility.GetStudentClassCodeSeatNo(_StudentIDList, _SchoolYear, _Semester, true);

            // 取得學生科別名稱
            Dictionary<string, string> StudeDeptNameDict = Utility.GetStudDeptNameDict(_StudentIDList, _SchoolYear, _Semester);

            // 取得學生異動資料
            List<SHUpdateRecordRecord> SHUpdateRecordRecordList = SHUpdateRecord.SelectByStudentIDs(_StudentIDList);
            Dictionary<string, List<SHUpdateRecordRecord>> SHUpdateRecordRecordDict = new Dictionary<string, List<SHUpdateRecordRecord>>();
            // 沒有核准日期文號不排入，並符合所選的學年度學期
            foreach(SHUpdateRecordRecord urRec in SHUpdateRecordRecordList)
            {   
                if (string.IsNullOrWhiteSpace(urRec.ADDate) || string.IsNullOrWhiteSpace(urRec.ADNumber))
                    continue;

                if(urRec.SchoolYear.HasValue && urRec.Semester.HasValue)
                {
                    if (urRec.SchoolYear <= _SchoolYear && urRec.Semester <= _Semester)
                    {
                        if (!SHUpdateRecordRecordDict.ContainsKey(urRec.StudentID))
                            SHUpdateRecordRecordDict.Add(urRec.StudentID, new List<SHUpdateRecordRecord>());

                        SHUpdateRecordRecordDict[urRec.StudentID].Add(urRec);
                    }
                }

            }

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

            // 部別對照
            if (cdDict.ContainsKey("部別代碼"))
            {
                foreach (ConfigDataItem cdi in cdDict["部別代碼"])
                {
                    if (!DepMappingDict.ContainsKey(cdi.TargetName))
                        DepMappingDict.Add(cdi.TargetName, cdi.Value);
                }
            }

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

            // 班級代碼對照
            ClassNoMappingDict = Utility.GetLHClassCodeDict();

            _bgWorker.ReportProgress(40);

            // 取得學生本資料
            List<SHStudentRecord> StudentRecordList = SHStudent.SelectByIDs(_StudentIDList);
            List<StudentBaseRec> StudentBaseRecList = new List<StudentBaseRec>();

            // 整理資料
            foreach (SHStudentRecord studRec in StudentRecordList)
            {
                // 身分證號,出生日期,所屬學校代碼,科/班/學程別代碼,部別,班別,班級座號代碼
                StudentBaseRec sbr = new StudentBaseRec();
                sbr.StudentNumber = studRec.StudentNumber;
                sbr.IDNumber = studRec.IDNumber.ToUpper();
                sbr.BirthDate = Utility.ConvertChDateString(studRec.Birthday);
                sbr.SchoolCode = _SchoolCode;
                sbr.Name = studRec.Name;
                if (studRec.Gender == "男")
                    sbr.GenderCode = "1";
                if (studRec.Gender == "女")
                    sbr.GenderCode = "2";

                // 科/班/學程別代碼
                sbr.DCLCode = "";
                if (StudeDeptNameDict.ContainsKey(studRec.ID))
                {
                    string name = StudeDeptNameDict[studRec.ID];
                    if (DeptMappingDict.ContainsKey(name))
                        sbr.DCLCode = DeptMappingDict[name];
                }

                // 部別
                sbr.DepCode = _DepCode;
                // 班別
                sbr.ClCode = _ClassCode;

                if (StudTagNameDict.ContainsKey(studRec.ID))
                {            


                    foreach (string str in StudTagNameDict[studRec.ID])
                    {
                        if (DepMappingDict.ContainsKey(str))
                            sbr.DepCode = DepMappingDict[str];

                        if (ClsMappingDict.ContainsKey(str))
                            sbr.ClCode = ClsMappingDict[str];
                    }

                }

                // 班級座號代碼
                sbr.ClassSeatCode = "";
                if (StudentSHClassSeatNoDict.ContainsKey(studRec.ID))
                {
                    sbr.ClassSeatCode = StudentSHClassSeatNoDict[studRec.ID];
                    if(sbr.ClassSeatCode.Length==5)
                        sbr.ClassCode=sbr.ClassSeatCode.Substring(0,3);
                }
                else
                {
                    //如果學年度學期和系統預設相同
                    if (K12.Data.School.DefaultSchoolYear == _SchoolYear.ToString() && K12.Data.School.DefaultSemester == _Semester.ToString())
                    {
                        if (ClassIDNameDict.ContainsKey(studRec.RefClassID))
                        {
                            string cName = ClassIDNameDict[studRec.RefClassID];
                            if (ClassNoMappingDict.ContainsKey(cName) && studRec.SeatNo.HasValue)
                            {
                                sbr.ClassSeatCode = ClassNoMappingDict[cName] + string.Format("{0:00}", studRec.SeatNo.Value);
                                sbr.ClassCode=ClassNoMappingDict[cName];
                            }
                        }
                    }
                }

                // 處理異動相關
                if(SHUpdateRecordRecordDict.ContainsKey(studRec.ID))
                {
                    // 排序
                    List<SHUpdateRecordRecord> UrDataList = (from data in SHUpdateRecordRecordDict[studRec.ID] orderby DateTime.Parse(data.ADDate) descending, int.Parse(data.ID) descending select data).ToList();

                    if(UrDataList.Count>0)
                    {
                        SHUpdateRecordRecord rec = UrDataList[0];
                        sbr.PermrecCode = rec.UpdateCode;
                        DateTime dt =DateTime.Parse(rec.UpdateDate);
                        sbr.PermrecDate = Utility.ConvertChDateString(dt);

                        sbr.OrBirthDate = "-1";
                        sbr.OrIDNumber = "-1";
                        sbr.Remak1 = rec.IDNumberComment;
                        sbr.SpCode = rec.SpecialStatus;
                    }
                    foreach(SHUpdateRecordRecord rec in UrDataList)
                    {
                        // 再次檢查身分證與生日是否有差異
                        int iUcode = int.Parse(rec.UpdateCode);
                        if(iUcode> 400 && iUcode <411)
                        {
                            if (rec.IDNumber.ToUpper() != studRec.IDNumber.ToUpper())
                                sbr.OrIDNumber = rec.IDNumber.ToUpper();

                            if(studRec.Birthday.HasValue)
                            {
                                if(studRec.Birthday.Value.ToShortDateString()!= rec.Birthdate)
                                {
                                    DateTime dto = DateTime.Parse(rec.Birthdate);
                                    sbr.OrBirthDate = Utility.ConvertChDateString(dto);
                                }
                            }
                        }
                    }
                }
                
                StudentBaseRecList.Add(sbr);
            }
            _bgWorker.ReportProgress(80);

            // 排序 班級座號代碼
            StudentBaseRecList = (from data in StudentBaseRecList orderby data.ClassSeatCode ascending select data).ToList();

            // 統計人數
            Dictionary<string, StudentBaseRecCount> StudentBaseRecCountDict = new Dictionary<string, StudentBaseRecCount>();

            foreach(StudentBaseRec rec in StudentBaseRecList)
            {
                string key = rec.ClCode + rec.DCLCode + rec.DepCode + rec.ClassCode;

                if(!StudentBaseRecCountDict.ContainsKey(key))
                {
                    StudentBaseRecCount sbrc = new StudentBaseRecCount();
                    sbrc.ClassCode = rec.ClassCode;
                    sbrc.ClCode = rec.ClCode;
                    sbrc.DCLCode = rec.DCLCode;
                    sbrc.DepCode = rec.DepCode;
                    sbrc.DocType = _DocType;
                    sbrc.SchoolCode = _SchoolCode;
                    sbrc.SchoolYear = _SchoolYear;
                    sbrc.Semester = _Semester;
                    sbrc.StudentCount = 0;
                    StudentBaseRecCountDict.Add(key, sbrc);
                }
                StudentBaseRecCountDict[key].StudentCount++;
            }

            List<StudentBaseRecCount> StudentBaseRecCountList = new List<StudentBaseRecCount>();
            foreach (string Key in StudentBaseRecCountDict.Keys)
                StudentBaseRecCountList.Add(StudentBaseRecCountDict[Key]);

            // 排序,依班級代碼
            StudentBaseRecCountList = (from data in StudentBaseRecCountList orderby data.ClassCode ascending select data).ToList();
            

            // 填值到 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.學生資料名冊樣板_非國教署主管學校_));
            Worksheet wst1 = _wb.Worksheets["學生資料名冊封面"];
            Worksheet wst2 = _wb.Worksheets["學生資料名冊"];

            int rowIdx = 1;
            foreach(StudentBaseRecCount sbrc in StudentBaseRecCountList)
            {
                //學校代碼 0,學年度 1,學期 2,名冊別 3,科/班/學程別代碼 4,年級班級代碼 5,部別 6,班別 7,班級人數 8
                wst1.Cells[rowIdx, 0].PutValue(sbrc.SchoolCode);
                wst1.Cells[rowIdx, 1].PutValue(sbrc.SchoolYear);
                wst1.Cells[rowIdx, 2].PutValue(sbrc.Semester);
                wst1.Cells[rowIdx, 3].PutValue(sbrc.DocType);
                wst1.Cells[rowIdx, 4].PutValue(sbrc.DCLCode);
                wst1.Cells[rowIdx, 5].PutValue(sbrc.ClassCode);
                wst1.Cells[rowIdx, 6].PutValue(sbrc.DepCode);
                wst1.Cells[rowIdx, 7].PutValue(sbrc.ClCode);
                wst1.Cells[rowIdx, 8].PutValue(sbrc.StudentCount);
                rowIdx++;
            }

            rowIdx = 1;
            foreach (StudentBaseRec sbr in StudentBaseRecList)
            {
                // 學號 0,身分證號 1,註1 2,姓名 3,性別代碼 4,出生日期 5,所屬學校代碼 6,科/班/學程別代碼 7,部別 8,
                //班別 9,班級座號代碼 10,特殊身分代碼 11,原身分證號 12,原出生日期 13,學籍狀態代碼 14,學籍狀態生效日期 15
                wst2.Cells[rowIdx, 0].PutValue(sbr.StudentNumber);
                wst2.Cells[rowIdx, 1].PutValue(sbr.IDNumber);
                wst2.Cells[rowIdx, 2].PutValue(sbr.Remak1);
                wst2.Cells[rowIdx, 3].PutValue(sbr.Name);
                wst2.Cells[rowIdx, 4].PutValue(sbr.GenderCode);
                wst2.Cells[rowIdx, 5].PutValue(sbr.BirthDate);
                wst2.Cells[rowIdx, 6].PutValue(sbr.SchoolCode);
                wst2.Cells[rowIdx, 7].PutValue(sbr.DCLCode);
                wst2.Cells[rowIdx, 8].PutValue(sbr.DepCode);                
                wst2.Cells[rowIdx, 9].PutValue(sbr.ClCode);
                wst2.Cells[rowIdx, 10].PutValue(sbr.ClassSeatCode);
                wst2.Cells[rowIdx, 11].PutValue(sbr.SpCode);
                wst2.Cells[rowIdx, 12].PutValue(sbr.OrIDNumber);
                wst2.Cells[rowIdx, 13].PutValue(sbr.OrBirthDate);
                wst2.Cells[rowIdx, 14].PutValue(sbr.PermrecCode);
                wst2.Cells[rowIdx, 15].PutValue(sbr.PermrecDate);
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
                btnExport.Enabled = false;
                _SchoolYear = iptSchoolYear.Value;
                _Semester = iptSemester.Value;
                _SchoolCode = K12.Data.School.Code;
                _DepCode = iptDepDefault.Value.ToString();
                _ClassCode = iptClassDefault.Value.ToString();

                _cd.ClearKeyValueItem();
                _cd.AddKeyValueItem("學年度", _SchoolYear.ToString());
                _cd.AddKeyValueItem("學期", _Semester.ToString());
                _cd.AddKeyValueItem("部別代碼預設值", _DepCode);
                _cd.AddKeyValueItem("班別代碼預設值", _ClassCode);
                _cd.SaveKeyValueItem(_ConfigName);

                _bgWorker.RunWorkerAsync();

            }
            catch (Exception ex)
            {

            }

        }


        private void StudentDataNReport_Load(object sender, EventArgs e)
        {
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);
            iptClassDefault.Value = iptDepDefault.Value = 1;

            this.MaximumSize = this.MinimumSize = this.Size;

            // 讀取預設值
            Dictionary<string, string> ds = _cd.GetKeyValueItem(_ConfigName);
            if (ds.ContainsKey("學年度"))
                if (ds["學年度"] != "")
                    iptSchoolYear.Value = int.Parse(ds["學年度"]);
            if (ds.ContainsKey("學期"))
                if (ds["學期"] != "")
                    iptSemester.Value = int.Parse(ds["學期"]);

            if (ds.ContainsKey("部別代碼預設值"))
                if (ds["部別代碼預設值"] != "")
                    iptDepDefault.Value = int.Parse(ds["部別代碼預設值"]);

            if (ds.ContainsKey("班別代碼預設值"))
                if (ds["班別代碼預設值"] != "")
                    iptClassDefault.Value = int.Parse(ds["班別代碼預設值"]);
        }

        private void lnkDepSetup_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkDepSetup.Enabled = false;
            Config.DepConfigForm dcf = new Config.DepConfigForm();
            dcf.ShowDialog();
            lnkDepSetup.Enabled = true;
        }

        private void StudentDataNNReport_Load(object sender, EventArgs e)
        {
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);
            iptClassDefault.Value = iptDepDefault.Value = 1;

            this.MaximumSize = this.MinimumSize = this.Size;

            // 讀取預設值
            Dictionary<string, string> ds = _cd.GetKeyValueItem(_ConfigName);
            if (ds.ContainsKey("學年度"))
                if (ds["學年度"] != "")
                    iptSchoolYear.Value = int.Parse(ds["學年度"]);
            if (ds.ContainsKey("學期"))
                if (ds["學期"] != "")
                    iptSemester.Value = int.Parse(ds["學期"]);

            if (ds.ContainsKey("部別代碼預設值"))
                if (ds["部別代碼預設值"] != "")
                    iptDepDefault.Value = int.Parse(ds["部別代碼預設值"]);

            if (ds.ContainsKey("班別代碼預設值"))
                if (ds["班別代碼預設值"] != "")
                    iptClassDefault.Value = int.Parse(ds["班別代碼預設值"]);
            
            // 說明
            txtDesc.Text = Properties.Resources.畫面說明_學生資料名冊;
            txtDesc.ReadOnly = true;
            txtDesc.BackColor = Color.White;
        }
    }
}
