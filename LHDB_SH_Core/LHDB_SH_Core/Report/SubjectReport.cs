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
    public partial class SubjectReport : BaseForm
    {
        private int _SchoolYear = 0, _Semester = 0;
        private string _SchoolCode = "";

        private string _ConfigName = "科目名冊_畫面設定";
        string _ConfigClassName = "特色班實驗班名稱對照";

        // 名冊別,學校種類
        private string _DocType = "4",_SchoolType="";

        BackgroundWorker _bgWorker;

        Dictionary<string, string> _SchoolTypeDict;
        ConfigData _cd;
        Workbook _wb;
        List<string> _StudentIDList;

        public SubjectReport(List<string> StudentIDList)
        {
            _cd = new ConfigData();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.WorkerReportsProgress = true;
            _SchoolTypeDict = GetSchoolTypeDict();
            _SchoolCode = K12.Data.School.Code;
            _StudentIDList = StudentIDList;
            InitializeComponent();
        }

        void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("科目名冊產生中..", e.ProgressPercentage);
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            
            FISCA.Presentation.MotherForm.SetStatusBarMessage("科目名冊產生完成");
            try
            {
                if(_wb !=null)
                {
                    Utility.ExprotXls("科目名冊", _wb);
                }
            }catch(Exception ex)
            { }
            btnExport.Enabled = true;
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);

            Dictionary<string, string> DeptMappingDict = new Dictionary<string, string>();
            // 取得科群對照
            Dictionary<string, string> GroupIDDict = Utility.GetGroupDeptIDDict();

            // 科別對照
            DeptMappingDict = Utility.GetDepartmetDict();

            
            // 取得學生學期對照班級
            Dictionary<string, string> classNameDict = new Dictionary<string, string>();
            List<SHSemesterHistoryRecord> SHSemesterHistoryRecordList = SHSemesterHistory.SelectByStudentIDs(_StudentIDList);
            foreach(SHSemesterHistoryRecord rec in SHSemesterHistoryRecordList)
            {
                foreach(K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                {
                    if(item.SchoolYear==_SchoolYear && item.Semester == _Semester)
                    {
                        if (!classNameDict.ContainsKey(rec.RefStudentID))
                            classNameDict.Add(rec.RefStudentID, item.ClassName);
                    }
                }
            }

            // 取得特色班對照
            Dictionary<string, string> SpecClassNameDict = _cd.GetKeyValueItem(_ConfigClassName);

            // 取得學生科別名稱
            Dictionary<string, string> StudeDeptNameDict = Utility.GetStudDeptNameDict(_StudentIDList,_SchoolYear,_Semester);

            Dictionary<string, SubjectRec> SubjectRecDict = new Dictionary<string, SubjectRec>();
            SmartSchool.Customization.Data.AccessHelper accHelper = new SmartSchool.Customization.Data.AccessHelper ();
            // 取得學生學期科目成績
            List<SmartSchool.Customization.Data.StudentRecord> StudentRecList = accHelper.StudentHelper.GetStudents(_StudentIDList);
            accHelper.StudentHelper.FillSemesterSubjectScore(true, StudentRecList);
            foreach(SmartSchool.Customization.Data.StudentRecord studRec in StudentRecList)
            {

                string studClassName = "";
                if(studRec.RefClass !=null)
                    studClassName = studRec.RefClass.ClassName;

                if (classNameDict.ContainsKey(studRec.StudentID))
                    studClassName = classNameDict[studRec.StudentID];

                // 科別
                string DeptName = "";
                if (StudeDeptNameDict.ContainsKey(studRec.StudentID))
                    DeptName = StudeDeptNameDict[studRec.StudentID];               

                foreach(SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo sssi in studRec.SemesterSubjectScoreList)
                {
                    if(sssi.SchoolYear==_SchoolYear && sssi.Semester== _Semester)
                    {
                        
                        // 需要加入科別，上傳測試過程科別也是key
                        string SubjCode = sssi.Detail.GetAttribute("科目") + "_" + sssi.Detail.GetAttribute("開課學分數") + "_" + sssi.Detail.GetAttribute("修課必選修") + "_" + sssi.Detail.GetAttribute("修課校部訂") + "_" + sssi.Detail.GetAttribute("開課分項類別") + "_" + sssi.Detail.GetAttribute("不計學分");
                        string key = DeptName + SubjCode;

                        if (!SubjectRecDict.ContainsKey(key))
                        {
                            SubjectRec sr = new SubjectRec();

                            // 科/班/學程別代碼
                            string DCLCode = "";
                            string GroupID = "";
                            if (StudeDeptNameDict.ContainsKey(studRec.StudentID))
                            {
                                string name = StudeDeptNameDict[studRec.StudentID];
                                if (DeptMappingDict.ContainsKey(name))
                                    DCLCode = DeptMappingDict[name];

                                if (GroupIDDict.ContainsKey(DCLCode))
                                    GroupID = GroupIDDict[DCLCode];
                            }
                            sr.Code = SubjCode;
                            sr.Name = sssi.Detail.GetAttribute("科目");
                            sr.CourseType = sssi.Detail.GetAttribute("修課校部訂") + sssi.Detail.GetAttribute("開課分項類別") + sssi.Detail.GetAttribute("修課必選修");
                            sr.DLCCode = DCLCode;
                            sr.DomainType = "";
                            sr.Group = GroupID;
                            sr.isCalc = sssi.Detail.GetAttribute("不計學分");
                            sr.Required = sssi.Detail.GetAttribute("修課必選修");
                            sr.SpcType = "";
                            if (SpecClassNameDict.ContainsKey(studClassName))
                                sr.SpcType = SpecClassNameDict[studClassName];

                            SubjectRecDict.Add(key, sr);
                        }

                        if (sssi.GradeYear == 1)
                            SubjectRecDict[key].CreditG1 = sssi.Detail.GetAttribute("開課學分數");

                        if (sssi.GradeYear == 2)
                            SubjectRecDict[key].CreditG2 = sssi.Detail.GetAttribute("開課學分數");

                        if (sssi.GradeYear == 3)
                            SubjectRecDict[key].CreditG3 = sssi.Detail.GetAttribute("開課學分數");
                        
                    }
                }
            }
            _bgWorker.ReportProgress(70);

            Dictionary<string, string> ClassTypeDict = GetClassTypeDict();
            List<SubjectRec> SubjectRecList = new List<SubjectRec>();
            foreach (SubjectRec rec in SubjectRecDict.Values)
            {
                // 資料轉換
                if (rec.Required == "必修")
                    rec.Required = "1";
                else if (rec.Required == "選修")
                    rec.Required = "2";
                else
                    rec.Required = "3";



                if (rec.isCalc == "否")
                    rec.isCalc = "1";
                else if (rec.isCalc == "是")
                    rec.isCalc = "2";
                else
                    rec.isCalc = "3";

                if(!string.IsNullOrEmpty(rec.CourseType))
                {
                    if (ClassTypeDict.ContainsKey(rec.CourseType))
                    {
                        rec.CourseType = ClassTypeDict[rec.CourseType];
                    }
                    else
                        rec.CourseType = "13";
                }

                SubjectRecList.Add(rec);
            }
            // 寫入 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.科目名冊樣板));
            Worksheet wst1 = _wb.Worksheets["科目名冊封面"];
            Worksheet wst2 = _wb.Worksheets["科目名冊"];

            // 學年度 0,學期 1,學校代碼 2,學校種類 3,名冊別 4
            wst1.Cells[1, 0].PutValue(_SchoolYear);
            wst1.Cells[1, 1].PutValue(_Semester);
            wst1.Cells[1, 2].PutValue(_SchoolCode);
            wst1.Cells[1, 3].PutValue(_SchoolType);
            wst1.Cells[1, 4].PutValue(_DocType);

            // 修課別 0,類別種類 1,領域分類 2,群別 3,科別 4,特色班/實驗班名稱 5,科目名稱 6,科目代碼 7,
            // 一年級學分 8,二年級學分 9,三年級學分 10,是否計算學分 11
            int rowIdx = 1;
            foreach(SubjectRec rec in SubjectRecList)
            {
                wst2.Cells[rowIdx, 0].PutValue(rec.Required);
                wst2.Cells[rowIdx, 1].PutValue(rec.CourseType);
                wst2.Cells[rowIdx, 2].PutValue(rec.DomainType);
                wst2.Cells[rowIdx, 3].PutValue(rec.Group);
                wst2.Cells[rowIdx, 4].PutValue(rec.DLCCode);
                wst2.Cells[rowIdx, 5].PutValue(rec.SpcType);
                wst2.Cells[rowIdx, 6].PutValue(rec.Name);
                wst2.Cells[rowIdx, 7].PutValue(rec.Code);
                wst2.Cells[rowIdx, 8].PutValue(rec.CreditG1);
                wst2.Cells[rowIdx, 9].PutValue(rec.CreditG2);
                wst2.Cells[rowIdx, 10].PutValue(rec.CreditG3);
                wst2.Cells[rowIdx, 11].PutValue(rec.isCalc);
                rowIdx++;
            }
            
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

                _cd.ClearKeyValueItem();
                _cd.AddKeyValueItem("學年度",_SchoolYear.ToString());
                _cd.AddKeyValueItem("學期", _Semester.ToString());
                
                List<string> ttList = new List<string>();
                List<string> ttNList = new List<string>();
                foreach (ListViewItem lvi in lvScType.CheckedItems)
                {
                    if (_SchoolTypeDict.ContainsKey(lvi.Text))
                        ttList.Add(_SchoolTypeDict[lvi.Text]);

                    ttNList.Add(lvi.Text);
                }

                _SchoolType = string.Join(",", ttList.ToArray());

                _cd.AddKeyValueItem("學校種類", string.Join(",", ttNList.ToArray()));
                _cd.SaveKeyValueItem(_ConfigName);

                _bgWorker.RunWorkerAsync();

            }catch(Exception ex)
            {

            }

            
        }

        private void SubjectReport_Load(object sender, EventArgs e)
        {
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);


            // 傳入預設值
            foreach (string key in _SchoolTypeDict.Keys)
                lvScType.Items.Add(key);


            // 讀取預設值
            Dictionary<string, string> ds = _cd.GetKeyValueItem(_ConfigName);
            if (ds.ContainsKey("學年度"))
                if (ds["學年度"] != "")
                    iptSchoolYear.Value = int.Parse(ds["學年度"]);
            if (ds.ContainsKey("學期"))
                if (ds["學期"] != "")
                    iptSemester.Value = int.Parse(ds["學期"]);

            if (ds.ContainsKey("學校種類"))
            {
                List<string> its = ds["學校種類"].Split(',').ToList();
                foreach (ListViewItem lvi in lvScType.Items)
                {
                    if (its.Contains(lvi.Text))
                        lvi.Checked = true;
                }
            }

            txtDesc.Text = Properties.Resources.畫面說明_科目名冊;
        }


        /// <summary>
        /// 取得學校種類
        /// </summary>
        /// <returns></returns>
        private Dictionary<string,string> GetSchoolTypeDict()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("普通高中", "1");
            value.Add("完全中學", "2");
            value.Add("高職", "3");
            value.Add("高中附設職業類科", "4");
            value.Add("高職附設普通科", "5");
            value.Add("進修學校(高中)", "6");
            value.Add("進修學校(高職)", "7");
            value.Add("特教學校", "8");
            return value;
        }

        /// <summary>
        /// 取得類別種類對照
        /// </summary>
        /// <returns></returns>
        private Dictionary<string,string> GetClassTypeDict()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("部訂學業必修", "1");
            value.Add("部訂實習科目必修", "2");
            value.Add("部訂專業科目必修", "3");
            value.Add("部訂學業選修", "4");
            value.Add("部訂活動必修", "5");
            value.Add("校訂學業必修", "6");
            value.Add("校訂學業選修", "7");
            value.Add("校訂專業科目必修", "8");
            value.Add("校訂專業科目選修", "9");
            value.Add("校訂實習科目必修", "10");
            value.Add("校訂實習科目選修", "11");
            value.Add("校定專精選修", "12");

            return value;
        }

        private void lnSpecClassName_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnSpecClassName.Enabled = false;
            Config.SpecClassNameConfigForm sc = new Config.SpecClassNameConfigForm();
            sc.ShowDialog();
            lnSpecClassName.Enabled = true;
        }
    }
}
