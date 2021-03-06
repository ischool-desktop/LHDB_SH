﻿using System;
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
    public partial class SubjectScoreReport :  BaseForm
    {
        private int _SchoolYear = 0, _Semester = 0;
        private string _SchoolCode = "";
        private string _ClassTypeCode = "";
        private ConfigData _cd;
        private string _ConfigName = "成績名冊_畫面設定";

        // 名冊別
        private string _DocType = "2";

        BackgroundWorker _bgWorker;
        Workbook _wb;
        List<string> _StudentIDList;


        public SubjectScoreReport(List<string> StudentIDList)
        {
            _cd = new ConfigData();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _StudentIDList = StudentIDList;
            InitializeComponent();
        }

        void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("成績名冊產生中..", e.ProgressPercentage);
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("成績名冊產生完成");
            btnExport.Enabled = true;
            try
            {
                if (_wb != null)
                {
                    Utility.ExprotXls("成績名冊", _wb);
                }
            }
            catch (Exception ex)
            { }
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);

            List<SubjectScoreRec> SubjectScoreRecList = new List<SubjectScoreRec>();
            Dictionary<string, string> DeptMappingDict = new Dictionary<string, string>();
            Dictionary<string, List<string>> StudTagNameDict = new Dictionary<string, List<string>>();
            
            // 科別對照
            DeptMappingDict = Utility.GetDepartmetDict();

            Dictionary<string, List<ConfigDataItem>> cdDict = _cd.GetConfigDataItemDict();
            Dictionary<string, string> ClsMappingDict = new Dictionary<string, string>();
            Dictionary<string, string> ClassNoMappingDict = new Dictionary<string, string>();
            List<SHStudentTagRecord> SHStudentTagRecordList = SHStudentTag.SelectByStudentIDs(_StudentIDList);

            // 取得學期對照轉成大學繁星代碼
            Dictionary<string, string> StudentSHClassCodDict = Utility.GetStudentClassCodeSeatNo(_StudentIDList, _SchoolYear, _Semester, false);

            // 班別對照
            if (cdDict.ContainsKey("班別代碼"))
            {
                foreach (ConfigDataItem cdi in cdDict["班別代碼"])
                {
                    if (!ClsMappingDict.ContainsKey(cdi.TargetName))
                        ClsMappingDict.Add(cdi.TargetName, cdi.Value);
                }
            }


            // 取得學生類別
            foreach (SHStudentTagRecord TRec in SHStudentTagRecordList)
            {
                if (!StudTagNameDict.ContainsKey(TRec.RefStudentID))
                    StudTagNameDict.Add(TRec.RefStudentID, new List<string>());

                StudTagNameDict[TRec.RefStudentID].Add(TRec.FullName);
            }

            // 班級代碼對照
            ClassNoMappingDict = Utility.GetLHClassCodeDict();

            _bgWorker.ReportProgress(10);

            // 取得學生科別名稱
            Dictionary<string, string> StudeDeptNameDict = Utility.GetStudDeptNameDict(_StudentIDList,_SchoolYear,_Semester);

            SmartSchool.Customization.Data.AccessHelper accHelper = new SmartSchool.Customization.Data.AccessHelper();
                 // 取得學生學期科目成績
            List<SmartSchool.Customization.Data.StudentRecord> StudentRecList = accHelper.StudentHelper.GetStudents(_StudentIDList);
            accHelper.StudentHelper.FillSemesterSubjectScore(true, StudentRecList);

            // 取得及格標準
            Dictionary<string, Dictionary<string, decimal>> passScoreDict = Utility.GetStudentApplyLimitDict(StudentRecList);
            Dictionary<string, string> StudGradYearDict = new Dictionary<string, string>();

            foreach (SmartSchool.Customization.Data.StudentRecord rec in StudentRecList)
            {
                if (!StudGradYearDict.ContainsKey(rec.StudentID))
                    if(rec.RefClass!=null)
                        StudGradYearDict.Add(rec.StudentID, rec.RefClass.GradeYear);
            }

            // 取得學期對照年為主
            List<SHSemesterHistoryRecord> SemsH = SHSemesterHistory.SelectByStudentIDs(_StudentIDList);
            foreach (SHSemesterHistoryRecord rec in SemsH)
            {
                foreach (K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                {
                    if (item.SchoolYear == _SchoolYear && item.Semester == _Semester)
                    {
                        if (!StudGradYearDict.ContainsKey(item.RefStudentID))
                            StudGradYearDict.Add(item.RefStudentID, item.GradeYear.ToString());
                        else
                            StudGradYearDict[item.RefStudentID] = item.GradeYear.ToString();
                    }
                }
            }

            _bgWorker.ReportProgress(30);
            
            foreach(SmartSchool.Customization.Data.StudentRecord studRec in StudentRecList)
            {
                string IDNumber = studRec.IDNumber.ToUpper();
                string BirthDate = "";
                DateTime dt;
                if(DateTime.TryParse(studRec.Birthday,out dt))
                    BirthDate = Utility.ConvertChDateString(dt);


                // 科/班/學程別代碼
                string DCLCode = "";
                if (StudeDeptNameDict.ContainsKey(studRec.StudentID))
                {
                    string name = StudeDeptNameDict[studRec.StudentID];
                    if (DeptMappingDict.ContainsKey(name))
                        DCLCode = DeptMappingDict[name];
                }


                // 修課班別
                string ClClassName = _ClassTypeCode;
                if (StudTagNameDict.ContainsKey(studRec.StudentID))
                {
                    foreach (string str in StudTagNameDict[studRec.StudentID])
                    {
                        if (ClsMappingDict.ContainsKey(str))
                            ClClassName = ClsMappingDict[str];
                    }
                }

                // 修課班級
                string ClassCode = "";
                if(StudentSHClassCodDict.ContainsKey(studRec.StudentID))
                {
                    ClassCode = StudentSHClassCodDict[studRec.StudentID];
                }
                else
                {
                    if (K12.Data.School.DefaultSchoolYear == _SchoolYear.ToString() && K12.Data.School.DefaultSemester == _Semester.ToString())
                    {
                        if (ClassNoMappingDict.ContainsKey(studRec.RefClass.ClassName))
                            ClassCode = ClassNoMappingDict[studRec.RefClass.ClassName];
                    }
                }


                #region 一般與補修
                foreach (SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo sssi in studRec.SemesterSubjectScoreList)
                {
                    if (sssi.SchoolYear == _SchoolYear && sssi.Semester == _Semester)
                    {
                        SubjectScoreRec ssr = new SubjectScoreRec();
                        ssr.IDNumber = IDNumber;
                        ssr.StudentID = studRec.StudentID;
                        ssr.BirthDate = BirthDate;
                        ssr.SubjectCode = sssi.Detail.GetAttribute("科目") + "_" + sssi.Detail.GetAttribute("開課學分數") + "_" + sssi.Detail.GetAttribute("修課必選修") + "_" + sssi.Detail.GetAttribute("修課校部訂") + "_" + sssi.Detail.GetAttribute("開課分項類別") + "_" + sssi.Detail.GetAttribute("不計學分");
                        ssr.ReSubjectCode = sssi.Detail.GetAttribute("科目") + "_" + sssi.Detail.GetAttribute("開課學分數") + "_" + sssi.Detail.GetAttribute("修課必選修") + "_" + sssi.Detail.GetAttribute("修課校部訂") + "_" + sssi.Detail.GetAttribute("開課分項類別") + "_" + sssi.Detail.GetAttribute("不計學分");
                        ssr.ScSubjectCode = sssi.Detail.GetAttribute("科目") + "_" + sssi.Detail.GetAttribute("開課學分數") + "_" + sssi.Detail.GetAttribute("修課必選修") + "_" + sssi.Detail.GetAttribute("修課校部訂") + "_" + sssi.Detail.GetAttribute("開課分項類別") + "_" + sssi.Detail.GetAttribute("不計學分");

                        ssr.SubjectCredit = sssi.Detail.GetAttribute("開課學分數");
                        // 預設值 -1
                        ssr.Score = ssr.ScScore = ssr.ReScore = ssr.MuScore = "-1";


                        string GrStr = "";
                        if (StudGradYearDict.ContainsKey(studRec.StudentID))
                            GrStr = StudGradYearDict[studRec.StudentID] + "_及";

                        decimal ds, dsre, dsmu, dssc, passScore = 60;

                        // 及格標準
                        if (passScoreDict[studRec.StudentID].ContainsKey(GrStr))
                            passScore = passScoreDict[studRec.StudentID][GrStr];

                        if (decimal.TryParse(sssi.Detail.GetAttribute("原始成績"), out ds))
                        {
                            if (ds < passScore)
                                ssr.Score = "*" + string.Format("{0:##0.00}", ds);
                            else
                                ssr.Score = string.Format("{0:##0.00}", ds);
                        }

                        if (decimal.TryParse(sssi.Detail.GetAttribute("重修成績"), out dsre))
                        {
                            if (dsre < passScore)
                                ssr.ReScore = "*" + string.Format("{0:##0.00}", dsre);
                            else
                                ssr.ReScore = string.Format("{0:##0.00}", dsre);
                        }

                        if (decimal.TryParse(sssi.Detail.GetAttribute("補考成績"), out dsmu))
                        {
                            if (dsmu < passScore)
                                ssr.MuScore = "*" + string.Format("{0:##0.00}", dsmu);
                            else
                                ssr.MuScore = string.Format("{0:##0.00}", dsmu);
                        }

                        if (decimal.TryParse(sssi.Detail.GetAttribute("原始成績"), out dssc))
                        {
                            if (dssc < passScore)
                                ssr.ScScore = "*" + string.Format("{0:##0.00}", dssc);
                            else
                                ssr.ScScore = string.Format("{0:##0.00}", dssc);
                        }


                        ssr.isScScore = ssr.isReScore = false;


                        if (sssi.Detail.GetAttribute("是否補修成績") == "是")
                        {
                            ssr.isScScore = true;
                            ssr.ScSubjectGradeYearSemester = sssi.GradeYear.ToString() + sssi.Semester.ToString();
                        }

                        ssr.DCLCode = DCLCode;
                        ssr.ClassName = ClassCode;
                        ssr.ClCode = ClClassName;

                        SubjectScoreRecList.Add(ssr);
                    }
                }

                #endregion

                #region 重修處理
                foreach (SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo sssi in studRec.SemesterSubjectScoreList)
                {
                    int chkSy = 0, chkSS = 0;

                    if (sssi.Detail.GetAttribute("重修學年度") != "" && sssi.Detail.GetAttribute("重修學期") != "")
                    {
                        int sy, ss;
                        if (int.TryParse(sssi.Detail.GetAttribute("重修學年度"), out sy))
                            chkSy = sy;

                        if (int.TryParse(sssi.Detail.GetAttribute("重修學期"), out ss))
                            chkSS = ss;

                        
                    }

                    // 以有填寫重修學年度學期為主
                    if (chkSy == _SchoolYear && chkSS == _Semester)
                    {
                        SubjectScoreRec ssr = new SubjectScoreRec();
                        ssr.IDNumber = IDNumber;
                        ssr.StudentID = studRec.StudentID;
                        ssr.BirthDate = BirthDate;
                        ssr.SubjectCode = sssi.Detail.GetAttribute("科目") + "_" + sssi.Detail.GetAttribute("開課學分數") + "_" + sssi.Detail.GetAttribute("修課必選修") + "_" + sssi.Detail.GetAttribute("修課校部訂") + "_" + sssi.Detail.GetAttribute("開課分項類別") + "_" + sssi.Detail.GetAttribute("不計學分");
                        ssr.ReSubjectCode = sssi.Detail.GetAttribute("科目") + "_" + sssi.Detail.GetAttribute("開課學分數") + "_" + sssi.Detail.GetAttribute("修課必選修") + "_" + sssi.Detail.GetAttribute("修課校部訂") + "_" + sssi.Detail.GetAttribute("開課分項類別") + "_" + sssi.Detail.GetAttribute("不計學分");
                        ssr.ScSubjectCode = sssi.Detail.GetAttribute("科目") + "_" + sssi.Detail.GetAttribute("開課學分數") + "_" + sssi.Detail.GetAttribute("修課必選修") + "_" + sssi.Detail.GetAttribute("修課校部訂") + "_" + sssi.Detail.GetAttribute("開課分項類別") + "_" + sssi.Detail.GetAttribute("不計學分");

                        ssr.SubjectCredit = sssi.Detail.GetAttribute("開課學分數");
                        // 預設值 -1
                        ssr.Score = ssr.ScScore = ssr.ReScore = ssr.MuScore = "-1";

                        string GrStr = "";
                        if (StudGradYearDict.ContainsKey(studRec.StudentID))
                            GrStr = StudGradYearDict[studRec.StudentID] + "_及";

                        decimal ds, dsre, dsmu, dssc, passScore = 60;

                        // 及格標準
                        if (passScoreDict[studRec.StudentID].ContainsKey(GrStr))
                            passScore = passScoreDict[studRec.StudentID][GrStr];

                        if (decimal.TryParse(sssi.Detail.GetAttribute("原始成績"), out ds))
                        {
                            if (ds < passScore)
                                ssr.Score = "*" + string.Format("{0:##0.00}", ds);
                            else
                                ssr.Score = string.Format("{0:##0.00}", ds);
                        }

                        if (decimal.TryParse(sssi.Detail.GetAttribute("重修成績"), out dsre))
                        {
                            if (dsre < passScore)
                                ssr.ReScore = "*" + string.Format("{0:##0.00}", dsre);
                            else
                                ssr.ReScore = string.Format("{0:##0.00}", dsre);
                        }

                        if (decimal.TryParse(sssi.Detail.GetAttribute("補考成績"), out dsmu))
                        {
                            if (dsmu < passScore)
                                ssr.MuScore = "*" + string.Format("{0:##0.00}", dsmu);
                            else
                                ssr.MuScore = string.Format("{0:##0.00}", dsmu);
                        }

                        if (decimal.TryParse(sssi.Detail.GetAttribute("原始成績"), out dssc))
                        {
                            if (dssc < passScore)
                                ssr.ScScore = "*" + string.Format("{0:##0.00}", dssc);
                            else
                                ssr.ScScore = string.Format("{0:##0.00}", dssc);
                        }

                        ssr.isScScore = ssr.isReScore = false;

                        // 使用學年度為該學年度學期
                        ssr.ReSubjectSchoolYearSemester = string.Format("{0:000}", sssi.SchoolYear) + sssi.Semester;
                        ssr.isReScore = true;


                        ssr.DCLCode = DCLCode;
                        ssr.ClassName = ClassCode;
                        ssr.ClCode = ClClassName;

                        SubjectScoreRecList.Add(ssr);
                    }
                }
                #endregion

                
            }

            _bgWorker.ReportProgress(80);
            
            // 讀取樣版並寫入 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.成績名冊樣版));
            Worksheet wst1 = _wb.Worksheets["成績名冊封面"];
            Worksheet wst2 = _wb.Worksheets["成績名冊"];
            Worksheet wst3 = _wb.Worksheets["補修成績名冊"];
            Worksheet wst4 = _wb.Worksheets["重修成績名冊"];

            // wst1:學校代碼 0,學年度 1,學期 2,名冊別 3
            wst1.Cells[1, 0].PutValue(_SchoolCode);
            wst1.Cells[1, 1].PutValue(_SchoolYear);
            wst1.Cells[1, 2].PutValue(_Semester);
            wst1.Cells[1, 3].PutValue(_DocType);
            
            // wst2:身分證號 0,出生日期 1,科目代碼 2,科目學分 3,修課科/班/學程別代碼 4,修課班級 5,修課班別 6,原始成績 7,補考成績 8
            int rowIdx = 1;
            foreach(SubjectScoreRec ssr in SubjectScoreRecList)
            {
                // 跳過補修
                if (ssr.isScScore )
                    continue;

                wst2.Cells[rowIdx, 0].PutValue(ssr.IDNumber);
                wst2.Cells[rowIdx, 1].PutValue(ssr.BirthDate);
                wst2.Cells[rowIdx, 2].PutValue(ssr.SubjectCode);
                wst2.Cells[rowIdx, 3].PutValue(ssr.SubjectCredit);
                wst2.Cells[rowIdx, 4].PutValue(ssr.DCLCode);
                wst2.Cells[rowIdx, 5].PutValue(ssr.ClassName);
                wst2.Cells[rowIdx, 6].PutValue(ssr.ClCode);
                wst2.Cells[rowIdx, 7].PutValue(ssr.Score);
                wst2.Cells[rowIdx, 8].PutValue(ssr.MuScore);
                rowIdx++;
            }            

            // wst3:身分證號 0,出生日期 1,補修科目代碼 2,補修科目開設年級及學期(選填) 3,科目學分 4,修課科/班/學程別代碼 5,修課班級 6,修課班別 7,補修成績 8,補考成績 9
            rowIdx = 1;
            foreach(SubjectScoreRec ssr in SubjectScoreRecList)
            {
                // 補修成績
                if(ssr.isScScore)
                {
                    wst3.Cells[rowIdx, 0].PutValue(ssr.IDNumber);
                    wst3.Cells[rowIdx, 1].PutValue(ssr.BirthDate);
                    wst3.Cells[rowIdx, 2].PutValue(ssr.ScSubjectCode);
                    wst3.Cells[rowIdx, 3].PutValue(ssr.ScSubjectGradeYearSemester);
                    wst3.Cells[rowIdx, 4].PutValue(ssr.SubjectCredit);
                    wst3.Cells[rowIdx, 5].PutValue(ssr.DCLCode);
                    wst3.Cells[rowIdx, 6].PutValue(ssr.ClassName);
                    wst3.Cells[rowIdx, 7].PutValue(ssr.ClCode);
                    wst3.Cells[rowIdx, 8].PutValue(ssr.ScScore);
                    wst3.Cells[rowIdx, 9].PutValue(ssr.MuScore);
                    rowIdx++;
                }
            }

            // wst4:身分證號 0,出生日期 1,重修科目代碼 2,重修科目開設年級及學期 3,科目學分 4,修課科/班/學程別代碼 5,修課班級 6,修課班別 7,重修成績 8
            rowIdx = 1;
            foreach (SubjectScoreRec ssr in SubjectScoreRecList)
            {
                // 有重修成績學年度學期
                if(ssr.isReScore)
                {
                    wst4.Cells[rowIdx, 0].PutValue(ssr.IDNumber);
                    wst4.Cells[rowIdx, 1].PutValue(ssr.BirthDate);
                    wst4.Cells[rowIdx, 2].PutValue(ssr.ReSubjectCode);
                    wst4.Cells[rowIdx, 3].PutValue(ssr.ReSubjectSchoolYearSemester);
                    wst4.Cells[rowIdx, 4].PutValue(ssr.SubjectCredit);
                    wst4.Cells[rowIdx, 5].PutValue(ssr.DCLCode);
                    wst4.Cells[rowIdx, 6].PutValue(ssr.ClassName);
                    wst4.Cells[rowIdx, 7].PutValue(ssr.ClCode);
                    wst4.Cells[rowIdx, 8].PutValue(ssr.ReScore);                    
                    rowIdx++;
                }
            }

            _bgWorker.ReportProgress(100);

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SubjectScoreReport_Load(object sender, EventArgs e)
        {
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);

            Dictionary<string, string> ds = _cd.GetKeyValueItem(_ConfigName);
            if (ds.ContainsKey("學年度"))
                if (ds["學年度"] != "")
                    iptSchoolYear.Value = int.Parse(ds["學年度"]);
            if (ds.ContainsKey("學期"))
                if (ds["學期"] != "")
                    iptSemester.Value = int.Parse(ds["學期"]);

            Dictionary<string, string> ds1 = _cd.GetKeyValueItem("學生名冊_畫面設定");

            if (ds1.ContainsKey("班別代碼預設值"))
                if (ds1["班別代碼預設值"] != "")
                    _ClassTypeCode =  ds1["班別代碼預設值"];

            // 載入說明
            txtDesc.Text = Properties.Resources.畫面說明_成績名冊;
            txtDesc.ReadOnly = true;
            txtDesc.BackColor = Color.White;

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            btnExport.Enabled = false;
            _SchoolCode = K12.Data.School.Code;
            _SchoolYear = iptSchoolYear.Value;
            _Semester = iptSemester.Value;
            _cd.ClearKeyValueItem();
            _cd.AddKeyValueItem("學年度", _SchoolYear.ToString());
            _cd.AddKeyValueItem("學期", _Semester.ToString());
            _cd.SaveKeyValueItem(_ConfigName);
      
            _bgWorker.RunWorkerAsync();

        }
    }
}
