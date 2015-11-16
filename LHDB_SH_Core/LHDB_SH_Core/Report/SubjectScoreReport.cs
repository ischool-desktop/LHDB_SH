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
            // 科別對照
            DeptMappingDict = Utility.GetDepartmetDict();

            // 取得學生科別名稱
            Dictionary<string, string> StudeDeptNameDict = Utility.GetStudDeptNameDict(_StudentIDList);

            SmartSchool.Customization.Data.AccessHelper accHelper = new SmartSchool.Customization.Data.AccessHelper();
                 // 取得學生學期科目成績
            List<SmartSchool.Customization.Data.StudentRecord> StudentRecList = accHelper.StudentHelper.GetStudents(_StudentIDList);
            accHelper.StudentHelper.FillSemesterSubjectScore(true, StudentRecList);
            foreach(SmartSchool.Customization.Data.StudentRecord studRec in StudentRecList)
            {
                string IDNumber = studRec.IDNumber.ToUpper();
                string BirthDate = "";
                DateTime dt;
                if(DateTime.TryParse(studRec.Birthday,out dt))
                    BirthDate = Utility.ConvertChDateString(dt);
                
                foreach (SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo sssi in studRec.SemesterSubjectScoreList)
                {
                    if(sssi.SchoolYear==_SchoolYear && sssi.Semester == _Semester)
                    {
                        SubjectScoreRec ssr = new SubjectScoreRec();
                        ssr.IDNumber = IDNumber;
                        ssr.StudentID = studRec.StudentID;
                        ssr.BirthDate = BirthDate;
                        ssr.SubjectCode = sssi.Subject;
                        ssr.SubjectCredit = sssi.Detail.GetAttribute("開課學分數");
                        ssr.Score = sssi.Detail.GetAttribute("原始成績");

                        // 科/班/學程別代碼
                        string DCLCode = "";                        
                        if (StudeDeptNameDict.ContainsKey(studRec.StudentID))
                        {
                            string name = StudeDeptNameDict[studRec.StudentID];
                            if (DeptMappingDict.ContainsKey(name))
                                DCLCode = DeptMappingDict[name];
                        }


                        SubjectScoreRecList.Add(ssr);
                    }
                }
            }




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
                wst2.Cells[rowIdx, 0].PutValue(ssr.IDNumber);
                wst2.Cells[rowIdx, 1].PutValue(ssr.BirthDate);
                wst2.Cells[rowIdx, 2].PutValue(ssr.SubjectCode);
                wst2.Cells[rowIdx, 3].PutValue(ssr.SubjectCredit);
                wst2.Cells[rowIdx, 4].PutValue(ssr.DCLCode);
                wst2.Cells[rowIdx, 5].PutValue(ssr.ClassName);
                wst2.Cells[rowIdx, 6].PutValue(ssr.ClCode);
                wst2.Cells[rowIdx, 7].PutValue(ssr.Score);
                wst2.Cells[rowIdx, 8].PutValue(ssr.ScScore);
                rowIdx++;
            }
            

            // wst3:身分證號 0,出生日期 1,補修科目代碼 2,補修科目開設年級及學期(選填) 3,科目學分 4,修課科/班/學程別代碼 5,修課班級 6,修課班別 7,補修成績 8,補考成績 9


            // wst4:身分證號 0,出生日期 1,重修科目代碼 2,重修科目開設年級及學期 3,科目學分 4,修課科/班/學程別代碼 5,修課班級 6,修課班別 7,重修成績 8

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SubjectScoreReport_Load(object sender, EventArgs e)
        {
            _SchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            _Semester = int.Parse(K12.Data.School.DefaultSemester);

            Dictionary<string, string> ds = _cd.GetKeyValueItem(_ConfigName);
            if (ds.ContainsKey("學年度"))
                if (ds["學年度"] != "")
                    iptSchoolYear.Value = int.Parse(ds["學年度"]);
            if (ds.ContainsKey("學期"))
                if (ds["學期"] != "")
                    iptSemester.Value = int.Parse(ds["學期"]);
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
