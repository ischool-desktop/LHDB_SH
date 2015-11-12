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
        
        // 名冊別
        private string _DocType = "5";

        BackgroundWorker _bgWorker;

        Dictionary<string, string> _ExamNameDict;
        string _ExamID = "";

        Workbook _wb;
        List<string> _StudentIDList;

        public StudentSCETakeScoreReport(List<string> StudentIDList)
        {
            _ExamNameDict = new Dictionary<string, string>();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.WorkerReportsProgress = true;
            InitializeComponent();
            _StudentIDList = StudentIDList;
        }

        void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("期考成績名冊產生中..");
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("期考成績名冊產生完成");
            btnExport.Enabled = true;
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
         
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
        }
    }
}
