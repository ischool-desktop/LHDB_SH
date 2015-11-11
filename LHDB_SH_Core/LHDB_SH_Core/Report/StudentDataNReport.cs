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
    public partial class StudentDataNReport : BaseForm
    {
        private int _SchoolYear = 0, _Semester = 0;
        private string _SchoolCode = "";
        // 名冊別
        private string _DocType = "1";

        BackgroundWorker _bgWorker;
        Workbook _wb;
        List<string> _StudentIDList;
        
        public StudentDataNReport(List<string> StudentIDList)
        {
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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學生資料名冊(國教署主管學校)產生完成");
            btnExport.Enabled = true;
            if(_wb !=null)
            {
                Utility.ExprotXls("學生資料名冊(國教署主管學校)", _wb);
            }
        }

        void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學生資料名冊(國教署主管學校)產生中..", e.ProgressPercentage);
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
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

            // 取得學生類別
            foreach(SHStudentTagRecord TRec in SHStudentTagRecordList)
            {
                if (!StudTagNameDict.ContainsKey(TRec.RefStudentID))
                    StudTagNameDict.Add(TRec.RefStudentID, new List<string>());

                StudTagNameDict[TRec.RefStudentID].Add(TRec.FullName);
            }



            foreach (SHClassRecord rec in SHClass.SelectAll())
                ClassIDNameDict.Add(rec.ID, rec.Name);

            // 部別對照
            if(cdDict.ContainsKey("部別代碼"))
            {
                foreach(ConfigDataItem cdi in cdDict["部別代碼"])
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
            ClassNoMappingDict = Utility.GetClassCodeDict();
        
            // 取得學生本資料
            List<SHStudentRecord> StudentRecordList = SHStudent.SelectByIDs(_StudentIDList);
            List<StudentBaseRec> StudentBaseRecList = new List<StudentBaseRec>();
            
            // 整理資料
            foreach(SHStudentRecord studRec in StudentRecordList)
            {
                // 身分證號,出生日期,所屬學校代碼,科/班/學程別代碼,部別,班別,班級座號代碼
                StudentBaseRec sbr = new StudentBaseRec();
                sbr.IDNumber = studRec.IDNumber;
                sbr.BirthDate = Utility.ConvertChDateString(studRec.Birthday);
                sbr.SchoolCode = _SchoolCode;
                
                if (StudTagNameDict.ContainsKey(studRec.ID))
                {
                    // 科/班/學程別代碼
                    sbr.DCLCode = "";
                    // 部別
                    sbr.DepCode = "";
                    // 班別
                    sbr.ClCode = "";

                    
                    foreach (string str in StudTagNameDict[studRec.ID])
                    {
                        if (DepMappingDict.ContainsKey(str))
                            sbr.DepCode = DepMappingDict[str];

                        if (DeptMappingDict.ContainsKey(str))
                            sbr.DCLCode = DeptMappingDict[str];

                        if (ClsMappingDict.ContainsKey(str))
                            sbr.ClCode = ClsMappingDict[str];
                    }

                }
                
                // 班級座號代碼
                sbr.ClassSeatCode = "";
                if(ClassIDNameDict.ContainsKey(studRec.RefClassID))
                {
                    string cName = ClassIDNameDict[studRec.RefClassID];
                    if(ClassNoMappingDict.ContainsKey(cName) && studRec.SeatNo.HasValue)
                        sbr.ClassSeatCode = ClassNoMappingDict[cName] + string.Format("{0:00}", studRec.SeatNo.Value);
                }


                StudentBaseRecList.Add(sbr);
            }

            // 排序 班級座號代碼
            StudentBaseRecList = (from data in StudentBaseRecList orderby data.ClassSeatCode ascending select data).ToList();

            // 填值到 Excel
            _wb = new Workbook(new MemoryStream(Properties.Resources.學生資料名冊樣板_國教署主管學校_));
            Worksheet wst1 = _wb.Worksheets["學生資料名冊封面"];
            Worksheet wst2 = _wb.Worksheets["學生資料名冊"];

            // 學校代碼,學年度,學期,名冊別
            wst1.Cells[1, 0].PutValue(_SchoolCode);
            wst1.Cells[1, 1].PutValue(_SchoolYear);
            wst1.Cells[1, 2].PutValue(_Semester);
            wst1.Cells[1, 3].PutValue(_DocType);

            // 身分證號,出生日期,所屬學校代碼,科/班/學程別代碼,部別,班別,班級座號代碼
            int rowIdx=1;
            foreach(StudentBaseRec sbr in StudentBaseRecList)
            {
                wst2.Cells[rowIdx, 0].PutValue(sbr.IDNumber);
                wst2.Cells[rowIdx, 1].PutValue(sbr.BirthDate);
                wst2.Cells[rowIdx, 2].PutValue(sbr.SchoolCode);
                wst2.Cells[rowIdx, 3].PutValue(sbr.DCLCode);
                wst2.Cells[rowIdx, 4].PutValue(sbr.DepCode);
                wst2.Cells[rowIdx, 5].PutValue(sbr.ClCode);
                wst2.Cells[rowIdx, 6].PutValue(sbr.ClassSeatCode);
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
                _SchoolCode = K12.Data.School.Code;

                _bgWorker.RunWorkerAsync();

            }catch(Exception ex)
            {

            }
            
        }

        private void lkDepSetup_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Config.DepConfigForm dcf = new Config.DepConfigForm();
            dcf.ShowDialog();
        }
    }
}
