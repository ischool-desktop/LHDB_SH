using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Aspose.Cells;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using SHSchool.Data;
using FISCA.Data;
using System.Data;
using LHDB_SH_Core.DAO;
using System.Xml;
using SmartSchool.Evaluation.ScoreCalcRule;
using FISCA.DSAUtil;
using SmartSchool.Customization.Data;

namespace LHDB_SH_Core
{
    public class Utility
    {
        /// <summary>
        /// 匯出 Excel
        /// </summary>
        /// <param name="inputReportName"></param>
        /// <param name="inputXls"></param>
        public static void ExprotXls(string ReportName, Workbook wbXls)
        {
            string reportName = ReportName;

            string path = Path.Combine(Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".xls");

            Workbook wb = wbXls;

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                wb.Save(path, SaveFormat.Excel97To2003);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".xls";
                sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        wb.Save(sd.FileName, SaveFormat.Excel97To2003);

                    }
                    catch
                    {
                        MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// 將西元日期轉成民國數字
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string ConvertChDateString(DateTime? dt)
        {
            string value = "";
            if(dt.HasValue)
            {
                value = string.Format("{0:000}",(dt.Value.Year - 1911)) + string.Format("{0:00}", dt.Value.Month) + string.Format("{0:00}", dt.Value.Day);
            }

            return value;
        }


        /// <summary>
        /// 建立使用到的 UDT Table
        /// </summary>
        public static void CreateUDTTable()
        {
            FISCA.UDT.SchemaManager Manager = new SchemaManager(new FISCA.DSAUtil.DSConnection(FISCA.Authentication.DSAServices.DefaultDataSource));
            Manager.SyncSchema(new DAO.udt_ConfigData());
        }


        public static string GetDgCellValue(DataGridViewCell cell)
        {
            if (cell.Value == null)
                return "";
            else
                return cell.Value.ToString();

        }

        /// <summary>
        /// 取得科別代碼,key:科別名稱，value:科別代碼
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string,string> GetDepartmetDict ()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            List<SHDepartmentRecord> RecList = SHDepartment.SelectAll();
            foreach(SHDepartmentRecord rec in RecList)
            {
                if (!value.ContainsKey(rec.FullName))
                    value.Add(rec.FullName, rec.Code);
            }
            return value;
        }

        /// <summary>
        /// 取得大學繁星班級代碼
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string,string> GetClassCodeDict()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            try
            {
                QueryHelper qh = new QueryHelper();
                string query = "select class_name,class_code from $sh.college.sat.class_code";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string cName = dr["class_name"].ToString();
                    string cCode = dr["class_code"].ToString();
                    if (!value.ContainsKey(cName))
                        value.Add(cName, cCode);
                }

            }
            catch (Exception ex) { }

            return value;
        }

        /// <summary>
        /// 取得學生科別名稱
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string,string> GetStudDeptNameDict(List<string> StudentIDList,int SchoolYear,int Semester)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();

            // 如果是系統預設學年度學期,不是讀取學期對照
            if (SchoolYear.ToString() == K12.Data.School.DefaultSchoolYear && Semester.ToString() == K12.Data.School.DefaultSemester)
            {
                if (StudentIDList.Count > 0)
                {
                    QueryHelper qhC1 = new QueryHelper();
                    // 取的學生所屬班級科別
                    string strC1 = "select student.id as sid,dept.name as deptName from student inner join class on student.ref_class_id=class.id inner join dept on class.ref_dept_id=dept.id where student.id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                    DataTable dtC1 = qhC1.Select(strC1);
                    foreach (DataRow dr in dtC1.Rows)
                    {
                        string sid = dr["sid"].ToString();
                        string deptName = dr["deptName"].ToString();
                        if (!value.ContainsKey(sid))
                            value.Add(sid, deptName);
                        else
                            value[sid] = deptName;
                    }

                    QueryHelper qhS1 = new QueryHelper();
                    string strS1 = "select student.id as sid,dept.name as deptName from student inner join dept on student.ref_dept_id=dept.id where student.id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                    DataTable dtS1 = qhS1.Select(strS1);
                    foreach (DataRow dr in dtS1.Rows)
                    {
                        string sid = dr["sid"].ToString();
                        string deptName = dr["deptName"].ToString();
                        if (!value.ContainsKey(sid))
                            value.Add(sid, deptName);
                        else
                            value[sid] = deptName;
                    }
                }
            }
            else
            {
                List<SHSemesterHistoryRecord> SHSemesterHistoryRecordList = SHSemesterHistory.SelectByStudentIDs(StudentIDList);
                foreach (SHSemesterHistoryRecord shrec in SHSemesterHistoryRecordList)
                {
                    foreach (K12.Data.SemesterHistoryItem item in shrec.SemesterHistoryItems)
                    {
                        if (item.SchoolYear == SchoolYear && item.Semester == Semester)
                        {
                            if (!value.ContainsKey(item.RefStudentID))
                                value.Add(item.RefStudentID, item.DeptName);
                        }
                    }
                }
            }          

            return value;
        }

        /// <summary>
        /// 取得所有試別
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string,string> GetExamNameDict()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            string query = "select id,exam_name from exam order by exam_name";
            DataTable dt = qh.Select(query);
            foreach(DataRow dr in dt.Rows)
            {
                string name = dr["exam_name"].ToString();
                string id = dr["id"].ToString();
                if (!value.ContainsKey(name))
                    value.Add(name, id);
            }
            return value;
        }

       public static Dictionary<string,List<StudentSCETakeRec>> GetStudentSCETakeDict(List<string> StudentIDList,string ExamID,int SchoolYear,int Semester)
        {
            Dictionary<string, List<StudentSCETakeRec>> value = new Dictionary<string, List<StudentSCETakeRec>>();
           if(StudentIDList.Count>0 && ExamID !="")
           {
               SmartSchool.Customization.Data.AccessHelper acc = new SmartSchool.Customization.Data.AccessHelper ();
               
               // 取得及格標準
               List<SmartSchool.Customization.Data.StudentRecord> StudentRecList = acc.StudentHelper.GetStudents(StudentIDList);
               Dictionary<string, Dictionary<string, decimal>> passScoreDict = GetStudentApplyLimitDict(StudentRecList);
               Dictionary<string, string> StudGradYearDict = new Dictionary<string, string>();

               foreach(SmartSchool.Customization.Data.StudentRecord rec in StudentRecList)
               {
                   if (!StudGradYearDict.ContainsKey(rec.StudentID))
                       if(rec.RefClass !=null)
                            StudGradYearDict.Add(rec.StudentID, rec.RefClass.GradeYear);
               }

               // 取得學期對照年為主
               List<SHSemesterHistoryRecord> SemsH = SHSemesterHistory.SelectByStudentIDs(StudentIDList);
               foreach(SHSemesterHistoryRecord rec in SemsH)
               {                   
                   foreach(K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems )
                   {
                       if(item.SchoolYear==SchoolYear && item.Semester== Semester)
                       {
                           if (!StudGradYearDict.ContainsKey(item.RefStudentID))
                               StudGradYearDict.Add(item.RefStudentID, item.GradeYear.ToString());
                           else
                               StudGradYearDict[item.RefStudentID] = item.GradeYear.ToString();
                       }
                   }
               }

               List<SHCourseRecord> CourseRecList = SHCourse.SelectBySchoolYearAndSemester(SchoolYear, Semester);
               Dictionary<string, SHCourseRecord> CourseRecDict = new Dictionary<string, SHCourseRecord>();
               foreach (SHCourseRecord rec in CourseRecList)
                   if (!CourseRecDict.ContainsKey(rec.ID))
                       CourseRecDict.Add(rec.ID, rec);

               QueryHelper qh = new QueryHelper();
               string query = "select sc_attend.ref_student_id as sid,course.id as cid,course.course_name,course.subject,sce_take.score as sce_score,course.credit as credit from sc_attend inner join course on sc_attend.ref_course_id=course.id inner join sce_take on sc_attend.id=sce_take.ref_sc_attend_id where sc_attend.ref_student_id in("+string.Join(",",StudentIDList.ToArray())+") and course.school_year="+SchoolYear+" and course.semester="+Semester+" and sce_take.ref_exam_id in("+ExamID+")";
               DataTable dt = qh.Select(query);
               foreach(DataRow dr in dt.Rows)
               {
                   string SubjCode = "";
                   string sid = dr["sid"].ToString();
                   if (!value.ContainsKey(sid))
                       value.Add(sid, new List<StudentSCETakeRec>());

                   string cid = dr["cid"].ToString();
                   StudentSCETakeRec scetRec = new StudentSCETakeRec();
                   scetRec.StudentID = sid;

                   if(CourseRecDict.ContainsKey(cid))
                   {
                       string Req="",notNq="";
                       if(CourseRecDict[cid].Required)
                           Req="必修";
                       else
                           Req="選修";

                       if(CourseRecDict[cid].NotIncludedInCredit)
                           notNq="是";
                       else
                           notNq="否";

                       SubjCode = CourseRecDict[cid].Subject + "_" + CourseRecDict[cid].Credit.Value + "_" + Req + "_" + CourseRecDict[cid].RequiredBy + "_" + CourseRecDict[cid].Entry + "_" + notNq;
                   }
                   scetRec.SubjectCode = SubjCode;
                   string GrStr="";
                   if (StudGradYearDict.ContainsKey(sid))
                       GrStr = StudGradYearDict[sid]+"_及";
                   
                   scetRec.Score = "-1";

                   decimal dd, passScore = 0;
                   if (decimal.TryParse(dr["sce_score"].ToString(),out dd))
                   {
                       // 取得及格標準
                       if(passScoreDict.ContainsKey(sid))
                           if (passScoreDict[sid].ContainsKey(GrStr))
                                passScore=passScoreDict[sid][GrStr];

                       // 不及格標 *
                       if(dd<passScore)
                           scetRec.Score ="*"+ string.Format("{0:###.00}", dd);
                       else
                        scetRec.Score = string.Format("{0:###.00}", dd);

                       scetRec.Status = "1";
                   }else
                   {                       
                       // 缺考
                       scetRec.Status = "4";
                   }
                   
                   scetRec.SubjectCredit = dr["credit"].ToString();
                   
                   value[sid].Add(scetRec);
               }
           }
            return value;
        }

       /// <summary>
       /// 取得學生及格與補考標準，參數用學生IDList,回傳:key:StudentID,1_及,數字
       /// </summary>
       /// <param name="StudRecList"></param>
       /// <returns></returns>
       public static Dictionary<string, Dictionary<string, decimal>> GetStudentApplyLimitDict(List<SmartSchool.Customization.Data.StudentRecord> StudRecList)
       {

           Dictionary<string, Dictionary<string, decimal>> retVal = new Dictionary<string, Dictionary<string, decimal>>();


           foreach (SmartSchool.Customization.Data.StudentRecord studRec in StudRecList)
           {
               //及格標準<年級,及格與補考標準>
               if (!retVal.ContainsKey(studRec.StudentID))
                   retVal.Add(studRec.StudentID, new Dictionary<string, decimal>());

               XmlElement scoreCalcRule = SmartSchool.Evaluation.ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(studRec.StudentID) == null ? null : SmartSchool.Evaluation.ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(studRec.StudentID).ScoreCalcRuleElement;
               if (scoreCalcRule == null)
               {

               }
               else
               {

                   DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                   decimal tryParseDecimal;
                   decimal tryParseDecimala;

                   foreach (XmlElement element in helper.GetElements("及格標準/學生類別"))
                   {
                       string cat = element.GetAttribute("類別");
                       bool useful = false;
                       //掃描學生的類別作比對
                       foreach (CategoryInfo catinfo in studRec.StudentCategorys)
                       {
                           if (catinfo.Name == cat || catinfo.FullName == cat)
                               useful = true;
                       }
                       //學生是指定的類別或類別為"預設"
                       if (cat == "預設" || useful)
                       {
                           for (int gyear = 1; gyear <= 4; gyear++)
                           {
                               switch (gyear)
                               {
                                   case 1:
                                       if (decimal.TryParse(element.GetAttribute("一年級及格標準"), out tryParseDecimal))
                                       {
                                           string k1s = gyear + "_及";

                                           if (!retVal[studRec.StudentID].ContainsKey(k1s))
                                               retVal[studRec.StudentID].Add(k1s, tryParseDecimal);

                                           if (retVal[studRec.StudentID][k1s] > tryParseDecimal)
                                               retVal[studRec.StudentID][k1s] = tryParseDecimal;
                                       }

                                       if (decimal.TryParse(element.GetAttribute("一年級補考標準"), out tryParseDecimala))
                                       {
                                           string k1a = gyear + "_補";

                                           if (!retVal[studRec.StudentID].ContainsKey(k1a))
                                               retVal[studRec.StudentID].Add(k1a, tryParseDecimala);

                                           if (retVal[studRec.StudentID][k1a] > tryParseDecimala)
                                               retVal[studRec.StudentID][k1a] = tryParseDecimala;
                                       }

                                       break;
                                   case 2:
                                       if (decimal.TryParse(element.GetAttribute("二年級及格標準"), out tryParseDecimal))
                                       {
                                           string k2s = gyear + "_及";

                                           if (!retVal[studRec.StudentID].ContainsKey(k2s))
                                               retVal[studRec.StudentID].Add(k2s, tryParseDecimal);

                                           if (retVal[studRec.StudentID][k2s] > tryParseDecimal)
                                               retVal[studRec.StudentID][k2s] = tryParseDecimal;
                                       }

                                       if (decimal.TryParse(element.GetAttribute("二年級補考標準"), out tryParseDecimala))
                                       {
                                           string k2a = gyear + "_補";

                                           if (!retVal[studRec.StudentID].ContainsKey(k2a))
                                               retVal[studRec.StudentID].Add(k2a, tryParseDecimala);

                                           if (retVal[studRec.StudentID][k2a] > tryParseDecimala)
                                               retVal[studRec.StudentID][k2a] = tryParseDecimala;

                                       }

                                       break;
                                   case 3:
                                       if (decimal.TryParse(element.GetAttribute("三年級及格標準"), out tryParseDecimal))
                                       {
                                           string k3s = gyear + "_及";

                                           if (!retVal[studRec.StudentID].ContainsKey(k3s))
                                               retVal[studRec.StudentID].Add(k3s, tryParseDecimal);

                                           if (retVal[studRec.StudentID][k3s] > tryParseDecimal)
                                               retVal[studRec.StudentID][k3s] = tryParseDecimal;
                                       }

                                       if (decimal.TryParse(element.GetAttribute("三年級補考標準"), out tryParseDecimala))
                                       {
                                           string k3a = gyear + "_補";

                                           if (!retVal[studRec.StudentID].ContainsKey(k3a))
                                               retVal[studRec.StudentID].Add(k3a, tryParseDecimala);

                                           if (retVal[studRec.StudentID][k3a] > tryParseDecimala)
                                               retVal[studRec.StudentID][k3a] = tryParseDecimala;
                                       }

                                       break;
                                   case 4:
                                       if (decimal.TryParse(element.GetAttribute("四年級及格標準"), out tryParseDecimal))
                                       {
                                           string k4s = gyear + "_及";

                                           if (!retVal[studRec.StudentID].ContainsKey(k4s))
                                               retVal[studRec.StudentID].Add(k4s, tryParseDecimal);

                                           if (retVal[studRec.StudentID][k4s] > tryParseDecimal)
                                               retVal[studRec.StudentID][k4s] = tryParseDecimal;
                                       }

                                       if (decimal.TryParse(element.GetAttribute("四年級補考標準"), out tryParseDecimala))
                                       {
                                           string k4a = gyear + "_補";

                                           if (!retVal[studRec.StudentID].ContainsKey(k4a))
                                               retVal[studRec.StudentID].Add(k4a, tryParseDecimala);

                                           if (retVal[studRec.StudentID][k4a] > tryParseDecimala)
                                               retVal[studRec.StudentID][k4a] = tryParseDecimala;
                                       }

                                       break;
                                   default:
                                       break;
                               }
                           }
                       }
                   }
               }
           }
           return retVal;
       }



        /// <summary>
        /// 取得科別、群代碼對照
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string,string> GetGroupDeptIDDict()
       {
           Dictionary<string, string> value = new Dictionary<string, string>();
           value.Add("099", "41");
           value.Add("101", "11");
           value.Add("102", "11");
           value.Add("103", "11");
           value.Add("104", "11");
           value.Add("105", "11");
           value.Add("106", "11");
           value.Add("107", "11");
           value.Add("108", "11");
           value.Add("109", "11");
           value.Add("110", "11");
           value.Add("111", "11");
           value.Add("121", "11");
           value.Add("122", "11");
           value.Add("201", "29");
           value.Add("202", "29");
           value.Add("204", "29");
           value.Add("205", "22");
           value.Add("206", "30");
           value.Add("214", "29");
           value.Add("215", "26");
           value.Add("216", "29");
           value.Add("217", "29");
           value.Add("301", "21");
           value.Add("302", "21");
           value.Add("303", "22");
           value.Add("304", "21");
           value.Add("305", "23");
           value.Add("306", "23");
           value.Add("307", "23");
           value.Add("308", "23");
           value.Add("309", "23");
           value.Add("311", "25");
           value.Add("312", "28");
           value.Add("315", "24");
           value.Add("316", "28");
           value.Add("318", "28");
           value.Add("319", "24");
           value.Add("331", "21");
           value.Add("332", "21");
           value.Add("335", "22");
           value.Add("337", "21");
           value.Add("338", "21");
           value.Add("352", "24");
           value.Add("360", "21");
           value.Add("361", "28");
           value.Add("363", "21");
           value.Add("364", "22");
           value.Add("365", "25");
           value.Add("366", "28");
           value.Add("367", "24");
           value.Add("369", "23");
           value.Add("371", "21");
           value.Add("372", "21");
           value.Add("373", "28");
           value.Add("374", "21");
           value.Add("380", "22");
           value.Add("381", "22");
           value.Add("384", "23");
           value.Add("392", "22");
           value.Add("394", "28");
           value.Add("397", "25");
           value.Add("398", "25");
           value.Add("399", "28");
           value.Add("401", "26");
           value.Add("402", "26");
           value.Add("403", "26");
           value.Add("404", "26");
           value.Add("405", "26");
           value.Add("406", "28");
           value.Add("407", "32");
           value.Add("408", "32");
           value.Add("418", "26");
           value.Add("419", "27");
           value.Add("421", "27");
           value.Add("424", "32");
           value.Add("425", "26");
           value.Add("426", "26");
           value.Add("427", "32");
           value.Add("430", "28");
           value.Add("431", "28");
           value.Add("501", "31");
           value.Add("502", "31");
           value.Add("503", "31");
           value.Add("504", "31");
           value.Add("505", "30");
           value.Add("512", "28");
           value.Add("513", "31");
           value.Add("514", "31");
           value.Add("515", "31");
           value.Add("516", "31");
           value.Add("517", "30");
           value.Add("701", "33");
           value.Add("702", "34");
           value.Add("703", "23");
           value.Add("705", "33");
           value.Add("706", "26");
           value.Add("708", "34");
           value.Add("717", "26");
           value.Add("718", "30");
           value.Add("801", "35");
           value.Add("802", "35");
           value.Add("803", "35");
           value.Add("804", "35");
           value.Add("806", "35");
           value.Add("807", "35");
           value.Add("808", "35");
           value.Add("809", "35");
           value.Add("813", "35");
           value.Add("816", "35");
           value.Add("817", "35");
           value.Add("820", "35");
           value.Add("821", "35");
           value.Add("822", "35");
           value.Add("823", "35");
           value.Add("824", "35");
           value.Add("825", "35");
           value.Add("901", "41");
           value.Add("00A", "11");
           value.Add("01A", "11");
           value.Add("03A", "11");
           value.Add("04A", "11");
           value.Add("05A", "11");
           value.Add("06A", "11");
           value.Add("09A", "11");
           value.Add("12A", "29");
           value.Add("18A", "26");
           value.Add("A01", "21");
           value.Add("A02", "21");
           value.Add("A03", "21");
           value.Add("A04", "21");
           value.Add("A05", "21");
           value.Add("A06", "21");
           value.Add("A07", "21");
           value.Add("A08", "21");
           value.Add("A09", "21");
           value.Add("A10", "21");
           value.Add("A11", "21");
           value.Add("A12", "21");
           value.Add("A13", "21");
           value.Add("A14", "21");
           value.Add("B01", "22");
           value.Add("B02", "22");
           value.Add("B03", "22");
           value.Add("B04", "22");
           value.Add("B05", "22");
           value.Add("B06", "22");
           value.Add("B07", "22");
           value.Add("B08", "22");
           value.Add("B09", "22");
           value.Add("B10", "22");
           value.Add("B11", "22");
           value.Add("C01", "23");
           value.Add("C02", "23");
           value.Add("C03", "23");
           value.Add("C04", "23");
           value.Add("C05", "23");
           value.Add("C06", "23");
           value.Add("C07", "23");
           value.Add("C08", "23");
           value.Add("C09", "23");
           value.Add("C10", "23");
           value.Add("C11", "23");
           value.Add("C12", "23");
           value.Add("C13", "23");
           value.Add("C14", "23");
           value.Add("C15", "23");
           value.Add("C16", "23");
           value.Add("C17", "23");
           value.Add("C18", "23");
           value.Add("C19", "23");
           value.Add("C20", "23");
           value.Add("D01", "24");
           value.Add("D02", "24");
           value.Add("D04", "24");
           value.Add("D05", "24");
           value.Add("E01", "25");
           value.Add("E02", "25");
           value.Add("E03", "25");
           value.Add("E04", "25");
           value.Add("E06", "25");
           value.Add("E07", "25");
           value.Add("F01", "26");
           value.Add("F02", "26");
           value.Add("F03", "26");
           value.Add("F04", "26");
           value.Add("F05", "26");
           value.Add("F06", "26");
           value.Add("F07", "26");
           value.Add("F08", "26");
           value.Add("F09", "26");
           value.Add("F10", "26");
           value.Add("F11", "26");
           value.Add("F12", "26");
           value.Add("F13", "26");
           value.Add("F14", "26");
           value.Add("F15", "26");
           value.Add("F16", "26");
           value.Add("F17", "26");
           value.Add("F18", "26");
           value.Add("F19", "26");
           value.Add("F20", "26");
           value.Add("F22", "26");
           value.Add("F23", "26");
           value.Add("G01", "29");
           value.Add("G02", "29");
           value.Add("G03", "29");
           value.Add("G04", "29");
           value.Add("G05", "29");
           value.Add("G06", "29");
           value.Add("G07", "29");
           value.Add("G08", "29");
           value.Add("G09", "29");
           value.Add("G10", "29");
           value.Add("G11", "29");
           value.Add("G12", "29");
           value.Add("H01", "31");
           value.Add("H02", "31");
           value.Add("H03", "31");
           value.Add("H04", "31");
           value.Add("H05", "31");
           value.Add("H06", "31");
           value.Add("H07", "31");
           value.Add("H08", "31");
           value.Add("H09", "31");
           value.Add("H10", "31");
           value.Add("H11", "31");
           value.Add("J01", "32");
           value.Add("J02", "32");
           value.Add("J03", "32");
           value.Add("J04", "32");
           value.Add("J05", "32");
           value.Add("J06", "32");
           value.Add("J07", "32");
           value.Add("J08", "32");
           value.Add("J09", "32");
           value.Add("J10", "32");
           value.Add("J11", "32");
           value.Add("J12", "32");
           value.Add("J13", "32");
           value.Add("J14", "32");
           value.Add("J15", "32");
           value.Add("J16", "32");
           value.Add("J17", "32");
           value.Add("J18", "32");
           value.Add("J19", "32");
           value.Add("K01", "29");
           value.Add("K03", "29");
           value.Add("K09", "29,31");
           value.Add("K11", "29");
           value.Add("K13", "29,30");
           value.Add("K15", "30");
           value.Add("K17", "29");
           value.Add("K19", "29");
           value.Add("K21", "30,32");
           value.Add("K25", "30");
           value.Add("K27", "29");
           value.Add("L01", "22");
           value.Add("L03", "23");
           value.Add("L05", "25");
           value.Add("L07", "28");
           value.Add("L09", "28");
           value.Add("L11", "28");
           value.Add("L13", "28");
           value.Add("L15", "21");
           value.Add("L17", "21");
           value.Add("L19", "22");
           value.Add("L20", "21");
           value.Add("L21", "22");
           value.Add("L23", "23");
           value.Add("L24", "23");
           value.Add("L25", "22,24");
           value.Add("L27", "28");
           value.Add("L29", "23");
           value.Add("L31", "23");
           value.Add("L33", "21");
           value.Add("L35", "25");
           value.Add("L37", "23");
           value.Add("L39", "28");
           value.Add("L41", "28");
           value.Add("L43", "28");
           value.Add("L45", "22");
           value.Add("L47", "21");
           value.Add("L49", "23");
           value.Add("L50", "23");
           value.Add("L51", "21");
           value.Add("L52", "21");
           value.Add("L53", "23");
           value.Add("L55", "21,25");
           value.Add("L57", "28");
           value.Add("L58", "28");
           value.Add("L59", "21");
           value.Add("L61", "21");
           value.Add("L62", "21");
           value.Add("L63", "22");
           value.Add("L65", "23");
           value.Add("L67", "23");
           value.Add("L69", "24");
           value.Add("L71", "24");
           value.Add("L75", "22");
           value.Add("L79", "28");
           value.Add("L81", "28");
           value.Add("M03", "26,28");
           value.Add("M05", "26");
           value.Add("M09", "26");
           value.Add("M11", "26");
           value.Add("M14", "32");
           value.Add("M15", "26");
           value.Add("M17", "32");
           value.Add("M19", "26");
           value.Add("M21", "32");
           value.Add("M23", "32");
           value.Add("M25", "28");
           value.Add("M27", "32");
           value.Add("M29", "26");
           value.Add("N01", "31");
           value.Add("N03", "31");
           value.Add("N05", "31,36");
           value.Add("N11", "31,36");
           value.Add("N15", "32");
           value.Add("N23", "31");
           value.Add("N25", "32");
           value.Add("P01", "33");
           value.Add("P09", "34");
           value.Add("P11", "34");
           value.Add("P12", "34");
           value.Add("P13", "33");
           value.Add("P15", "34");
           value.Add("P17", "30");
           value.Add("P19", "32");
           value.Add("P21", "34");
           value.Add("P23", "33");
           value.Add("Q01", "33");
           value.Add("Q02", "33");
           value.Add("R01", "34");
           value.Add("U01", "35");
           value.Add("U02", "35");
           value.Add("U03", "35");
           value.Add("U04", "35");
           value.Add("U05", "35");
           value.Add("V01", "28");
           value.Add("V02", "28");
           value.Add("V03", "28");
           value.Add("V04", "28");
           value.Add("V05", "28");
           value.Add("V06", "28");
           value.Add("V07", "28");
           value.Add("V08", "28");
           value.Add("V09", "28");
           value.Add("V10", "28");
           value.Add("V11", "28");
           value.Add("V12", "28");
           value.Add("V13", "28");
           value.Add("V14", "28");
           value.Add("V15", "28");
           value.Add("V16", "28");
           value.Add("V17", "28");
           value.Add("V18", "28");
           value.Add("V19", "28");
           value.Add("V20", "28");
           value.Add("V21", "28");
           value.Add("V22", "28");
           value.Add("W01", "30");
           value.Add("W02", "30");
           value.Add("W03", "30");
           value.Add("W04", "30");
           value.Add("X01", "27");
           value.Add("X02", "27");
           value.Add("X03", "27");
           value.Add("X04", "27");
           value.Add("X05", "27");
           value.Add("X06", "27");
           value.Add("X07", "27");
           value.Add("X08", "27");
           value.Add("X09", "27");
           value.Add("X10", "27");
           value.Add("X11", "27");
           value.Add("Y01", "41");
           value.Add("Y02", "41");

           return value;
       }

        /// <summary>
        /// 取得學期對照內班級座號，轉成大學繁星代碼與相對學習歷程資料庫需要班座
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="ReturnSeatNo"></param>
        /// <returns></returns>
        public static Dictionary<string,string> GetStudentClassCodeSeatNo(List<string> StudentIDList,int SchoolYear,int Semester,bool ReturnSeatNo)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            Dictionary<string, string> ClassCodeDict = new Dictionary<string, string>();            
            Dictionary<string, string> codeDict = new Dictionary<string, string>();
            Dictionary<string, string> SeatNoDict = new Dictionary<string, string>();

            // 取得學生學期對照
            List<SHSemesterHistoryRecord> SHSemesterHistoryRecordList = SHSemesterHistory.SelectByStudentIDs(StudentIDList);
            foreach (SHSemesterHistoryRecord rec in SHSemesterHistoryRecordList)
            {
                foreach(K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                {
                    if(item.SchoolYear== SchoolYear && item.Semester==Semester)
                    {
                        if (!ClassCodeDict.ContainsKey(rec.RefStudentID))
                        {
                            ClassCodeDict.Add(rec.RefStudentID, item.ClassName);
                            if (item.SeatNo.HasValue)
                                SeatNoDict.Add(rec.RefStudentID, string.Format("{0:00}", item.SeatNo.Value));
                        }
                    }
                }
            }

            // 取得學習歷程班級代碼
            try
            {
                codeDict = GetLHClassCodeDict();
            }
            catch (Exception ex) {  }
            
           // 比對資料
            foreach(string sid in ClassCodeDict.Keys)
            {
                string cName = ClassCodeDict[sid];

                if(codeDict.ContainsKey(cName))
                {
                    // +回傳座號
                    if(ReturnSeatNo)
                    {
                        if (SeatNoDict.ContainsKey(sid))
                            value.Add(sid, codeDict[cName] + SeatNoDict[sid]);
                    }
                    else
                    {
                        // 只有班代
                        value.Add(sid, codeDict[cName]);
                    }
                }
            }

            return value;
        }

        /// <summary>
        /// 取得學習歷程班級代碼
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string,string> GetLHClassCodeDict()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            ConfigData cd = new ConfigData();
            Dictionary<string, List<ConfigDataItem>> datas = cd.GetConfigDataItemDict();
            if(datas.ContainsKey("班級代碼"))
            {
                List<ConfigDataItem> items = (from data in datas["班級代碼"] orderby data.Name ascending select data).ToList();

                foreach(ConfigDataItem cdi in items)
                {
                    if (!value.ContainsKey(cdi.Name))
                        value.Add(cdi.Name, cdi.Value);
                }
            }
            return value;
        }
    }
}
