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
        /// 取得科別代碼,key:學生類別，value:code
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
            QueryHelper qh = new QueryHelper();
            string query = "select class_name,class_code from $sh.college.sat.class_code";
            DataTable dt = qh.Select(query);
            foreach(DataRow dr in dt.Rows)
            {
                string cName = dr["class_name"].ToString();
                string cCode = dr["class_code"].ToString();
                if(!value.ContainsKey(cName))
                    value.Add(cName,cCode);
            }

            return value;
        }

        /// <summary>
        /// 取得學生科別名稱
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string,string> GetStudDeptNameDict(List<string> StudentIDList)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            if(StudentIDList.Count>0)
            {
                QueryHelper qhC1 = new QueryHelper();
                // 取的學生所屬班級科別
                string strC1 = "select student.id as sid,dept.name as deptName from student inner join class on student.ref_class_id=class.id inner join dept on class.ref_dept_id=dept.id where student.id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                DataTable dtC1 = qhC1.Select(strC1);
                foreach(DataRow dr in dtC1.Rows)
                {
                    string sid = dr["sid"].ToString();
                    string deptName=dr["deptName"].ToString();
                    if (!value.ContainsKey(sid))
                        value.Add(sid, deptName);
                    else
                        value[sid] = deptName;
                }

                QueryHelper qhS1 = new QueryHelper();
                string strS1 = "select student.id as sid,dept.name as deptName from student inner join dept on student.ref_dept_id=dept.id where student.id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                DataTable dtS1 = qhS1.Select(strS1);
                foreach(DataRow dr in dtS1.Rows)
                {
                    string sid = dr["sid"].ToString();
                    string deptName = dr["deptName"].ToString();
                    if (!value.ContainsKey(sid))
                        value.Add(sid, deptName);
                    else
                        value[sid] = deptName;
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

               QueryHelper qh = new QueryHelper();
               string query = "select sc_attend.ref_student_id as sid,course.course_name,course.subject,sce_take.score as sce_score,course.credit as credit from sc_attend inner join course on sc_attend.ref_course_id=course.id inner join sce_take on sc_attend.id=sce_take.ref_sc_attend_id where sc_attend.ref_student_id in("+string.Join(",",StudentIDList.ToArray())+") and course.school_year="+SchoolYear+" and course.semester="+Semester+" and sce_take.ref_exam_id in("+ExamID+")";
               DataTable dt = qh.Select(query);
               foreach(DataRow dr in dt.Rows)
               {
                   string sid = dr["sid"].ToString();
                   if (!value.ContainsKey(sid))
                       value.Add(sid, new List<StudentSCETakeRec>());

                   StudentSCETakeRec scetRec = new StudentSCETakeRec();
                   scetRec.StudentID = sid;
                   scetRec.SubjectCode=dr["subject"].ToString();
                   string GrStr="";
                   if (StudGradYearDict.ContainsKey(sid))
                       GrStr = StudGradYearDict[sid]+"_及";

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
                       scetRec.Score = "00.00";
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

    }
}
