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
    }
}
