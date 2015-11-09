using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LHDB_SH_Core.DAO
{
    /// <summary>
    /// 學生缺勤報表資料
    /// </summary>
    public class StudAttendanceRec
    {

        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 身分證號
        /// </summary>
        public string IDNumber { get; set; }
        /// <summary>
        /// 出生日期
        /// </summary>
        public string BirthDate { get; set; }

        /// <summary>
        /// 缺勤種類代碼
        /// </summary>
        public string AttendType { get; set; }

        /// <summary>
        /// 缺勤節數
        /// </summary>
        public int AttendTypeCount { get; set; }

        /// <summary>
        /// 缺勤起始日期
        /// </summary>
        public string BeginDate { get; set; }

        /// <summary>
        /// 缺勤結束日期
        /// </summary>
        public string EndDate { get; set; }

    }
}
