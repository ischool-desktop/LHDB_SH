using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LHDB_SH_Core.DAO
{
    /// <summary>
    /// 定期評量成績
    /// </summary>
    public class StudentSCETakeRec
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
        /// 科目代碼
        /// </summary>
        public string SubjectCode { get; set; }

        /// <summary>
        /// 科目學分
        /// </summary>
        public string SubjectCredit { get; set; }

        /// <summary>
        /// 修課科/班/學程別代碼
        /// </summary>
        public string DCLCode { get; set; }

        /// <summary>
        /// 修課班級
        /// </summary>
        public string ClassName { get; set; }
        
        /// <summary>
        /// 修課班別
        /// </summary>
        public string ClClassName { get; set; }
        
        /// <summary>
        /// 段考成績
        /// </summary>
        public string Score { get; set; }

        /// <summary>
        /// 狀態代碼
        /// </summary>
        public string Status { get; set; }
    }
}
