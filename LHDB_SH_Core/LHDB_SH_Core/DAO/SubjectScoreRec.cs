using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LHDB_SH_Core.DAO
{
    /// <summary>
    /// 成績名冊紀錄
    /// </summary>
    public class SubjectScoreRec
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
        public string ClCode { get; set; }
        
        /// <summary>
        /// 原始成績
        /// </summary>
        public string Score { get; set; }

        /// <summary>
        /// 補考成績
        /// </summary>
        public string MuScore { get; set; }
        
        /// <summary>
        /// 補修科目代碼
        /// </summary>
        public string ScSubjectCode { get; set; }

        /// <summary>
        /// 補修科目開設年級及學期
        /// </summary>
        public string ScSubjectYearSemester { get; set; }
        
        /// <summary>
        /// 補修成績
        /// </summary>
        public string ScScore { get; set; }
                
        /// <summary>
        /// 重修科目代碼
        /// </summary>
        public string ReSubjectCode { get; set; }

        /// <summary>
        /// 重修科目開設年級及學期
        /// </summary>
        public string ReSubjectYearSemester { get; set; }
        
        /// <summary>
        /// 重修成績
        /// </summary>
        public string ReScore { get; set; }

    }
}
