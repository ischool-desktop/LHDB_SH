using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LHDB_SH_Core.DAO
{
    /// <summary>
    /// 學生基本資料
    /// </summary>
    public class StudentBaseRec
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        public string StudentNumber { get; set; }

        /// <summary>
        /// 身分證號
        /// </summary>
        public string IDNumber { get; set; }

        /// <summary>
        /// 註1
        /// </summary>
        public string Remak1 { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// 性別代碼
        /// </summary>
        public string GenderCode { get; set; }

        /// <summary>
        /// 出生日期
        /// </summary>
        public string BirthDate { get; set; }
        
        /// <summary>
        /// 所屬學校代碼
        /// </summary>
        public string SchoolCode { get; set; }
        
        /// <summary>
        /// 科/班/學程別代碼
        /// </summary>
        public string DCLCode { get; set; }
        
        /// <summary>
        /// 部別
        /// </summary>
        public string DepCode{ get; set; }
        
        /// <summary>
        /// 班別
        /// </summary>
        public string ClCode { get; set; }

        /// <summary>
        /// 班級座號代碼
        /// </summary>
        public string ClassSeatCode { get; set; }
        
        /// <summary>
        /// 特殊身分代碼
        /// </summary>
        public string SpCode { get; set; }
        
        /// <summary>
        /// 原身分證號
        /// </summary>
        public string OrIDNumber { get; set; }
        
        /// <summary>
        /// 原出生日期
        /// </summary>
        public string OrBirthDate { get; set; }
                
        /// <summary>
        /// 學籍狀態代碼
        /// </summary>
        public string PermrecCode { get; set; }
        
        /// <summary>
        /// 學籍狀態生效日期
        /// </summary>
        public string PermrecDate { get; set; }
    }
}
