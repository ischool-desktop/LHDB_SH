using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LHDB_SH_Core.DAO
{
    /// <summary>
    /// 學生資料統計
    /// </summary>
    public class StudentBaseRecCount
    {
        /// <summary>
        /// 學校代碼
        /// </summary>
        public string SchoolCode { get; set; }

        /// <summary>
        /// 學年度
        /// </summary>
        public int SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        public int Semester { get; set; }

        /// <summary>
        /// 名冊別
        /// </summary>
        public string DocType { get; set; }

        /// <summary>
        /// 科/班/學程別代碼
        /// </summary>
        public string DCLCode { get; set; }

        /// <summary>
        /// 年級班級代碼
        /// </summary>
        public string ClassCode { get; set; }
        
        /// <summary>
        /// 部別
        /// </summary>
        public string DepCode { get; set; }
        
        /// <summary>
        /// 班別
        /// </summary>
        public string ClCode { get; set; }

        /// <summary>
        /// 班級人數
        /// </summary>
        public int StudentCount { get; set; }
    }
}
