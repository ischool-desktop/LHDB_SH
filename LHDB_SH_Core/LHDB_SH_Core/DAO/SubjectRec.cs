using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LHDB_SH_Core.DAO
{
    /// <summary>
    /// 科目名冊使用
    /// </summary>
    public class SubjectRec
    {
        /// <summary>
        /// 修課別
        /// </summary>
        public string Required { get; set; }

        /// <summary>
        /// 類別種類
        /// </summary>
        public string CourseType { get; set; }
        
        /// <summary>
        /// 領域/分類
        /// </summary>
        public string DomainType { get; set; }
        
        /// <summary>
        /// 群別
        /// </summary>
        public string Group { get; set; }

        /// <summary>
        /// 科/班/學程別
        /// </summary>
        public string DLCCode { get; set; }
        
        /// <summary>
        /// 特色班/實驗班名稱
        /// </summary>
        public string SpcType { get; set; }

        /// <summary>
        /// 科目名稱
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// 科目代碼
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 一年級學分
        /// </summary>
        public string CreditG1 { get; set; }

        /// <summary>
        /// 二年級學分
        /// </summary>
        public string CreditG2 { get; set; }

        /// <summary>
        /// 三年級學分
        /// </summary>
        public string CreditG3 { get; set; }

        /// <summary>
        /// 是否計算學分
        /// </summary>
        public string isCalc { get; set; }

    }
}
