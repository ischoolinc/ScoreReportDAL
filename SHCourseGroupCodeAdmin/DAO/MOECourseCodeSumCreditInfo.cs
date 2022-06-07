using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class MOECourseCodeSumCreditInfo
    {
        public string EntryYear { get; set; }

        public string GroupCode { get; set; }

        public string GroupName { get; set; }

        /// <summary>
        /// 必修
        /// </summary>
        public decimal SumReqTCredit { get; set; }

        /// <summary>
        /// 選修
        /// </summary>
        public decimal SumReqFCredit { get; set; }

        /// <summary>
        /// 總計
        /// </summary>
        public decimal SumReqTotalCredit { get; set; }
    }
}
