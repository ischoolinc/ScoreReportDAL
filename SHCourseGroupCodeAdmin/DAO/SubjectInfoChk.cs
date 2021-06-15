using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class SubjectInfoChk
    {
        /// <summary>
        /// 學年度
        /// </summary>
        public string SchoolYear { get; set; }
        /// <summary>
        ///   學期
        /// </summary>
        public string Semester { get; set; }


        /// <summary>
        /// 科目名稱
        /// </summary>
        public string SubjectName { get; set; }
        /// <summary>
        /// 科目級別
        /// </summary>
        public string SubjectLevel { get; set; }
        /// <summary>
        /// 部定校訂
        /// </summary>
        public string RequireBy { get; set; }
        /// <summary>
        /// 必修選修
        /// </summary>
        public string IsRequired { get; set; }
        /// <summary>
        /// 學分數
        /// </summary>
        public string Credit { get; set; }
        
        /// <summary>
        /// 課程代碼
        /// </summary>
        public string course_code { get; set; }
        /// <summary>
        /// 授課學期學分節數
        /// </summary>
        public string credit_period { get; set; }
        /// <summary>
        /// Memo
        /// </summary>
        public string Memo { get; set; }

        /// <summary>
        /// 群科班代碼
        /// </summary>
        public string GroupCode { get; set; }

        /// 入學年      
        /// </summary>
        public string entry_year { get; set; }

        /// <summary>
        /// 檢查學分數
        /// </summary>
        /// <returns></returns>
        public bool CheckCreditPass(Dictionary<string, string> mappingTable)
        {
            int ey = 0;
            int idx = -1;
            bool value = false;

            char[] ret = credit_period.ToCharArray();

            if (int.TryParse(entry_year, out ey))
            {
                if (ey + "" == SchoolYear && Semester == "1")
                {
                    idx = 0;
                }

                if (ey + "" == SchoolYear && Semester == "2")
                {
                    idx = 1;
                }

                if (ey + 1 + "" == SchoolYear && Semester == "1")
                {
                    idx = 2;
                }

                if (ey + 1 + "" == SchoolYear && Semester == "2")
                {
                    idx = 3;
                }

                if (ey + 2 + "" == SchoolYear && Semester == "1")
                {
                    idx = 4;
                }

                if (ey + 2 + "" == SchoolYear && Semester == "2")
                {
                    idx = 5;
                }

                // 學分數相等
                if (idx > -1 && idx < ret.Count())
                {
                    string x = ret[idx] + "";

                    // 先比是否相同，不同在比對開
                    if (x == Credit)
                    {
                        value = true;
                    }
                    else
                    {
                        // 有對開
                        if (mappingTable.ContainsKey(x))
                        {
                            if (mappingTable[x] == Credit)
                            {
                                value = true;
                            }
                        }
                    }
                }
            }

            return value;
        }
    }
}
