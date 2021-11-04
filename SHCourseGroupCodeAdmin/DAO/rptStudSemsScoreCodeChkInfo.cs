using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
   public class rptStudSemsScoreCodeChkInfo
    {
        public string StudentID { get; set; }
        public string StudentNumber { get; set; }
        public string ClassName { get; set; }
        public string SeatNo { get; set; }
        public string GradeYear { get; set; }

        /// <summary>
        /// 成績年級
        /// </summary>
        public string SemsGradeYear { get; set; }

        public string StudentName { get; set; }
        public string SchoolYear { get; set; }
        public string Semester { get; set; }
        public string SubjectName { get; set; }
        public string SubjectLevel { get; set; }
        public string RequiredBy { get; set; }
        public string IsRequired { get; set; }
        public string Credit { get; set; }
        public string Period { get; set; }
        public string CourseCode { get; set; }
        public string credit_period { get; set; }

        public string entry_year { get; set; }

        public string gdc_code { get; set; }
        public List<string> ErrorMsgList = new List<string>();

        /// <summary>
        /// 檢查學分數
        /// </summary>
        /// <returns></returns>
        public bool CheckCreditPass(Dictionary<string, string> mappingTable)
        {
            int ey = 0;
            int idx = -1;
            bool value = false;

            if (credit_period == null)
                return false;
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
