using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class rptSCAttendCodeChkInfo
    {
        public string StudentID { get; set; }
        public string StudentNumber { get; set; }
        public string ClassName { get; set; }
        public string SeatNo { get; set; }
        public string GradeYear { get; set; }
        public string StudentName { get; set; }
        public string SchoolYear { get; set; }
        public string Semester { get; set; }
        public string CourseName { get; set; }
        public string CourseRefClass { get; set; }
        public string SubjectName { get; set; }
        public string SubjectLevel { get; set; }
        public string RequiredBy { get; set; }
        public string IsRequired { get; set; }
        public string Credit { get; set; }
        public string Period { get; set; }
        public string CourseCode { get; set; }

        public string SubjectCode { get; set; } // 原本修課紀錄的科目代碼 subject_code

        public string credit_period { get; set; }

        public string entry_year { get; set; }
        /// <summary>
        /// 分項類別
        /// </summary>
        public string ScoreType { get; set; }

        public string CourseID { get; set; }

        public string gdc_code { get; set; }
        public List<string> ErrorMsgList = new List<string>();

        // 開課方式
        public string open_type { get; set; }

        // 因成績冊新增
        public string IDNumber { get; set; }

        // 因成績冊新增
        public string BirthDayString { get; set; }
        // 因成績冊新增
        public bool hasCourseCode = false;
        // 因成績冊新增
        public bool CodePass = false;

        // 課程規劃表ID
        public string GraduationPlanID { get; set; }

        // 課程規劃表名稱
        public string GraduationPlanName { get; set; }

        // 報部科目名稱
        public string OfficialSubjectName { get; set; }

        /// <summary>
        /// 檢查學分數
        /// </summary>
        /// <returns></returns>
        public bool CheckCreditPass(Dictionary<string, string> mappingTable)
        {
            int ey = 0;
            int idx = -1;
            bool value = false;

            // 2022-03-23 Cynthia 先找 科目名稱、校部定、必選修、分項類別後，才會找到credit_period，
            // 所以當找不到 credit_period = null，也無法判斷學分數是否正確，那就不要出現學分數錯誤的提示，
            // 故直接return true，當成是正確來判斷。
            if (credit_period == null)
                return true;

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

                    // 加入使用節數來判斷，主要某些匯入課程只有節數沒有學分數
                    if (x == Period)
                    {
                        value = true;
                    }
                    else
                    {
                        // 有對開
                        if (mappingTable.ContainsKey(x))
                        {
                            if (mappingTable[x] == Period)
                            {
                                value = true;
                            }
                        }
                    }

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
