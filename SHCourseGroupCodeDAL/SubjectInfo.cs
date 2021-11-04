using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeDAL
{
    public class SubjectInfo
    {
        public SubjectInfo()
        {

        }

        public void SetSubjectInfo(string subjectName, string entryYear, string requireBy, string required, string courseAttr)
        {
            SubjectName = subjectName;
            EntryYear = entryYear;
            RequireBy = requireBy;
            Required = required;
            course_attr = courseAttr;

            if (courseAttr.Length > 2)
            {
                string code2 = course_attr.Substring(1, 1);
                if (code2 == "2")
                {
                    Entry = "專業科目";
                }
                else if (code2 == "3")
                {
                    Entry = "實習科目";
                }
                else
                    Entry = "學業";
            }
            ParseCourseAttr();
            // 科目名稱+校部定+必選修
            //  SubjectKey = SubjectName.Trim() + "_" + RequireBy + "_" + Required;
            SubjectKey = Entry + "_" + SubjectName.Trim() + "_" + RequireBy + "_" + Required;
        }

        /// <summary>
        /// 比對用
        /// </summary>
        private string SubjectKey = "";
        /// <summary>
        /// 科目名稱
        /// </summary>
        public string SubjectName = "";

        /// <summary>
        /// 入學年
        /// </summary>
        private string EntryYear = "";

        /// <summary>
        /// 校部定
        /// </summary>
        public string RequireBy = "";

        /// <summary>
        /// 必選修
        /// </summary>
        public string Required = "";

        /// <summary>
        /// 群組代碼
        /// </summary>
        private string GroupCode = "";

        public string Entry = "";

        //    private string credit_period = "";

        /// <summary>
        /// 課程屬性，新規格是將有異動資料寫在這，分別：課程類別(1)+科目屬性(1)+領域名稱(2)
        /// </summary>
        private string course_attr = "";

        /// <summary>
        /// 讀取最新修改回寫校部定必選修
        /// </summary>
        public void ParseCourseAttr()
        {
            // 解析課程屬性第1位
            if (course_attr.Length > 1)
            {
                // 1   部定必修
                // 2   校訂必修
                string co = course_attr.Substring(0, 1);
                if (co == "1")
                {
                    RequireBy = "部定";
                    Required = "必修";
                }
                else if (co == "2")
                {
                    RequireBy = "校訂";
                    Required = "必修";
                }
                else
                {
                    RequireBy = "校訂";
                    Required = "選修";
                }
            }
        }


        /// <summary>
        /// 課程代碼
        /// </summary>
        private string CourseCode = "";

        public string GetSubjectKey()
        {
            return SubjectKey;
        }

        public string GetCourseCode()
        {
            return CourseCode;
        }


        public string GetSubjectName()
        {
            return SubjectName;
        }

        public string GetRequireBy() { return RequireBy; }

        public string GetRequired() { return Required; }


        public void SetCode(string groupCode, string courseCode)
        {
            GroupCode = groupCode;
            CourseCode = courseCode;
        }

        /// <summary>
        ///  補修比對使用
        /// </summary>
        public string credit_period { get; set; }

    }
}
