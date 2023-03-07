using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class CourseGroupSetting
    {
        /// <summary>
        /// 課程群組名稱
        /// </summary>
        public string CourseGroupName { get; set; }

        /// <summary>
        /// 課程群組修課學分數
        /// </summary>
        public string CourseGroupCredit { get; set; }

        /// <summary>
        /// 課程群組顏色設定
        /// </summary>
        public Color CourseGroupColor { get; set; }

        /// <summary>
        /// 是否為學年課程群組
        /// </summary>
        public bool IsSchoolYearCourseGroup { get; set; }

        /// <summary>
        /// 課程群組設定Xml Element
        /// </summary>
        public XElement CourseGroupElement { get; set; }

    }
}
