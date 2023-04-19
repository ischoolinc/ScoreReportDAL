﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseCodeCheckAndUpdate.DAO
{
    public class StudSCAttendInfo
    {
        public string StudentID { get; set; } // 學生系統編號
        public string SCAttendID { get; set; } // 修課系統編號
        public string SchoolYear { get; set; } // 學年度
        public string Semester { get; set; } // 學期
        public string GradeYear { get; set; } // 班級年級
        public string ClassName { get; set; } // 班級
        public string SeatNo { get; set; } // 座號
        public string Name { get; set; } // 姓名
        public string CourseName { get; set; } // 課程名稱
        public string SubjectName { get; set; } // 科目名稱
        public string SubjectLevel { get; set; } // 科目級別
        public string RequiredBy { get; set; } // 校部定
        public string Required { get; set; } // 必選修
        public string Credit { get; set; } // 學分
        public string SC_CourseCode { get; set; } // 修課課程代碼
        public string GP_CourseCode { get; set; } // 課程規劃課程代碼
        public string GPName { get; set; } // 使用課程規畫表

        public string StudentNumber { get; set; } // 學號
        public string status { get; set; } // 學生狀態
    }
}