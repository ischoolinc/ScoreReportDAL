using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class CourseCodeRoot
    {
        public string 實施學年度 { get; set; }
        public string 課程類型 { get; set; }
        public List<CourseCodeInfo> 課程資料 { get; set; }
    }
}
