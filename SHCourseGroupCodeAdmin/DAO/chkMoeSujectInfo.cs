using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    // 課程代碼總表部分資料資料，比對時使用，因為111學年前當時不需要，如果沒有重新產生會缺少
    public class chkMoeSujectInfo
    {
        public string SubjectName { get; set; }
        public string CourseCode { get; set; }
        public string CreditPeriod { get; set; }
        public string CourseAttr { get; set; }
        public string OpenType { get; set; }
    }
}
