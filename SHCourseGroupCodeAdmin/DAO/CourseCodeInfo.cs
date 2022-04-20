using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class CourseCodeInfo
    {
        //"課程代碼": "108011302H11101A1010101",
        //        "科目名稱": "國語文",
        //        "授課學期學分節數": "444440",
        //        "授課學期開課方式": "00000-",
        //        "課程屬性": "1101"

        public string 課程代碼 { get; set; }      

        public string 科目名稱 {get;set; }
        public string 授課學期學分節數 { get; set; }
        public string 授課學期開課方式 { get; set; }
        public string 課程屬性 { get; set; }

    }
}
