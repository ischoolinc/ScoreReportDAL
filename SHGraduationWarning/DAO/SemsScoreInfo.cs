using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SHGraduationWarning.DAO
{
    // 學生學期成績
    public class SemsScoreInfo
    {
        //  學期成績系統編號
        public string id { get; set; }

        //  學生系統編號
        public string StudentID { get; set; }

        //  學年度
        public string SchoolYear { get; set; }

        //  學期
        public string Semester { get; set; }

        //  成績年級
        public string GradeYear { get; set; }

        //  成績內容
        public XElement ScoreInfo { get; set; }
    }
}
