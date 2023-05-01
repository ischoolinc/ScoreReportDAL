using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHGraduationWarning.DAO
{
    // 課程規劃表
    public class GPlanInfo
    {
        // id
        public string ID { get; set; }

        // name
        public string Name { get; set; }

        // 科目資料
        public Dictionary<string, GPlanSubjectInfo> SubjectsDict { get; set; } = new Dictionary<string, GPlanSubjectInfo>();
    }
}
