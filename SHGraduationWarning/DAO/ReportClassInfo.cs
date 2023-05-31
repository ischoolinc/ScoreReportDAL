using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHGraduationWarning.DAO
{
    // 報表使用班級資訊
    public class ReportClassInfo
    {
        // 班級系統編號
        public string ClassID { get; set; }
        // 班級名稱
        public string ClassName { get; set; }

        // 報表裡班級學生資料
        public List<ReportStudentInfo> StudentInfoList = new List<ReportStudentInfo>();

        // 班導師
        public string TeacherName { get; set; }

        // 報表欄位索引
        public Dictionary<string, int> dicColumnIndex = new Dictionary<string, int>();

        // 符合規則名稱
        public List<string> RuleList = new List<string>();
    }
}
