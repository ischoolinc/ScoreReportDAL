using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHSemsSubjectCheckEdit.DAO
{
    public class StudCheckDataInfo
    {
        public string StudentID { get; set; } // 學生系統編號

        // 上下學期科目名稱
        public List<string> SubjectNameList = new List<string>();

        // 上學期
        public Dictionary<string, Dictionary<string, List<StudSubjectInfo>>> StudSubjectInfoDict1 = new Dictionary<string, Dictionary<string, List<StudSubjectInfo>>>();

        // 下學期
        public Dictionary<string, Dictionary<string, List<StudSubjectInfo>>> StudSubjectInfoDict2 = new Dictionary<string, Dictionary<string, List<StudSubjectInfo>>>();

        // 檢查相同科目名稱1
        public Dictionary<string, List<StudSubjectInfo>> SameSubjectNameDict1 = new Dictionary<string, List<StudSubjectInfo>>();

        // 檢查相同科目名稱2
        public Dictionary<string, List<StudSubjectInfo>> SameSubjectNameDict2 = new Dictionary<string, List<StudSubjectInfo>>();
    }
}
