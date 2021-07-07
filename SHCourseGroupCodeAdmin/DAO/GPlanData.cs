using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SHCourseGroupCodeAdmin.DAO
{
    /// <summary>
    /// 班級課程規劃表資料
    /// </summary>
    public class GPlanData
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public Dictionary<string, string> UsedClassIDNameDict = new Dictionary<string, string>();
        public string MOEGroupCode { get; set; }
        public XElement ContentXML { get; set; }

    }
}
