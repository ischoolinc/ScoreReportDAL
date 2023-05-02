using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SHGraduationWarning.DAO
{
    public class SemsScoreInfo
    {
        public string id { get; set; }

        public XElement ScoreInfo { get; set; }
    }
}
