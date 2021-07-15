using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class rptGPlanClassChkInfo
    {
        public string GPID { get; set; }
        public string GPName { get; set; }
        public string GPMOECode { get; set; }
        public string GPMOEName { get; set; }
        public string ClassID { get; set; }
        public string ClassName { get; set; }
        public string ClassGDCCode { get; set; }
        public string ClassGDCName { get; set; }

        public List<string> ErrorMsgList = new List<string>();

        public void CheckData()
        {
            if (string.IsNullOrEmpty(GPMOECode))
            {
                ErrorMsgList.Add("課程規劃表缺少群科班代碼");
            }
            else if (string.IsNullOrEmpty(ClassGDCCode))
            {
                ErrorMsgList.Add("採用班級缺少群科班代碼");
            }
            else
            {
                if (GPMOECode != ClassGDCCode)
                {
                    ErrorMsgList.Add("課程規劃表與採用班級群科班代碼不同");
                }
            }
        }
    }
}
