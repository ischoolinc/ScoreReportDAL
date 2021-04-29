using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;


namespace SHCourseGroupCodeAdmin.DAO
{
    public class GPlanInfo
    {
        public string ID { get; set; }
        public string Name { get; set; }

        public List<GPCourseInfo> CourseInfoList = new List<GPCourseInfo>();

        /// <summary>
        /// 解析 XML
        /// </summary>
        /// <param name="strXML"></param>
        public void ParseContentToCourseInfoList()
        {
            try
            {
                /*
                 * <Subject Category="" Credit="4" Domain="語文" Entry="學業" GradeYear="1" Level="1" FullName="國語文 I" NotIncludedInCalc="False" NotIncludedInCredit="False" Required="必修" RequiredBy="部訂" Semester="1" SubjectName="國語文" 課程代碼="108041305H11101A1010101" 課程類別="部定必修" 開課方式="原班級" 科目屬性="一般科目" 領域名稱="語文" 課程名稱="國語文" 學分="4" 授課學期學分="444440">
		<Grouping RowIndex="1" startLevel="1"/>
	</Subject>
                 * */
                CourseInfoList.Clear();
                XElement elmRoot = XElement.Parse(Content);
                foreach (XElement elm in elmRoot.Elements("Subject"))
                {
                    GPCourseInfo data = new GPCourseInfo();
                    data.CourseName = GetAttribute(elm, "課程名稱");
                    data.Credit = GetAttribute(elm, "學分");
                    data.Entry = GetAttribute(elm, "Entry");
                    data.GPName = Name;
                    data.Required = GetAttribute(elm, "Required");
                    data.RequiredBy = GetAttribute(elm, "RequiredBy");

                    if (data.RequiredBy == "部訂")
                        data.RequiredBy = "部定";

                    data.SubjectName = GetAttribute(elm, "SubjectName");
                    data.GradeYear = GetAttribute(elm, "GradeYear");
                    data.Level = GetAttribute(elm, "Level");
                    data.NotIncludedInCalc = GetAttributeBool(elm, "NotIncludedInCalc");
                    data.NotIncludedInCredit = GetAttributeBool(elm, "NotIncludedInCredit");
                    CourseInfoList.Add(data);

                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        /// <summary>
        /// 原始 XML 資料
        /// </summary>
        public string Content { get; set; }

        private string GetAttribute(XElement elm, string attName)
        {
            string value = "";
            if (elm.Attribute(attName) != null)
                value = elm.Attribute(attName).Value;

            return value;
        }

        private bool GetAttributeBool(XElement elm, string attName)
        {
            bool value = false;

            if (elm.Attribute(attName) != null)
            {
                bool bo;
                if (bool.TryParse(elm.Attribute(attName).Value, out bo))
                {

                }
                value = bo;
            }

            return value;
        }

    }
}
