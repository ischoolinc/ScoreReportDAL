using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class chkGPlanInfo
    {
        // 課規系統編號
        public string ID { get; set; }

        // 課程規劃表名稱
        public string Name { get; set; }

        // 入學年
        public string EntryYear { get; set; }

        // 課規內容 XML
        public XElement ContentXML { get; set; }

        // 科目清單(xml)
        public Dictionary<string, XElement> SubjectXMLDict = new Dictionary<string, XElement>();

        // 科目清單
        public Dictionary<string, chkGPSubjectInfo> SubjectDict = new Dictionary<string, chkGPSubjectInfo>();

        // 轉換科目
        public void ParseSubjectDict()
        {
            SubjectXMLDict.Clear();
            SubjectDict.Clear();
            if (ContentXML != null)
            {
                foreach (XElement elm in ContentXML.Elements("Subject"))
                {
                    // 使用科目名稱+級別當 key
                    string SubjectName = elm.Attribute("SubjectName").Value;
                    string Level = "";
                    if (elm.Attribute("Level") != null)
                        Level = elm.Attribute("Level").Value;


                    string key = SubjectName + "_" + Level;

                    if (!SubjectXMLDict.ContainsKey(key))
                        SubjectXMLDict.Add(key, elm);

                    try
                    {
                        chkGPSubjectInfo subj = new chkGPSubjectInfo();
                        subj.SubjectName = elm.Attribute("SubjectName").Value;
                        subj.CourseCode = elm.Attribute("課程代碼").Value;
                        subj.Entry = elm.Attribute("Entry").Value;
                        subj.Domain = elm.Attribute("Domain").Value;
                        subj.isRequired = elm.Attribute("Required").Value;
                        subj.RequiredBy = elm.Attribute("RequiredBy").Value;
                        if (subj.RequiredBy == "部訂")
                            subj.RequiredBy = "部定";
                        subj.Credit = elm.Attribute("Credit").Value;
                        subj.credit_period = elm.Attribute("授課學期學分").Value;
                        
                        if (elm.Attribute("OpenType") != null)
                            subj.open_type = elm.Attribute("OpenType").Value;

                        if (elm.Attribute("OfficialSubjectName") != null)
                            subj.OfficialSubjectName = elm.Attribute("OfficialSubjectName").Value;


                        if (!SubjectDict.ContainsKey(key))
                            SubjectDict.Add(key, subj);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                }
            }

        }

        // 轉換ContentXML
        public void ParseContentXML(string content)
        {
            try
            {
                ContentXML = XElement.Parse(content);

                // 解析入學年
                if (ContentXML != null)
                {
                    if (ContentXML.Attribute("EntryYear") != null)
                    {
                        EntryYear = ContentXML.Attribute("EntryYear").Value;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

    }
}
