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

        public XElement MOEXml = null;

        public List<chkSubjectInfo> chkSubjectInfoList = new List<chkSubjectInfo>();

        public int calSubjDiffCount()
        {
            int value = 0;

            foreach (chkSubjectInfo subj in chkSubjectInfoList)
            {
                if (subj.ProcessStatus != "略過")
                    value++;
            }

            return value;
        }
        public int calSubjUpdateCount()
        {
            int value = 0;
            foreach (chkSubjectInfo subj in chkSubjectInfoList)
            {
                if (subj.ProcessStatus == "更新")
                    value++;
            }
            return value;
        }
        public int calSubjNoChangeCount()
        {
            int value = 0;
            foreach (chkSubjectInfo subj in chkSubjectInfoList)
            {
                if (subj.ProcessStatus == "略過")
                    value++;
            }
            return value;
        }
        public int calSubjDelCount()
        {
            int value = 0;
            foreach (chkSubjectInfo subj in chkSubjectInfoList)
            {
                if (subj.ProcessStatus == "刪除")
                    value++;
            }
            return value;
        }

        public void CheckData()
        {
            // 檢查資料
            if (MOEXml != null && ContentXML != null)
            {
                // 建立各別課程代碼索引
                Dictionary<string, List<XElement>> MOEDict = new Dictionary<string, List<XElement>>();
                Dictionary<string, List<XElement>> GPlanDict = new Dictionary<string, List<XElement>>();


                foreach (XElement elm in MOEXml.Elements("Subject"))
                {
                    if (elm.Attribute("課程代碼") != null)
                    {
                        if (!MOEDict.ContainsKey(elm.Attribute("課程代碼").Value))
                        {
                            MOEDict.Add(elm.Attribute("課程代碼").Value, new List<XElement>());
                        }

                        MOEDict[elm.Attribute("課程代碼").Value].Add(elm);
                    }
                    else
                    {
                        if (elm.Attribute("科目名稱") != null)
                        {
                            if (!MOEDict.ContainsKey(elm.Attribute("科目名稱").Value))
                            {
                                MOEDict.Add(elm.Attribute("科目名稱").Value, new List<XElement>());
                            }

                            MOEDict[elm.Attribute("科目名稱").Value].Add(elm);
                        }
                    }

                }

                foreach (XElement elm in ContentXML.Elements("Subject"))
                {
                    if (elm.Attribute("課程代碼") != null)
                    {
                        if (!GPlanDict.ContainsKey(elm.Attribute("課程代碼").Value))
                        {
                            GPlanDict.Add(elm.Attribute("課程代碼").Value, new List<XElement>());
                        }

                        GPlanDict[elm.Attribute("課程代碼").Value].Add(elm);
                    }
                    else
                    {
                        if (elm.Attribute("科目") != null)
                        {
                            if (!GPlanDict.ContainsKey(elm.Attribute("科目").Value))
                            {
                                GPlanDict.Add(elm.Attribute("科目").Value, new List<XElement>());
                            }

                            GPlanDict[elm.Attribute("科目").Value].Add(elm);
                        }

                    }
                }

                List<string> AddList = new List<string>();
                List<string> DelList = new List<string>();

                foreach (string mCo in MOEDict.Keys)
                {
                    if (!GPlanDict.ContainsKey(mCo))
                        AddList.Add(mCo);
                }

                foreach (string gCo in GPlanDict.Keys)
                {
                    if (!MOEDict.ContainsKey(gCo))
                        DelList.Add(gCo);
                }

                // 課程代碼大表有課程規劃表沒有
                foreach (string mCo in AddList)
                {
                    if (MOEDict.ContainsKey(mCo))
                    {
                        XElement elm = MOEDict[mCo][0];
                        chkSubjectInfo subj = new chkSubjectInfo();
                        subj.Domain = GetAttribute(elm, "Domain");
                        subj.Entry = GetAttribute(elm, "Entry");
                        subj.SubjectName = GetAttribute(elm, "SubjectName");

                        if (GetAttribute(elm, "RequiredBy") == "部訂")
                            subj.RequiredBy = "部定";
                        else
                            subj.RequiredBy = GetAttribute(elm, "RequiredBy");

                        subj.isRequired = GetAttribute(elm, "Required");
                        subj.CourseCode = GetAttribute(elm, "課程代碼");
                        subj.credit_period = GetAttribute(elm, "授課學期學分");
                        subj.OpenStatus = GetAttribute(elm, "開課方式");
                        subj.ProcessStatus = "更新";
                        subj.DiffStatusList.Add("多");

                        chkSubjectInfoList.Add(subj);
                    }
                }
            }
        }
        private string GetAttribute(XElement elm, string attrName)
        {
            string value = "";

            if (elm.Attribute(attrName) != null)
                value = elm.Attribute(attrName).Value;

            return value;
        }

    }
}
