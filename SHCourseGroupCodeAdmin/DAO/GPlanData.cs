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

        // 建立各別課程代碼索引
        public Dictionary<string, List<XElement>> MOEDict = new Dictionary<string, List<XElement>>();
        public Dictionary<string, List<XElement>> GPlanDict = new Dictionary<string, List<XElement>>();

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
        public int calSubjAddCount()
        {
            int value = 0;
            foreach (chkSubjectInfo subj in chkSubjectInfoList)
            {
                if (subj.ProcessStatus == "新增")
                    value++;
            }
            return value;
        }


        public void CheckData()
        {
            // 檢查資料
            if (MOEXml != null && ContentXML != null)
            {

                MOEDict.Clear();
                GPlanDict.Clear();

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

                // 用課程代碼比對，大表有課程規劃表沒有需要新增
                List<string> AddList = new List<string>();
                // 用課程規劃表比對大表沒有需要刪除
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
                        subj.ProcessStatus = "新增";
                        subj.DiffStatusList.Add("缺");
                        subj.MOEXml = MOEDict[mCo];

                        chkSubjectInfoList.Add(subj);
                    }
                }

                // 大表沒有，課程規劃表有 多出來
                foreach (string mCo in DelList)
                {
                    if (GPlanDict.ContainsKey(mCo))
                    {
                        XElement elm = GPlanDict[mCo][0];
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
                        subj.ProcessStatus = "刪除";
                        subj.DiffStatusList.Add("多");
                        subj.GPlanXml = GPlanDict[mCo];
                        chkSubjectInfoList.Add(subj);
                    }
                }

                // 以大表為主，當課程代碼相同，比對裡面資料是否相同
                foreach (string mCo in MOEDict.Keys)
                {
                    if (GPlanDict.ContainsKey(mCo))
                    {
                        XElement elm = GPlanDict[mCo][0];
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

                        if (MOEDict[mCo].Count != GPlanDict[mCo].Count)
                        {
                            subj.DiffStatusList.Add("開課學期數不同");
                            subj.DiffMessageList.Add("開課學期數：課程代碼表「" + MOEDict[mCo].Count + "」、課程規劃表「" + GPlanDict[mCo].Count + "」");
                        }

                        foreach (XElement elmM in MOEDict[mCo])
                        {
                            foreach (XElement elmG in GPlanDict[mCo])
                            {

                                if (elmM.Attribute("Domain").Value != elmG.Attribute("Domain").Value)
                                {
                                    if (!subj.DiffStatusList.Contains("領域不同"))
                                        subj.DiffStatusList.Add("領域不同");

                                    string msg = "領域：課程代碼表「" + elmM.Attribute("Domain").Value + "」、課程規劃表「" + elmG.Attribute("Domain").Value + "」";
                                    if (!subj.DiffMessageList.Contains(msg))
                                        subj.DiffMessageList.Add(msg);
                                }

                                if (elmM.Attribute("Entry").Value != elmG.Attribute("Entry").Value)
                                {
                                    if (!subj.DiffStatusList.Contains("分項類別不同"))
                                        subj.DiffStatusList.Add("分項類別不同");

                                    string msg = "分項類別：課程代碼表「" + elmM.Attribute("Entry").Value + "」、課程規劃表「" + elmG.Attribute("Entry").Value + "」";
                                    if (!subj.DiffMessageList.Contains(msg))
                                        subj.DiffMessageList.Add(msg);

                                }

                                if (elmM.Attribute("SubjectName").Value != elmG.Attribute("SubjectName").Value)
                                {
                                    if (!subj.DiffStatusList.Contains("科目不同"))
                                        subj.DiffStatusList.Add("科目不同");

                                    string msg = "科目名稱：課程代碼表「" + elmM.Attribute("SubjectName").Value + "」、課程規劃表「" + elmG.Attribute("SubjectName").Value + "」";
                                    if (!subj.DiffMessageList.Contains(msg))
                                        subj.DiffMessageList.Add(msg);

                                }

                                if (elmM.Attribute("RequiredBy").Value != elmG.Attribute("RequiredBy").Value)
                                {
                                    if (!subj.DiffStatusList.Contains("校訂部定不同"))
                                        subj.DiffStatusList.Add("校訂部定不同");

                                    string msg = "校部定：課程代碼表「" + elmM.Attribute("RequiredBy").Value + "」、課程規劃表「" + elmG.Attribute("RequiredBy").Value + "」";
                                    if (!subj.DiffMessageList.Contains(msg))
                                        subj.DiffMessageList.Add(msg);

                                }

                                if (elmM.Attribute("Required").Value != elmG.Attribute("Required").Value)
                                {
                                    if (!subj.DiffStatusList.Contains("必選修不同"))
                                        subj.DiffStatusList.Add("必選修不同");

                                    string msg = "必選修：課程代碼表「" + elmM.Attribute("Required").Value + "」、課程規劃表「" + elmG.Attribute("Required").Value + "」";
                                    if (!subj.DiffMessageList.Contains(msg))
                                        subj.DiffMessageList.Add(msg);

                                }

                                if (elmM.Attribute("開課方式").Value != elmG.Attribute("開課方式").Value)
                                {
                                    if (!subj.DiffStatusList.Contains("開課方式不同"))
                                        subj.DiffStatusList.Add("開課方式不同");

                                    string msg = "開課方式：課程代碼表「" + elmM.Attribute("開課方式").Value + "」、課程規劃表「" + elmG.Attribute("開課方式").Value + "」";
                                    if (!subj.DiffMessageList.Contains(msg))
                                        subj.DiffMessageList.Add(msg);

                                }

                                // 檢查學分數不同
                                if (elmM.Attribute("GradeYear").Value == elmG.Attribute("GradeYear").Value && elmM.Attribute("Semester").Value == elmG.Attribute("Semester").Value)
                                {
                                    string ss = "下";
                                    if (elmG.Attribute("Semester").Value == "1")
                                        ss = "上";
                                    string m1 = elmM.Attribute("GradeYear").Value + ss;
                                    if (elmM.Attribute("Credit").Value != elmG.Attribute("Credit").Value)
                                    {
                                        if (!subj.DiffStatusList.Contains("學分數不同"))
                                            subj.DiffStatusList.Add("學分數不同");

                                        string msg = m1 + "學分數：課程代碼表「" + elmM.Attribute("Credit").Value + "」、課程規劃表「" + elmG.Attribute("Credit").Value + "」";
                                        if (!subj.DiffMessageList.Contains(msg))
                                            subj.DiffMessageList.Add(msg);

                                    }

                                }
                            }
                        }

                        if (subj.DiffStatusList.Count > 0)
                            subj.ProcessStatus = "更新";
                        else
                            subj.ProcessStatus = "略過";


                        subj.GPlanXml = GPlanDict[mCo];
                        subj.MOEXml = MOEDict[mCo];
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
