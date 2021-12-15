﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Data;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class GPlanInfo108
    {
        /// <summary>
        /// 對應課程規劃表ID
        /// </summary>
        public string RefGPID { get; set; }
        /// <summary>
        /// 對應課程規劃表名稱
        /// </summary>
        public string RefGPName { get; set; }

        /// <summary>
        /// 對應課程規劃表XML Content
        /// </summary>
        public string RefGPContent { get; set; }

        /// <summary>
        /// 課程代碼代表轉換後
        /// </summary>
        public XElement MOEXml { get; set; }

        /// <summary>
        /// 課程規劃表內 XML
        /// </summary>
        public XElement RefGPContentXml { get; set; }

        /// <summary>
        /// 入學年
        /// </summary>
        public string EntrySchoolYear { get; set; }
        /// <summary>
        /// 群科班名稱
        /// </summary>
        public string GDCName { get; set; }
        /// <summary>
        /// 群科班代碼
        /// </summary>
        public string GDCCode { get; set; }
        /// <summary>
        /// 狀態
        /// </summary>
        public string Status { get; set; }
        /// <summary>
        /// 排序使用
        /// </summary>
        public int OrderByInt { get; set; }

        /// <summary>
        /// 對照到課程規劃表
        /// </summary>
        public List<DataRow> GPlanList = new List<DataRow>();
      
        /// <summary>
        /// 課程代碼代表原始資料
        /// </summary>
        public List<MOECourseCodeInfo> MOECourseCodeInfoList = new List<MOECourseCodeInfo>();

        // 建立各別課程代碼索引
        public Dictionary<string, List<XElement>> MOEDict = new Dictionary<string, List<XElement>>();
        public Dictionary<string, List<XElement>> GPlanDict = new Dictionary<string, List<XElement>>();
        public List<chkSubjectInfo> chkSubjectInfoList = new List<chkSubjectInfo>();
        /// <summary>
        /// 解析狀態
        /// </summary>
        public void ParseStatus()
        {
            Status = "無變動";
            if (string.IsNullOrEmpty(RefGPID))
            {
                Status = "新增";
            }
            else
            {
                if (calSubjDiffCount() > 0)
                    Status = "更新";
            }

        }

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

        /// <summary>
        /// 解析課程代碼大表ToXML
        /// </summary>
        public void ParseMOEXml()
        {
            MOEXml = new XElement("GraduationPlan");

            /*
             	<Subject Category="" Credit="4" Domain="語文" Entry="學業" GradeYear="1" Level="1" FullName="國語文 I" NotIncludedInCalc="False" NotIncludedInCredit="False" Required="必修" RequiredBy="部訂" Semester="1" SubjectName="國語文" 課程代碼="109041305H11101B1010101" 課程類別="部定必修" 開課方式="原班級" 科目屬性="一般科目" 領域名稱="語文" 課程名稱="國語文" 學分="4" 授課學期學分="444440">
		<Grouping RowIndex="1" startLevel="1"/>
	</Subject>
             
             */

            int RowIndex = 1, startLevel = 1;

            // 在課程群組代碼取前3，填入入學年
            if (GDCCode.Length > 3)
                MOEXml.SetAttributeValue("EntryYear", GDCCode.Substring(0, 3));

            Dictionary<string, string> CreditMappingTableDict = Utility.GetCreditMappingTable();

            Dictionary<string, int> chkCourseCodeCount = new Dictionary<string, int>();

            // 取得領域對照
            Dictionary<string, string> domainMappingDict = Utility.GetDomainNameMapping();


            // 讀取群科班資料
            foreach (MOECourseCodeInfo data in MOECourseCodeInfoList)
            {
                // 因為將來要將級別概念移除，開始級別都是1

                // 判斷開始級別，解析 credit_period
                char[] cp = data.credit_period.ToArray();

                int idx = 1;
                string strGearYear = "", strSemester = "";
                foreach (char c in cp)
                {
                    string credit = c + "";

                    if (idx == 1)
                    {
                        strGearYear = "1";
                        strSemester = "1";
                    }
                    else if (idx == 2)
                    {
                        strGearYear = "1";
                        strSemester = "2";
                    }
                    else if (idx == 3)
                    {
                        strGearYear = "2";
                        strSemester = "1";
                    }
                    else if (idx == 4)
                    {
                        strGearYear = "2";
                        strSemester = "2";
                    }
                    else if (idx == 5)
                    {
                        strGearYear = "3";
                        strSemester = "1";
                    }
                    else if (idx == 6)
                    {
                        strGearYear = "3";
                        strSemester = "2";
                    }
                    else
                    {

                    }

                    // 學分格式0 不放入
                    if (credit == "0")
                    {
                        idx++;
                        continue;
                    }


                    XElement subjElm = new XElement("Subject");
                    subjElm.SetAttributeValue("Category", "");
                    if (CreditMappingTableDict.ContainsKey(credit))
                    {
                        subjElm.SetAttributeValue("Credit", CreditMappingTableDict[credit]);
                    }
                    else
                    {
                        subjElm.SetAttributeValue("Credit", credit);
                    }

                    subjElm.SetAttributeValue("Domain", "");
                    // 處理領域對照
                    if (data.course_code.Length > 22)
                    {
                        string domainCode = data.course_code.Substring(19, 2);
                        if (domainMappingDict.ContainsKey(domainCode))
                        {
                            subjElm.SetAttributeValue("Domain", domainMappingDict[domainCode]);
                        }
                    }

                    subjElm.SetAttributeValue("Entry", "學業");
                    subjElm.SetAttributeValue("GradeYear", strGearYear);
                    subjElm.SetAttributeValue("Level", idx);
                    subjElm.SetAttributeValue("FullName", SubjFullName(data.subject_name, idx));
                    subjElm.SetAttributeValue("NotIncludedInCalc", "False");
                    subjElm.SetAttributeValue("NotIncludedInCredit", "False");
                    subjElm.SetAttributeValue("Required", data.is_required);
                    if (data.require_by == "部定")
                        subjElm.SetAttributeValue("RequiredBy", "部訂");
                    else
                        subjElm.SetAttributeValue("RequiredBy", data.require_by);

                    subjElm.SetAttributeValue("Semester", strSemester);
                    subjElm.SetAttributeValue("SubjectName", data.subject_name);
                    subjElm.SetAttributeValue("課程代碼", data.course_code);

                    if (!chkCourseCodeCount.ContainsKey(data.course_code))
                        chkCourseCodeCount.Add(data.course_code, 0);

                    chkCourseCodeCount[data.course_code]++;

                    subjElm.SetAttributeValue("課程類別", data.course_type);

                    string otStr = "跨班";
                    // 判斷是否原班上課
                    bool chkOt = false;
                    if (data.open_type != null && data.open_type.Length == 6)
                    {
                        char[] ot = data.open_type.ToArray();
                        int opIdx = 0;
                        if (strGearYear == "1" && strSemester == "1")
                            opIdx = 0;

                        if (strGearYear == "1" && strSemester == "2")
                            opIdx = 1;

                        if (strGearYear == "2" && strSemester == "1")
                            opIdx = 2;
                        if (strGearYear == "2" && strSemester == "2")
                            opIdx = 3;

                        if (strGearYear == "3" && strSemester == "1")
                            opIdx = 4;
                        if (strGearYear == "3" && strSemester == "2")
                            opIdx = 5;

                        if (ot[opIdx] == '0' || ot[opIdx] == 'A')
                        {
                            chkOt = true;
                        }

                    }

                    if (chkOt)
                        otStr = "原班";

                    subjElm.SetAttributeValue("開課方式", otStr);

                    subjElm.SetAttributeValue("領域名稱", "");
                    subjElm.SetAttributeValue("課程名稱", data.subject_name);
                    subjElm.SetAttributeValue("學分", credit);
                    subjElm.SetAttributeValue("授課學期學分", data.credit_period);

                    // 解析課程屬性，修正會放在這
                    if (data.course_attr != null && data.course_attr.Length == 4)
                    {
                        // 校部定必選修
                        string code1 = data.course_attr.Substring(0, 1);
                        if (code1 == "1")
                        {
                            // 部定必修
                            subjElm.SetAttributeValue("RequiredBy", "部訂");
                            subjElm.SetAttributeValue("Required", "必修");
                        }
                        else if (code1 == "2")
                        {
                            subjElm.SetAttributeValue("RequiredBy", "校訂");
                            subjElm.SetAttributeValue("Required", "必修");
                        }
                        else
                        {
                            subjElm.SetAttributeValue("RequiredBy", "校訂");
                            subjElm.SetAttributeValue("Required", "選修");
                        }

                        // 分項
                        string code2 = data.course_attr.Substring(1, 1);
                        if (code2 == "2")
                        {
                            subjElm.SetAttributeValue("Entry", "專業科目");
                        }
                        else if (code2 == "3")
                        {
                            subjElm.SetAttributeValue("Entry", "實習科目");
                        }
                        else
                        {
                            subjElm.SetAttributeValue("Entry", "學業");
                        }

                        string code3 = data.course_attr.Substring(2, 2);
                        // 領域
                        if (domainMappingDict.ContainsKey(code3))
                        {
                            subjElm.SetAttributeValue("Domain", domainMappingDict[code3]);
                        }

                    }



                    // row
                    XElement subjGroupElm = new XElement("Grouping");
                    subjGroupElm.SetAttributeValue("RowIndex", RowIndex);
                    subjGroupElm.SetAttributeValue("startLevel", startLevel);
                    subjElm.Add(subjGroupElm);

                    MOEXml.Add(subjElm);
                    idx++;
                }
                RowIndex++;
            }

            //// 重新排列科目級別
            //Dictionary<string, int> tmpSubjLevelDict = new Dictionary<string, int>();

            //foreach (XElement elm in MOEXml.Elements("Subject"))
            //{
            //    string subj = elm.Attribute("SubjectName").Value;

            //    if (!tmpSubjLevelDict.ContainsKey(subj))
            //        tmpSubjLevelDict.Add(subj, 0);

            //    tmpSubjLevelDict[subj] += 1;

            //    elm.SetAttributeValue("FullName", SubjFullName(subj, tmpSubjLevelDict[subj]));
            //    elm.SetAttributeValue("Level", tmpSubjLevelDict[subj]);

            //}

            //// 重新整理開始級別
            //Dictionary<string, string> tmpStartLevel = new Dictionary<string, string>();
            //foreach (XElement elm in MOEXml.Elements("Subject"))
            //{
            //    string subjName = elm.Attribute("SubjectName").Value;

            //    string rowIdx = elm.Element("Grouping").Attribute("RowIndex").Value;

            //    if (!tmpStartLevel.ContainsKey(subjName))
            //        tmpStartLevel.Add(subjName, rowIdx);
            //    else 
            //    {
            //        if (tmpStartLevel[subjName] != rowIdx)
            //        {
            //            // 設定開始級別是目前級別
            //            elm.Element("Grouping").SetAttributeValue("startLevel", elm.Attribute("Level").Value);
            //            tmpStartLevel[subjName] = rowIdx;
            //        }
            //    }
            //}
        }

        private string SubjFullName(string SubjectName, int level)
        {
            string lev = "";
            if (level == 1)
                lev = " I";

            if (level == 2)
                lev = " II";

            if (level == 3)
                lev = " III";

            if (level == 4)
                lev = " IV";

            if (level == 5)
                lev = " V";

            if (level == 6)
                lev = " VI";

            string value = SubjectName + lev;

            return value;
        }

        /// <summary>
        /// 解析目前課程規劃 ToXML
        /// </summary>
        public void ParseRefGPContentXml()
        {
            try
            {
                // 解析第一筆
                if (GPlanList.Count > 0 )
                {
                    RefGPID = GPlanList[0]["id"] + "";
                    RefGPName = GPlanList[0]["name"] + "";
                    RefGPContent = GPlanList[0]["content"] + "";
                    RefGPContentXml = XElement.Parse(RefGPContent);
                }               

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// 解析排序
        /// </summary>
        public void ParseOrderByInt()
        {
            if (Status == "更新")
                OrderByInt = 1;
            else if (Status == "新增")
            {
                OrderByInt = 2;
            }
            else if (Status == "無變動")
            {
                OrderByInt = 3;
            }
            else
            {
                OrderByInt = 4;
            }

        }

        public void CheckData()
        {
            // 檢查資料
            if (MOEXml != null && RefGPContentXml != null)
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

                foreach (XElement elm in RefGPContentXml.Elements("Subject"))
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
                        subj.GDCCode = mCo;
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
                        subj.GDCCode = mCo;
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
                        subj.GDCCode = mCo;
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
