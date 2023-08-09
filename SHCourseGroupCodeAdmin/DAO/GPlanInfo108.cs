using System;
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

        // 是否檢查科目級別
        bool _CheckSubjectLevel { get; set; }

        public bool needUpdateEntryYear = false;

        /// <summary>
        /// 課程代碼代表原始資料
        /// </summary>
        public List<MOECourseCodeInfo> MOECourseCodeInfoList = new List<MOECourseCodeInfo>();

        // 建立各別課程代碼索引
        public Dictionary<string, List<XElement>> MOEDict = new Dictionary<string, List<XElement>>();
        public Dictionary<string, List<XElement>> GPlanDict = new Dictionary<string, List<XElement>>();
        public List<chkSubjectInfo> chkSubjectInfoList = new List<chkSubjectInfo>();

        private string UserDefSubjectRoot = "使用者自訂科目";
        // 使用者自訂科目
        public Dictionary<string, List<XElement>> UserDefSubjectDict = new Dictionary<string, List<XElement>>();

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
                {
                    if (calSubjLevelCalCount() > 0)
                    {
                        Status = "級別更新";
                    }
                    else
                    {
                        Status = "更新";
                    }
                }
            }

        }

        public int calSubjLevelCalCount()
        {
            int value = 0;

            foreach (chkSubjectInfo subj in chkSubjectInfoList)
            {
                if (subj.ProcessStatus == "級別更新")
                    value++;
            }

            return value;
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
            {
                MOEXml.SetAttributeValue("EntryYear", GDCCode.Substring(0, 3));
                MOEXml.SetAttributeValue("SchoolYear", GDCCode.Substring(0, 3));
            }


            //// 支援就格式
            //if (MOEXml.Attribute("SchoolYear") != null)
            //{
            //    Console.WriteLine(MOEXml.Attribute("SchoolYear").Value);
            //}

            Dictionary<string, string> CreditMappingTableDict = Utility.GetCreditMappingTable();

            Dictionary<string, int> chkCourseCodeCount = new Dictionary<string, int>();

            // 取得領域對照
            Dictionary<string, string> domainMappingDict = Utility.GetDomainNameMapping();

            // 取得特殊課程類別對照
            Dictionary<string, string> courseTypeMappingDic = Utility.GetCourseTypeMapping();

            // 取得科目屬性對照
            Dictionary<string, string> subjectAttributeMappingDic = Utility.GetSubjectAttributeMapping();


            // 讀取群科班資料
            foreach (MOECourseCodeInfo data in MOECourseCodeInfoList)
            {

                // 判斷開始級別，解析 credit_period
                char[] cp = data.credit_period.ToArray();

                int idx = 1;
                string strGradeYear = "", strSemester = "";
                foreach (char c in cp)
                {
                    string credit = c + "";

                    if (idx == 1)
                    {
                        strGradeYear = "1";
                        strSemester = "1";
                    }
                    else if (idx == 2)
                    {
                        strGradeYear = "1";
                        strSemester = "2";
                    }
                    else if (idx == 3)
                    {
                        strGradeYear = "2";
                        strSemester = "1";
                    }
                    else if (idx == 4)
                    {
                        strGradeYear = "2";
                        strSemester = "2";
                    }
                    else if (idx == 5)
                    {
                        strGradeYear = "3";
                        strSemester = "1";
                    }
                    else if (idx == 6)
                    {
                        strGradeYear = "3";
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
                    subjElm.SetAttributeValue("GradeYear", strGradeYear);
                    subjElm.SetAttributeValue("Level", idx);  // 先塞學期別
                    subjElm.SetAttributeValue("FullName", SubjFullName(data.subject_name, idx));

                    subjElm.SetAttributeValue("NotIncludedInCalc", CheckNotIncludedInCredit(data.course_code));
                    subjElm.SetAttributeValue("NotIncludedInCredit", CheckNotIncludedInCredit(data.course_code));

                    subjElm.SetAttributeValue("Required", data.is_required);
                    if (data.require_by == "部定")
                        subjElm.SetAttributeValue("RequiredBy", "部訂");
                    else
                        subjElm.SetAttributeValue("RequiredBy", data.require_by);

                    subjElm.SetAttributeValue("Semester", strSemester);

                    subjElm.SetAttributeValue("OfficialSubjectName", data.subject_name);
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
                        if (strGradeYear == "1" && strSemester == "1")
                            opIdx = 0;

                        if (strGradeYear == "1" && strSemester == "2")
                            opIdx = 1;

                        if (strGradeYear == "2" && strSemester == "1")
                            opIdx = 2;
                        if (strGradeYear == "2" && strSemester == "2")
                            opIdx = 3;

                        if (strGradeYear == "3" && strSemester == "1")
                            opIdx = 4;
                        if (strGradeYear == "3" && strSemester == "2")
                            opIdx = 5;

                        // 調整只有代碼0，才是原班開課
                        if (ot[opIdx] == '0')
                        {
                            chkOt = true;
                        }

                    }

                    if (chkOt)
                        otStr = "原班";

                    subjElm.SetAttributeValue("開課方式", otStr);

                    subjElm.SetAttributeValue("領域名稱", "");
                    subjElm.SetAttributeValue("課程名稱", data.subject_name);
                    subjElm.SetAttributeValue("OpenType", data.open_type);
                    subjElm.SetAttributeValue("CourseAttr", data.course_attr);
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
                            subjElm.SetAttributeValue("特殊類別", "");
                        }
                        else if (code1 == "2")
                        {
                            subjElm.SetAttributeValue("RequiredBy", "校訂");
                            subjElm.SetAttributeValue("Required", "必修");
                            subjElm.SetAttributeValue("特殊類別", "");
                        }
                        else
                        {
                            subjElm.SetAttributeValue("RequiredBy", "校訂");
                            subjElm.SetAttributeValue("Required", "選修");
                            subjElm.SetAttributeValue("特殊類別", courseTypeMappingDic[code1]); // Mapping CourseType
                        }

                        // 分項
                        string code2 = data.course_attr.Substring(1, 1);
                        if (code2 == "2")
                        {
                            subjElm.SetAttributeValue("Entry", "專業科目");
                            subjElm.SetAttributeValue("科目屬性", subjectAttributeMappingDic[code2]);// Mapping SubjectAttribute
                        }
                        else if (code2 == "3")
                        {
                            subjElm.SetAttributeValue("Entry", "實習科目");
                            subjElm.SetAttributeValue("科目屬性", subjectAttributeMappingDic[code2]);// Mapping SubjectAttribute
                        }
                        else
                        {
                            subjElm.SetAttributeValue("Entry", "學業");
                            subjElm.SetAttributeValue("科目屬性", subjectAttributeMappingDic[code2]);// Mapping SubjectAttribute
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

            // 將級別中的學期別轉成冊別
            Dictionary<string, int> tmpSubjLevelDict = new Dictionary<string, int>();

            foreach (XElement elm in MOEXml.Elements("Subject"))
            {
                string subj = elm.Attribute("SubjectName").Value;

                if (!tmpSubjLevelDict.ContainsKey(subj))
                    tmpSubjLevelDict.Add(subj, 0);

                tmpSubjLevelDict[subj] += 1;

                elm.SetAttributeValue("FullName", SubjFullName(subj, tmpSubjLevelDict[subj]));
                elm.SetAttributeValue("Level", tmpSubjLevelDict[subj]);

            }

            // 依據冊別重新整理開始級別
            Dictionary<string, string> tmpStartLevel = new Dictionary<string, string>();
            string startLevelString = "";
            foreach (XElement elm in MOEXml.Elements("Subject"))
            {
                string subjName = elm.Attribute("SubjectName").Value;

                string rowIdx = elm.Element("Grouping").Attribute("RowIndex").Value;

                if (!tmpStartLevel.ContainsKey(subjName))
                {
                    tmpStartLevel.Add(subjName, rowIdx);
                    startLevelString = startLevel.ToString();
                }
                else
                {
                    if (tmpStartLevel[subjName] != rowIdx)
                    {
                        // 設定開始級別是目前級別
                        startLevelString = elm.Attribute("Level").Value.ToString();
                        tmpStartLevel[subjName] = rowIdx;
                    }

                    elm.Element("Grouping").SetAttributeValue("startLevel", startLevelString);
                }
            }

            // 處理多元選修級別計算
            foreach (XElement element in MOEXml.Elements("Subject"))
            {
                Utility.CalculateSubjectLevel(element);
            }
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

            if (level == 7)
                lev = " VII";

            if (level == 8)
                lev = " VIII";

            if (level == 9)
                lev = " IX";

            if (level == 10)
                lev = " X";

            if (level == 11)
                lev = " 11";

            if (level == 12)
                lev = " 12";


            string value = SubjectName + lev;

            return value;
        }


        public string CheckNotIncludedInCredit(string CourseCode)
        {
            string value = "False";
            // 判斷 8,9D 問題
            if (!string.IsNullOrWhiteSpace(CourseCode) && CourseCode.Length > 21)
            {
                string k1 = CourseCode.Substring(16, 1);
                string k2 = CourseCode.Substring(18, 1);

                // 8/團體活動時間、9/彈性學習時間 
                if (k1 == "8" || k1 == "9")
                {
                    value = "True";

                    // 需要評分
                    if (k2 == "D")
                    {
                        value = "False";
                    }
                }
            }


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
                if (GPlanList.Count > 0)
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
            // 用課程代碼比對，大表有課程規劃表沒有需要新增
            List<string> AddList = new List<string>();
            // 用課程規劃表比對大表沒有需要刪除
            List<string> DelList = new List<string>();

            Dictionary<string, List<XElement>> CourseCodeMDict = new Dictionary<string, List<XElement>>();

            // 綜合型高中學術學程1年級不分群不分班群，群科班代碼包含這串 M111960
            // 找出綜高 M            
            if (!GDCCode.Contains("M111960"))
            {
                if (GDCCode.Contains("M"))
                {
                    string key = GDCCode.Substring(0, 3);
                    // 檢查 是否有一年籍資料
                    bool hasGradeYear1 = false;

                    if (Global._GPlanInfo108MDict.ContainsKey(key))
                    {
                        if (RefGPContentXml != null)
                        {
                            foreach (XElement elm in RefGPContentXml.Elements("Subject"))
                            {
                                if (elm.Attribute("GradeYear") != null && elm.Attribute("GradeYear").Value == "1")
                                {
                                    // 表示有一年級的資料
                                    hasGradeYear1 = true;
                                }
                            }
                        }


                        if (hasGradeYear1)
                        {
                          //  Console.WriteLine("已有資料。");
                        }
                        else
                        {
                            GPlanInfo108 AddGPinfo = Global._GPlanInfo108MDict[key];
                            CourseCodeMDict.Clear();

                            // 整理新增資料
                            foreach (XElement elm in AddGPinfo.MOEXml.Elements("Subject"))
                            {
                                string coCode = GetAttribute(elm, "課程代碼");

                                if (!CourseCodeMDict.ContainsKey(coCode))
                                    CourseCodeMDict.Add(coCode, new List<XElement>());

                                CourseCodeMDict[coCode].Add(elm);
                            }

                            // 新增資料
                            foreach (string coCode in CourseCodeMDict.Keys)
                            {
                                if (CourseCodeMDict.ContainsKey(coCode))
                                {
                                    XElement elm = CourseCodeMDict[coCode][0];
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
                                    subj.open_type = GetAttribute(elm, "OpenType");
                                    subj.course_attr = GetAttribute(elm, "CourseAttr");
                                    subj.OfficialSubjectName = GetAttribute(elm, "OfficialSubjectName");
                                    subj.NotIncludedInCalc = CheckNotIncludedInCredit(subj.CourseCode);
                                    subj.NotIncludedInCredit = CheckNotIncludedInCredit(subj.CourseCode);
                                    elm.SetAttributeValue("NotIncludedInCalc", subj.NotIncludedInCalc);
                                    elm.SetAttributeValue("NotIncludedInCredit", subj.NotIncludedInCredit);

                                    subj.ProcessStatus = "新增";
                                    subj.DiffStatusList.Add("缺");
                                    subj.MOEXml = CourseCodeMDict[coCode];
                                    subj.GDCCode = Global._GPlanInfo108MDict[key].GDCCode;
                                    chkSubjectInfoList.Add(subj);
                                }
                            }
                        }

                    }
                }
            }



            // 新增資料
            if (MOEXml != null && RefGPContentXml == null)
            {
                this.RefGPID = "";

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

                foreach (string mCo in MOEDict.Keys)
                {
                    AddList.Add(mCo);
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
                        subj.open_type = GetAttribute(elm, "OpenType");
                        subj.course_attr = GetAttribute(elm, "CourseAttr");
                        subj.OfficialSubjectName = GetAttribute(elm, "OfficialSubjectName");
                        subj.NotIncludedInCalc = CheckNotIncludedInCredit(subj.CourseCode);
                        subj.NotIncludedInCredit = CheckNotIncludedInCredit(subj.CourseCode);
                        elm.SetAttributeValue("NotIncludedInCalc", subj.NotIncludedInCalc);
                        elm.SetAttributeValue("NotIncludedInCredit", subj.NotIncludedInCredit);

                        subj.ProcessStatus = "新增";
                        subj.DiffStatusList.Add("缺");
                        subj.MOEXml = MOEDict[mCo];
                        subj.GDCCode = mCo;
                        chkSubjectInfoList.Add(subj);
                    }
                }



            }
            else
            {
                // 檢查資料，當2邊都有資料
                if (MOEXml != null && RefGPContentXml != null)
                {

                    MOEDict.Clear();
                    GPlanDict.Clear();

                    if (RefGPContentXml.Attribute("SchoolYear") != null)
                    {
                        MOEXml.SetAttributeValue("SchoolYear", RefGPContentXml.Attribute("SchoolYear").Value);
                    }

                    if (RefGPContentXml.Attribute("EntryYear") == null)
                    {
                        needUpdateEntryYear = true;
                    }


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
                    AddList.Clear();
                    // 用課程規劃表比對大表沒有需要刪除
                    DelList.Clear();

                    foreach (string mCo in MOEDict.Keys)
                    {
                        if (!GPlanDict.ContainsKey(mCo))
                            AddList.Add(mCo);
                    }

                    foreach (string gCo in GPlanDict.Keys)
                    {
                        // 有綜高一年級部分班不刪除
                        if (gCo.Contains("M111960"))
                            continue;

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
                            subj.open_type = GetAttribute(elm, "OpenType");
                            subj.course_attr = GetAttribute(elm, "CourseAttr");
                            subj.OfficialSubjectName = GetAttribute(elm, "OfficialSubjectName");
                            subj.NotIncludedInCalc = CheckNotIncludedInCredit(subj.CourseCode);
                            subj.NotIncludedInCredit = CheckNotIncludedInCredit(subj.CourseCode);
                            elm.SetAttributeValue("NotIncludedInCalc", subj.NotIncludedInCalc);
                            elm.SetAttributeValue("NotIncludedInCredit", subj.NotIncludedInCredit);
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
                            subj.open_type = GetAttribute(elm, "OpenType");
                            subj.course_attr = GetAttribute(elm, "CourseAttr");
                            subj.OfficialSubjectName = GetAttribute(elm, "OfficialSubjectName");
                            subj.NotIncludedInCalc = CheckNotIncludedInCredit(subj.CourseCode);
                            subj.NotIncludedInCredit = CheckNotIncludedInCredit(subj.CourseCode);
                            subj.ProcessStatus = "刪除";
                            subj.DiffStatusList.Add("多");
                            subj.GPlanXml = GPlanDict[mCo];
                            subj.GDCCode = mCo;
                            chkSubjectInfoList.Add(subj);
                        }
                    }

                    Dictionary<string, string> MOELevelDict = new Dictionary<string, string>();
                    Dictionary<string, string> GPlanLevelDict = new Dictionary<string, string>();

                    Dictionary<string, string> MOEStartLevelDict = new Dictionary<string, string>();
                    Dictionary<string, string> GPlanStartLevelDict = new Dictionary<string, string>();

                    // 以大表為主，當課程代碼相同，比對裡面資料是否相同
                    foreach (string mCo in MOEDict.Keys)
                    {
                        if (GPlanDict.ContainsKey(mCo))
                        {
                            int index = 0;
                            foreach (XElement graduationPlanEle in GPlanDict[mCo])
                            {
                                if (index >= MOEDict[mCo].Count)
                                    continue;
                                
                                XElement moeEle = MOEDict[mCo][index];

                                // 複製原有課程規畫表群組設定
                                if (graduationPlanEle.Attribute("分組名稱") != null)
                                {
                                    string courseGroupName = graduationPlanEle.Attribute("分組名稱").Value.ToString();
                                    string courseGroupCredit = graduationPlanEle.Attribute("分組修課學分數").Value.ToString();

                                    if (!string.IsNullOrEmpty(courseGroupName))
                                    {
                                        moeEle.SetAttributeValue("分組名稱", courseGroupName);
                                        moeEle.SetAttributeValue("分組修課學分數", courseGroupCredit);
                                    }
                                }

                                // 複製對開設定，之後可以用課程群組取代對開設定時再移除
                                if (graduationPlanEle.Attribute("設定對開") != null)
                                {
                                    string openD = graduationPlanEle.Attribute("設定對開").Value.ToString();
                                    if (!string.IsNullOrEmpty(openD))
                                    {
                                        moeEle.SetAttributeValue("設定對開", openD);
                                    }
                                }

                                //// 複製級別設定
                                //if (graduationPlanEle.Attribute("Level") != null)
                                //{
                                //    string level = graduationPlanEle.Attribute("Level").Value.ToString();
                                //    if (!string.IsNullOrEmpty(level))
                                //    {
                                //        moeEle.SetAttributeValue("Level", level);
                                //    }
                                //}

                                index++;
                            }

                            XElement elm = GPlanDict[mCo][0];
                            XElement elmMoe = MOEDict[mCo][0];

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

                            subj.OpenStatus = GetAttribute(elm, "開課方式");

                            // 因為授課學期學分與OpenType，有可能總表調整需要更新，所以需要以總表為主
                            subj.credit_period = GetAttribute(elmMoe, "授課學期學分");
                            subj.open_type = GetAttribute(elmMoe, "OpenType");

                            subj.OfficialSubjectName = GetAttribute(elm, "OfficialSubjectName");
                            subj.course_attr = GetAttribute(elm, "CourseAttr");
                            subj.OfficialSubjectName = GetAttribute(elm, "OfficialSubjectName");
                            subj.NotIncludedInCalc = CheckNotIncludedInCredit(subj.CourseCode);
                            subj.NotIncludedInCredit = CheckNotIncludedInCredit(subj.CourseCode);
                            elm.SetAttributeValue("NotIncludedInCalc", subj.NotIncludedInCalc);
                            elm.SetAttributeValue("NotIncludedInCredit", subj.NotIncludedInCredit);


                            // 檢查開始級別與級別差異
                            if (_CheckSubjectLevel)
                            {
                                MOELevelDict.Clear();
                                GPlanLevelDict.Clear();
                                MOEStartLevelDict.Clear();
                                GPlanStartLevelDict.Clear();

                                foreach (XElement elmM in MOEDict[mCo])
                                {
                                    // 級別
                                    string key = GetAttribute(elmM, "GradeYear") + "_" + GetAttribute(elmM, "Semester");
                                    string level = GetAttribute(elmM, "Level");
                                    if (!MOELevelDict.ContainsKey(key))
                                        MOELevelDict.Add(key, level);

                                    // 開始級別
                                    string startLevel = GetAttribute(elmM.Element("Grouping"), "startLevel");

                                    if (!MOEStartLevelDict.ContainsKey(key))
                                        MOEStartLevelDict.Add(key, startLevel);

                                }


                                foreach (XElement elmM in GPlanDict[mCo])
                                {
                                    // 級別
                                    string key = GetAttribute(elmM, "GradeYear") + "_" + GetAttribute(elmM, "Semester");
                                    string level = GetAttribute(elmM, "Level");
                                    if (!GPlanLevelDict.ContainsKey(key))
                                        GPlanLevelDict.Add(key, level);

                                    // 開始級別
                                    string startLevel = GetAttribute(elmM.Element("Grouping"), "startLevel");
                                    if (!GPlanStartLevelDict.ContainsKey(key))
                                        GPlanStartLevelDict.Add(key, startLevel);
                                }

                                // 比對差異
                                foreach (string key in MOELevelDict.Keys)
                                {
                                    if (GPlanLevelDict.ContainsKey(key))
                                    {
                                        if (MOELevelDict[key] != GPlanLevelDict[key])
                                        {
                                            subj.DiffStatusList.Add("級別不同");
                                            subj.DiffMessageList.Add("級別：課程代碼表「" + MOELevelDict[key] + "」、課程規劃表「" + GPlanLevelDict[key] + "」");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        subj.DiffStatusList.Add("級別不同");
                                        subj.DiffMessageList.Add("級別：課程代碼表「" + MOELevelDict[key] + "」、課程規劃表「」");
                                        break;
                                    }
                                }

                                foreach (string key in MOEStartLevelDict.Keys)
                                {
                                    if (GPlanStartLevelDict.ContainsKey(key))
                                    {
                                        if (MOEStartLevelDict[key] != GPlanStartLevelDict[key])
                                        {
                                            subj.DiffStatusList.Add("開始級別不同");
                                            subj.DiffMessageList.Add("開始級別：課程代碼表「" + MOEStartLevelDict[key] + "」、課程規劃表「" + GPlanStartLevelDict[key] + "」");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        subj.DiffStatusList.Add("開始級別不同");
                                        subj.DiffMessageList.Add("開始級別：課程代碼表「" + MOEStartLevelDict[key] + "」、課程規劃表「」");
                                        break;
                                    }
                                }


                            }

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

                                    if (elmM.Attribute("OfficialSubjectName") != null && elmG.Attribute("OfficialSubjectName") != null)
                                    {
                                        if (elmM.Attribute("OfficialSubjectName").Value != elmG.Attribute("OfficialSubjectName").Value)
                                        {
                                            if (!subj.DiffStatusList.Contains("報部科目名稱不同"))
                                                subj.DiffStatusList.Add("報部科目名稱不同");

                                            string msg = "報部科目名稱：課程代碼表「" + elmM.Attribute("OfficialSubjectName").Value + "」、課程規劃表「" + elmG.Attribute("OfficialSubjectName").Value + "」";
                                            if (!subj.DiffMessageList.Contains(msg))
                                                subj.DiffMessageList.Add(msg);

                                        }
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

                                    if (elmM.Attribute("授課學期學分").Value != elmG.Attribute("授課學期學分").Value)
                                    {
                                        if (!subj.DiffStatusList.Contains("授課學期學分不同"))
                                            subj.DiffStatusList.Add("授課學期學分不同");

                                        string msg = "授課學期學分：課程代碼表「" + elmM.Attribute("授課學期學分").Value + "」、課程規劃表「" + elmG.Attribute("授課學期學分").Value + "」";
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

                                    // 檢查不需評分、不計學分不同
                                    if (elmM.Attribute("NotIncludedInCredit").Value != elmG.Attribute("NotIncludedInCredit").Value)
                                    {
                                        if (!subj.DiffStatusList.Contains("不需評分、不計學分不同"))
                                            subj.DiffStatusList.Add("不需評分、不計學分不同");

                                        string msg = "不需評分、不計學分：課程代碼表「" + elmM.Attribute("NotIncludedInCredit").Value + "」、課程規劃表「" + elmG.Attribute("NotIncludedInCredit").Value + "」";
                                        if (!subj.DiffMessageList.Contains(msg))
                                            subj.DiffMessageList.Add(msg);
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
        }
        private string GetAttribute(XElement elm, string attrName)
        {
            string value = "";

            if (elm.Attribute(attrName) != null)
                value = elm.Attribute(attrName).Value;

            return value;
        }

        public Dictionary<string, List<XElement>> GetUserDefSubjectDict()
        {
            UserDefSubjectDict.Clear();

            try
            {
                if (RefGPContentXml == null)
                {
                    ParseRefGPContentXml();
                }

                if (RefGPContentXml != null)
                {
                    XElement elmUDRoot = RefGPContentXml.Element(UserDefSubjectRoot);
                    if (elmUDRoot != null)
                    {
                        // 讀取科目名稱當 Key
                        foreach (XElement elm in elmUDRoot.Elements("Subject"))
                        {
                            string key = GetAttribute(elm, "Domain") + "_" + GetAttribute(elm, "Entry") + "_" + GetAttribute(elm, "Required") + "_" + GetAttribute(elm, "RequiredBy") + "_" + GetAttribute(elm, "SubjectName");

                            if (!UserDefSubjectDict.ContainsKey(key))
                                UserDefSubjectDict.Add(key, new List<XElement>());

                            UserDefSubjectDict[key].Add(elm);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return UserDefSubjectDict;
        }

        /// <summary>
        ///  設定使用者自訂
        /// </summary>
        /// <param name="value"></param>
        public void SetUserDefSubjectDict(Dictionary<string, List<XElement>> value)
        {
            try
            {
                if (RefGPContentXml != null)
                {
                    XElement elmUDRoot = new XElement(UserDefSubjectRoot);
                    foreach (string key in value.Keys)
                    {
                        foreach (XElement elm in value[key])
                            elmUDRoot.Add(elm);
                    }

                    if (RefGPContentXml.Element(UserDefSubjectRoot) != null)
                    {
                        // 先移除
                        RefGPContentXml.Element(UserDefSubjectRoot).Remove();
                    }

                    // 再新增
                    RefGPContentXml.Add(elmUDRoot);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void SetCheckSubjectLevel(bool value)
        {
            _CheckSubjectLevel = value;
        }
    }
}
