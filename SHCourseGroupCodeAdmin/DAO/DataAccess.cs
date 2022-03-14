using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using FISCA.Data;
using System.Xml.Linq;
using System.IO;
using System.Windows.Forms;

namespace SHCourseGroupCodeAdmin.DAO
{

    public class DataAccess
    {
        Dictionary<string, string> MOEGroupCodeDict = new Dictionary<string, string>();
        Dictionary<string, string> MOEGroupNameDict = new Dictionary<string, string>();
        Dictionary<string, string> MOEGPlanDict = new Dictionary<string, string>();
        Dictionary<string, string> MOEGPlanGroupNameDict = new Dictionary<string, string>();

        /// <summary>
        /// 取得課程代碼大表資料
        /// </summary>
        /// <returns></returns>
        public List<MOECourseCodeInfo> GetCourseGroupCodeList()
        {
            // 判斷日進校
            string courseType = Utility.GetCourseCodeWhereCond();

            List<MOECourseCodeInfo> value = new List<MOECourseCodeInfo>();
            try
            {
                string query = "" +
                   " SELECT  " +
" last_update AS 最後更新日期 " +
" , group_code AS 群組代碼 " +
" , course_code AS 課程代碼 " +
" , subject_name AS 科目名稱 " +
" , entry_year AS 入學年 " +
" , (CASE require_by WHEN '部訂' THEN '部定' ELSE require_by END) AS 部定校訂 " +
" , is_required AS 必修選修 " +
" , course_type AS 課程類型 " +
" , group_type AS 群別 " +
" , subject_type AS 科別 " +
" , class_type AS 班群 " +
" , credit_period AS 授課學期學分_節數 " +
" , open_type AS 授課開課方式 " +
" , course_attr AS 課程屬性 " +
"  FROM " +
"   $moe.subjectcode " + courseType + " ORDER BY group_code DESC,subject_name; ";
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    MOECourseCodeInfo data = new MOECourseCodeInfo();
                    data.group_code = dr["群組代碼"] + "";
                    data.course_code = dr["課程代碼"] + "";
                    data.subject_name = dr["科目名稱"] + "";
                    data.entry_year = dr["入學年"] + "";
                    data.require_by = dr["部定校訂"] + "";
                    data.is_required = dr["必修選修"] + "";
                    data.course_type = dr["課程類型"] + "";
                    data.group_type = dr["群別"] + "";
                    data.subject_type = dr["科別"] + "";
                    data.class_type = dr["班群"] + "";
                    data.credit_period = dr["授課學期學分_節數"] + "";
                    data.open_type = dr["授課開課方式"] + "";
                    data.course_attr = dr["課程屬性"] + "";

                    // 透過課程屬性更新校部定
                    if (data.course_attr.Length > 0)
                    {
                        string code = data.course_attr.Substring(0, 1);
                        if (code == "1")
                        {
                            data.is_required = "必修";
                            data.require_by = "部定";
                        }
                        else if (code == "2")
                        {
                            data.is_required = "必修";
                            data.require_by = "校訂";
                        }
                        else
                        {
                            data.is_required = "選修";
                            data.require_by = "校訂";
                        }
                    }

                    value.Add(data);
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }


        /// <summary>
        /// 檢查班級群科班設定
        /// </summary>
        /// <returns></returns>
        public List<ClassInfo> GetClassCourseGroup()
        {
            List<ClassInfo> value = new List<ClassInfo>();
            try
            {
                LoadMOEGroupCodeDict();


                string query = "" +
                    "SELECT " +
                    "id" +
                    ",COALESCE(grade_year,0) AS grade_year" +
                    ",class_name" +
                    ",COALESCE(gdc_code,'') AS gdc_code " +
                    " FROM class " +
                    " ORDER BY class.grade_year,display_order,class_name";
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    ClassInfo ci = new ClassInfo();
                    ci.ClassID = dr["id"] + "";
                    ci.ClassName = dr["class_name"] + "";
                    ci.GradeYear = dr["grade_year"] + "";
                    ci.ClassGroupCode = dr["gdc_code"] + "";
                    ci.ClassGroupName = GetGroupNameByCode(ci.ClassGroupCode);
                    value.Add(ci);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            return value;
        }

        /// <summary>
        /// 取得大表群組代碼與說明
        /// </summary>
        /// <returns></returns>
        public void LoadMOEGroupCodeDict()
        {
            MOEGroupCodeDict.Clear();
            MOEGroupNameDict.Clear();
            MOEGPlanDict.Clear();
            MOEGPlanGroupNameDict.Clear();

            string courseType = Utility.GetCourseCodeWhereCond();

            QueryHelper qh = new QueryHelper();
            string query = "" +
" SELECT DISTINCT group_code,(entry_year|| " +
" (CASE WHEN length(course_type)>0 THEN '_'||course_type END) || " +
" (CASE WHEN length(group_type)>0 THEN '_'||group_type END) || " +
" (CASE WHEN length(subject_type)>0 THEN '_'||subject_type END) || " +
" (CASE WHEN length(class_type)>0 THEN '_'||class_type END) " +
" ) AS group_name" +
" ,(entry_year||(CASE WHEN length(group_type)>0 THEN ''||group_type END)||(CASE WHEN length(subject_type)>0 THEN ''||subject_type END) || " +
" (CASE WHEN length(class_type) > 0 THEN '' || class_type END)) AS gplan_name " +
" FROM $moe.subjectcode " + courseType + " ORDER BY group_name DESC; ";

            DataTable dt = qh.Select(query);

            foreach (DataRow dr in dt.Rows)
            {
                string code = dr["group_code"] + "";
                string name = dr["group_name"] + "";
                string gpName = dr["gplan_name"] + "";

                if (!MOEGroupCodeDict.ContainsKey(code))
                    MOEGroupCodeDict.Add(code, name);

                if (!MOEGroupNameDict.ContainsKey(name))
                    MOEGroupNameDict.Add(name, code);

                if (!MOEGPlanDict.ContainsKey(code))
                    MOEGPlanDict.Add(code, gpName);

                if (!MOEGPlanGroupNameDict.ContainsKey(name))
                    MOEGPlanGroupNameDict.Add(name, gpName);
            }
        }

        public List<string> GetGroupNameList()
        {
            return MOEGroupNameDict.Keys.ToList();
        }

        public string GetGroupNameByCode(string code)
        {
            string name = "";
            if (MOEGroupCodeDict.ContainsKey(code))
                name = MOEGroupCodeDict[code];
            return name;
        }

        public string GetGroupCodeByName(string name)
        {
            string code = "";

            if (MOEGroupNameDict.ContainsKey(name))
                code = MOEGroupNameDict[name];

            return code;
        }

        /// <summary>
        ///  透過群組代碼取得課程規劃表預設名稱
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        public string GetGPlanNameByCode(string code)
        {
            string name = "";
            if (MOEGPlanDict.ContainsKey(code))
                name = MOEGPlanDict[code];
            return name;
        }

        /// <summary>
        /// 透過群組名稱取得課程規劃表預設名稱
        /// </summary>
        /// <param name="Gname"></param>
        /// <returns></returns>
        public string GetGPlanNameByGroupName(string Gname)
        {
            string name = "";
            if (MOEGPlanGroupNameDict.ContainsKey(Gname))
                name = MOEGPlanGroupNameDict[Gname];
            return name;
        }

        /// <summary>
        /// 取得已開課程，作資料檢查使用
        /// </summary>
        /// <param name="GradeYear"></param>
        /// <returns></returns>
        public List<CourseInfoChk> GetCourseCheckInfoListByGradeYear(int GradeYear)
        {
            List<CourseInfoChk> value = new List<CourseInfoChk>();

            try
            {
                QueryHelper qh = new QueryHelper();
                string query = "" +
                   " SELECT  " +
" DISTINCT course.id AS course_id " +
" , course_name " +
" , subject " +
" , subj_level " +
" , COALESCE(course.credit, null) AS c_credit " +
" , COALESCE(course.period, null) AS c_period " +
" , class.grade_year " +
" , COALESCE(course.credit,course.period) AS credit " +
" , course.school_year " +
" , course.semester " +
" , (CASE COALESCE(sc_attend.required_by,c_required_by) WHEN '1' THEN '部定' WHEN '2' THEN '校訂' ELSE '' END) AS required_by " +
" , (CASE COALESCE(sc_attend.is_required,c_is_required) WHEN '1' THEN '必修' WHEN '0' THEN '選修' ELSE '' END) AS required " +
" , COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code " +
" FROM course " +
" 	INNER JOIN sc_attend " +
"  ON course.id = sc_attend.ref_course_id  " +
" 	INNER JOIN student  " +
"  ON sc_attend.ref_student_id = student.id " +
" 	INNER JOIN class " +
" 	ON student.ref_class_id = class.id " +
" WHERE  " +
"  student.status = 1 AND class.grade_year in(" + GradeYear + ") " +
" ORDER BY school_year DESC,semester DESC ";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    CourseInfoChk data = new CourseInfoChk();
                    data.CourseID = dr["course_id"] + "";
                    data.CourseName = dr["course_name"] + "";
                    data.SchoolYear = dr["school_year"] + "";
                    data.Semester = dr["semester"] + "";
                    data.SubjectName = dr["subject"] + "";
                    data.IsRequired = dr["required"] + "";
                    data.RequireBy = dr["required_by"] + "";
                    data.SubjectLevel = dr["subj_level"] + "";
                    data.course_code = "";
                    data.Credit = "";
                    if (dr["c_credit"] != null)
                    {
                        data.Credit = dr["c_credit"] + "";
                    }

                    data.Period = "";
                    if (dr["c_period"] != null)
                    {
                        data.Period = dr["c_period"] + "";
                    }

                    data.credit_period = "";
                    data.GroupCode = dr["gdc_code"] + "";
                    value.Add(data);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }


        /// <summary>
        /// 取得群科班代碼科目內容
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, List<MOECourseCodeInfo>> GetCourseGroupCodeDict()
        {
            Dictionary<string, List<MOECourseCodeInfo>> value = new Dictionary<string, List<MOECourseCodeInfo>>();

            List<MOECourseCodeInfo> coCodeList = GetCourseGroupCodeList();
            foreach (MOECourseCodeInfo co in coCodeList)
            {
                if (!value.ContainsKey(co.group_code))
                    value.Add(co.group_code, new List<MOECourseCodeInfo>());

                value[co.group_code].Add(co);
            }
            return value;
        }

        public Dictionary<string, List<string>> GetCPIdGdcCodeDict(int GradeYear)
        {
            // 有可能一個課程規劃表對應到2個群組代碼，合理應該只有一個
            Dictionary<string, List<string>> value = new Dictionary<string, List<string>>();
            try
            {
                string query = "" +
                " SELECT DISTINCT  " +
                " 	COALESCE(student.gdc_code,class.gdc_code) AS gdc_code " +
                " 	,COALESCE(student.ref_graduation_plan_id,class.ref_graduation_plan_id) AS gp_id " +
                "  FROM student LEFT JOIN class " +
                " 	ON student.ref_class_id = class.id  " +
                "  WHERE class.grade_year IN (" + GradeYear + "); ";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string gp_id = "";
                    string gdc_code = "";
                    if (dr["gdc_code"] != null)
                        gdc_code = dr["gdc_code"] + "";

                    if (dr["gp_id"] != null)
                        gp_id = dr["gp_id"] + "";

                    if (gp_id != "")
                    {
                        if (!value.ContainsKey(gp_id))
                            value.Add(gp_id, new List<string>());

                        value[gp_id].Add(gdc_code);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        public Dictionary<string, GPlanInfo> GetGPlanInfoDictByGPID(List<string> GPIDList)
        {
            Dictionary<string, GPlanInfo> value = new Dictionary<string, GPlanInfo>();
            try
            {
                if (GPIDList.Count > 0)
                {
                    string query = "SELECT DISTINCT id,name,content FROM graduation_plan WHERE id IN (" + string.Join(",", GPIDList.ToArray()) + ") ORDER BY name;";

                    QueryHelper qh = new QueryHelper();
                    DataTable dt = qh.Select(query);
                    foreach (DataRow dr in dt.Rows)
                    {
                        GPlanInfo data = new GPlanInfo();
                        data.ID = dr["id"] + "";
                        data.Name = dr["name"] + "";
                        data.Content = dr["content"] + "";
                        data.ParseContentToCourseInfoList();

                        if (!value.ContainsKey(data.ID))
                            value.Add(data.ID, data);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        /// <summary>
        /// 透過群組代碼取得課程代碼大表資料
        /// </summary>
        /// <param name="GroupCode"></param>
        /// <returns></returns>
        public XElement CourseCodeConvertToGPlanByGroupCode(string GroupCode)
        {

            // 透過群組代碼取得課程代碼資料，轉換 XML 表內容。

            List<MOECourseCodeInfo> MOECourseCodeInfoList = GetCourseGroupCodeListByGroupCode(GroupCode);

            XElement GPlanXml = new XElement("GraduationPlan");

            /*
             	<Subject Category="" Credit="4" Domain="語文" Entry="學業" GradeYear="1" Level="1" FullName="國語文 I" NotIncludedInCalc="False" NotIncludedInCredit="False" Required="必修" RequiredBy="部訂" Semester="1" SubjectName="國語文" 課程代碼="109041305H11101B1010101" 課程類別="部定必修" 開課方式="原班級" 科目屬性="一般科目" 領域名稱="語文" 課程名稱="國語文" 學分="4" 授課學期學分="444440">
		<Grouping RowIndex="1" startLevel="1"/>
	</Subject>
             
             */

            int RowIndex = 1, startLevel = 1;

            // 課程規劃表預設取資料裡的入學年，如果沒有在透過課程群組代碼取前3
            if (MOECourseCodeInfoList.Count > 0)
            {
                GPlanXml.SetAttributeValue("SchoolYear", MOECourseCodeInfoList[0].entry_year);
            }
            else
            {
                if (GroupCode.Length > 3)
                    GPlanXml.SetAttributeValue("SchoolYear", GroupCode.Substring(0, 3));
            }

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

                    GPlanXml.Add(subjElm);
                    idx++;
                }
                RowIndex++;
            }

            // 重新排列科目級別
            Dictionary<string, int> tmpSubjLevelDict = new Dictionary<string, int>();

            foreach (XElement elm in GPlanXml.Elements("Subject"))
            {
                string subj = elm.Attribute("SubjectName").Value;

                if (!tmpSubjLevelDict.ContainsKey(subj))
                    tmpSubjLevelDict.Add(subj, 0);

                tmpSubjLevelDict[subj] += 1;

                elm.SetAttributeValue("FullName", SubjFullName(subj, tmpSubjLevelDict[subj]));
                elm.SetAttributeValue("Level", tmpSubjLevelDict[subj]);

            }

            // 重新整理開始級別
            Dictionary<string, string> tmpStartLevel = new Dictionary<string, string>();
            foreach (XElement elm in GPlanXml.Elements("Subject"))
            {
                string subjName = elm.Attribute("SubjectName").Value;

                string rowIdx = elm.Element("Grouping").Attribute("RowIndex").Value;

                if (!tmpStartLevel.ContainsKey(subjName))
                    tmpStartLevel.Add(subjName, rowIdx);
                else
                {
                    if (tmpStartLevel[subjName] != rowIdx)
                    {
                        // 設定開始級別是目前級別
                        elm.Element("Grouping").SetAttributeValue("startLevel", elm.Attribute("Level").Value);
                        tmpStartLevel[subjName] = rowIdx;
                    }
                }
            }

            return GPlanXml;
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


        public bool UpdateGPlanXML(string GPID, string XmlStr)
        {
            bool value = false;
            try
            {
                if (!string.IsNullOrWhiteSpace(GPID))
                {
                    string query = "UPDATE graduation_plan SET content ='" + XmlStr + "' WHERE id = " + GPID;
                    K12.Data.UpdateHelper uh = new K12.Data.UpdateHelper();
                    uh.Execute(query);
                    value = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }


        /// <summary>
        /// 透過群組代碼產生課程規劃表XML 並寫入資料庫
        /// </summary>
        /// <param name="GroupCode"></param>
        /// <returns></returns>
        public string WriteToGPlanByGroupCode(string GroupCode)
        {
            string value = "";
            // 課程規劃表XML
            try
            {
                if (MOEGPlanDict.Count == 0)
                {
                    LoadMOEGroupCodeDict();
                }

                string gpContent = CourseCodeConvertToGPlanByGroupCode(GroupCode).ToString(SaveOptions.DisableFormatting);

                //  Console.WriteLine(gpContent);

                // 取得課程規劃表名稱
                string GPlanName = GetGPlanNameByCode(GroupCode);

                string insertQry = "INSERT INTO graduation_plan(name,content,moe_group_code) VALUES('" + GPlanName + "','" + gpContent + "','" + GroupCode + "') RETURNING id;";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(insertQry);

            }
            catch (Exception ex)
            {
                value = ex.Message;
                return value;
            }
            return value;
        }

        /// <summary>
        /// 檢查課程規劃表是否已存在
        /// 
        /// 
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public bool CheckHasGPlan(string name)
        {
            bool value = false;
            try
            {
                string query = "SELECT name FROM graduation_plan WHERE name = '" + name + "';";
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                if (dt != null && dt.Rows.Count > 0)
                    value = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }



        public List<MOECourseCodeInfo> GetCourseGroupCodeListByGroupCode(string GroupCode)
        {
            List<MOECourseCodeInfo> value = new List<MOECourseCodeInfo>();
            try
            {
                string query = "" +
                   " SELECT  " +
" last_update AS 最後更新日期 " +
" , group_code AS 群組代碼 " +
" , course_code AS 課程代碼 " +
" , subject_name AS 科目名稱 " +
" , entry_year AS 入學年 " +
" , (CASE require_by WHEN '部訂' THEN '部定' ELSE require_by END) AS 部定校訂 " +
" , is_required AS 必修選修 " +
" , course_type AS 課程類型 " +
" , group_type AS 群別 " +
" , subject_type AS 科別 " +
" , class_type AS 班群 " +
" , credit_period AS 授課學期學分_節數 " +
" , open_type AS 授課開課方式 " +
" , course_attr AS 課程屬性 " +
"  FROM " +
"   $moe.subjectcode " +
"   WHERE group_code = '" + GroupCode + "'  " +
"   ORDER BY course_code; ";
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    MOECourseCodeInfo data = new MOECourseCodeInfo();
                    data.group_code = dr["群組代碼"] + "";
                    data.course_code = dr["課程代碼"] + "";
                    data.subject_name = dr["科目名稱"] + "";
                    data.entry_year = dr["入學年"] + "";
                    data.require_by = dr["部定校訂"] + "";
                    data.is_required = dr["必修選修"] + "";
                    data.course_type = dr["課程類型"] + "";
                    data.group_type = dr["群別"] + "";
                    data.subject_type = dr["科別"] + "";
                    data.class_type = dr["班群"] + "";
                    data.credit_period = dr["授課學期學分_節數"] + "";
                    data.open_type = dr["授課開課方式"] + "";
                    data.course_attr = dr["課程屬性"] + "";

                    // 透過課程屬性更新校部定
                    if (data.course_attr.Length > 0)
                    {
                        string code = data.course_attr.Substring(0, 1);
                        if (code == "1")
                        {
                            data.is_required = "必修";
                            data.require_by = "部定";
                        }
                        else if (code == "2")
                        {
                            data.is_required = "必修";
                            data.require_by = "校訂";
                        }
                        else
                        {
                            data.is_required = "選修";
                            data.require_by = "校訂";
                        }
                    }

                    value.Add(data);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        /// <summary>
        /// 透過年級取得學生學期科目成績比對資訊
        /// </summary>
        /// <param name="GradeYear"></param>
        /// <returns></returns>
        public List<StudentSubjectInfoChk> GetStudentSubjectInfoListByGradeYear(int GradeYear)
        {
            List<StudentSubjectInfoChk> value = new List<StudentSubjectInfoChk>();
            Dictionary<string, StudentSubjectInfoChk> studDict = new Dictionary<string, StudentSubjectInfoChk>();
            try
            {
                string query = "" +
               " WITH student_data AS ( " +
" 	SELECT " +
" 	student.id AS student_id " +
" 	,student_number " +
" 	,class_name " +
" 	,student.seat_no " +
" 	,student.name AS student_name " +
" 	, COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code " +
" FROM student  " +
" INNER JOIN class ON student.ref_class_id = class.id  " +
" 	WHERE student.status IN(1,2) AND class.grade_year  IN(" + GradeYear + ")" +
" ),sems_score_data AS( " +
" SELECT " +
" 	sems_subj_score_ext.ref_student_id " +
" 	, sems_subj_score_ext.grade_year " +
" 	, sems_subj_score_ext.semester " +
" 	, sems_subj_score_ext.school_year	 " +
" 	, array_to_string(xpath('//Subject/@科目', subj_score_ele), '')::text AS 科目 " +
" 	, array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '')::text AS 科目級別 " +
" 	, array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '')::text AS 學分數	 " +
" 	, array_to_string(xpath('//Subject/@修課必選修', subj_score_ele), '')::text AS 必選修 " +
" 	, array_to_string(xpath('//Subject/@修課校部訂', subj_score_ele), '')::text AS 校部訂	 " +
" FROM ( " +
" 		SELECT  " +
" 			sems_subj_score.* " +
" 			, 	unnest(xpath('//SemesterSubjectScoreInfo/Subject', xmlparse(content score_info))) as subj_score_ele " +
" 		FROM  " +
" 			sems_subj_score  " +
" 			INNER JOIN student_data ON sems_subj_score.ref_student_id = student_data.student_id " +
" 	) as sems_subj_score_ext  " +
" ) " +
" SELECT  " +
" student_id " +
" ,student_number " +
" ,class_name " +
" ,seat_no " +
" ,student_name " +
" ,gdc_code " +
" ,school_year " +
" ,semester " +
" ,grade_year " +
" ,科目 " +
" ,科目級別 " +
" ,學分數 " +
" ,必選修 " +
" ,(CASE 校部訂 WHEN '部訂' THEN '部定' ELSE 校部訂 END) AS 部定校訂 " +
"  FROM student_data INNER JOIN sems_score_data ON student_data.student_id = sems_score_data.ref_student_id; ";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    // 學生系統編號
                    string sid = dr["student_id"] + "";
                    if (!studDict.ContainsKey(sid))
                    {
                        StudentSubjectInfoChk ssi = new StudentSubjectInfoChk();
                        ssi.StudentID = sid;
                        // 學號
                        ssi.StudentNumber = dr["student_number"] + "";

                        // 班級
                        ssi.ClassName = dr["class_name"] + "";

                        // 座號
                        ssi.SeatNo = dr["seat_no"] + "";

                        // 姓名
                        ssi.StudentName = dr["student_name"] + "";

                        // gdc_code
                        ssi.gdc_code = dr["gdc_code"] + "";

                        studDict.Add(sid, ssi);
                    }

                    SubjectInfoChk SubjInfo = new SubjectInfoChk();
                    // 學年度
                    SubjInfo.SchoolYear = dr["school_year"] + "";
                    // 學期
                    SubjInfo.Semester = dr["semester"] + "";
                    // 科目名稱
                    SubjInfo.SubjectName = dr["科目"] + "";
                    // 科目級別
                    SubjInfo.SubjectLevel = dr["科目級別"] + "";

                    // 部定校訂
                    SubjInfo.RequireBy = dr["部定校訂"] + "";

                    // 必修選修
                    SubjInfo.IsRequired = dr["必選修"] + "";

                    // 學分數
                    SubjInfo.Credit = dr["學分數"] + "";

                    SubjInfo.credit_period = "";

                    studDict[sid].SubjectInfoChkList.Add(SubjInfo);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            value = studDict.Values.ToList();

            return value;
        }

        /// <summary>
        /// 取得所有班級課程規劃表
        /// </summary>
        /// <returns></returns>
        public DataTable GetAllGPlanData()
        {
            DataTable dt = null;
            try
            {
                string query = "" +
                    " SELECT " +
"     id " +
"     , name " +
"     , array_to_string(xpath('//Subject/@GradeYear', subject_ele), '')::text AS 年級 " +
"     , array_to_string(xpath('//Subject/@Semester', subject_ele), '')::text AS 學期 " +
"     , array_to_string(xpath('//Subject/@Entry', subject_ele), '')::text AS 分項類別 " +
"     , array_to_string(xpath('//Subject/@Domain', subject_ele), '')::text AS 領域 " +
"     , array_to_string(xpath('//Subject/@SubjectName', subject_ele), '')::text AS 科目名稱 " +
"     , array_to_string(xpath('//Subject/@Level', subject_ele), '')::text AS 科目級別 " +
"     , array_to_string(xpath('//Subject/@Credit', subject_ele), '')::text AS 學分數 " +
"     , array_to_string(xpath('//Subject/@Required', subject_ele), '')::text AS 必修選修 " +
"     , array_to_string(xpath('//Subject/@RequiredBy', subject_ele), '')::text AS 校訂部定 " +
" 	, array_to_string(xpath('//Subject/@課程代碼', subject_ele), '')::text AS 課程代碼 " +
" FROM " +
"     ( " +
"         SELECT  " +
"             id " +
"             , name " +
"             , unnest(xpath('//GraduationPlan/Subject', xmlparse(content content))) as subject_ele " +
"         FROM  " +
"             graduation_plan " +
"     ) AS graduation_plan " +
" ORDER BY  " +
"     id " +
"     , array_to_string(xpath('//Subject/@GradeYear', subject_ele), '')::text " +
"     , array_to_string(xpath('//Subject/@Semester', subject_ele), '')::text " +
"     , array_to_string(xpath('//Subject/@Credit', subject_ele), '')::text ";

                QueryHelper qh = new QueryHelper();
                dt = qh.Select(query);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return dt;
        }

        public bool UpdateGPlanMOECode()
        {
            bool value = false;
            // 建立群科班連結
            try
            {
                string linkQuery = "" +
                    " WITH moe_group_code_data AS(" +
" SELECT " +
"     DISTINCT id " +
"     , name  " +
" 	, substring(array_to_string(xpath('//Subject/@課程代碼', subject_ele), '')::text,0,17) AS 群組代碼 " +
" 	,moe_group_code " +
" 	,moe_group_code_1 " +
" FROM " +
"     ( " +
"         SELECT  " +
"             id " +
"             , name " +
" 			,moe_group_code  " +
" 			,moe_group_code_1 " +
"             , unnest(xpath('//GraduationPlan/Subject', xmlparse(content content))) as subject_ele " +
"         FROM  " +
"             graduation_plan " +
"     ) AS graduation_plan WHERE  substring(array_to_string(xpath('//Subject/@課程代碼', subject_ele), '')::text,0,17) <>'' " +
" ), " +
" update_data AS( " +
" UPDATE graduation_plan  " +
" SET moe_group_code = moe_group_code_data.群組代碼  " +
" FROM moe_group_code_data  " +
" WHERE graduation_plan.id = moe_group_code_data.id  " +
" RETURNING  moe_group_code_data.*  " +
" ) " +
" SELECT * FROM update_data ";

                QueryHelper qh = new QueryHelper();
                DataTable lkDt = qh.Select(linkQuery);
                value = true;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }


        /// <summary>
        /// 取得所有有群科班代碼班級課程規劃表
        /// </summary>
        /// <returns></returns>
        public DataTable GetAllHasMOECodeGPlanData()
        {
            DataTable dt = null;

            // 建立群科班連結
            try
            {
                string linkQuery = "" +
                    " WITH moe_group_code_data AS(" +
" SELECT " +
"     DISTINCT id " +
"     , name  " +
" 	, substring(array_to_string(xpath('//Subject/@課程代碼', subject_ele), '')::text,0,17) AS 群組代碼 " +
" 	,moe_group_code " +
" 	,moe_group_code_1 " +
" FROM " +
"     ( " +
"         SELECT  " +
"             id " +
"             , name " +
" 			,moe_group_code  " +
" 			,moe_group_code_1 " +
"             , unnest(xpath('//GraduationPlan/Subject', xmlparse(content content))) as subject_ele " +
"         FROM  " +
"             graduation_plan " +
"     ) AS graduation_plan WHERE  substring(array_to_string(xpath('//Subject/@課程代碼', subject_ele), '')::text,0,17) <>'' " +
" ), " +
" update_data AS( " +
" UPDATE graduation_plan  " +
" SET moe_group_code = moe_group_code_data.群組代碼  " +
" FROM moe_group_code_data  " +
" WHERE graduation_plan.id = moe_group_code_data.id  " +
" RETURNING  moe_group_code_data.*  " +
" ) " +
" SELECT * FROM update_data ";

                QueryHelper qh = new QueryHelper();
                DataTable lkDt = qh.Select(linkQuery);


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            try
            {
                string query = "" +
                    " SELECT " +
"     id " +
"     , name " +
"	  , moe_group_code " +
"     , array_to_string(xpath('//Subject/@GradeYear', subject_ele), '')::text AS 年級 " +
"     , array_to_string(xpath('//Subject/@Semester', subject_ele), '')::text AS 學期 " +
"     , array_to_string(xpath('//Subject/@Entry', subject_ele), '')::text AS 分項類別 " +
"     , array_to_string(xpath('//Subject/@Domain', subject_ele), '')::text AS 領域 " +
"     , array_to_string(xpath('//Subject/@SubjectName', subject_ele), '')::text AS 科目名稱 " +
"     , array_to_string(xpath('//Subject/@Level', subject_ele), '')::text AS 科目級別 " +
"     , array_to_string(xpath('//Subject/@Credit', subject_ele), '')::text AS 學分數 " +
"     , array_to_string(xpath('//Subject/@Required', subject_ele), '')::text AS 必修選修 " +
"     , array_to_string(xpath('//Subject/@RequiredBy', subject_ele), '')::text AS 校訂部定 " +
" 	, array_to_string(xpath('//Subject/@課程代碼', subject_ele), '')::text AS 課程代碼 " +
" FROM " +
"     ( " +
"         SELECT  " +
"             id " +
"             , name " +
"             , moe_group_code " +
"             , unnest(xpath('//GraduationPlan/Subject', xmlparse(content content))) as subject_ele " +
"         FROM  " +
"             graduation_plan WHERE moe_group_code <> '' " +
"     ) AS graduation_plan " +
" ORDER BY  " +
"     id " +
"     , array_to_string(xpath('//Subject/@GradeYear', subject_ele), '')::text " +
"     , array_to_string(xpath('//Subject/@Semester', subject_ele), '')::text " +
"     , array_to_string(xpath('//Subject/@Credit', subject_ele), '')::text ";

                QueryHelper qh = new QueryHelper();
                dt = qh.Select(query);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return dt;
        }

        /// <summary>
        /// 透過群科班代碼取得課程規劃表
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        public List<GPlanData> GetGPlanDataListByMOECode(string code)
        {
            List<GPlanData> value = new List<GPlanData>();

            try
            {
                string query = "SELECT id,name,moe_group_code,content FROM graduation_plan WHERE moe_group_code ='" + code + "' ORDER BY name";
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        GPlanData gd = new GPlanData();
                        gd.ID = dr["id"] + "";
                        gd.Name = dr["name"] + "";
                        gd.MOEGroupCode = code;

                        string content = dr["content"] + "";

                        try
                        {
                            gd.ContentXML = XElement.Parse(content);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        value.Add(gd);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        /// <summary>
        /// 設定課程規劃所採用班級
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public List<GPlanData> ParseGPlanDataRefClass(List<GPlanData> data)
        {
            try
            {
                if (data.Count > 0)
                {
                    // 取得課程規劃表 ID
                    List<string> GPIDList = (from data1 in data select data1.ID).ToList();

                    // 取得使用班級ID,Name
                    string query = "SELECT " +
                        "ref_graduation_plan_id AS gp_id" +
                        ",class.id AS class_id" +
                        ",class_name " +
                        "FROM " +
                        "class " +
                        "WHERE ref_graduation_plan_id IN(" + string.Join(",", GPIDList.ToArray()) + ") " +
                        "ORDER BY display_order,class_name";

                    QueryHelper qh = new QueryHelper();
                    DataTable dt = qh.Select(query);
                    Dictionary<string, List<DataRow>> gpcDict = new Dictionary<string, List<DataRow>>();
                    foreach (DataRow dr in dt.Rows)
                    {
                        string gpid = dr["gp_id"] + "";
                        if (!gpcDict.ContainsKey(gpid))
                            gpcDict.Add(gpid, new List<DataRow>());

                        gpcDict[gpid].Add(dr);
                    }

                    // 填入課程規劃物件
                    foreach (GPlanData gpd in data)
                    {
                        if (gpcDict.ContainsKey(gpd.ID))
                        {
                            foreach (DataRow dr in gpcDict[gpd.ID])
                            {
                                string class_id = dr["class_id"] + "";
                                if (!gpd.UsedClassIDNameDict.ContainsKey(class_id))
                                    gpd.UsedClassIDNameDict.Add(class_id, dr["class_name"] + "");
                            }
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return data;
        }

        /// <summary>
        /// 取得課程規劃表與採用班級群科班比對
        /// </summary>
        /// <returns></returns>
        public List<rptGPlanClassChkInfo> GetRptGPlanClassChkInfo(int SchoolYear, int Semester)
        {
            List<rptGPlanClassChkInfo> value = new List<rptGPlanClassChkInfo>();

            try
            {
                LoadMOEGroupCodeDict();

                QueryHelper qh = new QueryHelper();
                string query = "" +
                    " SELECT  " +
" graduation_plan.id AS gp_id " +
" ,graduation_plan.name AS gp_name " +
" ,graduation_plan.moe_group_code " +
" ,class.id AS class_id " +
" ,class.class_name " +
" ,class.gdc_code  " +
"   FROM  " +
"   class  " +
"   INNER JOIN  " +
"   graduation_plan  " +
"   ON class.ref_graduation_plan_id = graduation_plan.id WHERE class.id  " +
"   IN(SELECT  " +
"   	DISTINCT ref_class_id  " +
"   	FROM  " +
"   	course  " +
"   	WHERE  " +
"   	school_year = " + SchoolYear + "" +
"   	AND  " +
"   	semester = " + Semester + " " +
"   	AND ref_class_id IS NOT NULL " +
" ) ORDER BY graduation_plan.name,class_name ";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {

                    // 檢查課程規劃表如果是108前不處理
                    int gpsc;
                    string gpName = dr["gp_name"] + "";

                    if (gpName.Length > 3)
                    {
                        if (int.TryParse(gpName.Substring(0, 3), out gpsc))
                        {
                            if (gpsc < 108)
                                continue;
                        }
                    }

                    rptGPlanClassChkInfo data = new rptGPlanClassChkInfo();
                    data.GPID = dr["gp_id"] + "";
                    data.GPName = gpName;
                    data.GPMOECode = dr["moe_group_code"] + "";

                    if (MOEGroupCodeDict.ContainsKey(data.GPMOECode))
                    {
                        data.GPMOEName = MOEGroupCodeDict[data.GPMOECode];
                    }

                    data.ClassID = dr["class_id"] + "";
                    data.ClassName = dr["class_name"] + "";
                    data.ClassGDCCode = dr["gdc_code"] + "";
                    if (MOEGroupCodeDict.ContainsKey(data.ClassGDCCode))
                    {
                        data.ClassGDCName = MOEGroupCodeDict[data.ClassGDCCode];
                    }
                    data.CheckData();

                    value.Add(data);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }


        public List<rptGPlanCourseChkInfo> GetRptGPlanCourseChkInfo(int SchoolYear, int Semester)
        {
            List<rptGPlanCourseChkInfo> value = new List<rptGPlanCourseChkInfo>();
            try
            {
                LoadMOEGroupCodeDict();

                // 取得課程大表資料
                Dictionary<string, List<MOECourseCodeInfo>> MOECourseDict = GetCourseGroupCodeDict();

                string query = "" +
                    " SELECT  " +
" course.id AS course_id " +
" ,course.school_year " +
" ,course.semester " +
" ,course.course_name " +
" ,class.grade_year " +
" ,class.class_name " +
" ,class.id AS class_id " +
" ,subject " +
" ,subj_level " +
" ,c_required_by " +
" ,c_is_required " +
" ,credit" +
" ,period " +
" ,class.gdc_code " +
"  FROM course  " +
"  INNER JOIN class  " +
"  ON course.ref_class_id = class.id  " +
"  WHERE course.school_year = " + SchoolYear + " AND course.semester = " + Semester + "  " +
"  ORDER BY grade_year DESC,course_name ";


                List<string> errItem = new List<string>();

                // 取得學分對照表
                Dictionary<string, string> mappingTable = Utility.GetCreditMappingTable();

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    errItem.Clear();
                    errItem.Add("科目名稱");
                    errItem.Add("部定校訂");
                    errItem.Add("必修選修");
                    errItem.Add("學分數");

                    rptGPlanCourseChkInfo data = new rptGPlanCourseChkInfo();
                    data.ClassName = dr["class_name"] + "";
                    data.CourseID = dr["course_id"] + "";
                    data.CourseName = dr["course_name"] + "";
                    data.Credit = dr["credit"] + "";
                    data.RefClassID = dr["class_id"] + "";
                    data.SchoolYear = dr["school_year"] + "";
                    data.Semester = dr["semester"] + "";
                    data.SubjectName = dr["subject"] + "";
                    data.SubjectLevel = dr["subj_level"] + "";
                    data.GradeYear = dr["grade_year"] + "";
                    if (dr["c_required_by"] + "" == "1")
                    {
                        data.RequiredBy = "部定";
                    }

                    if (dr["c_required_by"] + "" == "2")
                    {
                        data.RequiredBy = "校訂";
                    }

                    if (dr["c_is_required"] + "" == "1")
                    {
                        data.isRequired = "必修";
                    }

                    if (dr["c_is_required"] + "" == "0")
                    {
                        data.isRequired = "選修";
                    }

                    data.gdc_code = dr["gdc_code"] + "";

                    // 比對大表資料
                    if (MOECourseDict.ContainsKey(data.gdc_code))
                    {
                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.isRequired == Mco.is_required && data.RequiredBy == Mco.require_by)
                            {
                                data.entry_year = Mco.entry_year;
                                data.credit_period = Mco.credit_period;
                                data.CourseCode = Mco.course_code;
                                data.open_type = Mco.open_type;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.isRequired == Mco.is_required)
                            {
                                errItem.Remove("科目名稱");
                                errItem.Remove("必修選修");
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.RequiredBy == Mco.require_by)
                            {
                                errItem.Remove("科目名稱");
                                errItem.Remove("部定校訂");
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name)
                            {
                                errItem.Remove("科目名稱");
                                break;
                            }
                        }

                        // 檢查學分數
                        if (data.CheckCreditPass(mappingTable))
                        {
                            errItem.Remove("學分數");
                        }


                        if (errItem.Count > 0)
                        {
                            foreach (string err in errItem)
                                data.ErrorMsgList.Add(err);
                        }
                    }
                    else
                    {
                        data.ErrorMsgList.Add("群科班代碼無法對照");
                    }

                    value.Add(data);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }


        /// <summary>
        /// 透過年級、學年度、學期，取得一般、延修 學期對照表 GDCCode
        /// </summary>
        /// <param name="GradeYear"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public Dictionary<string, string> GetStudentSemsHistoryGDCCodeByGradeYearSS(int GradeYear, int SchoolYear, int Semester)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            try
            {
                string qry = "" +
                    " WITH stud_data AS(" +
" SELECT history.id, " +
" 	name , " +
" 	('0'||array_to_string(xpath('//History/@SchoolYear', history_xml), '')::text)::integer as school_year, " +
" 	('0'||array_to_string(xpath('//History/@Semester', history_xml), '')::text)::integer as semester, " +
" 	('0'||array_to_string(xpath('//History/@GradeYear', history_xml), '')::text)::integer as grade_year, " +
" 	array_to_string(xpath('//History/@GDCCode', history_xml), '')::text as gdc_code " +
" FROM ( " +
" 		SELECT id,name, unnest(xpath('//root/History', xmlparse(content '<root>'||sems_history||'</root>'))) as history_xml " +
" 		FROM student WHERE student.status IN(1,2) " +
" 	) as history " +
" 	) " +
" 	SELECT id,name,gdc_code FROM stud_data WHERE school_year = " + SchoolYear + " AND semester = " + Semester + " AND grade_year =" + GradeYear + " ; ";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(qry);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["id"] + "";
                    string gdc_code = dr["gdc_code"] + "";
                    if (!value.ContainsKey(sid))
                        value.Add(sid, gdc_code);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        public List<rptSCAttendCodeChkInfo> GetGetStudentCourseInfoBySchoolYearSemester(int SchoolYear, int Semester, string strGrYear)
        {
            List<rptSCAttendCodeChkInfo> value = new List<rptSCAttendCodeChkInfo>();

            try
            {
                // 取得課程大表資料
                Dictionary<string, List<MOECourseCodeInfo>> MOECourseDict = GetCourseGroupCodeDict();
                List<string> errItem = new List<string>();

                // 取得學分對照表
                Dictionary<string, string> mappingTable = Utility.GetCreditMappingTable();

                QueryHelper qh = new QueryHelper();
                string query = "" +
                    " SELECT  " +
" student.id AS student_id " +
" , student.name AS student_name " +
" , student_number " +
" , student.seat_no " +
" , student.id_number " +
" , student.birthdate " +
" , class_name " +
" , class.grade_year AS grade_year " +
" , course.id AS course_id " +
" , course_name " +
" , subject " +
" , subj_level " +
" , course.ref_class_id AS c_ref_class_id " +
" , course.credit " +
" , course.period " +
" , course.school_year " +
" , course.semester " +
" , (CASE COALESCE(sc_attend.required_by,c_required_by) WHEN '1' THEN '部定' WHEN '2' THEN '校訂' ELSE '' END) AS required_by " +
" , (CASE COALESCE(sc_attend.is_required,c_is_required) WHEN '1' THEN '必修' WHEN '0' THEN '選修' ELSE '' END) AS required " +
" , COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code " +
" FROM course " +
" 	INNER JOIN sc_attend " +
"  ON course.id = sc_attend.ref_course_id  " +
" 	INNER JOIN student  " +
"  ON sc_attend.ref_student_id = student.id " +
" 	INNER JOIN class " +
" 	ON student.ref_class_id = class.id " +
" WHERE  " +
"  student.status IN(1,2) AND class.grade_year IN(" + strGrYear + ") " +
" AND course.school_year = " + SchoolYear + " AND course.semester = " + Semester + " " +
" ORDER BY class.grade_year DESC,class.display_order,class_name,seat_no,school_year,semester,course_name ";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    errItem.Clear();
                    errItem.Add("科目名稱");
                    errItem.Add("部定校訂");
                    errItem.Add("必修選修");
                    errItem.Add("學分數");

                    rptSCAttendCodeChkInfo data = new rptSCAttendCodeChkInfo();
                    data.StudentID = dr["student_id"] + "";
                    data.StudentName = dr["student_name"] + "";
                    data.StudentNumber = dr["student_number"] + "";
                    data.ClassName = dr["class_name"] + "";
                    data.SeatNo = dr["seat_no"] + "";
                    data.IDNumber = dr["id_number"] + "";
                    data.BirthDayString = ConvertChDateString(dr["birthdate"] + "");
                    data.CourseID = dr["course_id"] + "";
                    data.CourseName = dr["course_name"] + "";
                    data.SubjectName = dr["subject"] + "";
                    data.SubjectLevel = dr["subj_level"] + "";
                    data.SchoolYear = dr["school_year"] + "";
                    data.Semester = dr["semester"] + "";
                    data.RequiredBy = dr["required_by"] + "";
                    data.IsRequired = dr["required"] + "";
                    data.GradeYear = dr["grade_year"] + "";
                    data.Credit = dr["credit"] + "";
                    data.Period = dr["period"] + "";
                    if (dr["gdc_code"] != null)
                    {
                        data.gdc_code = dr["gdc_code"] + "";
                    }
                    else
                    {
                        data.gdc_code = "";
                    }


                    // 比對大表資料
                    if (MOECourseDict.ContainsKey(data.gdc_code))
                    {
                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.IsRequired == Mco.is_required && data.RequiredBy == Mco.require_by)
                            {
                                #region 2021-12-11 Cynthia 增加年級+學期條件比對大表中的open_type，取得課程代碼等資訊

                                int idx = -1;

                                if (data.GradeYear == "1" && data.Semester == "1")
                                    idx = 0;

                                if (data.GradeYear == "1" && data.Semester == "2")
                                    idx = 1;

                                if (data.GradeYear == "2" && data.Semester == "1")
                                    idx = 2;

                                if (data.GradeYear == "2" && data.Semester == "2")
                                    idx = 3;

                                if (data.GradeYear == "3" && data.Semester == "1")
                                    idx = 4;

                                if (data.GradeYear == "3" && data.Semester == "2")
                                    idx = 5;

                                char[] cp = Mco.open_type.ToArray();
                                if (idx != -1 && idx < cp.Length)
                                {
                                    if (cp[idx] != '-')
                                    {
                                        data.entry_year = Mco.entry_year;
                                        data.credit_period = Mco.credit_period;
                                        data.open_type = Mco.open_type;
                                        data.CourseCode = Mco.course_code;
                                    }
                                }

                                #endregion
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.IsRequired == Mco.is_required)
                            {
                                errItem.Remove("科目名稱");
                                errItem.Remove("必修選修");
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.RequiredBy == Mco.require_by)
                            {
                                errItem.Remove("科目名稱");
                                errItem.Remove("部定校訂");
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name)
                            {
                                errItem.Remove("科目名稱");
                                break;
                            }
                        }

                        // 檢查學分數
                        if (data.CheckCreditPass(mappingTable))
                        {
                            errItem.Remove("學分數");
                        }


                        if (errItem.Count > 0)
                        {
                            foreach (string err in errItem)
                                data.ErrorMsgList.Add(err);
                        }
                    }
                    else
                    {
                        data.ErrorMsgList.Add("群科班代碼無法對照");
                    }

                    value.Add(data);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }


        //        public List<rptSCAttendCodeChkInfo> GetStudentCourseInfoByGradeYear(int GradeYear, int SchoolYear, int Semester, bool useSemsHsitory)
        //        {
        //            List<rptSCAttendCodeChkInfo> value = new List<rptSCAttendCodeChkInfo>();
        //            try
        //            {
        //                // 取得課程大表資料
        //                Dictionary<string, List<MOECourseCodeInfo>> MOECourseDict = GetCourseGroupCodeDict();
        //                List<string> errItem = new List<string>();

        //                // 取得學分對照表
        //                Dictionary<string, string> mappingTable = Utility.GetCreditMappingTable();

        //                Dictionary<string, string> StudSemeHistoryGDCCode = new Dictionary<string, string>();

        //                if (useSemsHsitory)
        //                    StudSemeHistoryGDCCode = GetStudentSemsHistoryGDCCodeByGradeYearSS(GradeYear, SchoolYear, Semester);


        //                QueryHelper qh = new QueryHelper();
        //                string query = "" +
        //                    " SELECT  " +
        //" student.id AS student_id " +
        //" , student.name AS student_name " +
        //" , student_number " +
        //" , student.seat_no " +
        //" , class_name " +
        //" , class.grade_year AS grade_year " +
        //" , course.id AS course_id " +
        //" , course_name " +
        //" , subject " +
        //" , subj_level " +
        //" , course.ref_class_id AS c_ref_class_id " +
        //" , course.credit " +
        //" , course.period " +
        //" , course.school_year " +
        //" , course.semester " +
        //" , (CASE COALESCE(sc_attend.required_by,c_required_by) WHEN '1' THEN '部定' WHEN '2' THEN '校訂' ELSE '' END) AS required_by " +
        //" , (CASE COALESCE(sc_attend.is_required,c_is_required) WHEN '1' THEN '必修' WHEN '0' THEN '選修' ELSE '' END) AS required " +
        //" , COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code " +
        //" FROM course " +
        //" 	INNER JOIN sc_attend " +
        //"  ON course.id = sc_attend.ref_course_id  " +
        //" 	INNER JOIN student  " +
        //"  ON sc_attend.ref_student_id = student.id " +
        //" 	INNER JOIN class " +
        //" 	ON student.ref_class_id = class.id " +
        //" WHERE  " +
        //"  student.status IN(1,2) AND class.grade_year = " + GradeYear + "  " +
        //" AND course.school_year = " + SchoolYear + " AND course.semester = " + Semester + " " +
        //" ORDER BY class.display_order,class_name,seat_no,school_year,semester,course_name ";

        //                DataTable dt = qh.Select(query);
        //                foreach (DataRow dr in dt.Rows)
        //                {
        //                    errItem.Clear();
        //                    errItem.Add("科目名稱");
        //                    errItem.Add("部定校訂");
        //                    errItem.Add("必修選修");
        //                    errItem.Add("學分數");

        //                    rptSCAttendCodeChkInfo data = new rptSCAttendCodeChkInfo();
        //                    data.StudentID = dr["student_id"] + "";
        //                    data.StudentName = dr["student_name"] + "";
        //                    data.StudentNumber = dr["student_number"] + "";
        //                    data.ClassName = dr["class_name"] + "";
        //                    data.SeatNo = dr["seat_no"] + "";
        //                    data.CourseID = dr["course_id"] + "";
        //                    data.CourseName = dr["course_name"] + "";
        //                    data.SubjectName = dr["subject"] + "";
        //                    data.SubjectLevel = dr["subj_level"] + "";
        //                    data.SchoolYear = dr["school_year"] + "";
        //                    data.Semester = dr["semester"] + "";
        //                    data.RequiredBy = dr["required_by"] + "";
        //                    data.IsRequired = dr["required"] + "";
        //                    data.GradeYear = dr["grade_year"] + "";
        //                    data.Credit = dr["credit"] + "";
        //                    data.Period = dr["period"] + "";
        //                    if (dr["gdc_code"] != null)
        //                    {
        //                        data.gdc_code = dr["gdc_code"] + "";
        //                    }
        //                    else
        //                    {
        //                        data.gdc_code = "";
        //                    }


        //                    // 比對大表資料
        //                    if (MOECourseDict.ContainsKey(data.gdc_code))
        //                    {
        //                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
        //                        {
        //                            if (data.SubjectName == Mco.subject_name && data.IsRequired == Mco.is_required && data.RequiredBy == Mco.require_by)
        //                            {
        //                                data.entry_year = Mco.entry_year;
        //                                data.credit_period = Mco.credit_period;
        //                                data.CourseCode = Mco.course_code;
        //                                break;
        //                            }
        //                        }

        //                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
        //                        {
        //                            if (data.SubjectName == Mco.subject_name && data.IsRequired == Mco.is_required)
        //                            {
        //                                errItem.Remove("科目名稱");
        //                                errItem.Remove("必修選修");
        //                                break;
        //                            }
        //                        }

        //                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
        //                        {
        //                            if (data.SubjectName == Mco.subject_name && data.RequiredBy == Mco.require_by)
        //                            {
        //                                errItem.Remove("科目名稱");
        //                                errItem.Remove("部定校訂");
        //                                break;
        //                            }
        //                        }

        //                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
        //                        {
        //                            if (data.SubjectName == Mco.subject_name)
        //                            {
        //                                errItem.Remove("科目名稱");
        //                                break;
        //                            }
        //                        }

        //                        // 檢查學分數
        //                        if (data.CheckCreditPass(mappingTable))
        //                        {
        //                            errItem.Remove("學分數");
        //                        }


        //                        if (errItem.Count > 0)
        //                        {
        //                            foreach (string err in errItem)
        //                                data.ErrorMsgList.Add(err);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        data.ErrorMsgList.Add("群科班代碼無法對照");
        //                    }

        //                    value.Add(data);
        //                }

        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine(ex.Message);
        //            }
        //            return value;
        //        }

        public List<DataRow> GetHasGDCCodeStudent(string strGrYear)
        {
            List<DataRow> value = new List<DataRow>();
            try
            {
                QueryHelper qh = new QueryHelper();
                string query = "" +
                    " SELECT  " +
" student.id AS student_id " +
" , student.name AS student_name " +
" , student_number " +
" , student.seat_no " +
" , class_name " +
" , class.grade_year AS grade_year " +
" , COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code " +
" FROM student  " +
" 	INNER JOIN class " +
" 	ON student.ref_class_id = class.id " +
" WHERE  " +
"  student.status IN(1,2) AND class.grade_year IN(" + strGrYear + ") " +
"  AND COALESCE(student.gdc_code,class.gdc_code)  IS NOT NULL " +
" ORDER BY class.grade_year DESC,class.display_order,class_name,seat_no ";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                    value.Add(dr);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        public List<rptStudSemsScoreCodeChkInfo> GetStudentSemsScoreInfoByGradeYear(int GradeYear)
        {
            List<rptStudSemsScoreCodeChkInfo> value = new List<rptStudSemsScoreCodeChkInfo>();
            try
            {
                // 取得課程大表資料
                Dictionary<string, List<MOECourseCodeInfo>> MOECourseDict = GetCourseGroupCodeDict();
                List<string> errItem = new List<string>();

                // 取得學分對照表
                Dictionary<string, string> mappingTable = Utility.GetCreditMappingTable();

                QueryHelper qh = new QueryHelper();
                string query = "" +
                    " WITH student_data AS (  " +
"  	SELECT  " +
"  	student.id AS student_id  " +
"  	,student_number  " +
"  	,class_name  " +
"  	,student.seat_no  " +
"  	,student.name AS student_name  " +
"  	, COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code  " +
"  FROM student   " +
"  INNER JOIN class ON student.ref_class_id = class.id   " +
"  	WHERE student.status IN(1,2) AND class.grade_year  IN( " + GradeYear + " ) " +
"  ),sems_score_data AS(  " +
"  SELECT  " +
"  	sems_subj_score_ext.ref_student_id  " +
"  	, sems_subj_score_ext.grade_year  " +
"  	, sems_subj_score_ext.semester  " +
"  	, sems_subj_score_ext.school_year	  " +
"  	, array_to_string(xpath('//Subject/@科目', subj_score_ele), '')::text AS 科目  " +
"  	, array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '')::text AS 科目級別  " +
"  	, array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '')::text AS 學分數	  " +
"  	, array_to_string(xpath('//Subject/@修課必選修', subj_score_ele), '')::text AS 必選修  " +
"  	, array_to_string(xpath('//Subject/@修課校部訂', subj_score_ele), '')::text AS 校部訂	  " +
"  FROM (  " +
"  		SELECT   " +
"  			sems_subj_score.*  " +
"  			, 	unnest(xpath('//SemesterSubjectScoreInfo/Subject', xmlparse(content score_info))) as subj_score_ele  " +
"  		FROM   " +
"  			sems_subj_score   " +
"  			INNER JOIN student_data ON sems_subj_score.ref_student_id = student_data.student_id  " +
"  	) as sems_subj_score_ext   " +
"  )  " +
"  SELECT   " +
"  student_id  " +
"  ,student_number  " +
"  ,class_name  " +
"  ,seat_no  " +
"  ,student_name  " +
"  ,gdc_code  " +
"  ,school_year  " +
"  ,semester  " +
"  ,grade_year  " +
"  ,科目 AS subject " +
"  ,科目級別 AS subj_level " +
"  ,學分數 AS credit " +
"  ,必選修 AS required " +
"  ,(CASE 校部訂 WHEN '部訂' THEN '部定' ELSE 校部訂 END) AS required_by  " +
"   FROM student_data INNER JOIN sems_score_data ON student_data.student_id = sems_score_data.ref_student_id " +
"    ORDER BY class_name,seat_no,school_year,semester ";


                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    errItem.Clear();
                    errItem.Add("科目名稱");
                    errItem.Add("部定校訂");
                    errItem.Add("必修選修");
                    errItem.Add("學分數");

                    rptStudSemsScoreCodeChkInfo data = new rptStudSemsScoreCodeChkInfo();
                    data.StudentID = dr["student_id"] + "";
                    data.StudentName = dr["student_name"] + "";
                    data.StudentNumber = dr["student_number"] + "";
                    data.ClassName = dr["class_name"] + "";
                    data.SeatNo = dr["seat_no"] + "";
                    data.SubjectName = dr["subject"] + "";
                    data.SubjectLevel = dr["subj_level"] + "";
                    data.SchoolYear = dr["school_year"] + "";
                    data.Semester = dr["semester"] + "";
                    data.RequiredBy = dr["required_by"] + "";
                    data.IsRequired = dr["required"] + "";
                    data.GradeYear = dr["grade_year"] + "";
                    data.Credit = dr["credit"] + "";
                    if (dr["gdc_code"] != null)
                    {
                        data.gdc_code = dr["gdc_code"] + "";
                    }
                    else
                    {
                        data.gdc_code = "";
                    }


                    // 比對大表資料
                    if (MOECourseDict.ContainsKey(data.gdc_code))
                    {
                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.IsRequired == Mco.is_required && data.RequiredBy == Mco.require_by)
                            {
                                data.entry_year = Mco.entry_year;
                                data.credit_period = Mco.credit_period;
                                data.CourseCode = Mco.course_code;
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.IsRequired == Mco.is_required)
                            {
                                errItem.Remove("科目名稱");
                                errItem.Remove("必修選修");
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.RequiredBy == Mco.require_by)
                            {
                                errItem.Remove("科目名稱");
                                errItem.Remove("部定校訂");
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name)
                            {
                                errItem.Remove("科目名稱");
                                break;
                            }
                        }

                        // 檢查學分數
                        if (data.CheckCreditPass(mappingTable))
                        {
                            errItem.Remove("學分數");
                        }


                        if (errItem.Count > 0)
                        {
                            foreach (string err in errItem)
                                data.ErrorMsgList.Add(err);
                        }
                    }
                    else
                    {
                        data.ErrorMsgList.Add("群科班代碼無法對照");
                    }

                    value.Add(data);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        public List<GPlanInfo108> GPlanInfoOld108List()
        {
            List<GPlanInfo108> value = new List<GPlanInfo108>();

            try
            {
                // 取得群科班對照
                LoadMOEGroupCodeDict();

                // 取得課程代碼大表對照
                Dictionary<string, List<MOECourseCodeInfo>> MOECourseCodeDict = GetCourseGroupCodeDict();

                //// 取得課程規劃表，新格式
                //string query = "SELECT id,name,moe_group_code,content,array_to_string(xpath('//GraduationPlan/@EntryYear', xmlparse(content content)), '')::text AS entry_year FROM graduation_plan WHERE array_to_string(xpath('//GraduationPlan/@EntryYear', xmlparse(content content)), '')::text <>'' AND moe_group_code<>''";

                string query = "SELECT id,name,moe_group_code,content,array_to_string(xpath('//GraduationPlan/@SchoolYear', xmlparse(content content)), '')::text AS entry_year FROM graduation_plan WHERE array_to_string(xpath('//GraduationPlan/@SchoolYear', xmlparse(content content)), '')::text <>'' AND moe_group_code<>''";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);

                Dictionary<string, List<DataRow>> dtDict = new Dictionary<string, List<DataRow>>();
                foreach (DataRow dr in dt.Rows)
                {
                    //string moe_group_code = dr["moe_group_code"] + "";

                    string gname = dr["name"] + "";


                    if (!dtDict.ContainsKey(gname))
                        dtDict.Add(gname, new List<DataRow>());

                    dtDict[gname].Add(dr);
                }

                // 建立資料
                foreach (string name in dtDict.Keys)
                {
                    GPlanInfo108 data = new GPlanInfo108();
                    string code = "";

                    code = dtDict[name][0]["moe_group_code"] + "";
                    data.GPlanList = dtDict[name];
                    data.ParseRefGPContentXml();

                    data.GDCCode = code;
                    if (code.Length > 3)
                    {
                        data.EntrySchoolYear = code.Substring(0, 3);
                    }
                    data.GDCName = MOEGroupCodeDict[code];
                    if (MOEGPlanDict.ContainsKey(code))
                    {
                        // 解析出來課程規劃表名稱
                        data.RefGPName = MOEGPlanDict[code];
                    }


                    // 填入課程規劃表大表
                    if (MOECourseCodeDict.ContainsKey(code))
                    {
                        data.MOECourseCodeInfoList = MOECourseCodeDict[code].OrderBy(x => x.course_code).ToList();
                    }

                    data.Status = "無變動";
                    data.ParseOrderByInt();


                    value.Add(data);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            return value;
        }



        /// <summary>
        /// 取得所有群科班對照108新版
        /// </summary>
        /// <returns></returns>
        public List<GPlanInfo108> GPlanInfo108List()
        {
            List<GPlanInfo108> value = new List<GPlanInfo108>();

            try
            {
                // 取得群科班對照
                LoadMOEGroupCodeDict();

                // 取得課程代碼大表對照
                Dictionary<string, List<MOECourseCodeInfo>> MOECourseCodeDict = GetCourseGroupCodeDict();

                // 取得課程規劃表，新格式
                string query = "SELECT id,name,moe_group_code,content,array_to_string(xpath('//GraduationPlan/@EntryYear', xmlparse(content content)), '')::text AS entry_year FROM graduation_plan WHERE array_to_string(xpath('//GraduationPlan/@EntryYear', xmlparse(content content)), '')::text <>'' AND moe_group_code<>''";

                //string query = "SELECT id,name,moe_group_code,content,array_to_string(xpath('//GraduationPlan/@SchoolYear', xmlparse(content content)), '')::text AS entry_year FROM graduation_plan WHERE array_to_string(xpath('//GraduationPlan/@SchoolYear', xmlparse(content content)), '')::text <>'' AND moe_group_code<>''";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                Dictionary<string, List<DataRow>> dtDict = new Dictionary<string, List<DataRow>>();
                foreach (DataRow dr in dt.Rows)
                {
                    string moe_group_code = dr["moe_group_code"] + "";
                    if (!dtDict.ContainsKey(moe_group_code))
                        dtDict.Add(moe_group_code, new List<DataRow>());

                    dtDict[moe_group_code].Add(dr);
                }

                // 建立資料
                foreach (string code in MOEGroupCodeDict.Keys)
                {
                    GPlanInfo108 data = new GPlanInfo108();
                    data.GDCCode = code;
                    if (code.Length > 3)
                    {
                        data.EntrySchoolYear = code.Substring(0, 3);
                    }
                    data.GDCName = MOEGroupCodeDict[code];
                    if (MOEGPlanDict.ContainsKey(code))
                    {
                        // 解析出來課程規劃表名稱
                        data.RefGPName = MOEGPlanDict[code];
                    }


                    // 填入課程規劃表大表
                    if (MOECourseCodeDict.ContainsKey(code))
                    {
                        data.MOECourseCodeInfoList = MOECourseCodeDict[code].OrderBy(x => x.course_code).ToList();
                    }

                    // 放入課程規劃表原始
                    if (dtDict.ContainsKey(code))
                    {
                        data.GPlanList = dtDict[code];
                        // id 
                        data.RefGPID = dtDict[code][0] + "";
                    }


                    data.Status = "無變動";
                    data.ParseOrderByInt();


                    value.Add(data);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            return value;
        }

        /// <summary>
        /// 取得目前班級年級
        /// </summary>
        /// <returns></returns>
        public List<string> GetClassGradeYear()
        {
            List<string> value = new List<string>();

            try
            {
                QueryHelper qh = new QueryHelper();
                string query = "SELECT DISTINCT class.grade_year FROM student INNER JOIN class ON student.ref_class_id = class.id WHERE student.status = 1 AND class.grade_year IS NOT NULL ORDER BY class.grade_year ASC;";

                DataTable dt = qh.Select(query);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                        value.Add(dr["grade_year"] + "");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        /// 將西元日期轉成民國數字
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public string ConvertChDateString(string dtStr)
        {
            DateTime dt;

            string value = "";
            if (DateTime.TryParse(dtStr, out dt))
            {
                value = string.Format("{0:000}", (dt.Year - 1911)) + string.Format("{0:00}", dt.Month) + string.Format("{0:00}", dt.Day);
            }

            return value;
        }

        public Dictionary<string, List<chkCourseInfo>> GetCourseHasAttendBySchoolYearSemester(int SchoolYear, int Semester)
        {
            Dictionary<string, List<chkCourseInfo>> value = new Dictionary<string, List<chkCourseInfo>>();

            try
            {
                QueryHelper qh = new QueryHelper();
                string qry = "" +
                    " SELECT  " +
    " DISTINCT course.id AS course_id " +
    " , course_name " +
    " , subject " +
    " , COALESCE(course.credit,0) AS credit " +
    " , class.grade_year " +
    " , course.school_year " +
    " , course.semester " +
    " , (CASE c_required_by WHEN '1' THEN '部定' WHEN '2' THEN '校訂' ELSE '' END) AS required_by " +
    " , (CASE c_is_required WHEN '1' THEN '必修' WHEN '0' THEN '選修' ELSE '' END) AS required " +
    " , COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code " +
    " FROM course " +
    " 	INNER JOIN sc_attend " +
    "  ON course.id = sc_attend.ref_course_id  " +
    " 	INNER JOIN student  " +
    "  ON sc_attend.ref_student_id = student.id " +
    " 	INNER JOIN class " +
    " 	ON student.ref_class_id = class.id " +
    " WHERE  " +
    "  student.status = 1 AND  course.school_year = " + SchoolYear + " AND course.semester = " + Semester + "  " +
    " ORDER BY course_name ";

                DataTable dt = qh.Select(qry);
                foreach (DataRow dr in dt.Rows)
                {
                    string key = dr["gdc_code"] + "";

                    chkCourseInfo data = new chkCourseInfo();
                    data.CourseID = dr["course_id"] + "";
                    data.CourseName = dr["course_name"] + "";
                    data.credit = dr["credit"] + "";
                    data.required = dr["required"] + "";
                    data.required_by = dr["required_by"] + "";
                    data.SubjectName = dr["subject"] + "";

                    if (!value.ContainsKey(key))
                        value.Add(key, new List<chkCourseInfo>());

                    value[key].Add(data);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }


        public List<GPlanInfo108> GetGPlanInfo108All()
        {
            List<GPlanInfo108> value = new List<GPlanInfo108>();
            try
            {
                QueryHelper qh = new QueryHelper();
                string qry = "" +
                    " WITH entry_year_data AS(" +
" SELECT  " +
" 			id             " +
" 			, unnest(xpath('//GraduationPlan/@EntryYear', xmlparse(content content)))::TEXT AS entry_year " +
" 		FROM  " +
" 			graduation_plan) " +
" ,g_plan_data AS( " +
" SELECT entry_year_data.entry_year,graduation_plan.id,graduation_plan.name,graduation_plan.content,graduation_plan.moe_group_code FROM graduation_plan INNER JOIN entry_year_data ON graduation_plan.id = entry_year_data.id  " +
" ) " +
" SELECT * FROM g_plan_data ORDER BY entry_year,name;";

                DataTable dt = qh.Select(qry);

                foreach (DataRow dr in dt.Rows)
                {
                    GPlanInfo108 data = new GPlanInfo108();
                    data.EntrySchoolYear = dr["entry_year"] + "";
                    data.GDCCode = dr["moe_group_code"] + "";
                    data.RefGPContent = dr["content"] + "";
                    data.ParseRefGPContentXml();
                    data.RefGPID = dr["id"] + "";
                    data.RefGPName = dr["name"] + "";
                    value.Add(data);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        public Dictionary<string, List<DataRow>> GetGPlanRefClaasByID(string gp_id)
        {
            Dictionary<string, List<DataRow>> value = new Dictionary<string, List<DataRow>>();

            QueryHelper qh = new QueryHelper();
            string qry = " SELECT " +
" class.grade_year" +
" ,class_name" +
" ,COUNT(student.id) AS stud_cot " +
" 	FROM " +
" class " +
" LEFT JOIN " +
" student " +
" ON class.id = student.ref_class_id " +
" WHERE class.ref_graduation_plan_id = " + gp_id + " " +
"  AND student.status = 1 " +
"  GROUP BY class.grade_year,class_name " +
"  ORDER BY class.grade_year,class_name";

            DataTable dt = qh.Select(qry);

            foreach (DataRow dr in dt.Rows)
            {
                string gr = dr["grade_year"] + "";
                if (gr == "")
                    gr = "未分";
                if (!value.ContainsKey(gr))
                    value.Add(gr, new List<DataRow>());
                value[gr].Add(dr);
            }

            return value;
        }


        /// <summary>
        /// 透過班級系統編號取得相關課程規劃
        /// </summary>
        /// <param name="classIDs"></param>
        /// <returns></returns>

        public List<CClassCourseInfo> GetCClassCourseInfoList(List<string> classIDs)
        {
            List<CClassCourseInfo> value = new List<CClassCourseInfo>();

            List<string> gpidList = new List<string>();

            // 取得班級資訊
            QueryHelper qh = new QueryHelper();
            string qry = "" +
                " SELECT  " +
" 	id " +
" 	,class_name " +
" 	,grade_year " +
" 	,ref_graduation_plan_id " +
"  FROM  " +
" 	class  " +
"  WHERE id IN(" + string.Join(",", classIDs.ToArray()) + ")  " +
" 	ORDER BY  " +
"  grade_year DESC " +
"  ,display_order " +
"  ,class_name ";

            DataTable dt = qh.Select(qry);
            foreach (DataRow dr in dt.Rows)
            {
                CClassCourseInfo data = new CClassCourseInfo();
                data.ClassID = dr["id"] + "";
                data.ClassName = dr["class_name"] + "";
                data.GradeYear = dr["grade_year"] + "";
                data.RefGPID = dr["ref_graduation_plan_id"] + "";
                if (!gpidList.Contains(data.RefGPID))
                    gpidList.Add(data.RefGPID);

                value.Add(data);
            }

            // 取得課程規劃
            string qryGp = "" +
                " WITH entry_year_data AS(" +
" SELECT  " +
" 			id             " +
" 			, unnest(xpath('//GraduationPlan/@EntryYear', xmlparse(content content)))::TEXT AS entry_year " +
" 		FROM  " +
" 			graduation_plan WHERE id IN(" + string.Join(",", gpidList.ToArray()) + ")) " +
" ,g_plan_data AS( " +
" SELECT entry_year_data.entry_year,graduation_plan.id,graduation_plan.name,graduation_plan.content,graduation_plan.moe_group_code FROM graduation_plan INNER JOIN entry_year_data ON graduation_plan.id = entry_year_data.id  " +
" ) " +
" SELECT * FROM g_plan_data ORDER BY entry_year,name";

            Dictionary<string, DataRow> gpDr = new Dictionary<string, DataRow>();

            DataTable dtgp = qh.Select(qryGp);
            foreach (DataRow dr in dtgp.Rows)
            {
                string id = dr["id"] + "";
                if (!gpDr.ContainsKey(id))
                    gpDr.Add(id, dr);
            }

            foreach (CClassCourseInfo cc in value)
            {
                if (gpDr.ContainsKey(cc.RefGPID))
                {
                    cc.RefGPName = gpDr[cc.RefGPID]["name"] + "";
                    try
                    {
                        cc.RefGPlanXML = XElement.Parse(gpDr[cc.RefGPID]["content"] + "");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }



            return value;
        }


        public Dictionary<string, List<string>> GetClassStudentDict(List<string> classIDs)
        {
            Dictionary<string, List<string>> value = new Dictionary<string, List<string>>();

            try
            {
                QueryHelper qh = new QueryHelper();
                string qry = "SELECT class.id AS c_id,student.id AS s_id FROM class INNER JOIN student ON class.id = student.ref_class_id WHERE student.status IN(1,2) AND class.id IN(" + string.Join(",", classIDs.ToArray()) + ");";
                DataTable dt = qh.Select(qry);
                foreach (DataRow dr in dt.Rows)
                {
                    string cid = dr["c_id"] + "";
                    string sid = dr["s_id"] + "";
                    if (!value.ContainsKey(cid))
                        value.Add(cid, new List<string>());

                    value[cid].Add(sid);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }



        /// <summary>
        /// 取得目前系統內課程ID
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public Dictionary<string, string> GetHasCourseIDDict(string SchoolYear, string Semester)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            try
            {
                string qryCourse = "" +
                 " SELECT  " +
" 	course_name, " +
" 	id " +
"  FROM course  " +
"  WHERE  " +
"  school_year = " + SchoolYear + "  " +
"  AND semester = " + Semester + " " +
"  ORDER BY course_name; ";
                DataTable dtCourse = qh.Select(qryCourse);
                foreach (DataRow dr in dtCourse.Rows)
                {
                    string coname = (dr["course_name"] + "").Trim();
                    if (!value.ContainsKey(coname))
                        value.Add(coname, dr["id"] + "");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        /// <summary>
        /// 原班開課
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="dataList"></param>
        /// <returns></returns>
        public List<CClassCourseInfo> AddGPlanCourseBySchoolYearSemester(string SchoolYear, string Semester, List<CClassCourseInfo> dataList)
        {
            // 取得系統內課程
            Dictionary<string, string> hasCourseIDDict = new Dictionary<string, string>();
            List<string> insertSQLList = new List<string>();

            hasCourseIDDict = GetHasCourseIDDict(SchoolYear, Semester);

            try
            {
                QueryHelper qh = new QueryHelper();

                // 處理一般開課
                foreach (CClassCourseInfo data in dataList)
                {
                    data.AddCourseNameList.Clear();

                    foreach (XElement subjElm in data.OpenSubjectSourceList)
                    {
                        // 檢查課程名稱是否存在
                        string courseName = data.ClassName + " " + subjElm.Attribute("SubjectName").Value;
                        string chkName = courseName.Trim();
                        if (!data.AddCourseNameList.Contains(courseName))
                            data.AddCourseNameList.Add(courseName);

                        // 已存在跳過
                        if (hasCourseIDDict.ContainsKey(chkName))
                        {
                            continue;
                        }
                        string isReq = "", ReqBy = "";

                        if (subjElm.Attribute("RequiredBy").Value == "部訂" || subjElm.Attribute("RequiredBy").Value == "部定")
                        {
                            ReqBy = "1";
                        }
                        else
                        {
                            ReqBy = "2";
                        }

                        if (subjElm.Attribute("Required").Value == "必修")
                            isReq = "1";
                        else
                            isReq = "0";

                        string insStr = insertCourseSQL(courseName, subjElm.Attribute("Level").Value, subjElm.Attribute("SubjectName").Value, data.ClassID, SchoolYear, Semester, subjElm.Attribute("Credit").Value, subjElm.Attribute("Entry").Value, ReqBy, isReq, subjElm.Attribute("Credit").Value, subjElm.Attribute("Domain").Value);

                        if (!insertSQLList.Contains(insStr))
                            insertSQLList.Add(insStr);


                    }


                    // 處理對開課程
                    foreach (XElement subjElm in data.OpenSubjectSourceBList)
                    {
                        string subjName = subjElm.Attribute("SubjectName").Value;

                        if (data.SubjectBDict.ContainsKey(subjName))
                        {
                            if (data.SubjectBDict[subjName] == true)
                            {
                                // 檢查課程名稱是否存在
                                string courseName = data.ClassName + " " + subjElm.Attribute("SubjectName").Value;
                                string chkName = courseName.Trim();
                                if (!data.AddCourseNameList.Contains(courseName))
                                    data.AddCourseNameList.Add(courseName);
                                // 已存在跳過
                                if (hasCourseIDDict.ContainsKey(chkName))
                                {
                                    continue;
                                }
                                string isReq = "", ReqBy = "";

                                if (subjElm.Attribute("RequiredBy").Value == "部訂" || subjElm.Attribute("RequiredBy").Value == "部定")
                                {
                                    ReqBy = "1";
                                }
                                else
                                {
                                    ReqBy = "2";
                                }

                                if (subjElm.Attribute("Required").Value == "必修")
                                    isReq = "1";
                                else
                                    isReq = "0";


                                string inStr = insertCourseSQL(courseName, subjElm.Attribute("Level").Value, subjElm.Attribute("SubjectName").Value, data.ClassID, SchoolYear, Semester, subjElm.Attribute("Credit").Value, subjElm.Attribute("Entry").Value, ReqBy, isReq, subjElm.Attribute("Credit").Value, subjElm.Attribute("Domain").Value);

                                if (!insertSQLList.Contains(inStr))
                                    insertSQLList.Add(inStr);
                            }
                        }
                    }
                }


                // debug write text file
                using (StreamWriter sw = new StreamWriter(Application.StartupPath + "\\debug.txt", false))
                {
                    foreach (string sid in insertSQLList)
                    {
                        sw.WriteLine(sid);
                    }
                }


                if (insertSQLList.Count > 0)
                {
                    // 執行寫入
                    K12.Data.UpdateHelper uh = new K12.Data.UpdateHelper();
                    uh.Execute(insertSQLList);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return dataList;
        }

        /// <summary>
        /// 原班開課加入學生
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="dataList"></param>
        /// <returns></returns>
        public string AddCourseStudent(string SchoolYear, string Semester, List<CClassCourseInfo> dataList)
        {
            Dictionary<string, string> hasCourseIDDict = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            // 取得已有課程
            hasCourseIDDict = GetHasCourseIDDict(SchoolYear, Semester);
            List<string> insertCourseStudentSQL = new List<string>();
            Dictionary<string, List<string>> hasCourseStudent = new Dictionary<string, List<string>>();
            // 取得已有修課
            try
            {
                string qry = "" +
                    "SELECT " +
                    "course_name" +
                    ",sc_attend.ref_student_id " +
                    " FROM course " +
                    "INNER JOIN " +
                    "sc_attend ON course.id = sc_attend.ref_course_id " +
                    " WHERE course.school_year = " + SchoolYear + " AND course.semester = " + Semester +
                    " ORDER BY course_name;";

                DataTable dt = qh.Select(qry);
                foreach (DataRow dr in dt.Rows)
                {
                    string cname = dr["course_name"] + "";
                    string sid = dr["ref_student_id"] + "";
                    if (!hasCourseStudent.ContainsKey(cname))
                        hasCourseStudent.Add(cname, new List<string>());

                    hasCourseStudent[cname].Add(sid);
                }

                foreach (CClassCourseInfo data in dataList)
                {
                    // 檢查課程是否已有修課學生
                    foreach (string cname in data.AddCourseNameList)
                    {
                        string courseid = "";
                        if (hasCourseIDDict.ContainsKey(cname))
                            courseid = hasCourseIDDict[cname];

                        foreach (string sid in data.RefStudentIDList)
                        {
                            if (hasCourseStudent.ContainsKey(cname))
                            {
                                if (hasCourseStudent[cname].Contains(sid))
                                    continue;
                            }

                            string insSCattned = "INSERT INTO sc_attend(" +
                                "ref_course_id" +
                                ",ref_student_id) " +
                                "VALUES(" + courseid + "," + sid + ");";

                            insertCourseStudentSQL.Add(insSCattned);
                        }

                    }
                }

                // 新增資料
                if (insertCourseStudentSQL.Count > 0)
                {
                    using (StreamWriter sw = new StreamWriter(Application.StartupPath + "\\debug1.txt", false))
                    {
                        foreach (string sid in insertCourseStudentSQL)
                        {
                            sw.WriteLine(sid);
                        }
                    }
                    K12.Data.UpdateHelper uh = new K12.Data.UpdateHelper();
                    uh.Execute(insertCourseStudentSQL);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return ex.Message;
            }

            return "";
        }


        /// <summary>
        /// 新增課程資料
        /// </summary>
        /// <param name="course_name"></param>
        /// <param name="subj_level"></param>
        /// <param name="subject"></param>
        /// <param name="ref_class_id"></param>
        /// <param name="school_year"></param>
        /// <param name="semester"></param>
        /// <param name="credit"></param>
        /// <param name="score_type"></param>
        /// <param name="c_required_by"></param>
        /// <param name="c_is_required"></param>
        /// <param name="period"></param>
        /// <param name="domain"></param>
        /// <returns></returns>
        public string insertCourseSQL(string course_name, string subj_level, string subject, string ref_class_id, string school_year, string semester, string credit, string score_type, string c_required_by, string c_is_required, string period, string domain)
        {
            string refClassStr = "null";

            if (ref_class_id != "")
                refClassStr = "'" + ref_class_id + "'";



            string value = "" +
                " INSERT INTO course(" +
" 	course_name " +
" 	,subj_level " +
" 	,subject " +
" 	,ref_class_id " +
" 	,school_year " +
" 	,semester " +
" 	,credit " +
" 	,score_type " +
" 	,c_required_by " +
" 	,c_is_required " +
" 	,period " +
" 	,domain " +
" ) " +
" VALUES ( " +
" 	'" + course_name + "' " +
" 	,'" + subj_level + "' " +
" 	,'" + subject + "' " +
" 	," + refClassStr + " " +
" 	," + school_year + "" +
" 	," + semester + " " +
" 	," + credit + " " +
" 	,'" + score_type + "' " +
" 	,'" + c_required_by + "' " +
" 	,'" + c_is_required + "' " +
" 	,'" + period + "' " +
" 	,'" + domain + "' " +
" ); ";

            return value;
        }

        /// <summary>
        /// 跨班開課
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <param name="dataList"></param>
        public string AddGPlanCourse_C_BySchoolYearSemester(string SchoolYear, string Semester, List<SubjectCourseInfo> dataList)
        {
            // 取得系統內課程
            Dictionary<string, string> hasCourseIDDict = new Dictionary<string, string>();
            List<string> insertSQLList = new List<string>();

            hasCourseIDDict = GetHasCourseIDDict(SchoolYear, Semester);

            try
            {
                QueryHelper qh = new QueryHelper();
                // 處理跨班課程
                foreach (SubjectCourseInfo subj in dataList)
                {
                    // 開課數大於0才需要開課
                    if (subj.CourseCount > 0)
                    {
                        if (subj.CourseCount == 1)
                        {
                            string courseName = subj.SubjectXML.Attribute
                                ("SubjectName").Value;

                            string chkName = courseName.Trim();

                            if (hasCourseIDDict.ContainsKey(chkName))
                            {
                                continue;
                            }

                            string isReq = "", ReqBy = "";

                            if (subj.SubjectXML.Attribute("RequiredBy").Value == "部訂" || subj.SubjectXML.Attribute("RequiredBy").Value == "部定")
                            {
                                ReqBy = "1";
                            }
                            else
                            {
                                ReqBy = "2";
                            }

                            if (subj.SubjectXML.Attribute("Required").Value == "必修")
                                isReq = "1";
                            else
                                isReq = "0";

                            string insStr = insertCourseSQL(courseName, subj.SubjectXML.Attribute("Level").Value, subj.SubjectXML.Attribute("SubjectName").Value, "", SchoolYear, Semester, subj.SubjectXML.Attribute("Credit").Value, subj.SubjectXML.Attribute("Entry").Value, ReqBy, isReq, subj.SubjectXML.Attribute("Credit").Value, subj.SubjectXML.Attribute("Domain").Value);

                            if (!insertSQLList.Contains(insStr))
                                insertSQLList.Add(insStr);
                        }
                        else
                        {
                            for (int i = 1; i <= subj.CourseCount; i++)
                            {
                                string courseName = subj.SubjectXML.Attribute
                                ("SubjectName").Value + " " + Convert.ToChar(64 + i);
                                string chkName = courseName.Trim();

                                if (hasCourseIDDict.ContainsKey(chkName))
                                {
                                    continue;
                                }

                                string isReq = "", ReqBy = "";

                                if (subj.SubjectXML.Attribute("RequiredBy").Value == "部訂" || subj.SubjectXML.Attribute("RequiredBy").Value == "部定")
                                {
                                    ReqBy = "1";
                                }
                                else
                                {
                                    ReqBy = "2";
                                }

                                if (subj.SubjectXML.Attribute("Required").Value == "必修")
                                    isReq = "1";
                                else
                                    isReq = "0";

                                string insStr = insertCourseSQL(courseName, subj.SubjectXML.Attribute("Level").Value, subj.SubjectXML.Attribute("SubjectName").Value, "", SchoolYear, Semester, subj.SubjectXML.Attribute("Credit").Value, subj.SubjectXML.Attribute("Entry").Value, ReqBy, isReq, subj.SubjectXML.Attribute("Credit").Value, subj.SubjectXML.Attribute("Domain").Value);

                                if (!insertSQLList.Contains(insStr))
                                    insertSQLList.Add(insStr);
                            }
                        }
                    }
                }

                if (insertSQLList.Count > 0)
                {
                    // 執行寫入
                    K12.Data.UpdateHelper uh = new K12.Data.UpdateHelper();
                    uh.Execute(insertSQLList);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return ex.Message;
            }
            return "";
        }

        /// <summary>
        /// 取得目標學生中 學期對照表不完整的學生。
        /// </summary>
        /// <param name="strGrYear"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public DataTable GetStudentListNotHasSemsHistory(string strGrYear, int SchoolYear, int Semester)
        {
            DataTable value = new DataTable();
            try
            {
                QueryHelper qh = new QueryHelper();
                string query = @" WITH student_data AS(
 SELECT   
 student.id AS student_id  
 , student.name AS student_name  
 , student_number  
 , student.seat_no  
 , class_name  
 , class.grade_year AS grade_year  
 , COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code  
 , class.display_order
 FROM student   
 	INNER JOIN class  
 	ON student.ref_class_id = class.id  
 WHERE   
  student.status IN(1,2) AND class.grade_year IN(  {0}  )  
  AND COALESCE(student.gdc_code,class.gdc_code)  IS NOT NULL  
 ), history_data AS(
  SELECT   
 target_student.id AS student_id  
 , ('0'||array_to_string(xpath('//History/@GradeYear', history_xml), '')::text)::integer as his_gradeYear
 , ('0'||array_to_string(xpath('//History/@SchoolYear', history_xml), '')::text)::integer as school_year
 , ('0'||array_to_string(xpath('//History/@Semester', history_xml), '')::text)::integer as semester
         FROM (
                SELECT student.id
                    , unnest(xpath('//root/History', xmlparse(content '<root>'||sems_history||'</root>'))) as history_xml
                FROM student
                INNER JOIN class  ON student.ref_class_id = class.id  
 				WHERE   
       				student.status IN(1,2) AND class.grade_year IN(  {0} )  
            ) as target_student
WHERE
 		('0'||array_to_string(xpath('//History/@SchoolYear', history_xml), '')::text)::integer = {1}
         AND ('0'||array_to_string(xpath('//History/@Semester', history_xml), '')::text)::integer = {2} 
 ) SELECT 
 student_data.student_id
 , student_data.student_number AS 學號
 , student_data.class_name AS 班級
 , seat_no AS 座號
 , student_name AS 姓名
 --, school_year AS 學年度
-- , semester AS 學期
--  , his_gradeYear AS 當時年級
 FROM student_data
 LEFT JOIN history_data ON history_data.student_id=student_data.student_id
 WHERE his_gradeyear IS NULL
   ORDER BY student_data.grade_year DESC,student_data.display_order,class_name,seat_no
 
 ";
                query = string.Format(query, strGrYear, SchoolYear, Semester);
                DataTable dt = qh.Select(query);
                value = dt;
                //foreach (DataRow dr in dt.Rows)
                //    value.Add(dr);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        /// <summary>
        /// 取得學生學期成績
        /// </summary>
        /// <param name="GradeYear"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public List<rptStudSemsScoreCodeChkInfo> GetStudentSemsScoreInfo(string GradeYear, int SchoolYear, int Semester)
        {
            List<rptStudSemsScoreCodeChkInfo> value = new List<rptStudSemsScoreCodeChkInfo>();
            try
            {
                // 取得課程大表資料
                Dictionary<string, List<MOECourseCodeInfo>> MOECourseDict = GetCourseGroupCodeDict();
                List<string> errItem = new List<string>();

                // 取得學分對照表
                Dictionary<string, string> mappingTable = Utility.GetCreditMappingTable();

                QueryHelper qh = new QueryHelper();
                string query = "" +
                    " WITH student_data AS (  " +
"  	SELECT  " +
"  	student.id AS student_id  " +
"  	,student_number  " +
"  	,class_name  " +
"  	,student.seat_no  " +
"  	,student.name AS student_name  " +
"  	, COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code  " +
"  FROM student   " +
"  INNER JOIN class ON student.ref_class_id = class.id   " +
"  	WHERE student.status IN(1,2) AND class.grade_year  IN( " + GradeYear + " ) " +
"  ),sems_score_data AS(  " +
"  SELECT  " +
"  	sems_subj_score_ext.ref_student_id  " +
"  	, sems_subj_score_ext.grade_year  " +
"  	, sems_subj_score_ext.semester  " +
"  	, sems_subj_score_ext.school_year	  " +
"  	, array_to_string(xpath('//Subject/@科目', subj_score_ele), '')::text AS 科目  " +
"  	, array_to_string(xpath('//Subject/@原始成績', subj_score_ele), '')::text AS 原始成績  " +
"  	, array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '')::text AS 科目級別  " +
"  	, array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '')::text AS 學分數	  " +
"  	, array_to_string(xpath('//Subject/@修課必選修', subj_score_ele), '')::text AS 必選修  " +
"  	, array_to_string(xpath('//Subject/@修課校部訂', subj_score_ele), '')::text AS 校部訂	  " +
"  FROM (  " +
"  		SELECT   " +
"  			sems_subj_score.*  " +
"  			, 	unnest(xpath('//SemesterSubjectScoreInfo/Subject', xmlparse(content score_info))) as subj_score_ele  " +
"  		FROM   " +
"  			sems_subj_score   " +
"  			INNER JOIN student_data ON sems_subj_score.ref_student_id = student_data.student_id  " +
"  	) as sems_subj_score_ext   " +
"WHERE school_year = " + SchoolYear + " AND semester = " + Semester +
"  )  " +
"  SELECT   " +
"  student_id  " +
"  ,student_number  " +
"  ,class_name  " +
"  ,seat_no  " +
"  ,student_name  " +
"  ,gdc_code  " +
"  ,school_year  " +
"  ,semester  " +
"  ,grade_year  " +
"  ,科目 AS subject " +
"  ,原始成績 AS score" +
"  ,科目級別 AS subj_level " +
"  ,學分數 AS credit " +
"  ,必選修 AS required " +
"  ,(CASE 校部訂 WHEN '部訂' THEN '部定' ELSE 校部訂 END) AS required_by  " +
"   FROM student_data INNER JOIN sems_score_data ON student_data.student_id = sems_score_data.ref_student_id " +
"    ORDER BY class_name,seat_no,school_year,semester, 科目 ";


                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    errItem.Clear();
                    errItem.Add("科目名稱");
                    errItem.Add("部定校訂");
                    errItem.Add("必修選修");
                    errItem.Add("學分數");

                    rptStudSemsScoreCodeChkInfo data = new rptStudSemsScoreCodeChkInfo();
                    data.StudentID = dr["student_id"] + "";
                    data.StudentName = dr["student_name"] + "";
                    data.StudentNumber = dr["student_number"] + "";
                    data.ClassName = dr["class_name"] + "";
                    data.SeatNo = dr["seat_no"] + "";
                    data.SubjectName = dr["subject"] + "";
                    data.SubjectLevel = dr["subj_level"] + "";
                    data.SchoolYear = dr["school_year"] + "";
                    data.Semester = dr["semester"] + "";
                    data.RequiredBy = dr["required_by"] + "";
                    data.IsRequired = dr["required"] + "";
                    data.GradeYear = dr["grade_year"] + "";
                    data.Credit = dr["credit"] + "";

                    decimal score;
                    if (dr["score"].ToString() != "")
                        if (decimal.TryParse(dr["score"].ToString(), out score))
                        {
                            data.Score = score;
                        }

                    if (dr["gdc_code"] != null)
                    {
                        data.gdc_code = dr["gdc_code"] + "";
                    }
                    else
                    {
                        data.gdc_code = "";
                    }


                    // 比對大表資料
                    if (MOECourseDict.ContainsKey(data.gdc_code))
                    {
                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.IsRequired == Mco.is_required && data.RequiredBy == Mco.require_by)
                            {
                                data.entry_year = Mco.entry_year;
                                data.credit_period = Mco.credit_period;
                                data.CourseCode = Mco.course_code;
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.IsRequired == Mco.is_required)
                            {
                                errItem.Remove("科目名稱");
                                errItem.Remove("必修選修");
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.RequiredBy == Mco.require_by)
                            {
                                errItem.Remove("科目名稱");
                                errItem.Remove("部定校訂");
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name)
                            {
                                errItem.Remove("科目名稱");
                                break;
                            }
                        }

                        // 檢查學分數
                        if (data.CheckCreditPass(mappingTable))
                        {
                            errItem.Remove("學分數");
                        }


                        if (errItem.Count > 0)
                        {
                            foreach (string err in errItem)
                                data.ErrorMsgList.Add(err);
                        }
                    }
                    else
                    {
                        data.ErrorMsgList.Add("群科班代碼無法對照");
                    }

                    value.Add(data);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        /// <summary>
        /// 取得和學期對照表對應的學期成績
        /// </summary>
        /// <param name="GradeYear"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public List<rptStudSemsScoreCodeChkInfo> GetStudentSemsScoreInfoByExistingSemsHistory(string GradeYear, int SchoolYear, int Semester)
        {
            List<rptStudSemsScoreCodeChkInfo> value = new List<rptStudSemsScoreCodeChkInfo>();
            try
            {
                // 取得課程大表資料
                Dictionary<string, List<MOECourseCodeInfo>> MOECourseDict = GetCourseGroupCodeDict();
                List<string> errItem = new List<string>();

                // 取得學分對照表
                Dictionary<string, string> mappingTable = Utility.GetCreditMappingTable();

                QueryHelper qh = new QueryHelper();
                string query = @"
WITH student_data AS (   
  	SELECT   
  	student.id AS student_id   
  	,student_number
	, class.grade_year
  	,class_name
	,class.display_order
  	,student.seat_no   
  	,student.name AS student_name   
	, unnest(xpath('//root/History', xmlparse(content '<root>'||sems_history||'</root>'))) as history_xml
  	, COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code   
  FROM student    
  INNER JOIN class ON student.ref_class_id = class.id    
  	WHERE student.status IN(1,2) AND class.grade_year  IN(   {0}   ) 
  ) , ta_student_data AS(
 SELECT student_data.student_id
            , student_data.student_name
            , student_data.student_number
			, student_data.grade_year
            , class_name
			, display_order
            , seat_no
			, gdc_code
            , ('0'||array_to_string(xpath('//History/@SchoolYear', history_xml), '')::text)::integer as his_school_year
            , ('0'||array_to_string(xpath('//History/@Semester', history_xml), '')::text)::integer as his_semester
            , ('0'||array_to_string(xpath('//History/@GradeYear', history_xml), '')::text)::integer as his_gradeyear
            FROM student_data
            WHERE  ('0'||array_to_string(xpath('//History/@SchoolYear', history_xml), '')::text)::integer = {1}
            AND ('0'||array_to_string(xpath('//History/@Semester', history_xml), '')::text)::integer = {2}
 ) , sems_score_data AS(   
  SELECT   
  	sems_subj_score_ext.ref_student_id   
  	, sems_subj_score_ext.grade_year   AS sem_grade_year
  	, sems_subj_score_ext.semester   
  	, sems_subj_score_ext.school_year	   
  	, array_to_string(xpath('//Subject/@科目', subj_score_ele), '')::text AS 科目   
  	, array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '')::text AS 科目級別   
  	, array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '')::text AS 學分數	   
  	, array_to_string(xpath('//Subject/@修課必選修', subj_score_ele), '')::text AS 必選修   
  	, array_to_string(xpath('//Subject/@修課校部訂', subj_score_ele), '')::text AS 校部訂	   
  FROM (   
  		SELECT    
  			sems_subj_score.*   
  			, 	unnest(xpath('//SemesterSubjectScoreInfo/Subject', xmlparse(content score_info))) as subj_score_ele   
  		FROM    
  			sems_subj_score    
  			INNER JOIN ta_student_data ON sems_subj_score.ref_student_id = ta_student_data.student_id   
  	) as sems_subj_score_ext    
  )   
  SELECT    
  student_id   
  ,student_number   
  ,class_name
  ,grade_year
  ,seat_no   
  ,student_name   
  ,gdc_code   
  ,school_year   
  ,semester   
  ,sem_grade_year
  ,his_school_year
  ,his_semester
  ,his_gradeyear
  ,科目 AS subject  
  ,科目級別 AS subj_level  
  ,學分數 AS credit  
  ,必選修 AS required  
  ,(CASE 校部訂 WHEN '部訂' THEN '部定' ELSE 校部訂 END) AS required_by   
   FROM ta_student_data 
   INNER JOIN sems_score_data 
   		ON ta_student_data.student_id = sems_score_data.ref_student_id
		AND ta_student_data.his_school_year=sems_score_data.school_year
		AND ta_student_data.his_semester=sems_score_data.semester
		AND ta_student_data.his_gradeyear=sems_score_data.sem_grade_year
    ORDER BY grade_year DESC ,display_order, seat_no, school_year, semester
";

                query = string.Format(query, GradeYear, SchoolYear, Semester);
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    errItem.Clear();
                    errItem.Add("科目名稱");
                    errItem.Add("部定校訂");
                    errItem.Add("必修選修");
                    errItem.Add("學分數");

                    rptStudSemsScoreCodeChkInfo data = new rptStudSemsScoreCodeChkInfo();
                    data.StudentID = dr["student_id"] + "";
                    data.StudentName = dr["student_name"] + "";
                    data.StudentNumber = dr["student_number"] + "";
                    data.ClassName = dr["class_name"] + "";
                    data.SeatNo = dr["seat_no"] + "";
                    data.SubjectName = dr["subject"] + "";
                    data.SubjectLevel = dr["subj_level"] + "";
                    data.SchoolYear = dr["school_year"] + "";
                    data.Semester = dr["semester"] + "";
                    data.RequiredBy = dr["required_by"] + "";
                    data.IsRequired = dr["required"] + "";

                    /// 年級 
                    data.GradeYear = dr["grade_year"] + "";

                    /// 成績年級 
                    data.SemsGradeYear = dr["sem_grade_year"] + "";
                    data.Credit = dr["credit"] + "";
                    if (dr["gdc_code"] != null)
                    {
                        data.gdc_code = dr["gdc_code"] + "";
                    }
                    else
                    {
                        data.gdc_code = "";
                    }


                    // 比對大表資料
                    if (MOECourseDict.ContainsKey(data.gdc_code))
                    {
                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.IsRequired == Mco.is_required && data.RequiredBy == Mco.require_by)
                            {

                                #region 2022-03-07 Cynthia 增加年級+學期條件比對大表中的open_type，取得課程代碼等資訊

                                int idx = -1;

                                if (data.GradeYear == "1" && data.Semester == "1")
                                    idx = 0;

                                if (data.GradeYear == "1" && data.Semester == "2")
                                    idx = 1;

                                if (data.GradeYear == "2" && data.Semester == "1")
                                    idx = 2;

                                if (data.GradeYear == "2" && data.Semester == "2")
                                    idx = 3;

                                if (data.GradeYear == "3" && data.Semester == "1")
                                    idx = 4;

                                if (data.GradeYear == "3" && data.Semester == "2")
                                    idx = 5;

                                char[] cp = Mco.open_type.ToArray();
                                if (idx != -1 && idx < cp.Length)
                                {
                                    if (cp[idx] != '-')
                                    {
                                        data.entry_year = Mco.entry_year;
                                        data.credit_period = Mco.credit_period;
                                        data.CourseCode = Mco.course_code;
                                    }
                                }

                                #endregion
                                //break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.IsRequired == Mco.is_required)
                            {
                                errItem.Remove("科目名稱");
                                errItem.Remove("必修選修");
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name && data.RequiredBy == Mco.require_by)
                            {
                                errItem.Remove("科目名稱");
                                errItem.Remove("部定校訂");
                                break;
                            }
                        }

                        foreach (MOECourseCodeInfo Mco in MOECourseDict[data.gdc_code])
                        {
                            if (data.SubjectName == Mco.subject_name)
                            {
                                errItem.Remove("科目名稱");
                                break;
                            }
                        }

                        // 檢查學分數
                        if (data.CheckCreditPass(mappingTable))
                        {
                            errItem.Remove("學分數");
                        }


                        if (errItem.Count > 0)
                        {
                            foreach (string err in errItem)
                                data.ErrorMsgList.Add(err);
                        }
                    }
                    else
                    {
                        data.ErrorMsgList.Add("群科班代碼無法對照");
                    }

                    value.Add(data);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        public List<DataRow> GetHasGDCCodeAndSemHistoryStudent(string strGrYear, int SchoolYear, int Semester)
        {
            // 可以找到沒有學期歷程的學生，但用WHERE濾掉了
            List<DataRow> value = new List<DataRow>();
            try
            {
                QueryHelper qh = new QueryHelper();
                string query = @" WITH student_data AS(
 SELECT   
 student.id AS student_id  
 , student.name AS student_name  
 , student_number  
 , student.seat_no  
 , class_name  
 , class.grade_year AS grade_year  
 , COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code  
 , class.display_order
 FROM student   
 	INNER JOIN class  
 	ON student.ref_class_id = class.id  
 WHERE   
  student.status IN(1,2) AND class.grade_year IN(  {0}  )  
  AND COALESCE(student.gdc_code,class.gdc_code)  IS NOT NULL  
 ), history_data AS(
  SELECT   
 target_student.id AS student_id  
 , ('0'||array_to_string(xpath('//History/@GradeYear', history_xml), '')::text)::integer as his_grade_year
 , ('0'||array_to_string(xpath('//History/@SchoolYear', history_xml), '')::text)::integer as school_year
 , ('0'||array_to_string(xpath('//History/@Semester', history_xml), '')::text)::integer as semester
         FROM (
                SELECT student.id
                    , unnest(xpath('//root/History', xmlparse(content '<root>'||sems_history||'</root>'))) as history_xml
                FROM student
                INNER JOIN class  ON student.ref_class_id = class.id  
 				WHERE   
       				student.status IN(1,2) AND class.grade_year IN(  {0} )  
            ) as target_student
WHERE
 		('0'||array_to_string(xpath('//History/@SchoolYear', history_xml), '')::text)::integer = {1}
         AND ('0'||array_to_string(xpath('//History/@Semester', history_xml), '')::text)::integer = {2} 
 ) SELECT 
 student_data.* 
 , his_grade_year
 , school_year
 , semester
 FROM student_data
 LEFT JOIN history_data ON history_data.student_id=student_data.student_id
 WHERE his_gradeYear IS NOT NULL
   ORDER BY student_data.grade_year DESC,student_data.display_order,class_name,seat_no
 
 ";
                query = string.Format(query, strGrYear, SchoolYear, Semester);
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                    value.Add(dr);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        /// <summary>
        /// 取得學生基本資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<StudentInfo> GetStudentInfoListByIDs(List<string> StudentIDList, int schoolYear, int semester)
        {
            List<StudentInfo> StudentInfoList = new List<StudentInfo>();

            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string qry = @"
WITH student_data AS(
SELECT 
        student.id AS student_id
        , class.class_name
        , class.grade_year
        , student.seat_no
        , student.name AS student_name
        , id_number
		, COALESCE(student.ref_dept_id,class.ref_dept_id)  AS dept_id
		, COALESCE(student.gdc_code,class.gdc_code)  AS gdc_code
		, array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''學業(原始)'']/@成績', xmlparse(content score_info)), '')::text AS entry_score
        , display_order
FROM student 
LEFT JOIN class ON student.ref_class_id = class.id
LEFT JOIN sems_entry_score ON sems_entry_score.ref_student_id = student.id AND school_year=" + schoolYear + " AND semester = " + semester +
@"WHERE student.status IN (1, 2) 
AND student.id IN(" + string.Join(",", StudentIDList.ToArray()) + @")
ORDER BY class.grade_year, class.display_order, class.class_name, seat_no
)SELECT 
	student_id
	, class_name
	, seat_no
	, id_number
	, student_name
	, name AS dept_name
	, gdc_code
	, entry_score
	FROM student_data 
	LEFT JOIN dept ON student_data.dept_id = dept.id 
ORDER BY grade_year, display_order, class_name, seat_no, student_name
";
                try
                {
                    DataTable dt = qh.Select(qry);
                    if (dt != null)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            StudentInfo si = new StudentInfo();
                            si.StudentID = dr["student_id"].ToString();
                            si.ClassName = dr["class_name"].ToString();
                            si.SeatNo = dr["seat_no"].ToString();
                            si.Dept = dr["dept_name"].ToString();
                            si.IDNumber = dr["id_number"].ToString();
                            si.Name = dr["student_name"].ToString();
                            si.EntryScore= dr["entry_score"].ToString();

                            StudentInfoList.Add(si);
                        }
                    }

                }
                catch (Exception ex)
                {

                }
            }

            return StudentInfoList;
        }
    }
}
