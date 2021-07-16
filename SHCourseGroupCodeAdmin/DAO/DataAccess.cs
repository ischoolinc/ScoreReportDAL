﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using FISCA.Data;
using System.Xml.Linq;
using System.IO;

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
"   $moe.subjectcode ORDER BY group_code,subject_name; ";
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

            QueryHelper qh = new QueryHelper();
            string query = "" +
" SELECT DISTINCT group_code,(entry_year|| " +
" (CASE WHEN length(course_type)>0 THEN '_'||course_type END) || " +
" (CASE WHEN length(group_type)>0 THEN '_'||group_type END) || " +
" (CASE WHEN length(subject_type)>0 THEN '_'||subject_type END) || " +
" (CASE WHEN length(class_type)>0 THEN '_'||class_type END) " +
" ) AS group_name" +
" ,(entry_year||(CASE WHEN length(subject_type)>0 THEN ''||subject_type END) || " +
" (CASE WHEN length(class_type) > 0 THEN '' || class_type END)) AS gplan_name " +
" FROM $moe.subjectcode ORDER BY group_name; ";

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
                                break;
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
    }
}
