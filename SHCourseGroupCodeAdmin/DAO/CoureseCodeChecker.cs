using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.IO;
using System.Net;
using FISCA.Authentication;
using FISCA.Data;
using System.Data;

namespace SHCourseGroupCodeAdmin.DAO
{
    /// <summary>
    /// 課程計畫平台課程代碼檢查
    /// </summary>
    public class CoureseCodeChecker
    {
        string SchoolCode = "";
        string DSNS = "";
        Dictionary<int, string> jsonSourceDict;
        Dictionary<int, List<CourseCodeRoot>> CourseCodeRootDict;
        // 系統內 moe 資料
        Dictionary<int, Dictionary<string, string>> MOECourseCodeDict;


        public CoureseCodeChecker(string schoolCode, string dsns)
        {
            jsonSourceDict = new Dictionary<int, string>();
            CourseCodeRootDict = new Dictionary<int, List<CourseCodeRoot>>();
            MOECourseCodeDict = new Dictionary<int, Dictionary<string, string>>();
            SchoolCode = schoolCode;
            DSNS = dsns;
        }

        /// <summary>
        /// 傳入開始與結束實施學年度取得資料
        /// </summary>
        /// <param name="BeginYear"></param>
        /// <param name="EndYear"></param>
        public void LoadCouseCodeSourceData(int BeginYear, int EndYear)
        {
            jsonSourceDict.Clear();
            CourseCodeRootDict.Clear();

            try
            {
                for (int sc = BeginYear; sc <= EndYear; sc++)
                {
                    // 呼叫取得資料                
                    string jsonString = CallMOECourseSourceBySchoolYear(sc);

                    // 解析 json 資料
                    var courseDataList = new JavaScriptSerializer().Deserialize<List<CourseCodeRoot>>(jsonString);

                    // 放入暫存
                    if (!jsonSourceDict.ContainsKey(sc))
                        jsonSourceDict.Add(sc, jsonString);

                    if (!CourseCodeRootDict.ContainsKey(sc))
                        CourseCodeRootDict.Add(sc, courseDataList);
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
        }

        public List<string> CheckSystemMOEData(int BeginYear, int EndYear)
        {
            List<string> errorList = new List<string>();

            // 取得課程計畫平台資料
            LoadCouseCodeSourceData(BeginYear, EndYear);

            // 取得系統內 MOE 資料
            LoadSysMOECourseDict();

            // 比對資料
            for (int sy = BeginYear; sy <= EndYear; sy++)
            {
                if (CourseCodeRootDict.ContainsKey(sy))
                {
                    if (MOECourseCodeDict.ContainsKey(sy))
                    {
                        foreach (CourseCodeRoot ccr in CourseCodeRootDict[sy])
                        {
                            foreach (CourseCodeInfo cci in ccr.課程資料)
                            {
                                // key 課程代碼 科目名稱 課程屬性 授課學期學分節數 授課學期開課方式
                                string key = cci.課程代碼 + "_" + cci.科目名稱 + "_" + cci.課程屬性 + "_" + cci.授課學期學分節數 + "_" + cci.授課學期開課方式;
                                if (!MOECourseCodeDict[sy].ContainsKey(key))
                                {
                                    errorList.Add("系統內缺少:" + key);
                                }
                            }
                        }

                    }
                    else
                    {
                        errorList.Add("系統內沒有實施學年度" + sy + "資料");
                    }
                }
            }



            return errorList;
        }


        /// <summary>
        /// 取得課程計畫平台原始回傳 JSON檔案
        /// </summary>
        /// <returns></returns>
        public Dictionary<int, string> GetJSONStringDict()
        {
            return jsonSourceDict;
        }

        /// <summary>
        /// 取得課程計畫平台解析後物件
        /// </summary>
        /// <returns></returns>
        public Dictionary<int, List<CourseCodeRoot>> GetCourseCodeRootDict()
        {
            return CourseCodeRootDict;
        }




        /// <summary>
        /// 取得課程計畫平台原始資料
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <returns></returns>
        private string CallMOECourseSourceBySchoolYear(int SchoolYear)
        {
            string value = "";

            try
            {
               //  SchoolCode = "193316";

                String targetUrl = @"https://courseid.cloud.ncnu.edu.tw/api/GetAllCourses/" + SchoolCode + "/" + SchoolYear;
                HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(targetUrl);
                req.Method = "GET";
                //req.ContentType = "application/json";
                req.Headers.Add("ApiKey:" + CourseCodeAPIKey.ApiKey);
                req.Headers.Add("Secret:" + CourseCodeAPIKey.Secret);

                var response = req.GetResponse();
                Stream receiveStream = response.GetResponseStream();
                StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8);
                value = readStream.ReadToEnd();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Call MOE Course Error:" + ex.Message);
            }


            return value;
        }

        /// <summary>
        /// 載入系統內 MOE Data
        /// </summary>
        private void LoadSysMOECourseDict()
        {
            MOECourseCodeDict.Clear();
            try
            {
                string sql = @"
SELECT 
    uid
    ,course_attr
    ,course_code
    ,credit_period
    ,entry_year
    ,open_type
    ,subject_name
 FROM 
    $moe.subjectcode 
 ORDER BY 
     entry_year
     ,course_code 
";
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(sql);

                foreach (DataRow dr in dt.Rows)
                {
                    int ey;

                    if (int.TryParse(dr["entry_year"] + "", out ey))
                    {
                        if (!MOECourseCodeDict.ContainsKey(ey))
                            MOECourseCodeDict.Add(ey, new Dictionary<string, string>());

                        string key = dr["course_code"] + "_" + dr["subject_name"] + "_" + dr["course_attr"] + "_" + dr["credit_period"] + "_" + dr["open_type"];


                        if (!MOECourseCodeDict[ey].ContainsKey(key))
                            MOECourseCodeDict[ey].Add(key, dr["uid"] + "");
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
