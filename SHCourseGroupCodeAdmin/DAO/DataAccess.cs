using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using FISCA.Data;

namespace SHCourseGroupCodeAdmin.DAO
{
    public class DataAccess
    {
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
"  FROM " +
"   $moe.subjectcode ORDER BY group_code,subject_name; ";
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                foreach(DataRow dr  in dt.Rows)
                {
                    MOECourseCodeInfo data = new MOECourseCodeInfo();
                    data.last_update = dr["最後更新日期"] + "";
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
