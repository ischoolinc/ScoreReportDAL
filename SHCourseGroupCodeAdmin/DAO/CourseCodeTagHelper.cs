using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using K12.Data;
using FISCA.Data;
using System.Data;

namespace SHCourseGroupCodeAdmin.DAO
{
    /// <summary>
    /// 處理國教署規範標Tag
    /// </summary>
    public class CourseCodeTagHelper
    {        
        private string CourseCodeTagID = "";

        /// <summary>
        /// 檢查課程代碼類別是否存在
        /// </summary>
        /// <returns></returns>
        public string GetCourseCodeTagID()
        {
            string value = "";
            // SQL
            string strSQL = "SELECT id FROM tag WHERE name ='課程' AND prefix='課程計畫';";
            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(strSQL);
            if (dt != null && dt.Rows.Count > 0)
            {
                value = dt.Rows[0][0].ToString();
            }

            CourseCodeTagID = value;
            return value;
        }

        public CourseCodeTagHelper()
        {
            // 檢查課程Tag是否存在，如果不存在加入
            if (GetCourseCodeTagID() == "")
            {
                // 使用 K12 Tag 新增
                TagConfigRecord rec = new TagConfigRecord();
                rec.Name = "課程";
                rec.Prefix = "課程計畫";
                CourseCodeTagID = TagConfig.Insert(rec);
            }
        }

        /// <summary>
        /// 新增課程的標籤
        /// </summary>
        /// <param name="CourseIDList"></param>
        public void AddCourseCodeTag(List<string> CourseIDList)
        {
            if (CourseCodeTagID != "")
            {
                List<CourseTagRecord> recList = new List<CourseTagRecord>();
                foreach (string id in CourseIDList)
                {
                    CourseTagRecord rec = new CourseTagRecord();
                    rec.RefCourseID = id;
                    rec.RefTagID = CourseCodeTagID;
                }

                // 新增課程標籤
                if (recList.Count > 0)
                {
                    CourseTag.Insert(recList);
                }
            }
        }
    }
}
