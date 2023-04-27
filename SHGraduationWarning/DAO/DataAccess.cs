using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;

namespace SHGraduationWarning.DAO
{
    public class DataAccess
    {
        // 取得班級科別資訊
        public static List<ClassDeptInfo> GetClassDeptList()
        {
            List<ClassDeptInfo> value = new List<ClassDeptInfo>();

            try
            {
                string strSQL = @"
            SELECT
                DISTINCT dept.id AS dept_id,
                dept.name AS dept_name,
                class_name,
                class.id AS class_id
            FROM
                class
                LEFT JOIN dept ON class.ref_dept_id = dept.id
                INNER JOIN student ON class.id = student.ref_class_id
            WHERE
                student.status IN(1, 2)
            ORDER BY
                dept.name,
                class_name
            ";

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    ClassDeptInfo cd = new ClassDeptInfo();
                    cd.DeptID = dr["dept_id"] + "";
                    cd.DeptName = dr["dept_name"] + "";
                    cd.ClassID = dr["class_id"] + "";
                    cd.ClassName = dr["class_name"] + "";
                    value.Add(cd);
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
