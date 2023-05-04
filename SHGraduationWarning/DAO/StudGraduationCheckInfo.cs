using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHGraduationWarning.DAO
{
    // 學生畢業預警檢查資訊
    public class StudGraduationCheckInfo
    {
        // 學生系統編號
        public string StudentID { get; set; }
        // 學號
        public string StudentNumber { get; set; }
        // 班級系統編號
        public string ClassID { get; set; }
        // 班級名稱
        public string ClassName { get; set; }
        // 科別系統編號
        public string DeptID { get; set; }
        // 科別名稱
        public string DeptName { get; set; }
        // 座號
        public string SeatNo { get; set; }
        // 姓名
        public string StudentName { get; set; }
        // 班級年級
        public string ClassYear { get; set; }

        // 學生所有成績資訊
        public List<StudSubjectInfo> StudAllSubjectInfoList { get; set; }

        // 總學分
        public decimal TotalCredit { get; set; }
        // 實得總學分
        public decimal TotalCreditGet { get; set; }
        // 已修總學分
        public decimal TotalCreditComplete { get; set; }
        // 部定必修實得學分
        public decimal CoreCourseCreditGet { get; set; }
        // 部定必修已修學分
        public decimal CoreCourseCreditComplete { get; set; }
        // 校訂必修實得學分
        public decimal SchoolCoreCourseCreditGet { get; set; }
        // 校訂必修已修學分
        public decimal SchoolCoreCourseCreditComplete { get; set; }
        // 選修科目實得學分
        public decimal SelectCourseCreditGet { get; set; }
        // 選修科目已修學分
        public decimal SelectCourseCreditComplete { get; set; }
        // 專業及實習科目實得學分
        public decimal ProfessionalCreditGet { get; set; }
        // 專業及實習科目已修學分
        public decimal ProfessionalCreditComplete { get; set; }
        // 實習科目實得學分
        public decimal PracticeCreditGet { get; set; }
        // 實習科目已修學分
        public decimal PracticeCreditComplete { get; set; }
        // 是否畢業
        public bool IsGraduate { get; set; }
        // 未達畢業標準說明
        public List<string> NotGraduateReason { get; set; }

        // 部定科目均修習
        public bool IsCoreCourseAllPass { get; set; }
        // 必修均修習
        public bool IsSchoolCoreCourseAllPass { get; set; }

        

    }
}
