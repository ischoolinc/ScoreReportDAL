using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using SHCourseGroupCodeAdmin.DAO;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateCourseByGPlan108_C_Create : BaseForm
    {
        string _SchoolYear = "", _Semester = "";
        Dictionary<string, SubjectCourseInfo> _SubjectCourseInfoDict;

        public frmCreateCourseByGPlan108_C_Create()
        {
            InitializeComponent();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void SetSubjectCourseInfoDict(Dictionary<string, SubjectCourseInfo> data)
        {
            _SubjectCourseInfoDict = data;
        }

        public void SetSchoolYearSemester(string SchoolYear, string Semester)
        {
            _SchoolYear = SchoolYear;
            _Semester = Semester;
        }

    }
}
