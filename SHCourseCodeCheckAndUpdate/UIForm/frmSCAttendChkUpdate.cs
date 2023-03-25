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
using SHCourseCodeCheckAndUpdate.DAO;

namespace SHCourseCodeCheckAndUpdate.UIForm
{
    public partial class frmSCAttendChkUpdate : BaseForm
    {
        public frmSCAttendChkUpdate()
        {
            InitializeComponent();
        }

        private void frmSCAttendChkUpdate_Load(object sender, EventArgs e)
        {

        }

        private void btnExit_Click(object sender, EventArgs e)
        {                                          
            this.Close();
        }
                            
        private void btnQuery_Click(object sender, EventArgs e)
        {
                List<StudSCAttendInfo> sc = DataAccess.GetStudSCAttendBySchoolYearSems("111", "1", "1");
        }
    }
}
