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

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateGPlanMain108 : BaseForm
    {
        public frmCreateGPlanMain108()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmCreateGPlanMain108_Load(object sender, EventArgs e)
        {

        }

        private void btnCreate_Click(object sender, EventArgs e)
        {

        }

        private void btnQueryAndSet_Click(object sender, EventArgs e)
        {
            frmCreateGPlanQueryAndSetup108 fgq = new frmCreateGPlanQueryAndSetup108();
            fgq.ShowDialog();
        }
    }
}
