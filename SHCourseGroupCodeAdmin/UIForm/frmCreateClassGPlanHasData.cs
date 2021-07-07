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
    public partial class frmCreateClassGPlanHasData : BaseForm
    {
        string SelectGroupName = "";

        public frmCreateClassGPlanHasData()
        {
            InitializeComponent();
        }

        private void frmCreateClassGPlanHasData_Load(object sender, EventArgs e)
        {
            this.lblGroupName.Text = SelectGroupName;
        }

        public void SetGroupName(string name)
        {
            SelectGroupName = name;
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {

        }
    }
}
