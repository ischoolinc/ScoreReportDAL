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
    public partial class frmCreateGPlanItemSetup108 : BaseForm
    {
        GPlanInfo108 _GPlanInfo;
        public frmCreateGPlanItemSetup108()
        {
            InitializeComponent();
        }

        public void SetGPlanInfo(GPlanInfo108 data)
        {
            _GPlanInfo = data;
        }

        public GPlanInfo108 GetGPlanInfo()
        {
            return _GPlanInfo;
        }

        private void frmCreateGPlanItemSetup108_Load(object sender, EventArgs e)
        {
            lblGroupName.Text = _GPlanInfo.GDCName;

        }
    }
}
