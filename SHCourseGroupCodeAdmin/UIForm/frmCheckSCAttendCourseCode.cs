using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SHCourseGroupCodeAdmin.DAO;
using FISCA.Presentation.Controls;
using System.Xml.Linq;
using Aspose.Cells;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCheckSCAttendCourseCode : BaseForm
    {
        public frmCheckSCAttendCourseCode()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
