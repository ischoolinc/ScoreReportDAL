using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SHCourseGroupCodeSetup.DAO;

namespace SHCourseGroupCodeSetup.UIForm
{
    public partial class frmBatchCourseClassGroupCode : FISCA.Presentation.Controls.BaseForm
    {
        List<string> classIDList = new List<string>();
        List<string> GroupNameList = new List<string>();
        DataAccess da = new DataAccess();

        public frmBatchCourseClassGroupCode(List<string> ids)
        {
            InitializeComponent();
            classIDList = ids;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            da.SetClassGroupCodeByClassIDs(classIDList, da.GetGroupCodeByName(cbxCourseGroupCode.Text));
            MessageBox.Show("產生完成");
            this.Close();
        }

        private void frmBatchCourseClassGroupCode_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            cbxCourseGroupCode.Enabled = btnRun.Enabled = false;

            da.LoadMOEGroupCodeDict();
            GroupNameList = da.GetGroupNameList();

            cbxCourseGroupCode.Text = "";
            cbxCourseGroupCode.Items.Clear();
            cbxCourseGroupCode.Items.Add("");
            foreach (string name in GroupNameList)
                cbxCourseGroupCode.Items.Add(name);

            cbxCourseGroupCode.Enabled = btnRun.Enabled = true;
        }
    }
}
