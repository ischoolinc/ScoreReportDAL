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
    public partial class frmBatchCourseStudentGroupCode : FISCA.Presentation.Controls.BaseForm
    {
        List<string> studentIDList = new List<string>();
        List<string> GroupNameList = new List<string>();
        DataAccess da = new DataAccess();

        public frmBatchCourseStudentGroupCode(List<string> ids)
        {
            InitializeComponent();
            studentIDList = ids;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            da.SetStudentGroupCodeByStudentIDs(studentIDList, da.GetGroupCodeByName(cbxCourseGroupCode.Text));
            MessageBox.Show("產生完成");
            this.Close();
        }

        private void cbxCourseGroupCode_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void labelX1_Click(object sender, EventArgs e)
        {

        }

        private void frmBatchCourseStudentGroupCode_Load(object sender, EventArgs e)
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
