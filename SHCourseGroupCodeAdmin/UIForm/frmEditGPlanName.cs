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
    public partial class frmEditGPlanName : BaseForm
    {
        DataAccess _da;

        private string _GPNewPName = "";

        private GPlanInfo108 _GPlanInfo108;


        public frmEditGPlanName()
        {
            InitializeComponent();
            _da = new DataAccess();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            btnCreate.Enabled = false;

            // 檢查資料是否重複
            _GPNewPName = cbxSchoolYear.Text.Trim() + txtName.Text.Trim();
            Dictionary<string, string> chkNameDict = _da.GetAllGPNameDict();

            if (!chkNameDict.ContainsKey(_GPNewPName))
            {
                string _GPID = _da.UpdateGPlanByName(_GPNewPName, _GPlanInfo108);
                if (_GPID != "")
                {
                    MessageBox.Show("更新完成");
                    this.DialogResult = DialogResult.OK;
                }
                else
                {
                    MessageBox.Show("更新過程發生錯誤。");
                    btnCreate.Enabled = true;
                }

            }
            else
            {
                MessageBox.Show("名稱重複，無法更新。");
                btnCreate.Enabled = true;
            }
        }

        private void frmEditGPlanName_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.Size;
            lblOldName.Text = _GPlanInfo108.RefGPName;

            // 設定學年度選項
            cbxSchoolYear.ImeMode = ImeMode.Off;
            cbxSchoolYear.MaxLength = 3;

            cbxSchoolYear.Items.Add("");
            int sc;
            if (int.TryParse(K12.Data.School.DefaultSchoolYear, out sc))
            {
                for (int i = sc - 1; i <= sc + 1; i++)
                {
                    cbxSchoolYear.Items.Add(i);
                }
            }

        }

        public void SetGPlanInfo108(GPlanInfo108 info)
        {
            _GPlanInfo108 = info;
        }

        public string GetNewGPName()
        {
            return _GPNewPName;
        }
    }
}
