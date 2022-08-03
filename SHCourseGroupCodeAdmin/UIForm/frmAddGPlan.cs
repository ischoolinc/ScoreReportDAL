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
    public partial class frmAddGPlan : BaseForm
    {
        DataAccess _da;
        private string _GPName = "";
        private string _GPID = "";
        public frmAddGPlan()
        {
            _da = new DataAccess();
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            btnCreate.Enabled = false;

            // 檢查學年度是否數字
            string strSchoolYear = cbxSchoolYear.Text.Trim();
            if(!string.IsNullOrWhiteSpace(strSchoolYear))
            {
                int sc;
                if(int.TryParse(strSchoolYear,out sc) == false)
                {
                    MsgBox.Show("學年度 請輸入數字");
                    btnCreate.Enabled = true;
                    return;
                }
            }

            // 檢查資料是否重複
            _GPName = strSchoolYear + txtName.Text.Trim();
            Dictionary<string, string> chkNameDict = _da.GetAllGPNameDict();

            if (!chkNameDict.ContainsKey(_GPName))
            {
                _GPID = _da.AddGPlanByName(_GPName);
                if (_GPID != "")
                {
                    MessageBox.Show("新增完成");
                    this.DialogResult = DialogResult.OK;
                }
                else
                {
                    MessageBox.Show("新增過程發生錯誤。");
                    btnCreate.Enabled = true;
                }

            }
            else
            {
                MessageBox.Show("名稱重複，無法新增。");
                btnCreate.Enabled = true;
            }
        }

        public string GetGPName()
        {
            return _GPName;
        }

        public string GetGPID()
        {
            return _GPID;
        }

        private void frmAddGPlan_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.Size;

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
    }
}
