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
    public partial class frmCreateClassGPlan : BaseForm
    {
        DataAccess _da = new DataAccess();
        string _SelectGroupName = "";

        public frmCreateClassGPlan()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            // 取得畫面上選擇群科班名稱
            _SelectGroupName = cbxGroupName.Text;

            string gpName = _da.GetGPlanNameByGroupName(_SelectGroupName);

            bool hasGPName = _da.CheckHasGPlan(gpName);
            if (hasGPName)
            {
                MsgBox.Show(gpName + " 班級課程規劃表內已存在");
            }
            else
            {
                MsgBox.Show("Pass");
            }


        }

        private void frmCreateClassGPlan_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            cbxGroupName.Enabled = false;
            btnCreate.Enabled = false;

            // 取得課程代碼大表資料並填入群科班選項
            _da.LoadMOEGroupCodeDict();

            foreach (string name in _da.GetGroupNameList())
            {
                cbxGroupName.Items.Add(name);
            }

            if (cbxGroupName.Items.Count > 0)
            {
                cbxGroupName.SelectedIndex = 0;
                btnCreate.Enabled = true;
            }

            cbxGroupName.Enabled = true;
         
        }
    }
}
