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

namespace SHGraduationWarning.UIForm
{
    public partial class frmNoGraduationPlanStudent : BaseForm
    {
        List<string> NoGraduationPlanStudentList;

        public frmNoGraduationPlanStudent()
        {
            InitializeComponent();
            NoGraduationPlanStudentList = new List<string>();
        }

        private void frmNoGraduationPlanStudent_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            if (NoGraduationPlanStudentList.Count > 0)
            {
                lblMsg.Text = "有" + NoGraduationPlanStudentList.Count + "位學生沒有設定課程規劃，請設定後再進行查詢作業。";

                btnAddStudentTemp.Enabled = true;
            }
            else
            {
                btnAddStudentTemp.Enabled = false;
            }
        }

        public void SetNoGraduationPlanStudentList(List<string> StudentIDList)
        {
            NoGraduationPlanStudentList = StudentIDList;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAddStudentTemp_Click(object sender, EventArgs e)
        {
            try
            {
                // 加入待處理學生
                if (NoGraduationPlanStudentList.Count > 0)
                {
                    List<string> AddIdList = new List<string>();

                    foreach (string sid in NoGraduationPlanStudentList)
                    {
                        if (!K12.Presentation.NLDPanels.Student.TempSource.Contains(sid))
                        {
                            AddIdList.Add(sid);
                        }
                    }

                    if (AddIdList.Count > 0)
                    {
                        K12.Presentation.NLDPanels.Student.AddToTemp(AddIdList);
                        MsgBox.Show("已加入" + AddIdList.Count + "位學生至學生待處理");
                    }
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show(ex.Message);
            }

        }
    }
}
