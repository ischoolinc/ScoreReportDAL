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
using K12.Data;
using SHCourseGroupCodeAdmin.DAO;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateCourseByGPlan108_Detail : BaseForm
    {
        string _SchoolYear = "", _Semester = "";
        List<CClassCourseInfo> _CClassCourseInfoList;

        public frmCreateCourseByGPlan108_Detail()
        {
            InitializeComponent();
        }

        public void SetCClassCourseInfo(List<CClassCourseInfo> data)
        {
            _CClassCourseInfoList = data;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void frmCreateCourseByGPlan108_Detail_Load(object sender, EventArgs e)
        {
            lblSchoolYear.Text = "開設課程學年度學期：" + _SchoolYear + "學年度 第" + _Semester + "學期";

            LoadColumns();
            LoadData();
        }

        public void SetSchoolYearSemester(string SchoolYear, string Semester)
        {
            _SchoolYear = SchoolYear;
            _Semester = Semester;
        }

        private void LoadColumns()
        {
            DataGridViewTextBoxColumn tbClass = new DataGridViewTextBoxColumn();
            tbClass.Name = "班級";
            tbClass.Width = 100;
            tbClass.HeaderText = "班級";
            tbClass.ReadOnly = true;
            tbClass.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            List<string> colNameList = new List<string>();
            foreach (CClassCourseInfo data in _CClassCourseInfoList)
            {
                foreach (string name in data.SubjectBDict.Keys)
                {
                    if (!colNameList.Contains(name))
                        colNameList.Add(name);
                }
            }

            dgData.Columns.Add(tbClass);

            foreach (string name in colNameList)
            {

                DataGridViewCheckBoxColumn tbScoreType = new DataGridViewCheckBoxColumn();
                tbScoreType.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tbScoreType.Name = name;
                tbScoreType.Width = 70;
                tbScoreType.HeaderText = name;
                tbScoreType.ReadOnly = false;
                dgData.Columns.Add(tbScoreType);
            }

        }

        private void LoadData()
        {
            dgData.Rows.Clear();
            foreach (CClassCourseInfo data in _CClassCourseInfoList)
            {
                int rowIdx = dgData.Rows.Add();
                dgData.Rows[rowIdx].Tag = data;
                dgData.Rows[rowIdx].Cells["班級"].Value = data.ClassName;
                foreach (DataGridViewCell cell in dgData.Rows[rowIdx].Cells)
                {
                    if (cell.ColumnIndex > 0)
                    {
                        cell.ReadOnly = true;
                        cell.Style.BackColor = Color.LightGray;
                    }
                }
                foreach (string name in data.SubjectBDict.Keys)
                {
                    dgData.Rows[rowIdx].Cells[name].ReadOnly = false;
                    dgData.Rows[rowIdx].Cells[name].Style.BackColor = Color.White;
                }
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {

            Dictionary<string, List<string>> classSubjectDict = new Dictionary<string, List<string>>();


            // 將勾選預設清空
            foreach (DataGridViewRow drv in dgData.Rows)
            {
                CClassCourseInfo data = drv.Tag as CClassCourseInfo;
                if (data != null)
                {
                    List<string> x = data.SubjectBDict.Keys.ToList();

                                        
                    foreach (string d_name in x)
                    {
                        data.SubjectBDict[d_name] = false;
                    }
                    drv.Tag = data;
                }
            }


            // 整理畫面上有勾選的
            foreach (DataGridViewRow drv in dgData.Rows)
            {
                CClassCourseInfo data = drv.Tag as CClassCourseInfo;
                if (data != null)
                {
                    foreach (string d_name in data.SubjectBDict.Keys)
                    {
                        if (drv.Cells[d_name].Value != null && drv.Cells[d_name].Value.ToString() == "True")
                        {
                            if (!classSubjectDict.ContainsKey(data.ClassName))
                                classSubjectDict.Add(data.ClassName, new List<string>());

                            classSubjectDict[data.ClassName].Add(d_name);
                        }
                    }
                }
            }

            // 設定
            foreach (CClassCourseInfo data in _CClassCourseInfoList)
            {

                if (classSubjectDict.ContainsKey(data.ClassName))
                {
                    foreach (string sname in classSubjectDict[data.ClassName])
                    {
                        if (data.SubjectBDict.ContainsKey(sname))
                            data.SubjectBDict[sname] = true;
                    }
                }
            }

            frmCreateCourseByGPlan108_Create fccc = new frmCreateCourseByGPlan108_Create();
            fccc.SetCClassCourseInfo(_CClassCourseInfoList);
            fccc.SetSchoolYearSemester(_SchoolYear, _Semester);
            if (fccc.ShowDialog() == DialogResult.OK)
            {
                this.DialogResult = DialogResult.OK;
            }
            fccc.StartPosition = FormStartPosition.CenterScreen;


        }
    }
}
