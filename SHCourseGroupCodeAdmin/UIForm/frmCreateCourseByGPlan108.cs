﻿using System;
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
using K12.Presentation;
using SHCourseGroupCodeAdmin.DAO;
using System.Xml.Linq;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateCourseByGPlan108 : BaseForm
    {
        List<string> _ClassIDList = new List<string>();
        DataAccess da;
        string _SchoolYear = "";
        string _Semester = "";

        List<CClassCourseInfo> CClassCourseInfoList;

        BackgroundWorker _bwWorker;

        public frmCreateCourseByGPlan108(List<string> ClassIDs)
        {
            InitializeComponent();
            _bwWorker = new BackgroundWorker();
            da = new DataAccess();
            CClassCourseInfoList = new List<CClassCourseInfo>();
            _bwWorker.DoWork += _bwWorker_DoWork;
            _bwWorker.RunWorkerCompleted += _bwWorker_RunWorkerCompleted;
            _bwWorker.ProgressChanged += _bwWorker_ProgressChanged;
            _bwWorker.WorkerReportsProgress = true;
            _ClassIDList = ClassIDs;
        }

        private void _bwWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取資料中 ...", e.ProgressPercentage);
        }

        private void _bwWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ControlEnable(true);
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }

        private void _bwWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bwWorker.ReportProgress(1);
            CClassCourseInfoList = da.GetCClassCourseInfoList(_ClassIDList);

            Dictionary<string, List<string>> classStudentIDList = da.GetClassStudentDict(_ClassIDList);

            // 整理目前學年度學期年級，原班開課
            foreach (CClassCourseInfo data in CClassCourseInfoList)
            {
                // 加入班級學生
                if (classStudentIDList.ContainsKey(data.ClassID))
                {
                    data.RefStudentIDList = classStudentIDList[data.ClassID];
                }


                if (data.RefGPlanXML != null)
                {
                    foreach (XElement subjElm in data.RefGPlanXML.Elements("Subject"))
                    {
                        if (data.GradeYear == subjElm.Attribute("GradeYear").Value && _Semester == subjElm.Attribute("Semester").Value)
                        {
                            if (subjElm.Attribute("開課方式").Value == "原班")
                            {
                                if (subjElm.Attribute("Credit").Value == subjElm.Attribute("學分").Value)
                                {
                                    // 一般
                                    data.OpenSubjectSourceList.Add(subjElm);
                                }
                                else
                                {
                                    // 對開
                                    data.OpenSubjectSourceBList.Add(subjElm);
                                    string subjName = subjElm.Attribute("SubjectName").Value;
                                    if (!data.SubjectBDict.ContainsKey(subjName))
                                        data.SubjectBDict.Add(subjName, false);
                                }
                            }
                        }

                    }
                }
            }

            _bwWorker.ReportProgress(100);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ControlEnable(bool value)
        {
            cboSchoolYear.Enabled = cboSemester.Enabled = btnCreate.Enabled = value;
        }

        private void frmCreateCourseByGPlan108_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            // 預設值
            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");
            cboSemester.DropDownStyle = ComboBoxStyle.DropDownList;
            cboSemester.Text = K12.Data.School.DefaultSemester;
            int sy;
            if (int.TryParse(K12.Data.School.DefaultSchoolYear, out sy))
            {
                for (int i = sy - 3; i <= (sy + 3); i++)
                {
                    cboSchoolYear.Items.Add(i);
                }
            }
            cboSchoolYear.Text = K12.Data.School.DefaultSchoolYear;

            cboSchoolYear.DropDownStyle = ComboBoxStyle.DropDownList;

            _Semester = cboSemester.Text;
            _SchoolYear = cboSchoolYear.Text;

            ControlEnable(false);
            _bwWorker.RunWorkerAsync();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            frmCreateCourseByGPlan108_Detail fccd = new frmCreateCourseByGPlan108_Detail();
            fccd.SetCClassCourseInfo(CClassCourseInfoList);
            fccd.SetSchoolYearSemester(cboSchoolYear.Text, cboSemester.Text);
            if (fccd.ShowDialog() == DialogResult.OK)
            {
                this.Close();
            }
        }

        private void labelX2_Click(object sender, EventArgs e)
        {

        }
    }
}