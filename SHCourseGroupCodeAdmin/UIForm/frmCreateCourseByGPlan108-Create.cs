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
using SHCourseGroupCodeAdmin.DAO;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateCourseByGPlan108_Create : BaseForm
    {
        BackgroundWorker _bgWorker;
        string _SchoolYear = "", _Semester = "";
        List<CClassCourseInfo> _CClassCourseInfoList;
        DataAccess _da;
        StringBuilder _sb = new StringBuilder();

        public frmCreateCourseByGPlan108_Create()
        {
            InitializeComponent();
            _bgWorker = new BackgroundWorker();
            _da = new DataAccess();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.WorkerReportsProgress = true;
        }

        public void SetSchoolYearSemester(string SchoolYear, string Semester)
        {
            _SchoolYear = SchoolYear;
            _Semester = Semester;
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("產生課程中 ...", e.ProgressPercentage);
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_sb.Length > 2)
            {
                MsgBox.Show("錯誤：" + _sb.ToString());

            }
            else
            {
                // 呼叫課程同步
                FISCA.Features.Invoke("CourseSyncAllBackground");
                
                FISCA.Presentation.MotherForm.SetStatusBarMessage("產生完成。");
                MsgBox.Show("產生完成。");
                this.DialogResult = DialogResult.OK;
            }
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            _sb.Clear();
            // 新增課程
            _CClassCourseInfoList = _da.AddGPlanCourseBySchoolYearSemester(_SchoolYear, _Semester, _CClassCourseInfoList);
            _bgWorker.ReportProgress(50);

            // 加入課程修課學生
            _sb.AppendLine(_da.AddCourseStudent(_SchoolYear, _Semester, _CClassCourseInfoList));
            _bgWorker.ReportProgress(100);
        }

        public void SetCClassCourseInfo(List<CClassCourseInfo> data)
        {
            _CClassCourseInfoList = data;
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            btnCreate.Enabled = false;
            _bgWorker.RunWorkerAsync();
        }

        private void frmCreateCourseByGPlan108_Create_Load(object sender, EventArgs e)
        {
            this.StartPosition = FormStartPosition.CenterScreen;

            StringBuilder sb = new StringBuilder();
            Dictionary<string, List<string>> nameDict = new Dictionary<string, List<string>>();
            sb.AppendLine("班級開課清單：");
            foreach (CClassCourseInfo cc in _CClassCourseInfoList)
            {
                if (!nameDict.ContainsKey(cc.ClassName))
                    nameDict.Add(cc.ClassName, new List<string>());

                foreach (string key in cc.SubjectBDict.Keys)
                {
                    if (cc.SubjectBDict[key] == true)
                    {
                        nameDict[cc.ClassName].Add(key);
                    }
                }
            }

            foreach (string key in nameDict.Keys)
            {
                if (nameDict[key].Count > 0)
                    sb.AppendLine(key + "：" + string.Join(",", nameDict[key].ToArray()));
            }

            txtMsg.Text = sb.ToString();
            txtMsg.SelectionLength = 0;
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
