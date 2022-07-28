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
        List<string> _errClassList = new List<string>();
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
            if (e.Cancelled)
            {
                MsgBox.Show("班級：" + string.Join(",", _errClassList.ToArray()) + "，使用課程規劃非108適用，無法產生。");
            }
            else
            {
                ControlEnable(true);
                FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取完成");
                frmCreateCourseByGPlan108_Detail fccd = new frmCreateCourseByGPlan108_Detail();
                fccd.SetCClassCourseInfo(CClassCourseInfoList);
                fccd.SetSchoolYearSemester(cboSchoolYear.Text, cboSemester.Text);
                if (fccd.ShowDialog() == DialogResult.OK)
                {
                    this.Close();
                }
            }
        }

        private void _bwWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _errClassList.Clear();
            _bwWorker.ReportProgress(1);
            CClassCourseInfoList = da.GetCClassCourseInfoList(_ClassIDList);

            // 檢查課程規劃表
            foreach (CClassCourseInfo data in CClassCourseInfoList)
            {
                if (data.RefGPlanXML == null)
                    _errClassList.Add(data.ClassName);
            }

            if (_errClassList.Count > 0)
                e.Cancel = true;

            //取得班級學生
            Dictionary<string, List<string>> classStudentIDList = da.GetClassStudentDict(_ClassIDList);

            // 整理目前學年度學期年級，原班開課
            foreach (CClassCourseInfo data in CClassCourseInfoList)
            {
                // 加入班級學生
                if (classStudentIDList.ContainsKey(data.ClassID))
                {
                    data.RefStudentIDList = classStudentIDList[data.ClassID];
                }
                data.OpenSubjectSourceList.Clear();
                data.OpenSubjectSourceBList.Clear();
                List<string> tmpSubj = new List<string>();

                if (data.RefGPlanXML != null)
                {
                    // 原本課規
                    foreach (XElement subjElm in data.RefGPlanXML.Elements("Subject"))
                    {
                        if (data.GradeYear == subjElm.Attribute("GradeYear").Value && _Semester == subjElm.Attribute("Semester").Value)
                        {
                            if (subjElm.Attribute("開課方式").Value == "原班")
                            {
                                // 判斷對開
                                bool isOpenD = false;

                                // 使用者手動設定
                                if (subjElm.Attribute("設定對開") != null)
                                {
                                    if (subjElm.Attribute("設定對開").Value == "是")
                                        isOpenD = true;
                                }

                                // 處理對開
                                if (subjElm.Attribute("授課學期學分") != null)
                                {
                                    if (subjElm.Attribute("授課學期學分").Value.Length > 5)
                                    {
                                        string credit_period = subjElm.Attribute("授課學期學分").Value;
                                        string credit = "0";
                                        char[] cp = credit_period.ToArray();

                                        if (data.GradeYear == "1" && _Semester == "1")
                                            credit = cp[0] + "";

                                        if (data.GradeYear == "1" && _Semester == "2")
                                            credit = cp[1] + "";

                                        if (data.GradeYear == "2" && _Semester == "1")
                                            credit = cp[2] + "";

                                        if (data.GradeYear == "2" && _Semester == "2")
                                            credit = cp[3] + "";

                                        if (data.GradeYear == "3" && _Semester == "1")
                                            credit = cp[4] + "";

                                        if (data.GradeYear == "3" && _Semester == "2")
                                            credit = cp[5] + "";

                                        int x;
                                        if (int.TryParse(credit, out x))
                                        {
                                            // 一般學分
                                        }
                                        else
                                        {
                                            //對開
                                            isOpenD = true;
                                        }

                                    }
                                }


                                if (isOpenD)
                                {
                                    // 是對開
                                    data.OpenSubjectSourceBList.Add(subjElm);
                                    string subjName = subjElm.Attribute("SubjectName").Value;
                                    if (!data.SubjectBDict.ContainsKey(subjName))
                                        data.SubjectBDict.Add(subjName, false);
                                }
                                else
                                {
                                    // 一般
                                    data.OpenSubjectSourceList.Add(subjElm);
                                    string subjName = subjElm.Attribute("SubjectName").Value;
                                    if (!tmpSubj.Contains(subjName))
                                    {
                                        tmpSubj.Add(subjName);
                                    }
                                }
                            }
                        }

                    }

                    // 使用者自訂
                    if (data.RefGPlanXML.Element("使用者自訂科目") != null)
                    {
                        foreach (XElement subjElm in data.RefGPlanXML.Element("使用者自訂科目").Elements("Subject"))
                        {
                            if (data.GradeYear == subjElm.Attribute("GradeYear").Value && _Semester == subjElm.Attribute("Semester").Value)
                            {
                                if (subjElm.Attribute("開課方式").Value == "原班")
                                {
                                    // 判斷對開
                                    bool isOpenD = false;

                                    // 使用者手動設定
                                    if (subjElm.Attribute("設定對開") != null)
                                    {
                                        if (subjElm.Attribute("設定對開").Value == "是")
                                            isOpenD = true;
                                    }                              


                                    if (isOpenD)
                                    {
                                        // 是對開
                                        data.OpenSubjectSourceBList.Add(subjElm);
                                        string subjName = subjElm.Attribute("SubjectName").Value;
                                        if (!data.SubjectBDict.ContainsKey(subjName))
                                            data.SubjectBDict.Add(subjName, false);
                                    }
                                    else
                                    {
                                        // 一般
                                        data.OpenSubjectSourceList.Add(subjElm);
                                        string subjName = subjElm.Attribute("SubjectName").Value;
                                        if (!tmpSubj.Contains(subjName))
                                        {
                                            tmpSubj.Add(subjName);
                                        }
                                    }
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

            
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            _Semester = cboSemester.Text;
            _SchoolYear = cboSchoolYear.Text;

            ControlEnable(false);
            _bwWorker.RunWorkerAsync();         
        }

        private void labelX2_Click(object sender, EventArgs e)
        {

        }
    }
}
