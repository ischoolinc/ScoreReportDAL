using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevComponents.DotNetBar.Controls;
using FISCA.Presentation.Controls;
using SHGraduationWarning.DAO;
using System.Web.Script.Serialization;

namespace SHGraduationWarning.UIForm
{
    public partial class frmMain : BaseForm
    {
        string AllStr = "全部";
        BackgroundWorker bgWorkerLoadDefault;
        List<ClassDeptInfo> ClassDeptInfoList;
        string SelectedDeptName = "";
        string SelectedClassName = "";
        Dictionary<string, List<string>> DeptNameDict;

        public frmMain()
        {
            bgWorkerLoadDefault = new BackgroundWorker();
            bgWorkerLoadDefault.DoWork += BgWorkerLoadDefault_DoWork;
            bgWorkerLoadDefault.RunWorkerCompleted += BgWorkerLoadDefault_RunWorkerCompleted;
            bgWorkerLoadDefault.ProgressChanged += BgWorkerLoadDefault_ProgressChanged;
            DeptNameDict = new Dictionary<string, List<string>>();
            ClassDeptInfoList = new List<ClassDeptInfo>(); ;
            SelectedDeptName = SelectedClassName = AllStr;
            bgWorkerLoadDefault.WorkerReportsProgress = true;

            InitializeComponent();
        }

        private void BgWorkerLoadDefault_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("畢業資格檢查中...", e.ProgressPercentage);
        }

        private void BgWorkerLoadDefault_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            // 載入班級科別資訊
            LoadClassDeptToForm();
            ControlEnable(true);
        }

        private void BgWorkerLoadDefault_DoWork(object sender, DoWorkEventArgs e)
        {
            int rpInt = 1;
            bgWorkerLoadDefault.ReportProgress(rpInt);

            // 讀取班級科別資訊            
            ClassDeptInfoList = DataAccess.GetClassDeptList();
            DeptNameDict.Clear();
            foreach (ClassDeptInfo cd in ClassDeptInfoList)
            {
                if (!DeptNameDict.ContainsKey(cd.DeptName))
                    DeptNameDict.Add(cd.DeptName, new List<string>());

                DeptNameDict[cd.DeptName].Add(cd.ClassName);
            }

            rpInt = 10;
            bgWorkerLoadDefault.ReportProgress(rpInt);


            rpInt = 100;
            bgWorkerLoadDefault.ReportProgress(rpInt);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //清除班級科別資訊
        private void ClearClassDept()
        {
            ClassDeptInfoList.Clear();
            comboDept.Text = "";
            comboDept.Items.Clear();
            comboClass.Text = "";
            comboClass.Items.Clear();
        }

        // 載入科別班級資訊
        private void LoadClassDeptToForm()
        {
            this.SuspendLayout();
            // 科別、班級
            comboDept.Items.Add(AllStr);
            comboClass.Items.Add(AllStr);
            foreach (string name in DeptNameDict.Keys)
            {
                comboDept.Items.Add(name);
                foreach (string cname in DeptNameDict[name])
                    comboClass.Items.Add(cname);
            }

            comboDept.Text = AllStr;
            comboClass.Text = AllStr;
            this.ResumeLayout();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {

            ControlEnable(false);
            ClearClassDept();
            comboDept.DropDownStyle = ComboBoxStyle.DropDownList;
            comboClass.DropDownStyle = ComboBoxStyle.DropDownList;
            LoadDgGWColumns();
            // 讀取班級科別資訊
            bgWorkerLoadDefault.RunWorkerAsync();
         
            //DataGridViewButtonXColumn btn2 = new DataGridViewButtonXColumn();

            //btn2.HeaderText = "修改";
            //btn2.Text = "編輯";
            //btn2.UseColumnTextForButtonValue = true;

            //dgDataGW.Columns.Add(btn2);

        }

        private void ControlEnable(bool value)
        {
            btnQuery.Enabled = comboDept.Enabled = comboClass.Enabled = value;
        }

        private void comboDept_SelectedIndexChanged(object sender, EventArgs e)
        {            
            // 清除班級，填入相對值
            comboClass.Text = "";
            SelectedDeptName = comboDept.Text;
            
            if (comboClass.Text != AllStr)
                comboClass.Items.Clear();

            if (SelectedDeptName == AllStr)
            {
                if (comboDept.Items.Count == 0)
                {
                    LoadClassDeptToForm();
                }
            }
            else
            {
                if (DeptNameDict.ContainsKey(SelectedDeptName))
                {
                    foreach (string name in DeptNameDict[SelectedDeptName])
                    {
                        comboClass.Items.Add(name);
                    }
                    if (comboClass.Items.Count > 0)
                        comboClass.SelectedIndex = 0;
                }
            }

        }

        // 載入畢業預警檢查欄位
        private void LoadDgGWColumns()
        {
            try
            {
                string textColumnStrig = @"
            [
                {
                    ""HeaderText"": ""學號"",
                    ""Name"": ""學號"",
                    ""Width"": 80,
                    ""ReadOnly"": true
                },
                {
                    ""HeaderText"": ""班級"",
                    ""Name"": ""班級"",
                    ""Width"": 120,
                    ""ReadOnly"": true
                }
            ]
";
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                List<DataGridViewTextBoxColumnInfo> jsonObjArray = serializer.Deserialize<List<DataGridViewTextBoxColumnInfo>>(textColumnStrig);
                foreach (DataGridViewTextBoxColumnInfo jObj in jsonObjArray)
                {
                    DataGridViewTextBoxColumn dgt = new DataGridViewTextBoxColumn();
                    dgt.Name = jObj.Name;
                    dgt.Width = jObj.Width;
                    dgt.HeaderText = jObj.HeaderText;
                    dgt.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgt.ReadOnly = jObj.ReadOnly;
                    dgDataGW.Columns.Add(dgt);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
        }

    }
}
