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
using DevComponents.DotNetBar;
using System.Xml.Linq;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmGPlanConfig108 : BaseForm
    {
        BackgroundWorker _bgWorker;
        DataAccess _da;
        List<GPlanInfo108> GP108List;

        private ButtonItem _SelectButton;

        public frmGPlanConfig108()
        {
            InitializeComponent();
            GP108List = new List<GPlanInfo108>();

            _da = new DataAccess();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
            itemPanel1.Items.Clear();
            _SelectButton = null;
            foreach (GPlanInfo108 data in GP108List)
            {
                ButtonItem item = new ButtonItem(data.RefGPID, data.RefGPName);
                item.Tag = data;

                item.ImagePosition = eImagePosition.Left;
                item.ImageFixedSize = new Size(14, 14);
                item.Image = null;
                item.ButtonStyle = eButtonStyle.TextOnlyAlways;
                item.Click += new EventHandler(item_Click);
                itemPanel1.Items.Add(item);
                
            }

            LoadDataGridViewColumns();

            itemPanel1.Refresh();
        }

        private void item_Click(object sender, EventArgs e)
        {
            if (_SelectButton != null)
                _SelectButton.Checked = false;

            ButtonItem item = (ButtonItem)sender;
            GPlanInfo108 info = (GPlanInfo108)item.Tag;
            info.RefGPContentXml = XElement.Parse(info.RefGPContent);
            _SelectButton = item;
            lblGroupName.Text = info.RefGPName;
            item.Checked = true;

            // 資料整理
            Dictionary<string, List<XElement>> dataDict = new Dictionary<string, List<XElement>>();
            foreach (XElement elm in info.RefGPContentXml.Elements("Subject"))
            {
                string idx = elm.Element("Grouping").Attribute("RowIndex").Value;

                if (!dataDict.ContainsKey(idx))
                    dataDict.Add(idx, new List<XElement>());

                dataDict[idx].Add(elm);
            }

            // 取得採用班級
            Dictionary<string, List<DataRow>> classRows = _da.GetGPlanRefClaasByID(info.RefGPID);


            dgData.Rows.Clear();

            // 填入資料
            foreach (string idx in dataDict.Keys)
            {
                int rowIdx = dgData.Rows.Add();
                XElement firstElm = null;
                if (dataDict[idx].Count > 0)
                {
                    firstElm = dataDict[idx][0];
                }

                dgData.Rows[rowIdx].Cells["領域"].Value = firstElm.Attribute("Domain").Value;
                dgData.Rows[rowIdx].Cells["分項類別"].Value = firstElm.Attribute("Entry").Value;
                dgData.Rows[rowIdx].Cells["科目名稱"].Value = firstElm.Attribute("SubjectName").Value;

                if (firstElm.Attribute("RequiredBy").Value == "部訂")
                {
                    dgData.Rows[rowIdx].Cells["校訂部定"].Value = "部定";
                }
                else
                    dgData.Rows[rowIdx].Cells["校訂部定"].Value = firstElm.Attribute("RequiredBy").Value;

                dgData.Rows[rowIdx].Cells["必選修"].Value = firstElm.Attribute("Required").Value;

                foreach (XElement elmD in dataDict[idx])
                {
                    try
                    {

                        if (elmD.Attribute("GradeYear").Value == "1" && elmD.Attribute("Semester").Value == "1")
                        {
                            dgData.Rows[rowIdx].Cells["1上"].Value = elmD.Attribute("學分").Value;
                        }

                        if (elmD.Attribute("GradeYear").Value == "1" && elmD.Attribute("Semester").Value == "2")
                        {
                            dgData.Rows[rowIdx].Cells["1下"].Value = elmD.Attribute("學分").Value;
                        }

                        if (elmD.Attribute("GradeYear").Value == "2" && elmD.Attribute("Semester").Value == "1")
                        {
                            dgData.Rows[rowIdx].Cells["2上"].Value = elmD.Attribute("學分").Value;
                        }

                        if (elmD.Attribute("GradeYear").Value == "2" && elmD.Attribute("Semester").Value == "2")
                        {
                            dgData.Rows[rowIdx].Cells["2下"].Value = elmD.Attribute("學分").Value;
                        }

                        if (elmD.Attribute("GradeYear").Value == "3" && elmD.Attribute("Semester").Value == "1")
                        {
                            dgData.Rows[rowIdx].Cells["3上"].Value = elmD.Attribute("學分").Value;
                        }

                        if (elmD.Attribute("GradeYear").Value == "3" && elmD.Attribute("Semester").Value == "2")
                        {
                            dgData.Rows[rowIdx].Cells["3下"].Value = elmD.Attribute("學分").Value;
                        }


                    }
                    catch (Exception ex) { Console.WriteLine(ex.Message); }
                }

                dgData.Rows[rowIdx].Cells["開課方式"].Value = firstElm.Attribute("開課方式").Value;
                dgData.Rows[rowIdx].Cells["課程代碼"].Value = firstElm.Attribute("課程代碼").Value;
            }





            listViewEx1.SuspendLayout();
            listViewEx1.Items.Clear();
            listViewEx1.Groups.Clear();

            foreach (string key in classRows.Keys)
            {
                string groupKey;

                groupKey = key + "　年級";

                foreach(DataRow dr in classRows[key])
                {
                    ListViewGroup group = listViewEx1.Groups[groupKey];
                    if (group == null)
                        group = listViewEx1.Groups.Add(groupKey, groupKey);

                    string c_name = dr["class_name"] + "(" + dr["stud_cot"] + ")";
                    ListViewItem lvi = new ListViewItem(c_name, 0, group);
                    listViewEx1.Items.Add(lvi);
                }              
            }
            listViewEx1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            listViewEx1.ResumeLayout();
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("讀取課程規劃表...", e.ProgressPercentage);
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            GP108List = _da.GetGPlanInfo108All();
            _bgWorker.ReportProgress(100);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            MsgBox.Show("儲存完成");
        }

        private void frmGPlanConfig108_Load(object sender, EventArgs e)
        {
            _bgWorker.RunWorkerAsync();
        }

        private void LoadData()
        {
            dgData.Rows.Clear();
        }

        private void LoadDataGridViewColumns()
        {
            try
            {

                DataGridViewTextBoxColumn tbDomain = new DataGridViewTextBoxColumn();
                tbDomain.Name = "領域";
                tbDomain.Width = 70;
                tbDomain.HeaderText = "領域";
                tbDomain.ReadOnly = true;

                DataGridViewTextBoxColumn tbScoreType = new DataGridViewTextBoxColumn();
                tbScoreType.Name = "分項類別";
                tbScoreType.Width = 90;
                tbScoreType.HeaderText = "分項類別";
                tbScoreType.ReadOnly = true;

                DataGridViewTextBoxColumn tbSubjectName = new DataGridViewTextBoxColumn();
                tbSubjectName.Name = "科目名稱";
                tbSubjectName.Width = 150;
                tbSubjectName.HeaderText = "科目名稱";
                tbSubjectName.ReadOnly = true;

                DataGridViewTextBoxColumn tbRequiredBy = new DataGridViewTextBoxColumn();
                tbRequiredBy.Name = "校訂部定";
                tbRequiredBy.Width = 90;
                tbRequiredBy.HeaderText = "校訂部定";
                tbRequiredBy.ReadOnly = true;

                DataGridViewTextBoxColumn tbIsRequired = new DataGridViewTextBoxColumn();
                tbIsRequired.Name = "必選修";
                tbIsRequired.Width = 90;
                tbIsRequired.HeaderText = "必選修";
                tbIsRequired.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS11 = new DataGridViewTextBoxColumn();
                tbGS11.Name = "1上";
                tbGS11.Width = 60;
                tbGS11.HeaderText = "1上";
                tbGS11.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS12 = new DataGridViewTextBoxColumn();
                tbGS12.Name = "1下";
                tbGS12.Width = 60;
                tbGS12.HeaderText = "1下";
                tbGS12.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS21 = new DataGridViewTextBoxColumn();
                tbGS21.Name = "2上";
                tbGS21.Width = 60;
                tbGS21.HeaderText = "2上";
                tbGS21.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS22 = new DataGridViewTextBoxColumn();
                tbGS22.Name = "2下";
                tbGS22.Width = 60;
                tbGS22.HeaderText = "2下";
                tbGS22.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS31 = new DataGridViewTextBoxColumn();
                tbGS31.Name = "3上";
                tbGS31.Width = 60;
                tbGS31.HeaderText = "3上";
                tbGS31.ReadOnly = true;

                DataGridViewTextBoxColumn tbGS32 = new DataGridViewTextBoxColumn();
                tbGS32.Name = "3下";
                tbGS32.Width = 60;
                tbGS32.HeaderText = "3下";
                tbGS32.ReadOnly = true;

                DataGridViewTextBoxColumn tbOpenStatus = new DataGridViewTextBoxColumn();
                tbOpenStatus.Name = "開課方式";
                tbOpenStatus.Width = 90;
                tbOpenStatus.HeaderText = "開課方式";
                tbOpenStatus.ReadOnly = true;

                DataGridViewTextBoxColumn tbCourseCode = new DataGridViewTextBoxColumn();
                tbCourseCode.Name = "課程代碼";
                tbCourseCode.Width = 300;
                tbCourseCode.HeaderText = "課程代碼";
                tbCourseCode.ReadOnly = true;


                dgData.Columns.Add(tbDomain);
                dgData.Columns.Add(tbScoreType);
                dgData.Columns.Add(tbSubjectName);
                dgData.Columns.Add(tbRequiredBy);
                dgData.Columns.Add(tbIsRequired);
                dgData.Columns.Add(tbGS11);
                dgData.Columns.Add(tbGS12);
                dgData.Columns.Add(tbGS21);
                dgData.Columns.Add(tbGS22);
                dgData.Columns.Add(tbGS31);
                dgData.Columns.Add(tbGS32);
                dgData.Columns.Add(tbOpenStatus);
                dgData.Columns.Add(tbCourseCode);

                //// 因為自動排序有些問題，先將關閉
                //foreach(DataGridViewColumn col in dgData.Columns )
                //{
                //    col.SortMode = DataGridViewColumnSortMode.NotSortable;
                //}

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void tabControl1_Click(object sender, EventArgs e)
        {

        }
    }
}
