using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using DevComponents.DotNetBar;
using FISCA.Presentation.Controls;
using SHCourseGroupCodeAdmin.DAO;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmGPlanConfig108_MergeSubject : BaseForm
    {
        // 來源課程規畫表
        Dictionary<string, GPlanInfo108> GPlanDict;

        // 合併後課程規劃
        GPlanInfo108 TargetGPlanInfo = null;

        int AddSubejctCount = 0;

        public frmGPlanConfig108_MergeSubject()
        {
            InitializeComponent();
            GPlanDict = new Dictionary<string, GPlanInfo108>();
            TargetGPlanInfo = new GPlanInfo108();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            try
            {

                // 取得畫面上勾選
                List<string> SelectedGPlanNameList = new List<string>();
                foreach (ListViewItem lvi in lvData.CheckedItems)
                {
                    SelectedGPlanNameList.Add(lvi.Name);
                }

                if (SelectedGPlanNameList.Count == 0)
                {
                    MsgBox.Show("請勾選要合併的課程規畫表");
                    return;
                }

                if (cboTarget.Text == "")
                {
                    MsgBox.Show("請選擇合併後的課程規畫表");
                    return;
                }

                // 取得合併後課程規畫
                if (GPlanDict.ContainsKey(cboTarget.Text))
                {
                    TargetGPlanInfo = GPlanDict[cboTarget.Text];
                }
                else
                {
                    MsgBox.Show("合併後的課程規畫表不存在");
                    return;
                }

                btnRun.Enabled = false;

                // 合併資料
                // 取得 tagget 已有的課程代碼
                List<string> TargetCourseCodeList = new List<string>();

                // 最後一筆 RowIdx
                int LastRowIdx = 0;

                foreach (XElement elm in TargetGPlanInfo.RefGPContentXml.Elements("Subject"))
                {
                    string CourseCode = "";
                    if (elm.Attribute("課程代碼") != null)
                        CourseCode = elm.Attribute("課程代碼").Value;

                    // rowIdx，找最後一筆
                    if (elm.Element("Grouping") != null)
                    {
                        if (elm.Element("Grouping").Attribute("RowIndex") != null)
                        {
                            int idx;
                            if (int.TryParse(elm.Element("Grouping").Attribute("RowIndex").Value, out idx))
                            {
                                if (idx > LastRowIdx)
                                    LastRowIdx = idx;
                            }

                        }
                    }

                    if (!TargetCourseCodeList.Contains(CourseCode))
                        TargetCourseCodeList.Add(CourseCode);
                }

                LastRowIdx++;
                // 需要新增科目
                List<XElement> AddSubjectList = new List<XElement>();

                string strRowIdx = ""; // 比對 rowIdx使用

                foreach (string name in SelectedGPlanNameList)
                {
                    if (GPlanDict.ContainsKey(name))
                    {
                        GPlanInfo108 sourceGPlanInfo = GPlanDict[name];

                        // 檢查這目前課程代碼是否已經有，沒有就加入
                        foreach (XElement elm in sourceGPlanInfo.RefGPContentXml.Elements("Subject"))
                        {
                            string CourseCode = "";
                            if (elm.Attribute("課程代碼") != null)
                                CourseCode = elm.Attribute("課程代碼").Value;

                            if (!TargetCourseCodeList.Contains(CourseCode))
                            {
                                // 設定新的 rowIdx
                                XElement NewElm = new XElement(elm);
                                if (strRowIdx == "")
                                {
                                    strRowIdx = NewElm.Element("Grouping").Attribute("RowIndex").Value;
                                }

                                if (strRowIdx != NewElm.Element("Grouping").Attribute("RowIndex").Value)
                                {
                                    LastRowIdx++;
                                    strRowIdx = NewElm.Element("Grouping").Attribute("RowIndex").Value;
                                }

                                NewElm.Element("Grouping").SetAttributeValue("RowIndex", LastRowIdx);
                                AddSubjectList.Add(NewElm);

                            }
                        }
                    }
                }

                // 需要新增資料
                if (AddSubjectList.Count > 0)
                {
                    AddSubejctCount = AddSubjectList.Count;
                    foreach (XElement elm in AddSubjectList)
                        TargetGPlanInfo.RefGPContentXml.Add(elm);
                }

                this.DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void frmGPlanConfig108_MergeSubject_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;

            // 讀取課規名稱放入清單
            foreach (string key in GPlanDict.Keys)
            {
                ListViewItem item = new ListViewItem();
                item.Text = key;
                item.Name = key;
                item.Checked = false;
                lvData.Items.Add(item);
            }

            // 讀取課規名稱放入下選單
            foreach (string key in GPlanDict.Keys)
            {
                cboTarget.Items.Add(key);
            }
            cboTarget.DropDownStyle = ComboBoxStyle.DropDownList;


        }

        // 取得合併後課程規畫
        public GPlanInfo108 GetTargetGPlanInfo()
        {
            return TargetGPlanInfo;
        }

        // 設定來源課程規畫表
        public void SetGPlanDict(Dictionary<string, GPlanInfo108> GPlanDict)
        {
            this.GPlanDict = GPlanDict;
        }

        // 取得新增科目筆數
        public int GetAddSubjectCount()
        {
            return AddSubejctCount;
        }
    }
}
