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
using System.Xml.Linq;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCreateClassGPlanHasData : BaseForm
    {
        // 選擇的課程代碼大表名稱
        string SelectGroupName = "";

        // 選擇的課程代碼大表解析後 XML
        XElement SelectMOEXml = null;
        DataAccess da = new DataAccess();
        List<GPlanData> _GPlanData = new List<GPlanData>();

        public frmCreateClassGPlanHasData()
        {
            InitializeComponent();
        }

        private void frmCreateClassGPlanHasData_Load(object sender, EventArgs e)
        {
            this.lblGroupName.Text = SelectGroupName;
            LoadData();
        }

        private void LoadData()
        {
            foreach (GPlanData gData in _GPlanData)
            {
                gData.MOEXml = SelectMOEXml;
                gData.CheckData();

            }

            dgData.Rows.Clear();
            foreach (GPlanData gData in _GPlanData)
            {
                int rowIdx = dgData.Rows.Add();
                dgData.Rows[rowIdx].Tag = gData;
                // 名稱
                dgData.Rows[rowIdx].Cells[colGPName.Index].Value = gData.Name;
                // 採用班級
                dgData.Rows[rowIdx].Cells[colUsedClass.Index].Value = string.Join(",", gData.UsedClassIDNameDict.Values.ToArray());
                // 差異科目
                dgData.Rows[rowIdx].Cells[colDiffSujeCount.Index].Value = gData.calSubjDiffCount();
                // 更新科目
                dgData.Rows[rowIdx].Cells[colUpdateSubjCount.Index].Value = gData.calSubjUpdateCount();
                dgData.Rows[rowIdx].Cells[colSet.Index].Value = "設定";
            }
        }

        public void SetGPlanData(List<GPlanData> dataList)
        {
            _GPlanData = dataList;
        }


        public void SetGroupName(string name)
        {
            SelectGroupName = name;
        }

        public void SetSelectMOEXML(XElement data)
        {
            SelectMOEXml = data;
        }




        private void btnCreate_Click(object sender, EventArgs e)
        {
            btnCreate.Enabled = false;

            foreach(GPlanData data in _GPlanData)
            {
                //if (data.calSubjUpdateCount() == 0)
                //    continue;
                
                XElement GPlanXml = new XElement("GraduationPlan");
            
                foreach (chkSubjectInfo subj in data.chkSubjectInfoList)
                {                 

                    if (subj.ProcessStatus == "更新")
                    {
                        foreach (XElement elm in subj.MOEXml)
                        {
                            GPlanXml.Add(elm);
                        } 
                    }


                    if (subj.ProcessStatus == "略過")
                    {
                        foreach(XElement elm in subj.GPlanXml)
                        {
                            GPlanXml.Add(elm);
                        }
                    }
                }

                // 依課程代碼,科目名稱排序
                List<XElement> orderList = (from elmData in GPlanXml.Elements("Subject") orderby elmData.Attribute("課程代碼").Value ascending, elmData.Attribute("SubjectName").Value select elmData).ToList();
              //  List<XElement> orderList = GPlanXml.Elements("Subject").OrderBy(x => x.Attribute("課程代碼").Value).ToList();

               // 重整 idx                                
                int rowIdx = 0;              
                string tmpCode = "";
                foreach(XElement elm in orderList)
                {
                    string codeA = elm.Attribute("課程代碼").Value+ elm.Attribute("SubjectName").Value;
                    if (codeA != tmpCode)                    {
                        rowIdx++;
                        tmpCode = codeA;
                    }

                    if (elm.Element("Grouping") != null)
                    {
                        elm.Element("Grouping").SetAttributeValue("RowIndex", rowIdx);
                    }
                }

                // 重新排列科目級別
                Dictionary<string, int> tmpSubjLevelDict = new Dictionary<string, int>();
                
                foreach (XElement elm in GPlanXml.Elements("Subject"))
                {                    
                    string subj = elm.Attribute("SubjectName").Value;

                    if (!tmpSubjLevelDict.ContainsKey(subj))
                        tmpSubjLevelDict.Add(subj, 0);

                    tmpSubjLevelDict[subj] += 1;

                    elm.SetAttributeValue("FullName", SubjFullName(subj, tmpSubjLevelDict[subj]));
                    elm.SetAttributeValue("Level", tmpSubjLevelDict[subj]);
                    
                }

                // 重新整理開始級別
                Dictionary<string, string> tmpStartLevel = new Dictionary<string, string>();
                foreach (XElement elm in GPlanXml.Elements("Subject"))
                {
                    string subjName = elm.Attribute("SubjectName").Value;

                    string RowIndex = elm.Element("Grouping").Attribute("RowIndex").Value;

                    if (!tmpStartLevel.ContainsKey(subjName))
                        tmpStartLevel.Add(subjName, RowIndex);
                    else
                    {
                        if (tmpStartLevel[subjName] != RowIndex)
                        {
                            // 設定開始級別是目前級別
                            elm.Element("Grouping").SetAttributeValue("startLevel", elm.Attribute("Level").Value);
                            tmpStartLevel[subjName] = RowIndex;
                        }
                    }
                }


                GPlanXml.ReplaceAll(orderList);
                if (data.MOEGroupCode.Length > 3)
                    GPlanXml.SetAttributeValue("SchoolYear", data.MOEGroupCode.Substring(0, 3));

                //  Console.WriteLine(GPlanXml.ToString());
                //  data.ContentXML.ReplaceAll(GPlanXml);

                da.UpdateGPlanXML(data.ID, GPlanXml.ToString());

            }
            MsgBox.Show("完成");
            this.Close();
            
        }

        private string SubjFullName(string SubjectName, int level)
        {
            string lev = "";
            if (level == 1)
                lev = " I";

            if (level == 2)
                lev = " II";

            if (level == 3)
                lev = " III";

            if (level == 4)
                lev = " IV";

            if (level == 5)
                lev = " V";

            if (level == 6)
                lev = " VI";

            string value = SubjectName + lev;

            return value;
        }
        

        private void dgData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                if (e.ColumnIndex == colSet.Index)
                {
                    GPlanData data = dgData.Rows[e.RowIndex].Tag as GPlanData;
                    if (data != null)
                    {
                        frmCreateClassGPlanSetDetail fGPD = new frmCreateClassGPlanSetDetail();
                        fGPD.SetGPlanData(data);
                        fGPD.SetMOENameAndXml(SelectGroupName, SelectMOEXml);
                        if (fGPD.ShowDialog() == DialogResult.OK)
                        {
                            GPlanData newData = fGPD.GetGPlanData();
                            dgData.Rows[e.RowIndex].Tag = newData;
                            dgData.Rows[e.RowIndex].Cells[colDiffSujeCount.Index].Value = newData.calSubjDiffCount();
                            // 更新科目
                            dgData.Rows[e.RowIndex].Cells[colUpdateSubjCount.Index].Value = newData.calSubjUpdateCount();
                        }
                    }
                }
            }
        }
    }
}
