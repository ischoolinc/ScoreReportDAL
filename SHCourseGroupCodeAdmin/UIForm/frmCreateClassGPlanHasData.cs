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
                XElement GPlanXml = new XElement("GraduationPlan");
            
                foreach (chkSubjectInfo subj in data.chkSubjectInfoList)
                {
                    if (subj.ProcessStatus == "更新" && subj.DiffStatusList.Contains("缺"))
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

                List<XElement> orderList = GPlanXml.Elements("Subject").OrderBy(x => x.Attribute("課程代碼").Value).ToList();

                // 
                int rowIdx = 0;
                string tmpCode = "";
                foreach(XElement elm in orderList)
                {
                    if (elm.Attribute("課程代碼").Value != tmpCode)
                    {
                        rowIdx++;
                        tmpCode = elm.Attribute("課程代碼").Value;
                    }

                    if (elm.Element("Grouping") != null)
                    {
                        elm.Element("Grouping").SetAttributeValue("RowIndex", rowIdx);
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
            btnCreate.Enabled = true;
        }

        private void CreateData()
        {

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
