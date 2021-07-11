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
