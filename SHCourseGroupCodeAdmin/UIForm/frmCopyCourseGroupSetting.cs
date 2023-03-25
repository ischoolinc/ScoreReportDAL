using FISCA.Presentation.Controls;
using SHCourseGroupCodeAdmin.DAO;
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

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class frmCopyCourseGroupSetting : BaseForm
    {
        List<GPlanInfo108> _GraduationPlanList = new List<GPlanInfo108>();
        GPlanInfo108 _SelectedGraduationPlan = new GPlanInfo108();
        List<GPlanInfo108> _HasSettingGraduationPlanList = new List<GPlanInfo108>();

        public frmCopyCourseGroupSetting(List<GPlanInfo108> gradautionPlanList, GPlanInfo108 selectedGraduationPlan)
        {
            InitializeComponent();

            _GraduationPlanList = gradautionPlanList;
            _SelectedGraduationPlan = selectedGraduationPlan;
        }
        private void frmCopyCourseGroupSetting_Load(object sender, EventArgs e)
        {
            cboGraduationPlanName.Items.Clear();

            foreach (GPlanInfo108 graduationPlan in _GraduationPlanList)
            {
                if (graduationPlan.RefGPContent != null && graduationPlan.RefGPName != _SelectedGraduationPlan.RefGPName)
                {
                    XElement element = XElement.Parse(graduationPlan.RefGPContent);

                    // 有課程群組設定的課程規畫表才加進下拉式選單中
                    if (element.Element("CourseGroupSetting") != null)
                    {
                        if (element.Element("CourseGroupSetting").Elements("CourseGroup").Count() > 0)
                        {
                            _HasSettingGraduationPlanList.Add(graduationPlan);
                        }
                    }
                }
            }

            cboGraduationPlanName.Items.AddRange(_HasSettingGraduationPlanList.Select(x => x.RefGPName).ToArray());

            if (cboGraduationPlanName.Items.Count > 0)
            {
                cboGraduationPlanName.SelectedIndex = 0;
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cboGraduationPlanName.Text))
            {
                MessageBox.Show("請選擇包含課程群組設定的課程規畫表");
                return;
            }

            int index = cboGraduationPlanName.SelectedIndex;
            XElement selectedGradudationPlanElement = _SelectedGraduationPlan.RefGPContentXml;
            XElement copiedGraduationPlanElement = XElement.Parse(_HasSettingGraduationPlanList[index].RefGPContent);

            if (selectedGradudationPlanElement.Element("CourseGroupSetting") == null)
            {
                selectedGradudationPlanElement.Add(new XElement("CourseGroupSetting"));
            }

            bool hasDuplicate = false;
            string errMessage = "";
            List<XElement> selectedCourseGroupList = selectedGradudationPlanElement.Element("CourseGroupSetting").Elements("CourseGroup").ToList();
            List<XElement> copiedCourseGroupList = copiedGraduationPlanElement.Element("CourseGroupSetting").Elements("CourseGroup").ToList();

            foreach (XElement courseGroupSettingElement in copiedCourseGroupList)
            {
                if (selectedCourseGroupList.Where(x => x.Attribute("Name").Value == courseGroupSettingElement.Attribute("Name").Value).Count() > 0)
                {
                    errMessage = "欲複製的群組設定中包含重複的群組名稱";
                    hasDuplicate = true;
                    break;
                }

                if (selectedCourseGroupList.Where(x => x.Attribute("Color").Value == courseGroupSettingElement.Attribute("Color").Value).Count() > 0)
                {
                    errMessage = "欲複製的群組設定中包含重複的顯示顏色";
                    hasDuplicate = true;
                    break;
                }
            }

            if (hasDuplicate)
            {
                MessageBox.Show(errMessage);
                return;
            }

            foreach (XElement courseGroupSettingElement in copiedGraduationPlanElement.Element("CourseGroupSetting").Elements("CourseGroup"))
            {
                selectedGradudationPlanElement.Element("CourseGroupSetting").Add(courseGroupSettingElement);
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
