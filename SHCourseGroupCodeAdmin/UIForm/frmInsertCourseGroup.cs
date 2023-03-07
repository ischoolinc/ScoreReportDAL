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
    public partial class frmInsertCourseGroup : BaseForm
    {
        List<CourseGroupSetting> _CourseGroupSettingList = new List<CourseGroupSetting>();

        public frmInsertCourseGroup(List<CourseGroupSetting> courseGroupSettingList)
        {
            InitializeComponent();
            _CourseGroupSettingList = courseGroupSettingList;
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbCourseGroupName.Text))
            {
                MessageBox.Show("群組名稱不可為空");
                return;
            }
            if (string.IsNullOrEmpty(tbCredit.Text))
            {
                MessageBox.Show("群組修課學分數不可為空");
                return;
            }

            string courseGroupName = tbCourseGroupName.Text;
            string courseGroupCredit = tbCredit.Text;
            Color color = cpColorPicker.SelectedColor;
            bool isSchoolYearCourseGroup = cbIsSchoolYearCourseGroup.Checked;

            if (color == Color.Empty)
            {
                color = Color.LightBlue;
            }

            if (_CourseGroupSettingList.Where(x => x.CourseGroupColor.ToArgb() == color.ToArgb()).Count() > 0)
            {
                MessageBox.Show("課程群組顏色不可重複");
                return;
            }

            if (_CourseGroupSettingList.Where(x => x.CourseGroupName == courseGroupName).Count() > 0)
            {
                MessageBox.Show("課程群組名稱不可重複");
                return;
            }

            XElement element = new XElement("CourseGroup");
            element.SetAttributeValue("Name", courseGroupName);
            element.SetAttributeValue("Credit", courseGroupCredit);
            element.SetAttributeValue("Color", color.ToArgb());
            element.SetAttributeValue("IsSchoolYearCourseGroup", isSchoolYearCourseGroup);

            _CourseGroupSettingList.Add(new CourseGroupSetting()
            {
                CourseGroupName = courseGroupName,
                CourseGroupCredit = courseGroupCredit,
                CourseGroupColor = color,
                IsSchoolYearCourseGroup = isSchoolYearCourseGroup,
                CourseGroupElement = element
            });

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
