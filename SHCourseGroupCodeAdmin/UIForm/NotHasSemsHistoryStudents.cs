using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SHCourseGroupCodeAdmin.UIForm
{
    public partial class NotHasSemsHistoryStudents : BaseForm
    {

        public NotHasSemsHistoryStudents(DataTable dataTable, string GradeYear, int SchoolYear, int Semester)
        {
            InitializeComponent();
            dgv.DataSource = dataTable;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            string gradeYear = GradeYear;
            string schoolYear = SchoolYear.ToString();
            string semester = Semester.ToString();

            labelX1.Text = "下列"+ GradeYear + "年級學生在 "+schoolYear+"學年度 第"+semester+"學期 沒有完整的學期對照表資料，\r\n" +
                "請確認下列學生於 " + schoolYear + "學年度 第"+semester+"學期 是否在校，\r\n" +
                "若學生不在校，請忽略訊息，\r\n" +
                "若學生在校，請將學期對照表資料補齊後再檢核。";


            //id
            dgv.Columns[0].Visible = false;

            ////學號
            dgv.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            ////班級
            //dgv.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;    

            ////座號
            dgv.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            ////姓名
            dgv.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

        }

        private void NotHasSemsHistoryStudents_Load(object sender, EventArgs e)
        {

        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void cmdPrint_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void AddTemp_Click(object sender, EventArgs e)
        {
            List<string> studentlist = new List<string>();

            for (int i = dgv.Rows.Count - 1; i >= 0; i--)
            {
                studentlist.Add(dgv.Rows[i].Cells[0].Value.ToString());
            }
            K12.Presentation.NLDPanels.Student.AddToTemp(studentlist); //跨執行緒作業無效: 存取控制項 'btnTempory' 時所使用的執行緒與建立控制項的執行緒不同。'

        }
    }
}
