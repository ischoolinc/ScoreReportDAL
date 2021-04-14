using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Permission;
using FISCA.Presentation.Controls;
using K12.Data;

namespace SHCourseGroupCodeSetup
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void main()
        {


            // 新增學生群科班
            K12.Presentation.NLDPanels.Student.AddDetailBulider(new FISCA.Presentation.DetailBulider<DetailContent.UCStudentGroupCodeItem>());

            // 新增班級群科班
            K12.Presentation.NLDPanels.Class.AddDetailBulider(new FISCA.Presentation.DetailBulider<DetailContent.UCClassGroupCodeItem>());

            K12.Presentation.NLDPanels.Class.ListPaneContexMenu["批次產生群科班"].Click += delegate {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {
                    UIForm.frmBatchCourseClassGroupCode frm = new UIForm.frmBatchCourseClassGroupCode(K12.Presentation.NLDPanels.Class.SelectedSource);
                    frm.ShowDialog();
                }
            };

            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["批次產生群科班"].Click += delegate {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    UIForm.frmBatchCourseStudentGroupCode frm = new UIForm.frmBatchCourseStudentGroupCode(K12.Presentation.NLDPanels.Student.SelectedSource);
                    frm.ShowDialog();
                }
            };


            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["產生學生群科班清單"].Click += delegate {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    UIForm.frmBatchCourseStudentGroupCode frm = new UIForm.frmBatchCourseStudentGroupCode(K12.Presentation.NLDPanels.Student.SelectedSource);
                    frm.ShowDialog();
                }
            };
        }
    }
}
