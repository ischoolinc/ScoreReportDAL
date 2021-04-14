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

            K12.Presentation.NLDPanels.Class.ListPaneContexMenu["產生群科班"].Enable = FISCA.Permission.UserAcl.Current["B53F6CF5-4E20-40FE-BAA3-06668C68A322"].Executable;

            K12.Presentation.NLDPanels.Class.ListPaneContexMenu["產生群科班"].Click += delegate {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {
                    UIForm.frmBatchCourseClassGroupCode frm = new UIForm.frmBatchCourseClassGroupCode(K12.Presentation.NLDPanels.Class.SelectedSource);
                    frm.ShowDialog();
                }
            };

            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["產生群科班"].Enable = FISCA.Permission.UserAcl.Current["8C7E2AFE-F411-4F8F-94E0-AC529EBB2073"].Executable;
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["產生群科班"].Click += delegate {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    UIForm.frmBatchCourseStudentGroupCode frm = new UIForm.frmBatchCourseStudentGroupCode(K12.Presentation.NLDPanels.Student.SelectedSource);
                    frm.ShowDialog();
                }
            };

            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["學生群科班清單"].Enable = FISCA.Permission.UserAcl.Current["E012F3FC-0765-4F07-9548-9120094F2A60"].Executable;
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["學生群科班清單"].Click += delegate {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Reports.rptStudentCourseGroupList rgg = new Reports.rptStudentCourseGroupList(K12.Presentation.NLDPanels.Student.SelectedSource);
                    rgg.Run();
                }
            };

            Catalog ribbon1 = RoleAclSource.Instance["學生"]["功能按鈕"];
            ribbon1.Add(new RibbonFeature("8C7E2AFE-F411-4F8F-94E0-AC529EBB2073", "產生群科班"));
            ribbon1.Add(new RibbonFeature("E012F3FC-0765-4F07-9548-9120094F2A60", "學生群科班清單"));

            Catalog ribbon2 = RoleAclSource.Instance["班級"]["功能按鈕"];
            ribbon2.Add(new RibbonFeature("B53F6CF5-4E20-40FE-BAA3-06668C68A322", "產生群科班"));
        }
    }
}
