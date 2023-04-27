using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA;
using FISCA.Permission;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using SHGraduationWarning.DAO;
using SHGraduationWarning.UIForm;

namespace SHGraduationWarning
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void main()
        {
            // 教務作業>批次作業/檢視>成績作業>畢業預警
            Catalog ribbon1 = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            ribbon1.Add(new RibbonFeature("6460BE9A-3E82-43C7-ACE7-64E0C6100473", "畢業預警"));

            MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["畢業預警"].Enable = UserAcl.Current["6460BE9A-3E82-43C7-ACE7-64E0C6100473"].Executable;

            MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["畢業預警"].Click += delegate
            {
                // 畢業預警
                frmMain fm = new frmMain();
                fm.ShowDialog();
            };
        }
    }
}
