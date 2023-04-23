using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA;
using FISCA.Permission;
using FISCA.Presentation;
using SHSemsSubjectCheckEdit.UIForm;

namespace SHSemsSubjectCheckEdit
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void main()
        {
            // 教務作業>批次作業/檢視>排名作業>學期成績科目檢查與調整
            Catalog ribbon1 = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            ribbon1.Add(new RibbonFeature("3F82357D-B4CE-46CD-9992-B5936470D6C6", "學期成績科目檢查與調整"));

            MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["學期成績科目檢查與調整"].Enable = UserAcl.Current["3F82357D-B4CE-46CD-9992-B5936470D6C6"].Executable;

            MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["學期成績科目檢查與調整"].Click += delegate
            {
                // 學期成績科目檢查與調整
                frmSemsSubjectNameCheckEdit fss = new frmSemsSubjectNameCheckEdit();
                fss.ShowDialog();
            };


            // 教務作業>批次作業/檢視>排名作業>學期成績科目級別重複檢查與調整
            Catalog ribbon2 = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            ribbon2.Add(new RibbonFeature("5A1544EE-EC7B-4316-9B6C-C6C6B5022E35", "學期成績科目級別重複檢查與調整"));

            MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["學期成績科目級別重複檢查與調整"].Enable = UserAcl.Current["5A1544EE-EC7B-4316-9B6C-C6C6B5022E35"].Executable;

            MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["學期成績科目級別重複檢查與調整"].Click += delegate
            {
                // 學期成績科目級別重複檢查與調整
                frmSemsSubjectLevelDuplicate fss = new frmSemsSubjectLevelDuplicate();
                fss.ShowDialog();
            };

        }
    }
}
