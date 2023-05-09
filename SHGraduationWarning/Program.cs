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
using Campus.DocumentValidator;
using static System.Windows.Forms.LinkLabel;

namespace SHGraduationWarning
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void main()
        {
            // 教務作業>批次作業/檢視>成績作業>畢業預警
            Catalog ribbon1 = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            ribbon1.Add(new RibbonFeature("6460BE9A-3E82-43C7-ACE7-64E0C6100473", "畢業預警與資料合理檢查"));

            MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["畢業預警與資料合理檢查"].Enable = UserAcl.Current["6460BE9A-3E82-43C7-ACE7-64E0C6100473"].Executable;

            MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["畢業預警與資料合理檢查"].Click += delegate
            {
                // 畢業預警與資料合理檢查
                frmMain fm = new frmMain();
                fm.ShowDialog();
            };

            Catalog ribbon2 = RoleAclSource.Instance["學生"]["功能按鈕"];
            ribbon2.Add(new RibbonFeature("3DE8EA91-A176-4100-8401-1FBA9F44786A", "匯入更新學期科目級別"));


            MotherForm.RibbonBarItems["學生", "資料統計"]["匯入"]["成績相關匯入"]["匯入更新學期科目級別"].Enable = FISCA.Permission.UserAcl.Current["3DE8EA91-A176-4100-8401-1FBA9F44786A"].Executable;
            MotherForm.RibbonBarItems["學生", "資料統計"]["匯入"]["成績相關匯入"]["匯入更新學期科目級別"].Click += delegate
            {
                // 載入所有學生與狀態，資料匯入比對使用
                Utility._AllStudentNumberStatusIDTemp = Utility.GetAllStudenNumberStatusDict();
                ImportExport.ImportUpdateSubjectLevel importUpdateSubjectLevel = new ImportExport.ImportUpdateSubjectLevel();
                importUpdateSubjectLevel.Execute();
            };

            // 載入自訂驗證規則
            #region 自訂驗證規則
            FactoryProvider.RowFactory.Add(new ValidationRule.SemsScoreRowValidatorFactory());
            #endregion
        }
    }
}
