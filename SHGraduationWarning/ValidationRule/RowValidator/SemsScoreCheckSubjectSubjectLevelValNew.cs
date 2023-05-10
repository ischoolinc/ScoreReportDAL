using Campus.DocumentValidator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHGraduationWarning.ValidationRule.RowValidator
{
    public class SemsScoreCheckSubjectSubjectLevelValNew : IRowVaildator
    {
        public SemsScoreCheckSubjectSubjectLevelValNew()
        {

        }

        #region IRowVaildator 成員

        public string Correct(IRowStream Value)
        {
            return string.Empty;
        }

        public string ToString(string template)
        {
            return template;
        }

        public bool Validate(IRowStream Value)
        {
            bool retVal = true;
            if (Value.Contains("學生系統編號") && Value.Contains("學年度") && Value.Contains("學期") && Value.Contains("成績年級") && Value.Contains("新科目名稱") && Value.Contains("新科目級別"))
            {
                string key = Value.GetValue("學生系統編號") + "_" + Value.GetValue("學年度") + "_" + Value.GetValue("學期") + "_" + Value.GetValue("成績年級") + "_" + Value.GetValue("新科目名稱") + "_" + Value.GetValue("新科目級別");

                if (Utility._StudentSemesScoreSubjectLevelTemp.ContainsKey(key))
                    retVal = false;
            }

            return retVal;
        }

        #endregion
    }
}
