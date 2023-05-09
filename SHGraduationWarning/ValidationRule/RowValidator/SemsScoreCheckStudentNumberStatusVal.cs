using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Campus.DocumentValidator;

namespace SHGraduationWarning.ValidationRule.RowValidator
{
    public class SemsScoreCheckStudentNumberStatusVal : IRowVaildator
    {
        public SemsScoreCheckStudentNumberStatusVal()
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
            bool retVal = false;
            if (Value.Contains("學號") && Value.Contains("狀態"))
            {
                string key = Value.GetValue("學號") + "_" + Value.GetValue("狀態").Trim();

                if (Utility._AllStudentNumberStatusIDTemp.ContainsKey(key))
                    retVal = true;
            }

            return retVal;
        }

        #endregion
    }
}
