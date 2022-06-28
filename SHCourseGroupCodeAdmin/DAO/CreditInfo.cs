using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace SHCourseGroupCodeAdmin.DAO
{
    /// <summary>
    /// 學分值與樣式使用
    /// </summary>
    public class CreditInfo
    {
        /// <summary>
        /// 學分數字串
        /// </summary>
        public string StringValue = "";
        /// <summary>
        ///  是對開
        /// </summary>
        public bool isOpenD = false;
        /// <summary>
        /// 是使用這設定對開(Y/N/null 未設定)
        /// </summary>
        public bool? isSetOpenD;

        public Color BackgroundColor = Color.White;
    }
}
