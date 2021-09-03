using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHCourseGroupCodeDAL
{
    public class Utility
    {
        /// <summary>
        /// 取得課程類別代碼內容對照表
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetCourseTypeMappingTable()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("1", "部定必修");
            value.Add("2", "校訂必修");
            value.Add("3", "選修-加深加廣選修");
            value.Add("4", "選修-補強性選修");
            value.Add("5", "選修-多元選修");
            value.Add("6", "選修-其他");
            value.Add("7", "校訂選修");
            value.Add("8", "團體活動時間");
            value.Add("9", "彈性學習時間");
            value.Add("A", "大學預修課程");
            value.Add("B", "基礎訓練");
            value.Add("C", "職前訓練");
            value.Add("D", "寒暑假課程");
            value.Add("E", "返校課程");
            value.Add("F", "職業技能訓練");

            return value;
        }

        public static Dictionary<string, string> GetSubjectTypeMappingTable()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            value.Add("0", "不分屬性");
            value.Add("1", "一般科目");
            value.Add("2", "專業科目");
            value.Add("3", "實習科目");
            value.Add("4", "專精科目");
            value.Add("5", "專精科目(核心科目)");
            value.Add("6", "特殊需求領域");
            value.Add("A", "自主學習");
            value.Add("B", "選手培訓");
            value.Add("C", "充實(增廣)、補強性教學 [全學期、不授予學分]");
            value.Add("D", "充實(增廣)、補強性教學 [全學期、授予學分]");
            value.Add("E", "學校特色活動");
            return value;
        }
    }
}
