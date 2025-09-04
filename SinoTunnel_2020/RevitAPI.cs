using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace SinoTunnel
{
    public class RevitAPI
    {
        /// <summary>
        /// ElementId轉換為數字
        /// </summary>
        /// <param name="elemId"></param>
        /// <returns></returns>
        public static int GetValue(ElementId elemId)
        {
            return elemId.IntegerValue; // 2020
            //return ((int)elemId.Value); // 2024
        }
        /// <summary>
        /// 取得TaggedLocalElementId
        /// </summary>
        /// <param name="independentTags"></param>
        /// <returns></returns>
        public static List<ElementId> GetTaggedLocalElementId(List<IndependentTag> independentTags)
        {
            return independentTags.Select(x => x.TaggedLocalElementId).Distinct().ToList(); // 2020
            //return independentTags.Select(x => x.GetTaggedLocalElementIds().FirstOrDefault()).Distinct().ToList(); // 2022
        }
        /// <summary>
        /// 於識別資料中建立參數
        /// </summary>
        /// <param name="familyManager"></param>
        /// <returns></returns>
        public static FamilyParameter FamilyPara(FamilyManager familyManager)
        {
            return familyManager.AddParameter("指標ID", BuiltInParameterGroup.PG_IDENTITY_DATA, ParameterType.Text, true); // 2020
            //return familyManager.AddParameter("指標ID", GroupTypeId.IdentityData, SpecTypeId.String.Text, true); // 2022
        }
        /// <summary>
        /// 轉換單位
        /// </summary>
        /// <param name="number"></param>
        /// <param name="unit"></param>
        /// <returns></returns>        
        public static double ConvertFromInternalUnits(double number, string unit)
        {
            //if (unit.Equals("meters"))
            //{
            //    number = UnitUtils.ConvertFromInternalUnits(number, DisplayUnitType.DUT_METERS); // 2020
            //    //number = UnitUtils.ConvertFromInternalUnits(number, UnitTypeId.Meters); // 2022
            //}
            //else if (unit.Equals("millimeters"))
            //{
            //    number = UnitUtils.ConvertFromInternalUnits(number, DisplayUnitType.DUT_MILLIMETERS); // 2020
            //    //number = UnitUtils.ConvertFromInternalUnits(number, UnitTypeId.Millimeters); // 2022
            //}
            return number;
        }
        /// <summary>
        /// 轉換單位
        /// </summary>
        /// <param name="number"></param>
        /// <param name="unit"></param>
        /// <returns></returns>
        public static double ConvertToInternalUnits(double number, string unit)
        {
            //if (unit.Equals("meters"))
            //{
            //    number = UnitUtils.ConvertToInternalUnits(number, DisplayUnitType.DUT_METERS); // 2020
            //    //number = UnitUtils.ConvertToInternalUnits(number, UnitTypeId.Meters); // 2022
            //}
            //else if (unit.Equals("millimeters"))
            //{
            //    number = UnitUtils.ConvertToInternalUnits(number, DisplayUnitType.DUT_MILLIMETERS); // 2020
            //    //number = UnitUtils.ConvertToInternalUnits(number, UnitTypeId.Millimeters); // 2022
            //}
            return number;
        }

        ///// <summary>
        ///// 2022
        ///// </summary>
        ///// <param name="elemId"></param>
        ///// <returns></returns>
        //public static List<ElementId> GetTaggedLocalElementId(List<IndependentTag> independentTags)
        //{
        //    return independentTags.Select(x => x.GetTaggedLocalElementIds().FirstOrDefault()).Distinct().ToList(); // 2022
        //}
        //public static FamilyParameter FamilyPara(FamilyManager familyManager)
        //{
        //    return familyManager.AddParameter("指標ID", GroupTypeId.IdentityData, SpecTypeId.String.Text, true); // 2022
        //}
        //public static double ConvertFromInternalUnits(double number, string unit)
        //{
        //    if (unit.Equals("meters"))
        //    {
        //        number = UnitUtils.ConvertFromInternalUnits(number, UnitTypeId.Meters); // 2022
        //    }
        //    else if (unit.Equals("millimeters"))
        //    {
        //        number = UnitUtils.ConvertFromInternalUnits(number, UnitTypeId.Millimeters); // 2022
        //    }
        //    return number;
        //}
        //public static double ConvertToInternalUnits(double number, string unit)
        //{
        //    if (unit.Equals("meters"))
        //    {
        //        number = UnitUtils.ConvertToInternalUnits(number, UnitTypeId.Meters); // 2022
        //    }
        //    else if (unit.Equals("millimeters"))
        //    {
        //        number = UnitUtils.ConvertToInternalUnits(number, UnitTypeId.Millimeters); // 2022
        //    }
        //    return number;
        //}
    }
}