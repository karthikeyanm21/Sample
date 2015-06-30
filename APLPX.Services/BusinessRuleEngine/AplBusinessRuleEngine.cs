using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web;

namespace APLPX.Services.BusinessRuleEngine
{
    public class AplBusinessRuleEngine
    {
        public string Sku { get; set; }
        public string Name { get; set; }
        public string Company { get; set; }
        public double Price { get; set; }
        public double Shipping { get; set; }
        public double InStock { get; set; }
        public DateTime CrawlDate { get; set; }
        public string ErrorMsg { get; set; }
    }

    internal class APLConfigProperties
    {
        public string Scope { get; set; }
        public string ValidationRegEx { get; set; }
        public string ErrorMsg { get; set; }
    }

    static class AplBRuleEngine
    {
        private static Dictionary<string, Func<object, object, bool>> s_operators;
        private static Dictionary<string, PropertyInfo> s_properties;


        static AplBRuleEngine()
        {
            s_operators = new Dictionary<string, Func<object, object, bool>>();
            s_operators["greater_than"] = new Func<object, object, bool>(s_opGreaterThan);
            s_operators["less_than"] = new Func<object, object, bool>(s_opLessThan);
            s_operators["equal"] = new Func<object, object, bool>(s_opEqual);
            s_operators["greater_than_equalto"] = new Func<object, object, bool>(s_opGreaterThanEqualTo);
            s_operators["less_than_equalto"] = new Func<object, object, bool>(s_opLessThanEqualTo);
            s_operators["notnull"] = new Func<object, object, bool>(s_notnull);
            s_properties = typeof(AplBusinessRuleEngine).GetProperties().ToDictionary(propInfo => propInfo.Name);
        }

        public static string Apply(AplBusinessRuleEngine extractionVal, string op, string prop, object target)
        {
            bool result = s_operators[op](GetPropValue(extractionVal, prop), target);
            return result ? "Processed" : extractionVal.ErrorMsg;
        }

        private static object GetPropValue(AplBusinessRuleEngine user, string prop)
        {
            PropertyInfo propInfo = s_properties[prop];
            return propInfo.GetGetMethod(false).Invoke(user, null);
        }

        #region Operators

        static bool s_opGreaterThan(object o1, object o2)
        {
            if (o1 == null || o2 == null || o1.GetType() != o2.GetType() || !(o1 is IComparable))
                return false;
            return (o1 as IComparable).CompareTo(o2) > 0;
        }
        static bool s_opGreaterThanEqualTo(object o1, object o2)
        {
            if (o1 == null || o2 == null || o1.GetType() != o2.GetType() || !(o1 is IComparable))
                return false;
            return (o1 as IComparable).CompareTo(o2) >= 0;
        }
        static bool s_opLessThan(object o1, object o2)
        {
            if (o1 == null || o2 == null || o1.GetType() != o2.GetType() || !(o1 is IComparable))
                return false;
            return (o1 as IComparable).CompareTo(o2) < 0;
        }
        static bool s_opLessThanEqualTo(object o1, object o2)
        {
            if (o1 == null || o2 == null || o1.GetType() != o2.GetType() || !(o1 is IComparable))
                return false;
            return (o1 as IComparable).CompareTo(o2) <= 0;
        }
        static bool s_opEqual(object o1, object o2)
        {
            return o1 == o2;
        }
        static bool s_notnull(object o1, object o2)
        {
            if (o1 == null || o2 == null ||o2.ToString()==""||o1.ToString()=="")
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        #endregion

    }
}