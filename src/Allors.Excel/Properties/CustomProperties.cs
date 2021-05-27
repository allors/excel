using System;
using System.Collections.Generic;

namespace Allors.Excel
{

    /// <summary>
    /// Gets or sets the custom properties inside the Sheet.
    /// Does not support class instances or lists as values!
    /// </summary>
    public class CustomProperties : Dictionary<string, object>
    {
        public const string MagicNull = "{DDF8EA1F-C9D5-4A31-B05D-998A9E539D42}";
        public const string MagicDecimalMaxValue = "{0DADAEFB-3770-489A-A88F-ACD2D0D722E6}";
        public const string MagicDecimalMinValue = "{F52899DF-34E5-4C6D-BC44-E2928B4F41BE}";

        public const string MagicDateTimeMaxValue = "{291674F8-34A8-48D8-98B5-CA717EA67030}";
        public const string MagicDateTimeMinValue = "{1673E121-DCED-48AE-B7D2-908773CBE849}";

        public new void Add(string key, object value)
        {
            if (ContainsKey(key))
            {
                base[key] = value;
            }
            else
            {
                base.Add(key, value);
            }
        }

        public T Get<T>(string key)
        {
            var value = this[key];

            // From double to decimal, we will loose precision when precision is more than 6.
            if (value is double && typeof(T) == typeof(decimal))
            {
                var result = Math.Round(Convert.ToDecimal(value), 6);
                return (T)Convert.ChangeType(result, typeof(T));
            }
         
            if (value is T)
            {
                return (T)value;
            }

            if(value == null)
            {
                return default;
            }

            try
            {
                // Nullable types must be converted to their underlying type.
                var U = Nullable.GetUnderlyingType(typeof(T));
                if (U != null)
                {                   
                    var uV = Convert.ChangeType(value, U);

                    return (T) uV;
                }

                return (T)Convert.ChangeType(value, typeof(T));
            }
            catch (FormatException)
            {
                // bool seems to be stored as "0" false, and "-1" true
                // Converting "0" to bool throws FormatException
                if("0".Equals(value))
                {
                    return (T)Convert.ChangeType(false, typeof(bool));
                }

                return (T) Convert.ChangeType(true, typeof(bool));
            }
            catch (InvalidCastException)
            {
                return default;
            }
        }
    }
}
