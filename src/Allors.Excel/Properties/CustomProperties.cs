using System;
using System.Collections.Generic;
using System.Text;

namespace Allors.Excel
{

    /// <summary>
    /// Gets or sets the custom properties inside the Sheet.
    /// Does not support class instances or lists as values!
    /// </summary>
    public class CustomProperties : Dictionary<string, object>
    {
        public const string MagicNull = "{DDF8EA1F-C9D5-4A31-B05D-998A9E539D42}";

        public new void Add(string key, object value)
        {
            if (this.ContainsKey(key))
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
                else
                {
                    return (T)Convert.ChangeType(value, typeof(T));
                }
            }
            catch (FormatException)
            {
                // bool seems to be stored as "0" false, and "-1" true
                // Converting "0" to bool throws FormatException
                if("0".Equals(value))
                {
                    return (T)Convert.ChangeType(false, typeof(bool));
                }
                else
                {
                    return (T) Convert.ChangeType(true, typeof(bool));
                }
            }
            catch (InvalidCastException)
            {
                return default;
            }
        }
    }
}
