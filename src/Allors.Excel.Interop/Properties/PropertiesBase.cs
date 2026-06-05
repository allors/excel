namespace Allors.Excel.Interop
{
    using System;
    using System.Runtime.InteropServices;

    public abstract class PropertiesBase(dynamic properties)
    {
        protected dynamic Properties { get; } = properties;

        public bool? GetBoolean(string key)
        {
            var value = this.GetValue(key);
            return value != null ? Convert.ToBoolean(value) : null;
        }

        public DateTime? GetDate(string key)
        {
            var value = this.GetValue(key);
            return value != null ? Convert.ToDateTime(value) : null;
        }

        public double? GetFloat(string key)
        {
            var value = this.GetValue(key);
            return value != null ? Convert.ToDouble(value) : null;
        }

        public long? GetNumber(string key)
        {
            var value = this.GetValue(key);
            return value != null ? Convert.ToInt64(value) : null;
        }

        public string GetString(string key)
        {
            var value = this.GetValue(key);

            // null (absent or unset property) stays null; a present value is coerced to text.
            return value == null ? null : Convert.ToString(value);
        }

        public void Remove(string key) => this.Get(key)?.Delete();

        internal bool Exist(string key) => this.Get(key) != null;

        // A builtin document property can exist by name yet be unpopulated; reading its
        // Value then throws a COMException. Treat that (and a missing property) as "no value".
        private object GetValue(string key)
        {
            var property = this.Get(key);
            if (property == null)
            {
                return null;
            }

            try
            {
                return property.Value;
            }
            catch (COMException)
            {
                return null;
            }
        }

        protected dynamic? Get(string key)
        {
            foreach (var property in this.Properties)
            {
                if (Equals(property.Name, key))
                {
                    return property;
                }
            }

            return null;
        }
    }
}
