namespace Allors.Excel
{
    using System;
    using Microsoft.Office.Core;

    public class CustomProperties(object properties) : PropertiesBase(properties), ICustomProperties
    {
        public void SetBoolean(string key, bool? value) => this.Set(key, MsoDocProperties.msoPropertyTypeBoolean, value);

        public void SetDate(string key, DateTime? value) => this.Set(key, MsoDocProperties.msoPropertyTypeDate, value);

        public void SetFloat(string key, double? value) => this.Set(key, MsoDocProperties.msoPropertyTypeFloat, value);

        public void SetNumber(string key, long? value) => this.Set(key, MsoDocProperties.msoPropertyTypeNumber, value);

        public void SetString(string key, string? value) => this.Set(key, MsoDocProperties.msoPropertyTypeString, value);

        private void Set(string key, MsoDocProperties type, object? value)
        {
            if (value == null)
            {
                this.Remove(key);
                return;
            }

            var property = this.Get(key);
            if (property != null)
            {
                property.Value = value;
                return;
            }

            this.Properties.Add(key, false, type, value, null);
        }
    }
}
