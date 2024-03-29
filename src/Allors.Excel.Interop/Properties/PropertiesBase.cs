﻿namespace Allors.Excel.Interop
{
    using System;

    public abstract class PropertiesBase(dynamic properties)
    {
        protected dynamic Properties { get; } = properties;

        public bool? GetBoolean(string key)
        {
            var value = this.Get(key)?.Value;
            return value != null ? Convert.ToBoolean(value) : null;
        }

        public DateTime? GetDate(string key)
        {
            var value = this.Get(key)?.Value;
            return value != null ? Convert.ToDateTime(value) : null;
        }

        public double? GetFloat(string key)
        {
            var value = this.Get(key)?.Value;
            return value != null ? Convert.ToDouble(value) : null;
        }

        public long? GetNumber(string key)
        {
            var value = this.Get(key)?.Value;
            return value != null ? Convert.ToInt64(value) : null;
        }

        public string GetString(string key) => Convert.ToString(this.Get(key)?.Value);

        public void Remove(string key) => this.Get(key)?.Delete();

        internal bool Exist(string key) => this.Get(key) != null;

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
