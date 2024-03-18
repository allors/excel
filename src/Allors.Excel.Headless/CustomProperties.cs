// <copyright file="Workbook.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless
{
    using System;
    using System.Collections.Generic;
    using Allors.Excel;

    internal class CustomProperties : ICustomProperties
    {
        private Dictionary<string, object> properties = new();

        public bool? GetBoolean(string name)
        {
            _ = this.properties.TryGetValue(name, out var value);
            return value as bool?;
        }

        public DateTime? GetDate(string name)
        {
            _ = this.properties.TryGetValue(name, out var value);
            return value as DateTime?;
        }

        public double? GetFloat(string name)
        {
            _ = this.properties.TryGetValue(name, out var value);
            return value as double?;
        }

        public long? GetNumber(string name)
        {
            _ = this.properties.TryGetValue(name, out var value);
            return value as long?;
        }

        public string GetString(string name)
        {
            _ = this.properties.TryGetValue(name, out var value);
            return value as string;
        }

        public void SetBoolean(string name, bool? value) => this.properties[name] = value;

        public void SetDate(string name, DateTime? value) => this.properties[name] = value;

        public void SetFloat(string name, double? value) => this.properties[name] = value;

        public void SetNumber(string name, long? value) => this.properties[name] = value;

        public void SetString(string name, string value) => this.properties[name] = value;

        public void Remove(string name) => this.properties.Remove(name);
    }
}
