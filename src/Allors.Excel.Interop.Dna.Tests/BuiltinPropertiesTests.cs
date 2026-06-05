// <copyright file="BuiltinPropertiesTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop.Dna.Tests
{
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using Xunit;

    public class BuiltinPropertiesTests
    {
        // Mimics an Excel BuiltinDocumentProperty: it exists (has a Name) but reading Value
        // throws a COMException when the property has not been populated.
        public sealed class FakeProperty
        {
            private readonly object value;
            private readonly bool throwsOnValue;

            public FakeProperty(string name, object value, bool throwsOnValue = false)
            {
                this.Name = name;
                this.value = value;
                this.throwsOnValue = throwsOnValue;
            }

            public string Name { get; }

            public object Value => this.throwsOnValue
                ? throw new COMException("Builtin property has no value.")
                : this.value;
        }

        [Fact]
        public void GetStringReturnsNullForUnsetBuiltinProperty()
        {
            var properties = new List<FakeProperty> { new FakeProperty("Manager", null, throwsOnValue: true) };
            var builtin = new BuiltinProperties(properties);

            Assert.Null(builtin.Manager);
        }

        [Fact]
        public void GetNumberReturnsNullForUnsetBuiltinProperty()
        {
            var properties = new List<FakeProperty> { new FakeProperty("Number of pages", null, throwsOnValue: true) };
            var builtin = new BuiltinProperties(properties);

            Assert.Null(builtin.NumberOfPages);
        }

        [Fact]
        public void GetDateReturnsNullForUnsetBuiltinProperty()
        {
            var properties = new List<FakeProperty> { new FakeProperty("Last print date", null, throwsOnValue: true) };
            var builtin = new BuiltinProperties(properties);

            Assert.Null(builtin.LastPrintDate);
        }

        [Fact]
        public void GetterReturnsValueWhenSet()
        {
            var properties = new List<FakeProperty> { new FakeProperty("Title", "MyTitle") };
            var builtin = new BuiltinProperties(properties);

            Assert.Equal("MyTitle", builtin.Title);
        }
    }
}
