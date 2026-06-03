// <copyright file="InteropWorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop.Tests.Shared
{
    /// <summary>
    /// Base for the interop WorkbookTests suites: runs the shared workbook facts
    /// against a real Excel instance hosted by <see cref="InteropExcelFixture"/>.
    /// </summary>
    public abstract class InteropWorkbookTests : Allors.Excel.Tests.WorkbookTests
    {
        private readonly InteropExcelFixture fixture;

        protected InteropWorkbookTests()
        {
            // Created here, not in a field initializer, so Excel starts after the
            // base constructor has run (field initializers precede base constructors).
            this.fixture = new InteropExcelFixture();
        }

        public override void Dispose() => this.fixture.Dispose();

        protected override IAddIn NewAddIn() => this.fixture.NewAddIn();

        protected override void AddWorkbook() => this.fixture.AddWorkbook();
    }
}
