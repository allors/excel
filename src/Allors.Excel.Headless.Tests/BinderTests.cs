// <copyright file="BinderTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless.Tests
{
    using Allors.Excel;
    using Moq;
    using Xunit;

    public class BinderTests
    {
        [Fact]
        public void DisposeUnsubscribesFromCellsChanged()
        {
            var worksheet = new Mock<IWorksheet>();
            var binder = new Binder(worksheet.Object);

            var count = 0;
            binder.ToDomained += (sender, e) => count++;

            var args = new CellChangedEvent(new ICell[0]);

            worksheet.Raise(w => w.CellsChanged += null, worksheet.Object, args);
            Assert.Equal(1, count);

            binder.Dispose();

            // After Dispose the Binder must no longer react to the worksheet's event.
            worksheet.Raise(w => w.CellsChanged += null, worksheet.Object, args);
            Assert.Equal(1, count);
        }
    }
}
