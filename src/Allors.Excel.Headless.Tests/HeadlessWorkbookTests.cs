// <copyright file="HeadlessWorkbookTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Headless.Tests
{
    using System;
    using System.Threading.Tasks;
    using Allors.Excel;
    using Allors.Excel.Headless;
    using Allors.Excel.Tests;
    using Moq;
    using Xunit;

    public class HeadlessWorkbookTests
    {
        [Fact]
        public async Task AddWorksheetAtZeroWithNoActiveSheetDoesNotThrow()
        {
            var addIn = await AddIn.CreateAsync(new TestProgram(), new Mock<IRibbon>().Object);

            // A fresh workbook with no worksheets (hence none active): AddWorksheet(0) must
            // insert at the front, not call Insert(-1) (IndexOf of a null "active" sheet).
            var workbook = new Workbook(addIn);

            var worksheet = await workbook.AddWorksheet(0);

            Assert.Single(workbook.Worksheets);
            Assert.Same(worksheet, workbook.Worksheets[0]);
        }

        [Fact]
        public async Task AddWorksheetAwaitsOnNew()
        {
            // OnNew(worksheet) returns a gate that stays pending; AddWorksheet must not complete
            // until the gate is released. (A fire-and-forget OnNew would complete immediately.)
            Func<Task> onNewWorksheet = () => Task.CompletedTask;
            var program = new Mock<IProgram>();
            program.Setup(p => p.OnNew(It.IsAny<IWorksheet>())).Returns(() => onNewWorksheet());
            program.Setup(p => p.OnNew(It.IsAny<IWorkbook>())).Returns(Task.CompletedTask);
            program.Setup(p => p.OnStart(It.IsAny<IAddIn>())).Returns(Task.CompletedTask);

            var addIn = await AddIn.CreateAsync(program.Object, new Mock<IRibbon>().Object);
            var workbook = new Workbook(addIn);

            var gate = new TaskCompletionSource<bool>();
            onNewWorksheet = () => gate.Task;

            var addTask = workbook.AddWorksheet();
            Assert.False(addTask.IsCompleted);

            gate.SetResult(true);
            await addTask;

            Assert.True(addTask.IsCompleted);
        }
    }
}
