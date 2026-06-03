// <copyright file="TestProgram.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Tests
{
    using System.Threading.Tasks;

    /// <summary>
    /// Minimal IProgram for tests: every new workbook gets an extra worksheet and
    /// every new worksheet is renamed to a sequential number ("1", "2", ...).
    /// All members complete synchronously.
    /// </summary>
    public class TestProgram : IProgram
    {
        private int counter;

        public IAddIn AddIn { get; private set; }

        public Task OnStart(IAddIn addIn)
        {
            this.AddIn = addIn;
            return Task.CompletedTask;
        }

        public Task OnStop() => Task.CompletedTask;

        public Task OnNew(IWorkbook workbook)
        {
            var sheet = workbook.AddWorksheet();

            // Deterministic content: the GetUsedRange test asserts a used range of
            // 50 rows by 15 columns on this sheet (rows 0-49, columns 0-14).
            sheet[0, 0].Value = "0.0";
            sheet[49, 14].Value = "49.14";
            sheet.Flush();

            return Task.CompletedTask;
        }

        public void OnClose(IWorkbook workbook, ref bool cancel)
        {
        }

        public Task OnNew(IWorksheet worksheet)
        {
            worksheet.Name = $"{++this.counter}";
            return Task.CompletedTask;
        }

        public Task OnBeforeDelete(IWorksheet worksheet) => Task.CompletedTask;

        public Task OnHandle(string handle, params object[] argument) => Task.CompletedTask;

        public Task OnLogin() => Task.CompletedTask;

        public Task OnLogout() => Task.CompletedTask;

        public bool IsEnabled(string controlId, string controlTag) => true;
    }
}
