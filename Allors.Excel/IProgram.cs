// <copyright file="IProgram.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    using System.Threading.Tasks;

    public interface IProgram
    {
        Task OnStart(IAddIn addIn);

        Task OnStop();

        Task OnNew(IWorkbook workbook);

        void OnClose(IWorkbook workbook, ref bool cancel);

        Task OnNew(IWorksheet worksheet);

        Task OnBeforeDelete(IWorksheet worksheet);

        Task OnHandle(string handle, params object[] argument);

        Task OnLogin();

        Task OnLogout();

        bool IsEnabled(string controlId, string controlTag);
    }
}
