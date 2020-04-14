// <copyright file="IBinding.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    public interface IBinding
    {
        void ToCell(ICell cell);

        void ToDomain(ICell cell);

        bool OneWayBinding { get; }

        bool TwoWayBinding { get; }
    }
}