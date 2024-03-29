// <copyright file="Binding.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    using System;

    public class Binding : IBinding
    {
        private readonly Action<ICell> toCell;

        private readonly Action<ICell> toDomain;

        public Binding(Action<ICell> toCell = null, Action<ICell> toDomain = null)
        {
            this.toCell = toCell;
            this.toDomain = toDomain;
        }

        public bool OneWayBinding => this.toDomain == null;

        public bool TwoWayBinding => !this.OneWayBinding;

        public void ToCell(ICell cell) => this.toCell?.Invoke(cell);

        public void ToDomain(ICell cell) => this.toDomain?.Invoke(cell);
    }
}
