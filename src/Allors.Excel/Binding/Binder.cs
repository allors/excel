// <copyright file="Binder.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq;

    public class Binder
    {
        private readonly IDictionary<ICell, IBinding> bindingByCell = new ConcurrentDictionary<ICell, IBinding>();

        private readonly IDictionary<ICell, Style> changedCells;

        private IList<ICell> boundCells = new List<ICell>();

        private IList<ICell> bindingCells = new List<ICell>();

        public Binder(IWorksheet worksheet, Style changedStyle = null)
        {
            this.Worksheet = worksheet;
            this.Worksheet.CellsChanged += this.Worksheet_CellsChanged;

            this.ChangedStyle = changedStyle;
            if (this.ChangedStyle != null)
            {
                this.changedCells = new Dictionary<ICell, Style>();
            }
        }

        public event EventHandler ToDomained;

        public IWorksheet Worksheet { get; }

        public Style ChangedStyle { get; }

        public void Set(int row, int column, IBinding binding) => this.Set(this.Worksheet[row, column], binding);

        public void Set(ICell cell, IBinding binding)
        {
            this.bindingByCell[cell] = binding;
            this.bindingCells.Add(cell);
        }

        public ICell[] ToCells()
        {
            var obsoleteCells = this.boundCells.Except(this.bindingCells).ToArray();
            this.boundCells = this.bindingCells;
            this.bindingCells = new List<ICell>();

            foreach (var obsoleteCell in obsoleteCells)
            {
                this.bindingByCell.Remove(obsoleteCell);
            }

            foreach (var kvp in this.bindingByCell)
            {
                var cell = kvp.Key;
                var binding = kvp.Value;
                binding.ToCell(cell);
            }

            return obsoleteCells;
        }

        public void ResetChangedCells()
        {
            if (this.changedCells != null)
            {
                foreach (var kvp in this.changedCells)
                {
                    var cell = kvp.Key;
                    var style = kvp.Value;
                    cell.Style = style;
                }

                this.changedCells.Clear();
            }
        }

        public bool ExistBinding(int row, int column) => this.bindingCells.Any(v => v.Row.Index == row && v.Column.Index == column);

        private void Worksheet_CellsChanged(object sender, CellChangedEvent e)
        {
            foreach (var cell in e.Cells)
            {
                if (this.bindingByCell.TryGetValue(cell, out var binding))
                {
                    if (binding.TwoWayBinding)
                    {
                        binding.ToDomain(cell);

                        if (this.ChangedStyle != null && this.changedCells != null)
                        {
                            if (!this.changedCells.ContainsKey(cell))
                            {
                                this.changedCells.Add(cell, cell.Style);
                            }

                            cell.Style = this.ChangedStyle;
                        }
                    }
                    else
                    {
                        binding.ToCell(cell);
                    }
                }
            }

            this.ToDomained?.Invoke(this, EventArgs.Empty);
        }
    }
}
