// <copyright file="Binder.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;

namespace Allors.Excel
{
    public class Binder
    {
        public IWorksheet Worksheet { get; }

        public event EventHandler ToDomained;

        private IDictionary<ICell, IBinding> bindingByCell = new ConcurrentDictionary<ICell, IBinding>();

        private IList<ICell> boundCells = new List<ICell>();

        private IList<ICell> bindingCells = new List<ICell>();

        public readonly Style changedStyle;

        private readonly IDictionary<ICell, Style> changedCells;

        public Binder(IWorksheet worksheet, Style changedStyle = null)
        {
            Worksheet = worksheet;
            Worksheet.CellsChanged += Worksheet_CellsChanged;
 
            this.changedStyle = changedStyle;
            if (this.changedStyle != null)
            {
                changedCells = new Dictionary<ICell, Style>();
            }
        }

        public void Set(int row, int column, IBinding binding)
        {
            Set(Worksheet[row, column], binding);
        }

        public void Set(ICell cell, IBinding binding)
        {
            bindingByCell[cell] = binding;
            bindingCells.Add(cell);
        }

        public ICell[] ToCells()
        {
            var obsoleteCells = boundCells.Except(bindingCells).ToArray();
            boundCells = bindingCells;
            bindingCells = new List<ICell>();

            foreach (var obsoleteCell in obsoleteCells)
            {
                bindingByCell.Remove(obsoleteCell);
            }

            foreach (var kvp in bindingByCell)
            {
                var cell = kvp.Key;
                var binding = kvp.Value;
                binding.ToCell(cell);
            }

            return obsoleteCells;
        }

        private void Worksheet_CellsChanged(object sender, CellChangedEvent e)
        {
            foreach (var cell in e.Cells)
            {
                if (bindingByCell.TryGetValue(cell, out var binding))
                {
                    if (binding.TwoWayBinding)
                    {
                        binding.ToDomain(cell);

                        if (changedStyle != null)
                        {
                            if (!changedCells.ContainsKey(cell))
                            {
                                changedCells.Add(cell, cell.Style);
                            }
                            cell.Style = changedStyle;
                        }
                    }
                    else
                    {
                        binding.ToCell(cell);
                    }
                }
            }

            ToDomained?.Invoke(this, EventArgs.Empty);
        }

        public void ResetChangedCells()
        {
            if (changedStyle != null)
            {
                foreach (var kvp in changedCells)
                {
                    var cell = kvp.Key;
                    var style = kvp.Value;
                    cell.Style = style;
                }

                changedCells.Clear();
            }
        }

        public bool ExistBinding(int row, int column)
        {
            return bindingCells.Any(v => v.Row.Index == row && v.Column.Index == column);
        }
    }
}