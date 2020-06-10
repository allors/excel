// <copyright file="Cells.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using System.Collections.Generic;
using System.Linq;

namespace Allors.Excel.Interop
{
    public static class Cells
    {
        public static IEnumerable<IList<IList<Cell>>> Chunks(this IEnumerable<Cell> @this, Func<Cell, Cell, bool> combine)
        {
            return @this
                .GroupBy(v => v.Row)
                .SelectMany(v =>
                {
                    return v.OrderBy(w => w.Column).Aggregate(new List<IList<Cell>> { new List<Cell>() },
                        (acc, w) =>
                        {
                            var list = acc[acc.Count - 1];
                            if (list.Count == 0 || (list[list.Count - 1].Column.Index + 1 == w.Column.Index && combine(list[list.Count - 1], w)))
                            {
                                list.Add(w);
                            }
                            else
                            {
                                list = new List<Cell> { w };
                                acc.Add(list);
                            }

                            return acc;
                        });
                })
                .GroupBy(v => v[0].Column)
                .SelectMany(v =>
                {
                    return v.OrderBy(w => w[0].Row).Aggregate(
                        new List<IList<IList<Cell>>> { new List<IList<Cell>>() },
                        (acc, w) =>
                        {
                            var list = acc[acc.Count - 1];
                            if (list.Count == 0 || (list[list.Count - 1].Count == w.Count && list[list.Count - 1][0].Row.Index + 1 == w[0].Row.Index && combine(list[list.Count - 1][0], w[0])))
                            {
                                list.Add(w);
                            }
                            else
                            {
                                list = new List<IList<Cell>> { w };
                                acc.Add(list);
                            }

                            return acc;
                        });
                });
        }
    }
}
