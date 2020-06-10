// <copyright file="InteropRangeExtension.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System;
using InteropRange = Microsoft.Office.Interop.Excel.Range;

namespace Allors.Excel.Interop
{
    public static class InteropRangeExtension
    {
        public static (int, int) Coordinates(this InteropRange @this) => (@this.Row - 1, @this.Column - 1);
    }
}
