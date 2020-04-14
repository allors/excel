// <copyright file="Style.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

using System.Collections.Generic;
using System.Drawing;

namespace Allors.Excel
{
    public class Style
    {
        public Style(Color backgroundColor, Color textColor)
        {
            BackgroundColor = backgroundColor;
            TextColor = textColor;
        }

        public Color BackgroundColor { get; }

        public Color TextColor { get; }

        public override bool Equals(object obj)
        {
            var that = obj as Style;
            return this.BackgroundColor == that?.BackgroundColor && this.TextColor == that?.TextColor;
        }

        public override int GetHashCode()
        {
            return this.BackgroundColor.GetHashCode() + this.TextColor.GetHashCode();
        }
    }
}
