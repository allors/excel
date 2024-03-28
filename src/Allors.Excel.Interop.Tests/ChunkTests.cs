// <copyright file="ChunkTests.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Allors.Excel.Interop.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Allors.Excel.Interop;
    using Moq;
    using Xunit;

    public class ChunkTests
    {
        private readonly Dictionary<int, Row> rowByIndex = new Dictionary<int, Row>();

        private readonly Dictionary<int, Column> columnByIndex = new Dictionary<int, Column>();

        public Row Row(int index)
        {
            if (!this.rowByIndex.TryGetValue(index, out var row))
            {
                row = new Row(null, index);
                this.rowByIndex.Add(index, row);
            }

            return row;
        }

        public Column Column(int index)
        {
            if (!this.columnByIndex.TryGetValue(index, out var column))
            {
                column = new Column(null, index);
                this.columnByIndex.Add(index, column);
            }

            return column;
        }

        [Fact]
        public void OneChunk_OneRow_TwoCells()
        {
            var cells = new[]
            {
                new Cell(null, this.Row(0), this.Column(0)),
                new Cell(null, this.Row(0), this.Column(1)),
            };

            var chunks = cells.Chunks((v, w) => true).ToArray();
            Assert.Single(chunks);

            chunks = cells.Chunks((v, w) => false).ToArray();
            Assert.Equal(2, chunks.Length);
        }

        [Fact]
        public void OneChunk_OneRow_FourCells()
        {
            var cells = new[]
                {
                    new Cell(null, this.Row(0), this.Column(0)),
                    new Cell(null, this.Row(0), this.Column(1)),
                    new Cell(null, this.Row(0), this.Column(2)),
                    new Cell(null, this.Row(0), this.Column(3)),
            };

            var chunks = cells.Chunks((v, w) => true).ToArray();
            Assert.Single(chunks);
        }

        [Fact]
        public void OneChunk_TwoRows_OneCell()
        {
            var cells = new[]
                {
                    new Cell(null, this.Row(0), this.Column(0)),
                    new Cell(null, this.Row(1), this.Column(0)),
            };

            var chunks = cells.Chunks((v, w) => true).ToArray();
            Assert.Single(chunks);
        }

        [Fact]
        public void OneChunk_TwoRows_TwoCells()
        {
            var cells = new[]
                {
                    new Cell(null, this.Row(0), this.Column(0)),
                    new Cell(null, this.Row(0), this.Column(1)),
                    new Cell(null, this.Row(1), this.Column(0)),
                    new Cell(null, this.Row(1), this.Column(1)),
            };

            var chunks = cells.Chunks((v, w) => true).ToArray();
            Assert.Single(chunks);
        }

        [Fact]
        public void TwoChunks_OneRow_TwoCells()
        {
            var cells = new[]
                {
                    new Cell(null, this.Row(0), this.Column(0)),
                    new Cell(null, this.Row(0), this.Column(1)),
                    new Cell(null, this.Row(0), this.Column(3)),
                    new Cell(null, this.Row(0), this.Column(4)),
            };

            var chunks = cells.Chunks((v, w) => true).ToArray();
            Assert.Equal(2, chunks.Length);
        }

        [Fact]
        public void Square()
        {
            var raster = new[]
            {
                "###",
                "# #",
                "###",
            };

            var worksheet = new Mock<Excel.Interop.IWorksheet>().Object;
            var cells = this.CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(5, chunks.Length);
        }

        [Fact]
        public void Cross()
        {
            var raster = new[]
            {
                "# #",
                " # ",
                "# #",
            };

            var worksheet = new Mock<Excel.Interop.IWorksheet>().Object;
            var cells = this.CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(9, chunks.Length);
        }

        [Fact]
        public void HorizontalLines()
        {
            var raster = new[]
            {
                "###",
                "   ",
                "###",
            };

            var worksheet = new Mock<Excel.Interop.IWorksheet>().Object;
            var cells = this.CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(3, chunks.Length);
        }

        [Fact]
        public void VerticalLines()
        {
            var raster = new[]
            {
                "# #",
                "# #",
                "# #",
            };

            var worksheet = new Mock<Excel.Interop.IWorksheet>().Object;
            var cells = this.CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(3, chunks.Length);
        }

        [Fact]
        public void LShape()
        {
            var raster = new[]
            {
                "#  ",
                "#  ",
                "###",
            };

            var worksheet = new Mock<Excel.Interop.IWorksheet>().Object;
            var cells = this.CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(3, chunks.Length);
        }

        [Fact]
        public void ReverseLShape()
        {
            var raster = new[]
            {
                "  #",
                "  #",
                "###",
            };

            var worksheet = new Mock<Excel.Interop.IWorksheet>().Object;
            var cells = this.CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(3, chunks.Length);
        }

        private IList<Cell> CellsFromRaster(Excel.Interop.IWorksheet worksheet, string[] raster, Action<ICell, string> setup)
        {
            var cells = new List<Cell>();
            for (var i = 0; i < raster.Length; i++)
            {
                var line = raster[i];
                for (var j = 0; j < 3; j++)
                {
                    var cell = new Cell(worksheet, this.Row(i), this.Column(j));
                    setup(cell, line[j].ToString());
                    cells.Add(cell);
                }
            }

            return cells;
        }
    }
}
