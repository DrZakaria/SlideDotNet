﻿using System.Linq;
using SlideDotNet.Spreadsheet;
using Xunit;

// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Tests
{
    /// <summary>
    /// Contains test for <see cref="CellFormulaParser"/> class.
    /// </summary>
    public class CellFormulaTest
    {
        [Fact]
        public void GetCellAddresses_Test()
        {
            // Arrange
            var testExpression = "B10:B12";
            var cellRangeParser = new CellFormulaParser(testExpression);
            var testExpression2 = "B10:B12,B14";
            var cellRangeParser2 = new CellFormulaParser(testExpression2);
            var testExpression3 = "B10";
            var cellRangeParser3 = new CellFormulaParser(testExpression3);

            // Act
            var result = cellRangeParser.GetCellAddresses();
            var result2 = cellRangeParser2.GetCellAddresses();
            var result3 = cellRangeParser3.GetCellAddresses();

            // Assert
            Assert.Equal(@"B10", result[0]);
            Assert.Equal(@"B11", result[1]);
            Assert.Equal(@"B12", result[2]);
            Assert.Equal(@"B14", result2[3]);
            Assert.Equal(@"B10", result3.Single());
        }
    }
}
