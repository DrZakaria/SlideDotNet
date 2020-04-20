using System.Linq;
using SlideDotNet.Enums;
using SlideDotNet.Models;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace SlideDotNet.Tests
{
    public class TestFile_025
    {
        [Fact]
        public void Chart_Test()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._025);
            var sld1 = pre.Slides[0];
            var chart8 = sld1.Shapes.First(x => x.Id == 8).Chart;
            var chart3 = sld1.Shapes.First(x => x.Id == 3).Chart;

            // Act
            var hasXValues = chart8.HasXValues;
            var d = chart3.HasCategories;

            // Assert
            Assert.False(hasXValues);
        }
    }
}
