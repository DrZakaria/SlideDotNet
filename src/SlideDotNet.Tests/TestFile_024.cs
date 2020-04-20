﻿using System.Linq;
using SlideDotNet.Enums;
using SlideDotNet.Models;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace SlideDotNet.Tests
{
    public class TestFile_024
    {
        [Fact]
        public void Slide_Shapes_Test()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._024);
            var sld2 = pre.Slides[1];
            var chart = sld2.Shapes.First(x => x.Id == 5).Chart;

            // Act
            var hasXValues = chart.HasXValues;
            var XValue = chart.XValues[0];

            // Assert
            var shapes = pre.Slides[0].Shapes; //should not throw exception
            Assert.True(hasXValues);
            Assert.Equal(10, XValue);
        }
    }
}
