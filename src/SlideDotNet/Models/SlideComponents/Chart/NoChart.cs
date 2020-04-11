using System;
using SlideDotNet.Collections;
using SlideDotNet.Enums;
using SlideDotNet.Exceptions;

namespace SlideDotNet.Models.SlideComponents.Chart
{
    public class NoChart : IChart
    {
        public ChartType Type => throw new NotSupportedException(ExceptionMessages.NoChart);

        public string Title => throw new NotSupportedException(ExceptionMessages.NoChart);

        public bool HasTitle => throw new NotSupportedException(ExceptionMessages.NoChart);

        public SeriesCollection SeriesCollection => throw new NotSupportedException(ExceptionMessages.NoChart);

        public CategoryCollection Categories => throw new NotSupportedException(ExceptionMessages.NoChart);

        public bool HasCategories => throw new NotSupportedException(ExceptionMessages.NoChart);
    }
}