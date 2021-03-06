﻿using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Enums;
using SlideDotNet.Extensions;
using SlideDotNet.Services.Placeholders;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleNullReferenceException

namespace SlideDotNet.Services
{
    /// <summary>
    /// Represents a font size manager for placeholder elements
    /// </summary>
    public class PlaceholderFontService
    {
        private readonly SlidePart _sdkSldPart;
        private readonly Lazy<HashSet<PlaceholderFontData>> _layoutPlaceholders;
        private readonly Lazy<HashSet<PlaceholderFontData>> _masterPlaceholders;
        private readonly Lazy<Dictionary<int, int>> _masterBodyFontHeights;

        public PlaceholderFontService(SlidePart sdkSldPart)
        {
            _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));

            var layoutSldData = _sdkSldPart.SlideLayoutPart.SlideLayout.CommonSlideData;
            var masterSldData = _sdkSldPart.SlideLayoutPart.SlideMasterPart.SlideMaster.CommonSlideData;
            _layoutPlaceholders = new Lazy<HashSet<PlaceholderFontData>>(InitLayoutMaster(layoutSldData));
            _masterPlaceholders = new Lazy<HashSet<PlaceholderFontData>>(InitLayoutMaster(masterSldData));
            _masterBodyFontHeights = new Lazy<Dictionary<int, int>>(InitBodyTypePlaceholder(_sdkSldPart));
        }

        /// <summary>
        /// Tries gets font height. Return null if font height is not defined.
        /// </summary>
        /// <param name="sdkCompositeElement">Placeholder element.</param>
        /// <param name="pLvl">Paragraph level.</param>
        /// <returns></returns>
        public int? TryGetFontHeight(OpenXmlCompositeElement sdkCompositeElement, int pLvl) //TODO: use pattern Try
        {
            Check.NotNull(sdkCompositeElement, nameof(sdkCompositeElement));

            var paramPlaceholderData = PlaceholderLocationService.CreatePlaceholderData(sdkCompositeElement);
            
            // From slide layout element
            var lPlaceholder = _layoutPlaceholders.Value.FirstOrDefault(e => e.Equals(paramPlaceholderData));
            if (lPlaceholder != null && lPlaceholder.LvlFontHeights.ContainsKey(pLvl))
            {
                return lPlaceholder.LvlFontHeights[pLvl];
            }

            // From slide master element
            var mPlaceholder = _masterPlaceholders.Value.FirstOrDefault(e => e.Equals(paramPlaceholderData));
            if (mPlaceholder != null && mPlaceholder.LvlFontHeights.ContainsKey(pLvl))
            {
                return mPlaceholder.LvlFontHeights[pLvl];
            }

            // Title type
            var masterGlobalTextStyle = _sdkSldPart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles;
            if (paramPlaceholderData.PlaceholderType == PlaceholderType.Title)
            {
                return masterGlobalTextStyle.TitleStyle.Level1ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value;
            }

            // Master body type placeholder settings
            if (_masterBodyFontHeights.Value.ContainsKey(pLvl))
            {
                return _masterBodyFontHeights.Value[pLvl];
            }

            return null;
        }

        #region Private Methods

        private HashSet<PlaceholderFontData> InitLayoutMaster(P.CommonSlideData layoutMasterCommonSlideData)
        {
            var fontDataPlaceholders = new HashSet<PlaceholderFontData>();
            foreach (var sdkShape in layoutMasterCommonSlideData.ShapeTree.Elements<P.Shape>().Where(e => e.IsPlaceholder()))
            {
                var fontDataPlaceholder = FromLayoutMasterElement(sdkShape);
                fontDataPlaceholders.Add(fontDataPlaceholder);
            }

            return fontDataPlaceholders;
        }

        private static Dictionary<int, int> InitBodyTypePlaceholder(SlidePart xmlSldPart)
        {
            return FontHeightParser.FromCompositeElement(xmlSldPart.SlideLayoutPart.SlideMasterPart.SlideMaster.TextStyles.BodyStyle);
        }

        private static PlaceholderFontData FromLayoutMasterElement(P.Shape sdkShape)
        {
            var placeholderFontData = PlaceholderLocationService.PlaceholderFontDataFromCompositeElement(sdkShape);
            placeholderFontData.LvlFontHeights = FontHeightParser.FromCompositeElement(sdkShape.TextBody.ListStyle);

            if (!placeholderFontData.LvlFontHeights.Any()) // font height is still not known
            {
                var endParaRunPrFs = sdkShape.TextBody.GetFirstChild<A.Paragraph>().GetFirstChild<A.EndParagraphRunProperties>()?.FontSize;
                if (endParaRunPrFs != null)
                {
                    placeholderFontData.LvlFontHeights.Add(1, endParaRunPrFs.Value);
                }
            }

            return placeholderFontData;
        }

        #endregion Private Methods
    }
}
