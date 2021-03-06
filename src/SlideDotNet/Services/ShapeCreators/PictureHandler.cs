﻿using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services.Builders;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services.ShapeCreators
{
    /// <summary>
    /// Represents a picture handler for p:pic and picture p:graphicFrame element.
    /// </summary>
    public class PictureHandler : OpenXmlElementHandler
    {
        #region Fields

        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly InnerTransformFactory _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;
        private readonly SlidePart _sdkSldPart;
        private readonly IGeometryFactory _geometryFactory;

        #endregion Fields

        #region Constructors

        public PictureHandler(ShapeContext.Builder shapeContextBuilder,
                              InnerTransformFactory transformFactory,
                              IGeometryFactory geometryFactory,
                              SlidePart sdkSldPart) :
            this(shapeContextBuilder, transformFactory, geometryFactory, sdkSldPart, new ShapeEx.Builder())
        {

        }

        public PictureHandler(ShapeContext.Builder shapeContextBuilder,
                              InnerTransformFactory transformFactory,
                              IGeometryFactory geometryFactory,
                              SlidePart sdkSldPart,
                              IShapeBuilder shapeBuilder)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _shapeBuilder = shapeBuilder ?? throw new ArgumentNullException(nameof(shapeBuilder));
            _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _geometryFactory = geometryFactory ?? throw new ArgumentNullException(nameof(geometryFactory));
        }

        #endregion Constructors

        #region Constructors

        public override ShapeEx Create(OpenXmlElement sdkElement)
        {
            Check.NotNull(sdkElement, nameof(sdkElement));

            P.Picture sdkPicture;
            if (sdkElement is P.Picture treePic)
            {
                sdkPicture = treePic;
            }
            else
            {
                var framePic = sdkElement.Descendants<P.Picture>().FirstOrDefault();
                sdkPicture = framePic;
            }
            if (sdkPicture != null)
            {
                var pBlipFill = sdkPicture.GetFirstChild<P.BlipFill>();
                var blipRelateId = pBlipFill?.Blip?.Embed?.Value;
                if (blipRelateId == null)
                {
                    return null;
                }
                var pictureEx = new PictureEx(_sdkSldPart, blipRelateId);
                var spContext = _shapeContextBuilder.Build(sdkElement);
                var innerTransform = _transformFactory.FromComposite(sdkPicture);
                var geometry = _geometryFactory.ForPicture(sdkPicture);
                var shape = _shapeBuilder.WithPicture(innerTransform, spContext, pictureEx, geometry);

                return shape;
            }

            if (Successor != null)
            {
                return Successor.Create(sdkElement);
            }

            return null;
        }

        #endregion Constructors
    }
}