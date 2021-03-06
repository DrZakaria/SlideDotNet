﻿using System;
using System.Collections.Generic;
using System.Linq;
using SlideDotNet.Collections;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.TableComponents;
using SlideDotNet.Validation;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Models.SlideComponents
{
    /// <summary>
    /// Represents a collection of table rows.
    /// </summary>
    public class RowCollection : EditAbleCollection<RowEx>
    {
        private readonly Dictionary<RowEx, A.TableRow> _innerSdkDic;

        #region Constructors

        public RowCollection(IEnumerable<A.TableRow> sdkTblRows, IShapeContext spContext)
        {
            Check.NotNull(sdkTblRows, nameof(sdkTblRows));
            Check.NotNull(spContext, nameof(spContext));

            var count = sdkTblRows.Count();
            CollectionItems = new List<RowEx>(count);
            _innerSdkDic = new Dictionary<RowEx, A.TableRow>(count);
            foreach (var sdkRow in sdkTblRows)
            {
                var innerRow = new RowEx(sdkRow, spContext);

                _innerSdkDic.Add(innerRow, sdkRow);
                CollectionItems.Add(innerRow);
            }
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Removes the specified table row.
        /// </summary>
        /// <param name="item"></param>
        public override void Remove(RowEx item)
        {
            if (!_innerSdkDic.ContainsKey(item))
            {
                throw new ArgumentNullException(nameof(item));
            }

            _innerSdkDic[item].Remove();
            CollectionItems.Remove(item);
        }

        /// <summary>
        /// Removes table row by index.
        /// </summary>
        /// <param name="index"></param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= CollectionItems.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var innerRow = CollectionItems[index];
            Remove(innerRow);
        }

        #endregion Public Methods
    }
}