﻿using System;
using System.Collections.Generic;

namespace QuickSpread.Client
{
    /// <summary>
    /// Quick spread client interface.
    /// </summary>
    public interface IQuickClient
    {
        /// <summary>
        /// Export to a specified spreadsheet.
        /// </summary>
        /// <param name="exportCollections">export collections</param>
        /// <exception cref="ArgumentNullException"></exception>
        void Export<T>(IList<T> exportCollections);
    }
}
