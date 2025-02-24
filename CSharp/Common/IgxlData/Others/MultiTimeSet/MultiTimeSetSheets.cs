﻿using System;
using System.Collections.Generic;
using System.IO;

namespace IgxlData.Others.MultiTimeSet
{
    public class MultiTimeSetSheets
    {
        public MultiTimeSetSheets()
        {
            TimeSetBasicSheetsList = new List<ComTimeSetBasicSheet>();
        }

        #region Propety

        public List<ComTimeSetBasicSheet> TimeSetBasicSheetsList { get; set; }

        #endregion


        #region Member Function

        public void AddTimeSetSheet(ComTimeSetBasicSheet tsetSheet)
        {
            TimeSetBasicSheetsList.Add(tsetSheet);
        }

        public ComTimeSetBasicSheet FindTimeSetSheet(string name)
        {
            name = Path.GetFileNameWithoutExtension(name);
            return TimeSetBasicSheetsList.Find(x =>
                x.SheetName.Equals(name, StringComparison.CurrentCultureIgnoreCase));
        }

        #endregion
    }
}