using System;
using System.Collections.Generic;
using System.Text;

namespace IntNovAction.Utils.Importer.Tests.SampleClasses
{
    class SampleImportInto
    {
        public int IntColumn { get; set; }

        public int? NullableIntColumn { get; set; }

        public decimal DecimalColumn { get; set; }

        public decimal? NullableDecimalColumn { get; set; }

        public float FloatColumn { get; set; }

        public float? NullableFloatColumn { get; set; }

        public string StringColumn { get; set; } 

        public DateTime DateColumn { get; set; }

        public DateTime? NullableDateColumn { get; set; }


        public string BoolColumn { get; set; }
    }
}
