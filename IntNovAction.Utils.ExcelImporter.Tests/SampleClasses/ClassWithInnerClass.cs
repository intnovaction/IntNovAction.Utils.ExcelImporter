using System;
using System.Collections.Generic;
using System.Text;

namespace IntNovAction.Utils.ExcelImporter.Tests.SampleClasses
{
    public class ClassWithInnerClass
    {
        public ClassWithInnerClass()
        {
            Inner = new InnerClass();
        }

        public InnerClass Inner { get; set; }

        public int TestInt { get; set; }
    }

    public class InnerClass
    {
        public int PropInt { get; set; }

    }
}
