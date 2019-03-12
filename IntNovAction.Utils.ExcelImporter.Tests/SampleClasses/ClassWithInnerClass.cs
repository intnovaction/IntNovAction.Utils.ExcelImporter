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
    }

    public class InnerClass
    {
        public int Prop1 { get; set; }

    }
}
