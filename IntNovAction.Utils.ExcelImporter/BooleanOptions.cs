using System.Collections.Generic;

namespace IntNovAction.Utils.ExcelImporter
{
    internal class BooleanOptions
    {
        public BooleanOptions()
        {
            this.TrueStrings = new List<string>()
            {
                "y",
                "s",
                "1",
                "yes",
                "si",
                "sí",
                "true",
                "verdadero"
            };

            this.FalseStrings = new List<string>(){
                "n",
                "0",
                "no",
                "false",
                "falso"
            };
        }

        public readonly List<string> TrueStrings;
        public readonly List<string> FalseStrings;
    }
}