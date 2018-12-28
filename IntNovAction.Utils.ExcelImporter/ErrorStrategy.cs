namespace IntNovAction.Utils.Importer
{
    public enum ErrorStrategy
    {
        /// <summary>
        /// When an error is found in a column, the corresponding object is not included in the results
        /// </summary>
        DoNotAddElement = 0,

        /// <summary>
        /// When an error is found in a column, the corresponding object is included in the results, the field has
        /// its default value
        /// </summary>
        AddElement = 1
    }
}