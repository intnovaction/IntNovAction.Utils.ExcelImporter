namespace IntNovAction.Utils.ExcelImporter
{
    public enum DuplicatedColumnStrategy
    {
        /// <summary>
        /// Take the values from the first column with the duplicated name 
        /// </summary>
        TakeFirst = 0,

        /// <summary>
        /// Take the values from the last column with the duplicated name 
        /// </summary>
        TakeLast = 1,

        /// <summary>
        /// Raise an error and do not import
        /// </summary>
        RaiseError = 3,
    }
}