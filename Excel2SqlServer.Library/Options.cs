using System.Collections.Generic;

namespace Excel2SqlServer.Library
{
    public class Options
    {
        /// <summary>
        /// when importing multiple worksheets, schema name to use on generated tables
        /// </summary>
        public string SchemaName { get; set; }
        public bool TruncateFirst { get; set; }
        public bool AutoTrimStrings { get; set; }
        public bool RemoveNonPrintingChars { get; set; }
        public IEnumerable<string> CustomColumns { get; set; }
    }
}
