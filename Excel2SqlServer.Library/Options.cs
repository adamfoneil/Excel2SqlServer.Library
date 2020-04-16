using System.Collections.Generic;

namespace Excel2SqlServer.Library
{
    public class Options
    {
        public bool TruncateFirst { get; set; }
        public bool AutoTrimStrings { get; set; }
        public IEnumerable<string> CustomColumns { get; set; }
    }
}
