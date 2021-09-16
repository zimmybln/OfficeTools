using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeTools.Types
{
    public record WordFileInfo
    {
        public string DisplayName { get; init; }

        public string Name { get; init; }
    }
}
