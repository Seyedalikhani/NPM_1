using System.Collections.Generic;

namespace NPM_1.Controllers
{
    public class FilterRequest
    {
        public IEnumerable<object> Filters { get; internal set; }
    }
}