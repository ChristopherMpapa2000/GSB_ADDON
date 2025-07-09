using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnGs.Models.OpenAPI.Request
{
    internal class ExtCreateMemorandumRequest
    {
        public class CreateMemorandumRequest
        {
            public string UserPrincipalName { get; set; }
            public string Subject { get; set; }
            public string Action { get; set; }
            public string DocumentCode { get; set; }
            public IEnumerable<FormControl> Form { get; set; }
            public string EFixCreator { get; set; } = string.Empty;
            public string EFixRequestor { get; set; } = string.Empty;
        }
    }
}
