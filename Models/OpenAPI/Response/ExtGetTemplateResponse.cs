using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnGs.Models.OpenAPI.Response
{
    internal class ExtGetTemplateResponse
    {
        public string documentCode { get; set; }
        public string subject { get; set; }
        public IEnumerable<FormControl> form { get; set; }
    }
}
