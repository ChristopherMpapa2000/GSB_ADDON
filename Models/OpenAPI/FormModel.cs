using AddOnGs.Models.OpenAPI.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnGs.Models.OpenAPI
{
    public class Attribute
    {
        public bool Required { get; set; }
        public List<FormControl> Column { get; set; }
    }

    public class Template
    {
        public string Label { get; set; }
        public string Alter { get; set; }
        public string Type { get; set; }
        public Attribute Attribute { get; set; }
    }

    public class Data
    {
        public dynamic Value { get; set; }
        public List<List<Data>> Row { get; set; }
    }

    public class FormControl : IFormControl
    {
        public Template Template { get; set; }
        public Data Data { get; set; }
        public string Guid { get; set; }
    }

    namespace Interface
    {
        public interface IFormControl
        {
            Template Template { get; set; }
            Data Data { get; set; }
            string Guid { get; set; }
        }
    }
}
