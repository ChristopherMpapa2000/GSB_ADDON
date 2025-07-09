using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnGs.Models.OpenAPI.Response
{
    internal class ExtCreateMemorandumResponse
    {

        public class ExtCreateMemorandumResult
        {
            public ExtCreateMemorandumData data { get; set; }
        }

        public class ExtCreateMemorandumData
        {
            public int Id { get; set; }
            public string DocumentNo { get; set; }
            public string FileUrl { get; set; }
        }


        //public class CreateMemorandumResult : ICreateMemorandumResult
        //{
        //    public int Id { get; set; }
        //    public string DocumentNo { get; set; }
        //    public string FileUrl { get; set; }
        //}
        public class CreateMemorandumResponse
        {
            public ICreateMemorandumResult Data { get; set; }
        }
        public interface ICreateMemorandumResult
        {
            int Id { get; set; }
            string DocumentNo { get; set; }
            string FileUrl { get; set; }
        }
    }
}
