using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using WolfApprove.API2.Extension;
using System.Runtime.CompilerServices;

namespace AddOnGs.Extenstion
{
    public static class CallOpenAPI
    {
        public static Task<HttpResponseMessage> PostAPI<T>(this HttpClient httpClient ,string url ,T data)
        {
            var dataAsString = JsonConvert.SerializeObject(data);
            var content = new StringContent(dataAsString);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            return httpClient.PostAsync(url, content);
        }

        public static Task<HttpResponseMessage> GetAPI(this HttpClient httpClient, string url)
        {
            return httpClient.GetAsync(url);
        }

        public static async Task<T> ReadContentAs<T>(this HttpResponseMessage response)
        {
            if (!response.IsSuccessStatusCode)
            {
                Ext.WriteLogFile("Something went wrong calling the API: {response.ReasonPhrase}");
            }

            var dataAsString = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            return JsonConvert.DeserializeObject<T>(dataAsString);  
        }
    }
}
