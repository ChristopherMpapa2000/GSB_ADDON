using AddOnGs.Extenstion;
using AddOnGs.Models;
using AddOnGs.Service;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Http.Results;
using WolfApprove.API2.Controllers.Bean;
using WolfApprove.API2.Extension;
using WolfApprove.Model.CustomClass;

namespace AddOnGs.Handlers
{
    public class AddonApiVersionRedirectHandler : DelegatingHandler
    {
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            try
            {
                if (request.RequestUri.AbsolutePath.Contains("api/services/submitform"))
                {
                    var requestContent = await request.Content.ReadAsStringAsync();
                    var requestModel = JsonConvert.DeserializeObject<AddonFormModel.Form_Model>(requestContent);


                    #region [-- SE-F-SHE-020 To SE-F-SHE-020-01 --]
                    if (requestModel.memoPage.memoDetail.template_code == "SE-F-SHE-020")
                    {
                        var resultFromMainService = await base.SendAsync(request, cancellationToken);
                        if (resultFromMainService.IsSuccessStatusCode)
                        {
                            // response From Standard Process
                            var responseString = await resultFromMainService.Content.ReadAsStringAsync();
                            var responseBean = JsonConvert.DeserializeObject<ResponseBean>(responseString);
                            requestModel.memoPage.memoDetail.memoid = responseBean.memoid;

                            var _serviceSourcing = new AutoCreated(requestModel.memoPage.memoDetail.connectionString);
                            var result = await _serviceSourcing.AutoGenerate01(requestModel);
                        }
                        return resultFromMainService;
                    }
                    #endregion
                    #region [-- SE-F-SHE-020-01 To SE-F-SHE-0211 --]
                    if (requestModel.memoPage.memoDetail.template_code == "SE-F-SHE-020-01")
                    {
                        var resultFromMainService = await base.SendAsync(request, cancellationToken);
                        if (resultFromMainService.IsSuccessStatusCode)
                        {
                            // response From Standard Process
                            var responseString = await resultFromMainService.Content.ReadAsStringAsync();
                            var responseBean = JsonConvert.DeserializeObject<ResponseBean>(responseString);
                            requestModel.memoPage.memoDetail.memoid = responseBean.memoid;

                            var _serviceSourcing = new AutoCreated(requestModel.memoPage.memoDetail.connectionString);
                            var result = await _serviceSourcing.AutoGenerate021(requestModel);
                        }
                        return resultFromMainService;
                    }
                    #endregion
                    #region [ -- F-QMS-008 To F-QMS-009 -- ]
                    if (requestModel.memoPage.memoDetail.template_code == "F-QMS-008")
                    {
                        var resultFromMainService = await base.SendAsync(request, cancellationToken);
                        if (resultFromMainService.IsSuccessStatusCode)
                        {
                            // response From Standard Process
                            var responseString = await resultFromMainService.Content.ReadAsStringAsync();
                            var responseBean = JsonConvert.DeserializeObject<ResponseBean>(responseString);
                            requestModel.memoPage.memoDetail.memoid = responseBean.memoid;

                            var _serviceSourcing = new AutoCreated(requestModel.memoPage.memoDetail.connectionString);
                            var result = await _serviceSourcing.AutoGenerate009(requestModel);
                        }
                        return resultFromMainService;
                    }
                    #endregion
                    #region [ -- F-QMS-009 To F-QMS-010 -- ]
                    if (requestModel.memoPage.memoDetail.template_code == "F-QMS-009")
                    {
                        var resultFromMainService = await base.SendAsync(request, cancellationToken);
                        if (resultFromMainService.IsSuccessStatusCode)
                        {
                            // response From Standard Process
                            var responseString = await resultFromMainService.Content.ReadAsStringAsync();
                            var responseBean = JsonConvert.DeserializeObject<ResponseBean>(responseString);
                            requestModel.memoPage.memoDetail.memoid = responseBean.memoid;

                            var _serviceSourcing = new AutoCreated(requestModel.memoPage.memoDetail.connectionString);
                            var result = await _serviceSourcing.AutoGenerate010(requestModel);
                        }
                        return resultFromMainService;
                    }
                    #endregion
                    #region [ -- F-QMS-010 To F-QMS-004 -- ]
                    if (requestModel.memoPage.memoDetail.template_code == "F-QMS-010")
                    {
                        var resultFromMainService = await base.SendAsync(request, cancellationToken);
                        if (resultFromMainService.IsSuccessStatusCode)
                        {
                            // response From Standard Process
                            var responseString = await resultFromMainService.Content.ReadAsStringAsync();
                            var responseBean = JsonConvert.DeserializeObject<ResponseBean>(responseString);
                            requestModel.memoPage.memoDetail.memoid = responseBean.memoid;

                            var _serviceSourcing = new AutoCreated(requestModel.memoPage.memoDetail.connectionString);
                            var result = await _serviceSourcing.AutoGenerate004(requestModel);
                        }
                        return resultFromMainService;
                    }
                    #endregion

                    #region [-- SE-F-SHE-021 To SE-F-SHE-022 --]
                    if (requestModel.memoPage.memoDetail.template_code == "SE-F-SHE-021")
                    {
                        var resultFromMainService = await base.SendAsync(request, cancellationToken);
                        if (resultFromMainService.IsSuccessStatusCode)
                        {
                            // response From Standard Process
                            var responseString = await resultFromMainService.Content.ReadAsStringAsync();
                            var responseBean = JsonConvert.DeserializeObject<ResponseBean>(responseString);
                            requestModel.memoPage.memoDetail.memoid = responseBean.memoid;

                            var _serviceSourcing = new AutoCreated(requestModel.memoPage.memoDetail.connectionString);
                            var result = await _serviceSourcing.AutoGenerate022(requestModel);
                        }
                        return resultFromMainService;
                    }
                    #endregion
                    #region [-- SE-F-SHE-022 To SE-F-SHE-024 --]
                    if (requestModel.memoPage.memoDetail.template_code == "SE-F-SHE-022")
                    {
                        var resultFromMainService = await base.SendAsync(request, cancellationToken);
                        if (resultFromMainService.IsSuccessStatusCode)
                        {
                            // response From Standard Process
                            var responseString = await resultFromMainService.Content.ReadAsStringAsync();
                            var responseBean = JsonConvert.DeserializeObject<ResponseBean>(responseString);
                            requestModel.memoPage.memoDetail.memoid = responseBean.memoid;

                            var _serviceSourcing = new AutoCreated(requestModel.memoPage.memoDetail.connectionString);
                            var result = await _serviceSourcing.AutoGenerateCar(requestModel);
                        }
                        return resultFromMainService;
                    }
                    #endregion

                }
                return await base.SendAsync(request, cancellationToken);
            }
            catch (Exception ex)
            {
                Ext.ErrorLog(ex, "SendAsync");
                throw;
            }
        }
    }
}
