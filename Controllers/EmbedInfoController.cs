using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using PowerBi.Models;
using PowerBi.Services;
using System;
using System.Text.Json;

namespace PowerBi.Controllers
{
    public class EmbedInfoController : Controller
    {
        private readonly PbiEmbedService pbiEmbedService;
        private readonly IOptions<AzureAd> azureAd;

        public EmbedInfoController(PbiEmbedService pbiEmbedService, IOptions<AzureAd> azureAd)
        {
            this.pbiEmbedService = pbiEmbedService;
            this.azureAd = azureAd;
        }

        [HttpGet]
        public string GetEmbedInfo()
        {
            try
            {
                EmbedParams embedParams = pbiEmbedService.GetEmbedParams(new Guid(azureAd.Value.WorkspaceId), new Guid(azureAd.Value.ReportId));
                return JsonSerializer.Serialize(embedParams);
            }
            catch (Exception ex)
            {
                HttpContext.Response.StatusCode = 500;
                return ex.Message + "\n\n" + ex.StackTrace;
            }
        }
    }
}
