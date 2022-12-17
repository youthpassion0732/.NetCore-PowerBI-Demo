using Microsoft.PowerBI.Api;
using Microsoft.PowerBI.Api.Models;
using Microsoft.Rest;
using PowerBi.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace PowerBi.Services
{
    public class PbiEmbedService
    {
        private readonly AadService aadService;
        private readonly string powerBiApiUrl = "https://api.powerbi.com";

        public PbiEmbedService(AadService aadService)
        {
            this.aadService = aadService;
        }

        #region Public
        public EmbedParams GetEmbedParams(Guid workspaceId, Guid reportId, [Optional] Guid additionalDatasetId)
        {
            PowerBIClient pbiClient = this.GetPowerBIClient();

            Report pbiReport = pbiClient.Reports.GetReportInGroup(workspaceId, reportId);

            bool isRDLReport = String.IsNullOrEmpty(pbiReport.DatasetId);

            EmbedToken embedToken;

            if (isRDLReport)
            {
                embedToken = GetEmbedTokenForRDLReport(workspaceId, reportId);
            }
            else
            {
                List<Guid> datasetIds = new List<Guid>
                {
                    Guid.Parse(pbiReport.DatasetId)
                };

                if (additionalDatasetId != Guid.Empty)
                {
                    datasetIds.Add(additionalDatasetId);
                }

                embedToken = GetEmbedToken(reportId, datasetIds, workspaceId);
            }

            List<EmbedReport> embedReports = new List<EmbedReport>() {
                new EmbedReport
                {
                    ReportId = pbiReport.Id, ReportName = pbiReport.Name, EmbedUrl = pbiReport.EmbedUrl
                }
            };

            EmbedParams embedParams = new EmbedParams
            {
                EmbedReport = embedReports,
                Type = "Report",
                EmbedToken = embedToken
            };

            return embedParams;
        }
        #endregion

        #region Private
        private EmbedToken GetEmbedToken(Guid reportId, IList<Guid> datasetIds, [Optional] Guid targetWorkspaceId)
        {
            PowerBIClient pbiClient = this.GetPowerBIClient();

            // Create a request for getting Embed token 
            // This method works only with new Power BI V2 workspace experience
            GenerateTokenRequestV2 tokenRequest = new GenerateTokenRequestV2(

                reports: new List<GenerateTokenRequestV2Report>() { new GenerateTokenRequestV2Report(reportId) },

                datasets: datasetIds.Select(datasetId => new GenerateTokenRequestV2Dataset(datasetId.ToString())).ToList(),

                targetWorkspaces: targetWorkspaceId != Guid.Empty ? new List<GenerateTokenRequestV2TargetWorkspace>() { new GenerateTokenRequestV2TargetWorkspace(targetWorkspaceId) } : null
            );

            // Generate Embed token
            EmbedToken embedToken = pbiClient.EmbedToken.GenerateToken(tokenRequest);

            return embedToken;
        }

        public EmbedParams GetEmbedParams(Guid workspaceId, IList<Guid> reportIds, [Optional] IList<Guid> additionalDatasetIds)
        {
            // Note: This method is an example and is not consumed in this sample app

            PowerBIClient pbiClient = this.GetPowerBIClient();

            // Create mapping for reports and Embed URLs
            List<EmbedReport> embedReports = new List<EmbedReport>();

            // Create list of datasets
            List<Guid> datasetIds = new List<Guid>();

            // Get datasets and Embed URLs for all the reports
            foreach (Guid reportId in reportIds)
            {
                // Get report info
                Report pbiReport = pbiClient.Reports.GetReportInGroup(workspaceId, reportId);

                datasetIds.Add(Guid.Parse(pbiReport.DatasetId));

                // Add report data for embedding
                embedReports.Add(new EmbedReport { ReportId = pbiReport.Id, ReportName = pbiReport.Name, EmbedUrl = pbiReport.EmbedUrl });
            }

            // Append to existing list of datasets to achieve dynamic binding later
            if (additionalDatasetIds != null)
            {
                datasetIds.AddRange(additionalDatasetIds);
            }

            // Get Embed token multiple resources
            EmbedToken embedToken = GetEmbedToken(reportIds, datasetIds, workspaceId);

            // Capture embed params
            EmbedParams embedParams = new EmbedParams
            {
                EmbedReport = embedReports,
                Type = "Report",
                EmbedToken = embedToken
            };

            return embedParams;
        }

        private EmbedToken GetEmbedToken(IList<Guid> reportIds, IList<Guid> datasetIds, [Optional] Guid targetWorkspaceId)
        {
            // Note: This method is an example and is not consumed in this sample app

            PowerBIClient pbiClient = this.GetPowerBIClient();

            // Convert report Ids to required types
            List<GenerateTokenRequestV2Report> reports = reportIds.Select(reportId => new GenerateTokenRequestV2Report(reportId)).ToList();

            // Convert dataset Ids to required types
            List<GenerateTokenRequestV2Dataset> datasets = datasetIds.Select(datasetId => new GenerateTokenRequestV2Dataset(datasetId.ToString())).ToList();

            // Create a request for getting Embed token 
            // This method works only with new Power BI V2 workspace experience
            GenerateTokenRequestV2 tokenRequest = new GenerateTokenRequestV2(

                datasets: datasets,

                reports: reports,

                targetWorkspaces: targetWorkspaceId != Guid.Empty ? new List<GenerateTokenRequestV2TargetWorkspace>() { new GenerateTokenRequestV2TargetWorkspace(targetWorkspaceId) } : null
            );

            // Generate Embed token
            EmbedToken embedToken = pbiClient.EmbedToken.GenerateToken(tokenRequest);

            return embedToken;
        }

        private EmbedToken GetEmbedToken(IList<Guid> reportIds, IList<Guid> datasetIds, [Optional] IList<Guid> targetWorkspaceIds)
        {
            // Note: This method is an example and is not consumed in this sample app

            PowerBIClient pbiClient = this.GetPowerBIClient();

            // Convert report Ids to required types
            List<GenerateTokenRequestV2Report> reports = reportIds.Select(reportId => new GenerateTokenRequestV2Report(reportId)).ToList();

            // Convert dataset Ids to required types
            List<GenerateTokenRequestV2Dataset> datasets = datasetIds.Select(datasetId => new GenerateTokenRequestV2Dataset(datasetId.ToString())).ToList();

            // Convert target workspace Ids to required types
            IList<GenerateTokenRequestV2TargetWorkspace> targetWorkspaces = null;
            if (targetWorkspaceIds != null)
            {
                targetWorkspaces = targetWorkspaceIds.Select(targetWorkspaceId => new GenerateTokenRequestV2TargetWorkspace(targetWorkspaceId)).ToList();
            }

            // Create a request for getting Embed token 
            // This method works only with new Power BI V2 workspace experience
            GenerateTokenRequestV2 tokenRequest = new GenerateTokenRequestV2(

                datasets: datasets,

                reports: reports,

                targetWorkspaces: targetWorkspaceIds != null ? targetWorkspaces : null
            );

            // Generate Embed token
            EmbedToken embedToken = pbiClient.EmbedToken.GenerateToken(tokenRequest);

            return embedToken;
        }

        private EmbedToken GetEmbedTokenForRDLReport(Guid targetWorkspaceId, Guid reportId, string accessLevel = "view")
        {
            PowerBIClient pbiClient = this.GetPowerBIClient();

            // Generate token request for RDL Report
            GenerateTokenRequest generateTokenRequestParameters = new GenerateTokenRequest(
                accessLevel: accessLevel
            );

            // Generate Embed token
            EmbedToken embedToken = pbiClient.Reports.GenerateTokenInGroup(targetWorkspaceId, reportId, generateTokenRequestParameters);

            return embedToken;
        }

        private PowerBIClient GetPowerBIClient()
        {
            TokenCredentials tokenCredentials = new TokenCredentials(aadService.GetAccessToken(), "Bearer");
            return new PowerBIClient(new Uri(powerBiApiUrl), tokenCredentials);
        }
        #endregion
    }
}
