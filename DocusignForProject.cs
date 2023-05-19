//// <copyright file="DocusignForProject.cs" company="Microsoft">
//// Copyright (c) 2020 All Rights Reserved
//// </copyright>
//// <author></author>
//// <date>07/03/2020 06:00:47 PM</date>
//// <summary>Docusign for Project entity</summary>
namespace MCS.PSA.Plugins
{
    using System;
    using MCS.PSA.Plugins.BusinessLogic;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Client;
    using Microsoft.Xrm.Sdk.Query;

    /// <summary>
    /// DOCUSIGH for project.
    /// </summary>
    public class DocusignForProject : IPlugin
    {
        /// <summary>
        /// Execute Method
        ///  Execute Method
        /// </summary>
        /// <param name="serviceProvider">service provider</param>
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService orgService = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            Guid currentInitiatingUser = context.InitiatingUserId;

            // Access Token 
            if (!context.InputParameters.Contains("AccessToken") || !(context.InputParameters["AccessToken"] is string) || context.InputParameters["AccessToken"] == null)
            {
                string errorMessage = LocalizationHelper.GetLocalizedString(orgService, currentInitiatingUser, "DocuSign.AccessTokenErrorMessage");
                throw new Exception(errorMessage, new Exception("Access Token is null or empty."));
            }

            string accessToken = context.InputParameters["AccessToken"].ToString();

            // Envelope ID
            if (!context.InputParameters.Contains("Envelope") || !(context.InputParameters["Envelope"] is string) || context.InputParameters["Envelope"] == null)
            {
                string errorMessage = LocalizationHelper.GetLocalizedString(orgService, currentInitiatingUser, "DocuSign.GeneralErrorMessage");
                throw new Exception(errorMessage, new Exception("Envelope input parameter is null or empty."));
            }

            string envelope = context.InputParameters["Envelope"].ToString();
            string envelopeId = envelope.Split(',')[0];

            Docusign.CreateDocusignForProject(orgService, tracingService, context, currentInitiatingUser, accessToken, envelopeId);
        }
    }
}
