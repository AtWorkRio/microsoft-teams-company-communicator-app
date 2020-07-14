// <copyright file="CompanyCommunicatorBotFilterMiddleware.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

using System.Collections.Generic;
using System.Net.Http;
using IdentityModel.Client;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Teams.Apps.CompanyCommunicator.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Microsoft.Teams.Apps.CompanyCommunicator.Bot
{
    using System;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Extensions.Configuration;

    /// <summary>
    /// The bot's general filter middleware.
    /// </summary>
    public class CompanyCommunicatorBotFilterMiddleware : IMiddleware
    {
        private static readonly string MsTeamsChannelId = "msteams";

        private readonly IConfiguration configuration;
        private readonly DiscoveryCache discoveryCache;
        private readonly AtWorkRioIdentityOptions atWorkRioIdentityOptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyCommunicatorBotFilterMiddleware"/> class.
        /// </summary>
        /// <param name="configuration">ASP.NET Core <see cref="IConfiguration"/> instance.</param>
        public CompanyCommunicatorBotFilterMiddleware(IConfiguration configuration, DiscoveryCache discoveryCache,
            AtWorkRioIdentityOptions atWorkRioIdentityOptions)
        {
            this.configuration = configuration;
            this.discoveryCache = discoveryCache;
            this.atWorkRioIdentityOptions = atWorkRioIdentityOptions;
        }
        
        public const string DocumentCommand1 = "/doc ";
        public const string DocumentCommand2 = "/docs ";

        /// <summary>
        /// Processes an incoming activity.
        /// If the activity's channel id is not "msteams", or its conversation's tenant is not an allowed tenant,
        /// then the middleware short circuits the pipeline, and skips the middlewares and handlers
        /// that are listed after this filter in the pipeline.
        /// </summary>
        /// <param name="turnContext">Context object containing information for a single turn of a conversation.</param>
        /// <param name="next">The delegate to call to continue the bot middleware pipeline.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        public async Task OnTurnAsync(ITurnContext turnContext, NextDelegate next,
            CancellationToken cancellationToken = default)
        {
            CancellationTokenSource cts = null;
            try
            {
                cts = new CancellationTokenSource();
                cancellationToken.Register(() => cts.Cancel());
                var text = turnContext.Activity.Text.Trim().ToLower();

                if (text.Contains(DocumentCommand1) || text.Contains(DocumentCommand2))
                {
                    var commandParam = text
                        .Replace(DocumentCommand1, string.Empty)
                        .Replace(DocumentCommand2, string.Empty)
                        .Trim();
                    var docs = await SearchDocuments(commandParam, discoveryCache, atWorkRioIdentityOptions);

                    if (docs.Count == 1)
                    {
                        await SendDocument(turnContext, docs[0], cancellationToken).ConfigureAwait(false);
                    }
                    else if (docs.Count > 1)
                    {
                        await SendDocumentOptions(turnContext, docs, cancellationToken).ConfigureAwait(false);
                    }
                }
            }
            catch
            {
            }

            var isMsTeamsChannel = this.ValidateBotFrameworkChannelId(turnContext);
            if (!isMsTeamsChannel)
            {
                return;
            }

            var isAllowedTenant = this.ValidateTenant(turnContext);
            if (!isAllowedTenant)
            {
                return;
            }

            await next(cancellationToken).ConfigureAwait(false);
        }

        private bool ValidateBotFrameworkChannelId(ITurnContext turnContext)
        {
            var disableChannelFilter = this.configuration.GetValue<bool>("DisableChannelFilter", false);
            if (disableChannelFilter)
            {
                return true;
            }

            return CompanyCommunicatorBotFilterMiddleware.MsTeamsChannelId.Equals(
                turnContext?.Activity?.ChannelId,
                StringComparison.OrdinalIgnoreCase);
        }

        private bool ValidateTenant(ITurnContext turnContext)
        {
            var disableTenantFilter = this.configuration.GetValue<bool>("DisableTenantFilter", false);
            if (disableTenantFilter)
            {
                return true;
            }

            var allowedTenantIds = this.configuration
                ?.GetValue<string>("AllowedTenants", string.Empty)
                ?.Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                ?.Select(p => p.Trim());
            if (allowedTenantIds == null || allowedTenantIds.Count() == 0)
            {
                var exceptionMessage = "AllowedTenants setting is not set properly in the configuration file.";
                Console.WriteLine(exceptionMessage);
                throw new ApplicationException(exceptionMessage);
            }

            var tenantId = turnContext?.Activity?.Conversation?.TenantId;
            return allowedTenantIds.Contains(tenantId);
        }

        private static async Task SendDocument(ITurnContext turnContext, SgcpDocumentListDTO doc,
            CancellationToken cancellationToken)
        {
            var message = MessageFactory.Attachment(new Attachment(
                doc.ContentType,
                doc.ContentUrl, // "http://localhost:5002/api/projects/c87fa576-567c-4fa4-adfa-9dc96df27165/documents/801982fb-2891-4711-ae9f-31abad94a66d/download?access_token=c9ad03cf5e6034d948cb7bdc04a09cbe3e604a047ed799ffe773b4b034385f8b",
                null,
                doc.DocumentName
            ), $"Project: {doc.ProjectName}");

            await turnContext.SendActivityAsync(message, cancellationToken).ConfigureAwait(false);
        }

        private static async Task SendDocumentOptions(ITurnContext turnContext, List<SgcpDocumentListDTO> docs,
            CancellationToken cancellationToken)
        {.
            var card = new HeroCard
            {
                Title = $"I've found ({docs.Count}) documents",
                Buttons = docs.Select(x => new CardAction
                {
                    Type = ActionTypes.OpenUrl,
                    Title = x.DocumentName,
                    Value = x.ContentUrl
                }).ToList()
            };

            var response = MessageFactory.Attachment(card.ToAttachment());

            await turnContext.SendActivityAsync(response, cancellationToken).ConfigureAwait(false);

            //var message = MessageFactory.SuggestedActions(
            //    docs.Select(x => new CardAction
            //    {
            //        Title = $"{x.DocumentName}",
            //        Type = ActionTypes.MessageBack,
            //        Value = $"/doc {x.DocumentName}"
            //    }).ToList(), "Suggested Documents");

            //await turnContext.SendActivityAsync(message, cancellationToken).ConfigureAwait(false);
        }


        public static async Task<List<SgcpDocumentListDTO>> SearchDocuments(string code, DiscoveryCache discoveryCache, AtWorkRioIdentityOptions atWorkRioIdentityOptions)
        {
            var disco = await discoveryCache.GetAsync();
            if (disco.IsError) throw new Exception(disco.Error);

            var tokenClient = new HttpClient();
            var tokenResponse = await tokenClient.RequestClientCredentialsTokenAsync(new ClientCredentialsTokenRequest
            {
                Address = disco.TokenEndpoint,
                ClientId = atWorkRioIdentityOptions.SgcpTeamsClientId,
                ClientSecret = atWorkRioIdentityOptions.SgcpTeamsClientSecret,
                Scope = "https://sgcp-teams.atworkrio.com https://plantra.atworkrio.com"
            });

            if (tokenResponse.IsError) throw new Exception(tokenResponse.Error);

            // call API
            var apiClient = new HttpClient();
            apiClient.SetBearerToken(tokenResponse.AccessToken);

            try
            {
                var json = await apiClient.GetStringAsync(
                    $"{atWorkRioIdentityOptions.SgcpTeamsApiUrl}/teams/docSearch?code={Uri.EscapeUriString(code)}");
                var docs = JsonConvert.DeserializeObject<Envelope<List<SgcpDocumentListDTO>>>(json);
                if (docs.Result != null)
                {
                    return docs.Result;
                }
            }
            catch
            {
            }

            return new List<SgcpDocumentListDTO>();
        }
    }

    public class SgcpDocumentListDTO
    {
        public string ProjectName { get; set; }
        public string DocumentName { get; set; }
        public string ContentUrl { get; set; }
        public string ContentType { get; set; }
    }
}