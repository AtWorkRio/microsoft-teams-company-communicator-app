// <copyright file="CompanyCommunicatorBotFilterMiddleware.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

using System.Collections.Generic;
using Microsoft.Bot.Schema;

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

        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyCommunicatorBotFilterMiddleware"/> class.
        /// </summary>
        /// <param name="configuration">ASP.NET Core <see cref="IConfiguration"/> instance.</param>
        public CompanyCommunicatorBotFilterMiddleware(IConfiguration configuration)
        {
            this.configuration = configuration;
        }

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
        public async Task OnTurnAsync(ITurnContext turnContext, NextDelegate next, CancellationToken cancellationToken = default)
        {
            //CancellationTokenSource cts = null;
            //try
            //{
            //    cts = new CancellationTokenSource();
            //    cancellationToken.Register(() => cts.Cancel());
            //    string content = (string) turnContext.Activity.Attachments[0].Content;
            //    if (content.Contains("/document"))
            //    {
            //        await SendTypingAsync(turnContext, TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(2), cancellationToken);
            //    }
            //}
            //finally
            //{
            //    if (cts != null)
            //    {
            //        cts.Cancel();
            //    }
            //}

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

        
        private static async Task SendTypingAsync(ITurnContext turnContext, TimeSpan delay, TimeSpan period, CancellationToken cancellationToken)
        {
            try
            {
                await Task.Delay(delay, cancellationToken).ConfigureAwait(false);

                while (!cancellationToken.IsCancellationRequested)
                {
                    if (!cancellationToken.IsCancellationRequested)
                    {
                        await SendTypingActivityAsync(turnContext, cancellationToken).ConfigureAwait(false);
                    }

                    // if we happen to cancel when in the delay we will get a TaskCanceledException
                    await Task.Delay(period, cancellationToken).ConfigureAwait(false);
                }
            }
            catch (TaskCanceledException)
            {
                // do nothing
            }
        }

        private static async Task SendTypingActivityAsync(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            // create a TypingActivity, associate it with the conversation and send immediately
            var typingActivity = new Activity
            {
                Type = ActivityTypes.Message,
                Attachments = new List<Attachment>
                {
                    new Attachment("application/pdf", "")
                },
                RelatesTo = turnContext.Activity.RelatesTo,
            };

            // sending the Activity directly on the Adapter avoids other Middleware and avoids setting the Responded
            // flag, however, this also requires that the conversation reference details are explicitly added.
            var conversationReference = turnContext.Activity.GetConversationReference();
            typingActivity.ApplyConversationReference(conversationReference);

            // make sure to send the Activity directly on the Adapter rather than via the TurnContext
            await turnContext.Adapter.SendActivitiesAsync(turnContext, new Activity[] { typingActivity }, cancellationToken).ConfigureAwait(false);
        }
    }
}
