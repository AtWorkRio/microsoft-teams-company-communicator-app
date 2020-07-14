// <copyright file="CompanyCommunicatorBot.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

using System.Collections.Generic;
using IdentityModel.Client;
using Newtonsoft.Json.Linq;

namespace Microsoft.Teams.Apps.CompanyCommunicator.Bot
{
    using System;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;

    /// <summary>
    /// Company Communicator Bot.
    /// </summary>
    public class CompanyCommunicatorBot : ActivityHandler
    {
        private static readonly string TeamRenamedEventType = "teamRenamed";

        private readonly TeamsDataCapture teamsDataCapture;
        private readonly DiscoveryCache discoveryCache;
        private readonly AtWorkRioIdentityOptions atWorkRioIdentityOptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyCommunicatorBot"/> class.
        /// </summary>
        /// <param name="teamsDataCapture">Teams data capture service.</param>
        public CompanyCommunicatorBot(TeamsDataCapture teamsDataCapture, DiscoveryCache discoveryCache, AtWorkRioIdentityOptions atWorkRioIdentityOptions)
        {
            this.teamsDataCapture = teamsDataCapture;
            this.discoveryCache = discoveryCache;
            this.atWorkRioIdentityOptions = atWorkRioIdentityOptions;
        }

        /// <summary>
        /// Invoked when a conversation update activity is received from the channel.
        /// </summary>
        /// <param name="turnContext">The context object for this turn.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnConversationUpdateActivityAsync(
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            // base.OnConversationUpdateActivityAsync is useful when it comes to responding to users being added to or removed from the conversation.
            // For example, a bot could respond to a user being added by greeting the user.
            // By default, base.OnConversationUpdateActivityAsync will call <see cref="OnMembersAddedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been added or <see cref="OnMembersRemovedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been removed. base.OnConversationUpdateActivityAsync checks the member ID so that it only responds to updates regarding members other than the bot itself.
            await base.OnConversationUpdateActivityAsync(turnContext, cancellationToken);

            var activity = turnContext.Activity;
            var botId = activity.Recipient.Id;

            var isTeamRenamed = this.IsTeamInformationUpdated(activity);
            if (isTeamRenamed)
            {
                await this.teamsDataCapture.OnTeamInformationUpdatedAsync(activity);
            }

            // Take action if this event includes the bot being added
            if (activity.MembersAdded?.FirstOrDefault(p => p.Id == botId) != null)
            {
                await this.teamsDataCapture.OnBotAddedAsync(activity);
            }

            // Take action if this event includes the bot being removed
            if (activity.MembersRemoved?.FirstOrDefault(p => p.Id == botId) != null)
            {
                await this.teamsDataCapture.OnBotRemovedAsync(activity);
            }
        }

        public override Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = new CancellationToken())
        {
            return base.OnTurnAsync(turnContext, cancellationToken);
        }

        //protected override async Task<InvokeResponse> OnInvokeActivityAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        //{
        //    switch (turnContext.Activity.Name)
        //    {
        //        case "composeExtension/query":
        //            return CreateInvokeResponse(await OnTeamsMessagingExtensionQueryAsync(turnContext, (MessagingExtensionQuery)turnContext.Activity.Value, cancellationToken).ConfigureAwait(false));

        //        case "composeExtension/selectItem":
        //            return CreateInvokeResponse(await OnTeamsMessagingExtensionSelectItemAsync(turnContext, turnContext.Activity.Value as JObject, cancellationToken).ConfigureAwait(false));

        //        default:
        //            return await base.OnInvokeActivityAsync(turnContext, cancellationToken).ConfigureAwait(false);
        //    }
        //}

        //protected async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
        //{
        //    var text = query?.Parameters?[0]?.Value as string ?? string.Empty;

        //    var docs = await CompanyCommunicatorBotFilterMiddleware.SearchDocuments(text, discoveryCache, atWorkRioIdentityOptions);

        //    // We take every row of the results and wrap them in cards wrapped in in MessagingExtensionAttachment objects.
        //    // The Preview is optional, if it includes a Tap, that will trigger the OnTeamsMessagingExtensionSelectItemAsync event back on this bot.
        //    var attachments = docs.Select(package =>
        //    {
        //        var previewCard = new ThumbnailCard { Title = package.DocumentName, Tap = new CardAction { Type = "invoke", Value = package } };
        //        //if (!string.IsNullOrEmpty(package.Item5))
        //        //    previewCard.Images = new List<CardImage>() { new CardImage(package.Item5, "Icon") };

        //        var attachment = new MessagingExtensionAttachment
        //        {
        //            ContentType = HeroCard.ContentType,
        //            Content = new HeroCard { Title = package.DocumentName },
        //            Preview = previewCard.ToAttachment(),
        //        };

        //        return attachment;
        //    }).ToList();

        //    // The list of MessagingExtensionAttachments must we wrapped in a MessagingExtensionResult wrapped in a MessagingExtensionResponse.
        //    return new MessagingExtensionResponse
        //    {
        //        ComposeExtension = new MessagingExtensionResult
        //        {
        //            Type = "result",
        //            AttachmentLayout = "list",
        //            Attachments = attachments
        //        }
        //    };
        //}

        //protected Task<MessagingExtensionResponse> OnTeamsMessagingExtensionSelectItemAsync(ITurnContext<IInvokeActivity> turnContext, JObject query, CancellationToken cancellationToken)
        //{
        //    // The Preview card's Tap should have a Value property assigned, this will be returned to the bot in this event. 
        //    var (packageId, version, description, projectUrl, iconUrl) = query.ToObject<(string, string, string, string, string)>();

        //    // We take every row of the results and wrap them in cards wrapped in in MessagingExtensionAttachment objects.
        //    // The Preview is optional, if it includes a Tap, that will trigger the OnTeamsMessagingExtensionSelectItemAsync event back on this bot.
        //    var card = new ThumbnailCard
        //    {
        //        Title = $"{packageId}, {version}",
        //        Subtitle = description,
        //        Buttons = new List<CardAction>
        //            {
        //                new CardAction { Type = ActionTypes.OpenUrl, Title = "Nuget Package", Value = $"https://www.nuget.org/packages/{packageId}" },
        //                new CardAction { Type = ActionTypes.OpenUrl, Title = "Project", Value = projectUrl },
        //            },
        //    };

        //    if (!string.IsNullOrEmpty(iconUrl))
        //    {
        //        card.Images = new List<CardImage>() { new CardImage(iconUrl, "Icon") };
        //    }

        //    var attachment = new MessagingExtensionAttachment
        //    {
        //        ContentType = ThumbnailCard.ContentType,
        //        Content = card,
        //    };

        //    return Task.FromResult(new MessagingExtensionResponse
        //    {
        //        ComposeExtension = new MessagingExtensionResult
        //        {
        //            Type = "result",
        //            AttachmentLayout = "list",
        //            Attachments = new List<MessagingExtensionAttachment> { attachment }
        //        }
        //    });
        //}


        private bool IsTeamInformationUpdated(IConversationUpdateActivity activity)
        {
            if (activity == null)
            {
                return false;
            }

            var channelData = activity.GetChannelData<TeamsChannelData>();
            if (channelData == null)
            {
                return false;
            }

            return CompanyCommunicatorBot.TeamRenamedEventType.Equals(channelData.EventType, StringComparison.OrdinalIgnoreCase);
        }
    }
}