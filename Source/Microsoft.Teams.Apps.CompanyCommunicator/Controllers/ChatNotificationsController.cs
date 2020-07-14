
namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.NotificationDelivery;
    using Microsoft.Teams.Apps.CompanyCommunicator.Repositories.Extensions;


    [Authorize(AuthenticationSchemes = PolicyNames.AtWorkRioIdentity, Policy = PolicyNames.MustBeValidUpnPolicy)]
    [Route("api/chatNotifications")]
    public class ChatNotificationsController : ControllerBase
    {
        private readonly NotificationDataRepository notificationDataRepository;
        private readonly TeamDataRepository teamDataRepository;
        private readonly NotificationDelivery notificationDelivery;
        private readonly DraftNotificationPreviewService draftNotificationPreviewService;

        /// <summary>
        /// Initializes a new instance of the <see cref="ChatNotificationsController"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository instance.</param>
        /// <param name="teamDataRepository">Team data repository instance.</param>
        /// <param name="notificationDelivery">TODO</param>
        /// <param name="draftNotificationPreviewService">Draft notification preview service.</param>
        public ChatNotificationsController(
            NotificationDataRepository notificationDataRepository,
            TeamDataRepository teamDataRepository,
            NotificationDelivery notificationDelivery,
            DraftNotificationPreviewService draftNotificationPreviewService)
        {
            this.notificationDataRepository = notificationDataRepository;
            this.teamDataRepository = teamDataRepository;
            this.notificationDelivery = notificationDelivery;
            this.draftNotificationPreviewService = draftNotificationPreviewService;
        }

        /// <summary>
        /// Create a new draft notification.
        /// </summary>
        /// <param name="notification">A new Draft Notification to be created.</param>
        /// <returns>The newly created notification's id.</returns>
        [HttpPost]
        public async Task<string> Create([FromBody] DraftNotification notification)
        {
            if(string.IsNullOrWhiteSpace(notification.Title))
            {
                return null;
            }

            var draftNotificationId = await this.notificationDataRepository.CreateDraftNotificationAsync(
                notification,
                this.HttpContext.User?.Identity?.Name);

            var draftNotificationEntity = await this.notificationDataRepository.GetAsync(
                PartitionKeyNames.NotificationDataTable.DraftNotificationsPartition,
                draftNotificationId);
            if (draftNotificationEntity == null)
            {
                return null;
            }

            await this.notificationDelivery.SendAsync(draftNotificationEntity);

            return draftNotificationId;
        }
    }
}
