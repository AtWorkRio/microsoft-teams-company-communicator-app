using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.Teams.Apps.CompanyCommunicator
{
    public class AtWorkRioIdentityOptions
    {
        public string Authority { get; set; }
        public string AuthorizationUrl { get; set; }
        public string ApiName { get; set; }
        public string ApiSecret { get; set; }

    }
}
