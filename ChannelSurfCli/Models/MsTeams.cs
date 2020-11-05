﻿using System.Collections.Generic;

namespace ChannelSurfCli.Models
{
    public class MsTeams
    {
        public class Team
        {
            public List<Channel> value { get; set; }
        }

        public class Channel
        {
            public string id { get; set; }
            public string displayName { get; set; }
            public string description { get; set; } = "";
            public string folderId { get; set; } = "";
        }

        public class User
        {
            public string id { get; set; }
            public string mail { get; set; }
            public string userPrincipalName { get; set; }
            public string displayName { get; set; }
        }
    }
}
