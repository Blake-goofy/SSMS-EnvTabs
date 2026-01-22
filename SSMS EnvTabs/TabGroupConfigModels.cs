using System.Collections.Generic;
using System.Runtime.Serialization;

namespace SSMS_EnvTabs
{
    [DataContract]
    internal sealed class TabGroupConfig
    {
        [DataMember(Name = "groups", IsRequired = false)]
        public List<TabGroupRule> Groups { get; set; } = new List<TabGroupRule>();

        [DataMember(Name = "settings", IsRequired = false)]
        public TabGroupSettings Settings { get; set; } = new TabGroupSettings();
    }

    [DataContract]
    internal sealed class TabGroupSettings
    {
        [DataMember(Name = "enableAutoRename", IsRequired = false)]
        public bool EnableAutoRename { get; set; } = true;

        [DataMember(Name = "enableAutoColor", IsRequired = false)]
        public bool EnableAutoColor { get; set; } = false;
    }

    [DataContract]
    internal sealed class TabGroupRule
    {
        [DataMember(Name = "groupName", IsRequired = false)]
        public string GroupName { get; set; }

        [DataMember(Name = "server", IsRequired = false)]
        public string Server { get; set; }

        [DataMember(Name = "database", IsRequired = false)]
        public string Database { get; set; }

        [DataMember(Name = "priority", IsRequired = false)]
        public int Priority { get; set; } = 0;

        [DataMember(Name = "colorIndex", IsRequired = false)]
        public int ColorIndex { get; set; } = 0;
    }
}
