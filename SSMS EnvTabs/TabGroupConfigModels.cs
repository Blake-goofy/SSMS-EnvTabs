using System.Collections.Generic;
using System.Runtime.Serialization;

namespace SSMS_EnvTabs
{
    [DataContract]
    internal sealed class TabGroupConfig
    {
        [DataMember(Name = "description", IsRequired = false, Order = 0)]
        public string Description { get; set; }

        [DataMember(Name = "version", IsRequired = false, Order = 1)]
        public string Version { get; set; }

        [DataMember(Name = "settings", IsRequired = false, Order = 2)]
        public TabGroupSettings Settings { get; set; } = new TabGroupSettings();

        [DataMember(Name = "documentation", IsRequired = false, Order = 3)]
        public Dictionary<string, string> Documentation { get; set; }

        [DataMember(Name = "colors", IsRequired = false, Order = 4)]
        public Dictionary<string, string> Colors { get; set; }

        [DataMember(Name = "groups", IsRequired = false, Order = 5)]
        public List<TabGroupRule> Groups { get; set; } = new List<TabGroupRule>();
    }

    [DataContract]
    internal sealed class TabGroupSettings
    {
        [DataMember(Name = "enableAutoRename", IsRequired = false, Order = 0)]
        public bool EnableAutoRename { get; set; } = true;

        [DataMember(Name = "enableAutoColor", IsRequired = false, Order = 1)]
        public bool EnableAutoColor { get; set; } = false;

        [DataMember(Name = "autoConfigure", IsRequired = false, Order = 2)]
        public string AutoConfigure { get; set; }

        [DataMember(Name = "enableConfigurePrompt", IsRequired = false, Order = 3)]
        public bool EnableConfigurePrompt { get; set; } = true;
    }

    [DataContract]
    internal sealed class TabGroupRule
    {
        [DataMember(Name = "groupName", IsRequired = false, Order = 0)]
        public string GroupName { get; set; }

        [DataMember(Name = "server", IsRequired = false, Order = 1)]
        public string Server { get; set; }

        [DataMember(Name = "database", IsRequired = false, Order = 2)]
        public string Database { get; set; }

        [DataMember(Name = "priority", IsRequired = false, Order = 3)]
        public int Priority { get; set; } = 0;

        [DataMember(Name = "colorIndex", IsRequired = false, Order = 4)]
        public int ColorIndex { get; set; } = 0;
    }
}
