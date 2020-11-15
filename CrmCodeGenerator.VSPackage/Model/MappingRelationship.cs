#region Version
// Last Change: Rodriguez Mustelier Angel (rodang) - 2019-01-28 15:39
// 
#endregion
using System;

namespace CrmCodeGenerator.VSPackage.Model
{
    [Serializable]
    public class MappingRelationship
    {
        public CrmRelationshipAttribute Attribute { get; set; }
        public string DisplayName { get; set; }
        public string ForeignKey { get; set; }
        public string SchemaName { get; set; }
        public string HybridName { get; set; }
        public string PrivateName { get; set; }
        public string EntityRole { get; set; }
        public string Type { get; set; }
        public MappingEntity ToEntity { get; set; }
    }
}