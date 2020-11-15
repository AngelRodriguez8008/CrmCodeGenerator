using System;
using System.Linq;
using Microsoft.Xrm.Sdk.Metadata;
using CrmCodeGenerator.VSPackage.Helpers;

namespace CrmCodeGenerator.VSPackage.Model
{
    [Serializable]
    public class MappingRelationshipN1 : MappingRelationship
    {
        public string LogicalName { get; set; }

        public MappingField Property { get; set; }

        public static MappingRelationshipN1 Parse(OneToManyRelationshipMetadata rel, MappingField[] properties)
        {
            var property = properties.FirstOrDefault(p => string.Equals(p.Attribute.LogicalName, rel.ReferencingAttribute, StringComparison.OrdinalIgnoreCase));
            if (property == null)
                return null;

            var propertyName = property.DisplayName;

            var result = new MappingRelationshipN1
            {
                Attribute = new CrmRelationshipAttribute
                {
                    ToEntity = rel.ReferencedEntity,
                    ToKey = rel.ReferencedAttribute,
                    FromEntity = rel.ReferencingEntity,
                    FromKey = rel.ReferencingAttribute,
                    IntersectingEntity = ""
                },

                DisplayName = Naming.GetProperVariableName(rel.SchemaName),
                SchemaName = Naming.GetProperVariableName(rel.SchemaName),
                LogicalName = rel.ReferencingAttribute,
                HybridName = Naming.GetProperVariableName(rel.SchemaName) + "_N1",
                PrivateName = "_n1"+ Naming.GetEntityPropertyPrivateName(rel.SchemaName),
                ForeignKey = propertyName,
                Type = Naming.GetProperVariableName(rel.ReferencedEntity),
                Property = property,
                EntityRole = "null"
            };

            if (rel.ReferencedEntity == rel.ReferencingEntity)
            {
                result.EntityRole = "Microsoft.Xrm.Sdk.EntityRole.Referencing";
                result.DisplayName = "Referencing" + result.DisplayName;
            }

            return result;
        }
    }
}
