using System;
using Microsoft.Xrm.Sdk.Metadata;
using CrmCodeGenerator.VSPackage.Helpers;

namespace CrmCodeGenerator.VSPackage.Model
{
    [Serializable]
    public class MappingRelationshipMN : MappingRelationship, ICloneable
    {
        public bool IsSelfReferenced { get; set; }
     
        public static MappingRelationshipMN Parse(ManyToManyRelationshipMetadata rel, string entityLogicalName)
        {
            var result = new MappingRelationshipMN();
            if (rel.Entity1LogicalName == entityLogicalName)
            {
                result.Attribute = new CrmRelationshipAttribute
                {
                    FromEntity = rel.Entity1LogicalName,
                    FromKey = rel.Entity1IntersectAttribute,
                    ToEntity = rel.Entity2LogicalName,
                    ToKey = rel.Entity2IntersectAttribute,
                    IntersectingEntity = rel.IntersectEntityName
                };
            }
            else
            {
                result.Attribute = new CrmRelationshipAttribute
                {
                    ToEntity = rel.Entity1LogicalName,
                    ToKey = rel.Entity1IntersectAttribute,
                    FromEntity = rel.Entity2LogicalName,
                    FromKey = rel.Entity2IntersectAttribute,
                    IntersectingEntity = rel.IntersectEntityName
                };
            }

            result.EntityRole = "null";
            result.SchemaName = Naming.GetProperVariableName(rel.SchemaName);
            result.DisplayName = Naming.GetProperVariableName(rel.SchemaName);
            if (rel.Entity1LogicalName == rel.Entity2LogicalName && rel.Entity1LogicalName == entityLogicalName)
            {
                result.DisplayName = "Referenced" + result.DisplayName;
                result.EntityRole = "Microsoft.Xrm.Sdk.EntityRole.Referenced";
                result.IsSelfReferenced = true;
            }
            if (result.DisplayName == entityLogicalName)
            {
                result.DisplayName += "1";   // this is what CrmSvcUtil does
            }

            result.HybridName = Naming.GetProperVariableName(rel.SchemaName) + "_NN";  
            result.PrivateName = "_nn" + Naming.GetEntityPropertyPrivateName(rel.SchemaName);
            result.ForeignKey =  Naming.GetProperVariableName(result.Attribute.ToKey);
            result.Type = Naming.GetProperVariableName(result.Attribute.ToEntity);

            return result;
        }

        public object Clone()
        {
            var newPerson = (MappingRelationshipMN)MemberwiseClone();
            newPerson.Attribute = (CrmRelationshipAttribute)Attribute.Clone();
            return newPerson;
        }
    }
}
