using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk.Metadata;
using CrmCodeGenerator.VSPackage.Helpers;

namespace CrmCodeGenerator.VSPackage.Model
{
    [Serializable]
    public class MappingEntity
    {
        public CrmEntityAttribute Attribute { get; set; }
        public bool IsIntersect { get; set; }
        public int? TypeCode { get; set; }
        public MappingField[] Fields { get; set; }
        public MappingEnum States { get; set; }
        public MappingEnum Status { get; set; }      
        public MappingEnum[] Enums { get; set; }
        public MappingRelationship1N[] RelationshipsOneToMany { get; set; }
        public MappingRelationshipN1[] RelationshipsManyToOne { get; set; }

        public string LogicalName => Attribute.LogicalName;

        private string mappedName;
        public string MappedName
        {
            get => string.IsNullOrWhiteSpace(mappedName) ? HybridName : mappedName;
            set => mappedName = value;
        }
        public string DisplayName { get; set; }
        public string HybridName { get; set; }
        public string StateName { get; set; }
        public MappingField PrimaryKey { get; set; }
        public string PrimaryKeyProperty { get; set; }
        public string PrimaryNameAttribute { get; set; }
        public string Description { get; set; }
        public string DescriptionXmlSafe => Naming.XmlEscape(Description);
        public string Plural => Naming.GetPluralName(DisplayName);
        public MappingEntity()
        {
            Description = "";
        }

        public static MappingEntity Parse(EntityMetadata entityMetadata, EntityMappingSetting entityMapping)
        {
            var primaryKey = entityMetadata.PrimaryIdAttribute;
            var primaryName = entityMetadata.PrimaryNameAttribute;

            var entity = new MappingEntity
            {
                Attribute = new CrmEntityAttribute
                {
                    PrimaryKey = primaryKey,
                    LogicalName = entityMetadata.LogicalName
                },
                TypeCode = entityMetadata.ObjectTypeCode,
                IsIntersect = entityMetadata.IsIntersect == true
            };
            // entity.DisplayName = Helper.GetProperVariableName(entityMetadata.SchemaName);
            entity.DisplayName = Naming.GetProperEntityName(entityMetadata.SchemaName);
            entity.HybridName = Naming.GetProperHybridName(entityMetadata.SchemaName, entityMetadata.LogicalName);
            entity.StateName = entity.MappedName + "State";
            entity.MappedName = entityMapping?.CodeName;

            if (entityMetadata.Description?.UserLocalizedLabel != null)
                entity.Description = entityMetadata.Description.UserLocalizedLabel.Label;

            IEnumerable<AttributeMetadata> attributes = entityMetadata.Attributes;

            if (entityMapping?.Attributes != null)
                attributes = attributes.Where(a => /*a is StateAttributeMetadata || 
                                                   a is StatusAttributeMetadata || */
                                                   a.LogicalName == primaryKey ||
                                                   a.LogicalName == primaryName ||
                                                   entityMapping.Attributes.ContainsKey(a.LogicalName));

            var fields = attributes
                .Where(a => a.AttributeOf == null)
                .Select(a => MappingField.Parse(a, entity))
                .ToList();

            fields.ForEach(f =>
                    {
                        if (f.DisplayName == entity.DisplayName)
                            f.DisplayName += "1";

                        if (entityMapping?.Attributes != null)
                        {
                            entityMapping.Attributes.TryGetValue(f.LogicalName, out var mappedName);
                            f.MappedName = mappedName;
                        }
                        //f.HybridName = Naming.GetProperHybridFieldName(f.DisplayName, f.Attribute);
                    }
                );

            AddLookupFields(fields);
            AddEnityImageCRM2013(fields);
            entity.Fields = fields.ToArray();

            var mappingEnums = new List<MappingEnum>();
            var stateAttribute = attributes.FirstOrDefault(a => a is StateAttributeMetadata);
            if (stateAttribute != null)
            {
                entity.States = MappingEnum.Parse(stateAttribute as EnumAttributeMetadata);
                mappingEnums.Add(entity.States);
            }
            var statusAttribute = attributes.FirstOrDefault(a => a is StatusAttributeMetadata);
            if (statusAttribute != null)
            {
                entity.Status = MappingEnum.Parse(statusAttribute as StatusAttributeMetadata, entity.States);
                mappingEnums.Add(entity.Status);
            }

            var picklists = attributes
                .Where(a => a != null && (a is PicklistAttributeMetadata || a is MultiSelectPicklistAttributeMetadata || a is BooleanAttributeMetadata))
                .Select(MappingEnum.Parse)
                .Where(enm => enm != null);

            mappingEnums.AddRange(picklists);
            mappingEnums.ForEach(enm =>
            {
                string mappedName = null;
                entityMapping?.Attributes?.TryGetValue(enm.LogicalName, out mappedName);
                enm.MappedName = mappedName;
            });

            entity.Enums = mappingEnums.ToArray();

            entity.PrimaryKey = entity.Fields.FirstOrDefault(f => f.Attribute.LogicalName == entity.Attribute.PrimaryKey);
            entity.PrimaryKeyProperty = entity.PrimaryKey?.DisplayName;
            entity.PrimaryNameAttribute = primaryName;

            entity.RelationshipsOneToMany = entityMetadata.OneToManyRelationships
                .Select(r => MappingRelationship1N.Parse(r, entity.Fields))
                .Where(r => r != null)
                .ToArray();

            entity.RelationshipsOneToMany.ToList().ForEach(r => RenameDuplicates(r, entity));

            entity.RelationshipsManyToOne = entityMetadata.ManyToOneRelationships
                .Select(r => MappingRelationshipN1.Parse(r, entity.Fields))
                .Where(r => r != null)
                .ToArray();

            entity.RelationshipsManyToOne.ToList().ForEach(r => RenameDuplicates(r, entity));

            var relationshipsManyToMany = entityMetadata.ManyToManyRelationships
                .Select(r => MappingRelationshipMN.Parse(r, entity.LogicalName))
                .ToList();

            var selfReferenced = relationshipsManyToMany.Where(r => r.IsSelfReferenced).ToList();
            foreach (var referecned in selfReferenced)
            {
                var referencing = (MappingRelationshipMN)referecned.Clone();
                referencing.DisplayName = "Referencing" + Naming.GetProperVariableName(referecned.SchemaName);
                referencing.EntityRole = "Microsoft.Xrm.Sdk.EntityRole.Referencing";
                relationshipsManyToMany.Add(referencing);
            }
            relationshipsManyToMany.ForEach(r => RenameDuplicates(r, entity));
            entity.RelationshipsManyToMany = relationshipsManyToMany.OrderBy(r => r.DisplayName).ToArray();

            return entity;
        }

        public bool HasCommonStates =>
            States != null && States.LogicalName == "statecode" && States.Items.Length == 2 &&
            States.Items.Any(s => s.Value == 0 && s.Name == "Active") &&
            States.Items.Any(s => s.Value == 1 && s.Name == "Inactive");
        public bool HasCommonStatus =>
            Status != null && Status.LogicalName == "statuscode" && Status.Items.Length == 2 &&
                   Status.Items.Any(s => s.Value == 1 && s.Name == "Active_Active") &&
                   Status.Items.Any(s => s.Value == 2 && s.Name == "Inactive_Inactive");
    public bool IsProcessInstance => Attribute.PrimaryKey == "businessprocessflowinstanceid";
        public bool IsActivity => Attribute.PrimaryKey == "activityid";

        private static void RenameDuplicates(MappingRelationship relation, MappingEntity entity)
        {
            var newName = relation.DisplayName;

            if (newName == entity.DisplayName || newName == entity.HybridName)
                newName = relation.DisplayName += "1";

            if (entity.Fields.Any(e => e.DisplayName == newName))
            {
                newName = relation.DisplayName += "2";
            }
            relation.DisplayName = newName;
        }

        private static void AddLookupFields(List<MappingField> fields)
        {
            var fieldsIterator = fields.Where(e => e.Attribute.IsLookup).ToArray();
            foreach (var lookup in fieldsIterator)
            {
                var nameField = new MappingField
                {
                    Attribute = new CrmPropertyAttribute
                    {
                        IsLookup = false,
                        Type = lookup.Attribute.Type,
                        LogicalName = lookup.Attribute.LogicalName + "Name",
                        IsEntityReferenceHelper = true
                    },
                    DisplayName = lookup.DisplayName + "Name",
                    HybridName = lookup.HybridName + "Name",
                    MappedName = lookup.MappedName + "Name",
                    FieldType = AttributeTypeCode.EntityName,
                    IsValidForUpdate = false,
                    GetMethod = "",
                    PrivatePropertyName = lookup.PrivatePropertyName + "Name"
                };

                if (fields.Count(f => f.DisplayName == nameField.DisplayName) == 0)
                    fields.Add(nameField);

                if (!string.IsNullOrEmpty(lookup.LookupSingleType))
                    continue;

                var typeField = new MappingField
                {
                    Attribute = new CrmPropertyAttribute
                    {
                        IsLookup = false,
                        Type = lookup.Attribute.Type,
                        LogicalName = lookup.Attribute.LogicalName + "Type",
                        IsEntityReferenceHelper = true
                    },
                    DisplayName = lookup.DisplayName + "Type",
                    HybridName = lookup.HybridName + "Type",
                    MappedName = lookup.MappedName + "Type",
                    FieldType = AttributeTypeCode.EntityName,
                    IsValidForUpdate = false,
                    GetMethod = "",
                    PrivatePropertyName = lookup.PrivatePropertyName + "Type"
                };

                if (fields.Count(f => f.DisplayName == typeField.DisplayName) == 0)
                    fields.Add(typeField);
            }
        }
        private static void AddEnityImageCRM2013(List<MappingField> fields)
        {
            var hasImage = fields.Any(f => f.DisplayName.Equals("EntityImageId"));
            if (!hasImage)
                return;

            var image = new MappingField
            {
                Attribute = new CrmPropertyAttribute
                {
                    IsLookup = false,
                    Type = "Image",
                    LogicalName = "entityimage",
                    IsEntityReferenceHelper = false
                },
                DisplayName = "EntityImage",
                HybridName = "EntityImage",
                TargetTypeForCrmSvcUtil = "byte[]",
                IsValidForUpdate = true,
                Description = "",  // TODO there is an Description for this entityimage, Need to figure out how to read it from the server
                GetMethod = ""
            };
            SafeAddField(fields, image);

            var imageTimestamp = new MappingField
            {
                Attribute = new CrmPropertyAttribute
                {
                    IsLookup = false,
                    Type =  AttributeTypeCode.BigInt.ToString(),
                    LogicalName = "entityimage_timestamp",
                    IsEntityReferenceHelper = false
                },
                DisplayName = "EntityImage_Timestamp",
                HybridName = "EntityImage_Timestamp",
                TargetTypeForCrmSvcUtil = "System.Nullable<long>",
                FieldType = AttributeTypeCode.BigInt,
                IsValidForUpdate = false,
                IsValidForCreate = false,
                Description = " ",  // CrmSvcUtil provides an empty description for this EntityImage_TimeStamp
                GetMethod = ""
            };
            SafeAddField(fields, imageTimestamp);

            var imageUrl = new MappingField
            {
                Attribute = new CrmPropertyAttribute
                {
                    IsLookup = false,
                    Type =  AttributeTypeCode.String.ToString(),
                    LogicalName = "entityimage_url",
                    IsEntityReferenceHelper = false
                },
                DisplayName = "EntityImage_URL",
                HybridName = "EntityImage_URL",
                TargetTypeForCrmSvcUtil = "string",
                FieldType = AttributeTypeCode.String,
                IsValidForUpdate = false,
                IsValidForCreate = false,
                Description = " ",   // CrmSvcUtil provides an empty description for this EntityImage_URL
                GetMethod = ""
            };
            SafeAddField(fields, imageUrl);
        }
        private static void SafeAddField(List<MappingField> fields, MappingField image)
        {
            if (fields.All(f => f.DisplayName != image.DisplayName))
                fields.Add(image);
        }
        public MappingRelationshipMN[] RelationshipsManyToMany { get; set; }
    }
}
