using CrmCodeGenerator.VSPackage.Helpers;
using Microsoft.Xrm.Sdk.Metadata;
using System;

namespace CrmCodeGenerator.VSPackage.Model
{
    [Serializable]
    public class MappingField
    {
        public CrmPropertyAttribute Attribute { get; set; }
        public MappingEntity Entity { get; set; }
        public string AttributeOf { get; set; }
        public MappingEnum EnumData { get; set; }
        public AttributeTypeCode FieldType { get; set; }
        public string FieldTypeString { get; set; }
        public bool IsValidForCreate { get; set; }
        public bool IsValidForRead { get; set; }
        public bool IsValidForUpdate { get; set; }
        public bool IsActivityParty { get; set; }
        public bool IsStateCode { get; set; }
        public bool IsStatusCode { get; set; }
        public bool IsDeprecated { get; set; }
        public bool IsOptionSet { get; private set; }
        public bool IsOptionSetCollection { get; private set; }
        public bool IsTwoOption { get; private set; }
        public string DeprecatedVersion {get ; set; }
        public string LookupSingleType { get; set; }
        public bool IsPrimaryKey { get; set; }
        public bool IsRequired { get; set; }
        public int? MaxLength { get; set; }
        public decimal? Min { get; set; }
        public decimal? Max { get; set; }
        public string PrivatePropertyName { get; set; }
        public string DisplayName { get; set; }
        public string HybridName { get; set; }
        public string LogicalName { get; set; }
        public string StateName { get; set; }
        public string TargetTypeForCrmSvcUtil { get; set; }
        public string Description { get; set; }
        public string DescriptionXmlSafe => Naming.XmlEscape(Description);
        public string Label { get; set; }

        private string mappedName;
        public string MappedName
        {
            get => string.IsNullOrWhiteSpace(mappedName) ? HybridName : mappedName;
            set => mappedName = value;
        }

        public MappingField()
        {
            IsValidForUpdate = false;
            IsValidForCreate = false;
            IsDeprecated = false;
            Description = "";
        }
        public static MappingField Parse(AttributeMetadata attribute, MappingEntity entity)
        {
            var result = new MappingField
            {
                Entity = entity,
                LogicalName = attribute.LogicalName,
                AttributeOf = attribute.AttributeOf,
                IsPrimaryKey = attribute.IsPrimaryId == true,
                IsValidForCreate = attribute.IsValidForCreate == true,
                IsValidForRead = attribute.IsValidForRead == true,
                IsValidForUpdate = attribute.IsValidForUpdate == true,
                IsActivityParty = attribute.AttributeType == AttributeTypeCode.PartyList,
                IsStateCode = attribute.AttributeType == AttributeTypeCode.State,
                IsStatusCode = attribute.AttributeType == AttributeTypeCode.Status,
                IsOptionSet = attribute.AttributeType == AttributeTypeCode.Picklist,
                IsTwoOption = attribute.AttributeType == AttributeTypeCode.Boolean,
                IsOptionSetCollection = attribute.AttributeType == AttributeTypeCode.Virtual &&  attribute.AttributeTypeName == AttributeTypeDisplayName.MultiSelectPicklistType,
                DeprecatedVersion = attribute.DeprecatedVersion,
                IsDeprecated = !string.IsNullOrWhiteSpace(attribute.DeprecatedVersion)
            };
        
            result.DisplayName = Naming.GetProperVariableName(attribute);
            result.PrivatePropertyName = Naming.GetEntityPropertyPrivateName(attribute.SchemaName);
            result.HybridName = Naming.GetProperHybridFieldName(result.DisplayName, result.Attribute);
            result.IsRequired = attribute.RequiredLevel?.Value == AttributeRequiredLevel.ApplicationRequired;

            var typeName = attribute.AttributeTypeName?.Value;
            if(typeName.EndsWith("Type"))
                typeName = typeName.Remove(typeName.Length- 4);
           
            result.Attribute = 
                new CrmPropertyAttribute 
                {
                    LogicalName = attribute.LogicalName,
                    IsLookup = attribute.AttributeType == AttributeTypeCode.Lookup ||
                               attribute.AttributeType == AttributeTypeCode.Customer || 
                               attribute.AttributeType == AttributeTypeCode.Owner,
                    Type = typeName
                };

            if (attribute is PicklistAttributeMetadata picklist)
                result.EnumData = MappingEnum.Parse(picklist);  
            
            if (attribute is MultiSelectPicklistAttributeMetadata multiSelect) 
                result.EnumData = MappingEnum.Parse(multiSelect);

            if (attribute is LookupAttributeMetadata lookup)
                if (lookup.Targets.Length == 1)
                    result.LookupSingleType = lookup.Targets[0];

            ParseMinMaxValues(attribute, result);

            if (attribute.AttributeType != null)
            {
                result.FieldType = attribute.AttributeType.Value;
                result.FieldTypeString = attribute.AttributeType.Value.ToString("F");
            }
            
            if(attribute.Description?.UserLocalizedLabel != null)
                result.Description = attribute.Description.UserLocalizedLabel.Label;

            if (attribute.DisplayName?.UserLocalizedLabel != null)
                result.Label = attribute.DisplayName.UserLocalizedLabel.Label;

            string typeForCrmSvcUtil = GetTargetType(result);
            result.TargetTypeForCrmSvcUtil = typeForCrmSvcUtil;
            
            return result;
        }

        private static void ParseMinMaxValues(AttributeMetadata attribute, MappingField result)
        {
            switch (attribute)
            {
                case StringAttributeMetadata metaStr:
                    result.MaxLength = metaStr.MaxLength ?? -1;
                    break;
                case MemoAttributeMetadata metaMemo:
                    result.MaxLength = metaMemo.MaxLength ?? -1;
                    break;
                case IntegerAttributeMetadata metaInt:
                    result.Min = metaInt.MinValue ?? -1;
                    result.Max = metaInt.MaxValue ?? -1;
                    break;
                case DecimalAttributeMetadata metaDecimal:
                    result.Min = metaDecimal.MinValue ?? -1;
                    result.Max = metaDecimal.MaxValue ?? -1;
                    break;
                case MoneyAttributeMetadata metaMoney:
                    result.Min = metaMoney.MinValue != null ? (decimal)metaMoney.MinValue.Value : -1;
                    result.Max = metaMoney.MaxValue != null ? (decimal)metaMoney.MaxValue.Value : -1;
                    break;
                case DoubleAttributeMetadata metaDouble:
                    result.Min = metaDouble.MinValue != null ? (decimal)metaDouble.MinValue.Value : -1;
                    result.Max = metaDouble.MaxValue != null ? (decimal)metaDouble.MaxValue.Value : -1;
                    break;
            }
        }

        private static string GetTargetType(MappingField field)
        {
            if (field.IsPrimaryKey)
                return "System.Nullable<System.Guid>";

            switch (field.FieldType)
            {
                case AttributeTypeCode.Virtual when field.Attribute.Type == "MultiSelectPicklist":
                    return "Microsoft.Xrm.Sdk.OptionSetValueCollection";
                case AttributeTypeCode.Picklist:
                    return "Microsoft.Xrm.Sdk.OptionSetValue";
                case AttributeTypeCode.BigInt:
                    return "System.Nullable<long>";
                case AttributeTypeCode.Integer:
                    return "System.Nullable<int>";
                case AttributeTypeCode.Boolean:
                    return "System.Nullable<bool>";
                case AttributeTypeCode.DateTime:
                    return "System.Nullable<System.DateTime>";
                case AttributeTypeCode.Decimal:
                    return "System.Nullable<decimal>";
                case AttributeTypeCode.Money:
                    return "Microsoft.Xrm.Sdk.Money";
                case AttributeTypeCode.Double:
                    return "System.Nullable<double>";
                case AttributeTypeCode.Uniqueidentifier:
                    return "System.Nullable<System.Guid>";
                case AttributeTypeCode.Lookup:
                case AttributeTypeCode.Owner:
                case AttributeTypeCode.Customer:
                    return "Microsoft.Xrm.Sdk.EntityReference";
                case AttributeTypeCode.State:
                    return "System.Nullable<" + field.Entity.StateName + ">";
                case AttributeTypeCode.Status:
                    return "Microsoft.Xrm.Sdk.OptionSetValue";
                case AttributeTypeCode.Memo:
                case AttributeTypeCode.Virtual:
                case AttributeTypeCode.EntityName:
                case AttributeTypeCode.String:
                    return "string";
                case AttributeTypeCode.PartyList:
                    return "System.Collections.Generic.IEnumerable<ActivityParty>";
                case AttributeTypeCode.ManagedProperty:
                    return "Microsoft.Xrm.Sdk.BooleanManagedProperty";
                default:
                    return "object";
            }
        }
        
        public string TargetType
        {
            get
            {
                if (IsPrimaryKey)
                    return "Guid";

                switch (FieldType)
                {
                    case AttributeTypeCode.Picklist:
                        return $"Enums.{EnumData.MappedName}?";

                    case AttributeTypeCode.BigInt:
                    case AttributeTypeCode.Integer:
                        return "int?";

                    case AttributeTypeCode.Boolean:
                        return "bool?";

                    case AttributeTypeCode.DateTime:
                        return "DateTime?";

                    case AttributeTypeCode.Decimal:
                    case AttributeTypeCode.Money:
                        return "decimal?";

                    case AttributeTypeCode.Double:
                        return "double?";

                    case AttributeTypeCode.Uniqueidentifier:
                    case AttributeTypeCode.Lookup:
                    case AttributeTypeCode.Owner:
                    case AttributeTypeCode.Customer:
                        return "Guid?";

                    case AttributeTypeCode.State:
                    case AttributeTypeCode.Status:
                        return "int";

                    case AttributeTypeCode.Memo:
                    case AttributeTypeCode.Virtual:
                    case AttributeTypeCode.EntityName:
                    case AttributeTypeCode.String:
                        return "string";

                    default:
                        return "object";
                }
            }
        }

        public string GetMethod { get; set; }

        public string SetMethodCall
        {
            get
            {
                var methodName = "";

                switch (FieldType)
                {
                    case AttributeTypeCode.Picklist:
                        methodName = "SetPicklist";
                        break;
                    case AttributeTypeCode.BigInt:
                    case AttributeTypeCode.Integer:
                        methodName = "SetValue<int?>";
                        break;
                    case AttributeTypeCode.Boolean:
                        methodName = "SetValue<bool?>";
                        break;
                    case AttributeTypeCode.DateTime:
                        methodName = "SetValue<DateTime?>";
                        break;
                    case AttributeTypeCode.Decimal:
                        methodName = "SetValue<decimal?>";
                        break;
                    case AttributeTypeCode.Money:
                        methodName = "SetMoney";
                        break;
                    case AttributeTypeCode.Memo:
                    case AttributeTypeCode.String:
                        methodName = "SetValue<string>";
                        break;
                    case AttributeTypeCode.Double:
                        methodName = "SetValue<double?>";
                        break;
                    case AttributeTypeCode.Uniqueidentifier:
                        methodName = "SetValue<Guid?>";
                        break;
                    case AttributeTypeCode.Lookup:
                        methodName = "SetLookup";
                        break;
                    //methodName = "SetLookup"; break;
                    case AttributeTypeCode.Virtual:
                        methodName = "SetValue<string>";
                        break;
                    case AttributeTypeCode.Customer:
                        methodName = "SetCustomer";
                        break;
                    case AttributeTypeCode.Status:
                        methodName = "";
                        break;
                    case AttributeTypeCode.EntityName:
                        methodName = "SetEntityNameReference";
                        break;
                    case AttributeTypeCode.State:
                    case AttributeTypeCode.Owner:
                    default:
                        return "";
                }

                if (methodName == "" || !IsValidForUpdate)
                    return "";

                if (FieldType == AttributeTypeCode.Picklist)
                    return $"{methodName}(\"{Attribute.LogicalName}\", (int?)value);";

                if (FieldType == AttributeTypeCode.Lookup || FieldType == AttributeTypeCode.Customer)
                    if (string.IsNullOrEmpty(LookupSingleType))
                        return $"{methodName}(\"{Attribute.LogicalName}\", {DisplayName}Type, value);";
                    else
                        return $"{methodName}(\"{Attribute.LogicalName}\", \"{LookupSingleType}\", value);";

                return $"{methodName}(\"{Attribute.LogicalName}\", value);";
            }
        }
    }
}
