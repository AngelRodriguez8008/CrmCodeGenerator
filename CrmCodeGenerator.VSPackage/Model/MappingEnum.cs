using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk.Metadata;
using CrmCodeGenerator.VSPackage.Helpers;

namespace CrmCodeGenerator.VSPackage.Model
{
    [Serializable]
    public class MappingEnum
    {
        public string LogicalName { get; set; }
        public string DisplayName { get; set; }
        public string GlobalName { get; set; }
        public bool IsGlobal { get; set; }   
        public bool IsTwoOption { get; set; }
        public MapperEnumItem[] Items { get; set; }

        private string mappedName;
        public string MappedName
        {
            get => string.IsNullOrWhiteSpace(mappedName) ? DisplayName : mappedName;
            set => mappedName = value;
        }


        public static MappingEnum Parse(object attribute)
        {
            if (attribute == null)
                return null;
            
            if (attribute is EnumAttributeMetadata metadata)
                return Parse(metadata);

            if (attribute is BooleanAttributeMetadata option)
                return Parse(option);

            return null;
        }
        
        public static MappingEnum Parse(StatusAttributeMetadata status, MappingEnum entityStates)
        {
            var enumItems = status.OptionSet.Options
                .OfType<StatusOptionMetadata>()
                .Select(
                    o =>
                    {
                        var label = o.Label.UserLocalizedLabel.Label;
                        var stateLabel = entityStates?.Items.FirstOrDefault(i => i.Value == o.State)?.Name ?? string.Empty;
                        return new MapperEnumItem
                        {
                            Attribute = new CrmPicklistAttribute
                            {
                                DisplayName = label,
                                Value = o.Value ?? 1,
                                State = o.State
                            },
                            Name = $"{stateLabel}_{Naming.GetEnumItemName(label)}"
                        };
                    })
                .ToArray();


            var enm = new MappingEnum
            {
                LogicalName = status.LogicalName,
                DisplayName = Naming.GetProperVariableName(Naming.GetProperVariableName(status.SchemaName)),
                IsGlobal = status.OptionSet.IsGlobal.GetValueOrDefault(false),
                IsTwoOption = false,
                GlobalName = status.OptionSet.Name,
                Items = enumItems
            };
            RenameDuplicates(enm);
            return enm;
        }
    
        public static MappingEnum Parse(EnumAttributeMetadata picklist)
        {
            var enumItems = picklist.OptionSet.Options.Select(
                o =>
                {
                    var label = o.Label.UserLocalizedLabel.Label;
                    return new MapperEnumItem
                    {
                        Attribute = new CrmPicklistAttribute
                        {
                            DisplayName = label,
                            Value = o.Value ?? 1
                        },
                        Name = Naming.GetEnumItemName(label)
                    };
                }).ToArray();

            var enm = new MappingEnum
            {
                LogicalName = picklist.LogicalName,
                DisplayName = Naming.GetProperVariableName(Naming.GetProperVariableName(picklist.SchemaName)),
                IsGlobal = picklist.OptionSet.IsGlobal.GetValueOrDefault(false),
                IsTwoOption = false,
                GlobalName = picklist.OptionSet.Name,
                Items = enumItems
            };
            RenameDuplicates(enm);
            return enm;
        }
        public static MappingEnum Parse(BooleanAttributeMetadata twoOption)
        {
            var enm = new MappingEnum
            { 
                LogicalName = twoOption.LogicalName,
                DisplayName = Naming.GetProperVariableName(Naming.GetProperVariableName(twoOption.SchemaName)),
                IsGlobal = false,
                IsTwoOption = true,
                GlobalName = null,
                Items = new []
                {
                    MapBoolOption(twoOption.OptionSet.TrueOption),
                    MapBoolOption(twoOption.OptionSet.FalseOption)
                }
            };
            RenameDuplicates(enm);
            return enm;
        }
        private static void RenameDuplicates(MappingEnum enm)
        {
            var duplicates = new Dictionary<string, int>();
            foreach (var i in enm.Items)
                if (duplicates.ContainsKey(i.Name))
                {
                    duplicates[i.Name] = duplicates[i.Name] + 1;
                    i.Name += "_" + duplicates[i.Name];
                }
                else
                    duplicates[i.Name] = 1;
        }

        private static MapperEnumItem MapBoolOption(OptionMetadata option)
        {
            var label = option.Label.UserLocalizedLabel.Label;
            var results = new MapperEnumItem
            {
                Attribute = new CrmPicklistAttribute
                {
                    DisplayName = label,
                    Value = option.Value ?? -1
                },
                Name = Naming.GetEnumItemName(label)
            };
            return results;
        }
    }

    [Serializable]
    public class MapperEnumItem
    {
        public CrmPicklistAttribute Attribute { get; set; }
        public string Name { get; set; }
        public int Value => Attribute.Value;
    }
}
