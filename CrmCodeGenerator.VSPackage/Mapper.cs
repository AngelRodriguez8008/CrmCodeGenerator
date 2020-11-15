using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk;
using CrmCodeGenerator.VSPackage.Model;
using CrmCodeGenerator.VSPackage.Xrm;

namespace CrmCodeGenerator.VSPackage
{
    public delegate void MapperHandler(object sender, MapperEventArgs e);

    public class Mapper
    {
        public Settings Settings { get; set; }

        private EntityMetadata[] allEntities;

        public Mapper() { }
        public Mapper(Settings settings)
        {
            Settings = settings;
        }

        #region event handler
        public event MapperHandler Message;

        protected void OnMessage(string message, string extendedMessage = "")
        {
            Message?.Invoke(this, new MapperEventArgs { Message = message, MessageExtended = extendedMessage });
        }
        #endregion

        public Context CreateContext()
        { 
            var entitiesMapping = Settings.MappingSettings?.Entities;
            if (allEntities == null)
            {
                using (var service = QuickConnection.Connect(Settings))
                {
                    OnMessage("Gathering metadata, this may take a few minutes...");
                    allEntities = GetMetadataFromServer(service);
                    OnMessage($"Entities Metadata retrived from server: {allEntities.Length}");
                }
            }

            var selectedEntities = GetSelectedEntities(allEntities);
            OnMessage($"Selected Entities Metadata: {selectedEntities.Count}");
           
            var entities = GetEntities(selectedEntities, entitiesMapping);
            var enums = GetGlobalEnums(entities);
            var context = new Context
            {
                Entities = entities,
                Enums = enums
            };
            SortEntities(context);
            SortEnums(context);
            return context;
        }

        public EntityMetadata[] GetMetadataFromServer(IOrganizationService service)
        {
            bool includeUnpublish = Settings.IncludeUnpublish;
            var selection = Settings.EntitiesSelected;
            if (selection.Count > 20)
                return service.GetAllEntitiesMetadata(includeUnpublish);

            var selectedEntities = selection.ToList();
            var isActivityPartySelected = selectedEntities.Any(x => x.Equals("activityparty"));
            if (isActivityPartySelected == false)
            {
                selectedEntities.Add("activityparty");
            }

            var result = service.GetEntitiesMetadata(selectedEntities, includeUnpublish);
            return result;
        }

        public void SortEntities(Context context)
        {
            context.Entities = context.Entities.OrderBy(e => e.DisplayName).ToArray();

            foreach (var e in context.Entities)
                e.Enums = e.Enums.OrderBy(en => en.DisplayName).ToArray();

            foreach (var e in context.Entities)
                e.Fields = e.Fields.OrderBy(f => f.DisplayName).ToArray();

            foreach (var e in context.Entities)
                e.RelationshipsOneToMany = e.RelationshipsOneToMany.OrderBy(r => r.LogicalName).ToArray();

            foreach (var e in context.Entities)
                e.RelationshipsManyToOne = e.RelationshipsManyToOne.OrderBy(r => r.LogicalName).ToArray();
        }

        public void SortEnums(Context context)
        {
            context.Enums = context.Enums.OrderBy(e => e.DisplayName).ToArray();
        }

        public MappingEnum[] GetGlobalEnums(MappingEntity[] entities)
        {
            var uniqueNames = new List<string>();
            List<MappingEnum> globalEnums = new List<MappingEnum>();

            foreach (MappingEntity entity in entities)
            {
                foreach (MappingEnum e in entity.Enums.Where(w => w.IsGlobal))
                {
                    if (!uniqueNames.Contains(e.GlobalName))
                    {
                        globalEnums.Add(e);
                        uniqueNames.Add(e.GlobalName);
                    }
                }
            }

            return globalEnums.ToArray();
        }

        public static MappingEntity[] GetEntities(IEnumerable<EntityMetadata> entitiesMetadata, Dictionary<string, EntityMappingSetting> entitiesMapping)
        {
            List<MappingEntity> mappedEntities = entitiesMetadata.Select(e =>
                                                    {
                                                        EntityMappingSetting mapping = null;
                                                        entitiesMapping?.TryGetValue(e.LogicalName, out mapping);
                                                        var result = MappingEntity.Parse(e, mapping);
                                                        return result;
                                                    })
                                                  .OrderBy(e => e.DisplayName)
                                                  .ToList();

            ExcludeRelationshipsNotIncluded(mappedEntities);
            foreach (MappingEntity ent in mappedEntities)
            {
                foreach (var field in ent.Fields)
                {
                    var lookupSingleType = field.LookupSingleType;
                    if (string.IsNullOrWhiteSpace(lookupSingleType))
                        continue;

                    var mappedEntity = mappedEntities.FirstOrDefault(e => e.LogicalName == lookupSingleType);
                    var mappedName = mappedEntity?.MappedName;
                    if (string.IsNullOrWhiteSpace(mappedName) == false)
                        field.LookupSingleType = mappedName;
                }
                foreach (var rel in ent.RelationshipsOneToMany)
                {
                    rel.ToEntity = mappedEntities.FirstOrDefault(e => e.LogicalName.Equals(rel.Attribute.ToEntity));
                }
                foreach (var rel in ent.RelationshipsManyToOne)
                {
                    rel.ToEntity = mappedEntities.FirstOrDefault(e => e.LogicalName.Equals(rel.Attribute.ToEntity));
                }
                foreach (var rel in ent.RelationshipsManyToMany)
                {
                    rel.ToEntity = mappedEntities.FirstOrDefault(e => e.LogicalName.Equals(rel.Attribute.ToEntity));
                }
            }

            return mappedEntities.ToArray();
        }

        public List<EntityMetadata> GetSelectedEntities(EntityMetadata[] entitiesMetadata)
        {
            string[] inMapping = Settings.MappingSettings?.Entities.Keys.ToArray();

            IEnumerable<string> selection = inMapping ?? Settings.EntitiesSelected.ToArray();

            if (Settings.IncludeNonStandard == false)
            {
                var nonStandard = EntityHelper.NonStandard.Except(selection);
                selection = selection.Except(nonStandard);
            }

            var selectedMetadatas = entitiesMetadata.Where(e => selection.Contains(e.LogicalName)).ToList();
            var activityParty = "activityparty";
            if (selection.Contains(activityParty) &&
                selectedMetadatas.Any(e => e.IsActivity == true || e.IsActivityParty == true))
            {
                if (!selectedMetadatas.Any(e => e.LogicalName.Equals(activityParty)))
                {
                    var activityparty = entitiesMetadata.SingleOrDefault(r => r.LogicalName.Equals(activityParty));
                    if (activityparty != null)
                        selectedMetadatas.Add(activityparty);
                }
            }
            return selectedMetadatas;
        }

        private static void ExcludeRelationshipsNotIncluded(List<MappingEntity> mappedEntities)
        {
            foreach (var ent in mappedEntities)
            {
                ent.RelationshipsOneToMany = ent.RelationshipsOneToMany.ToList().Where(r => mappedEntities.Select(m => m.LogicalName).Contains(r.Type)).ToArray();
                ent.RelationshipsManyToOne = ent.RelationshipsManyToOne.ToList().Where(r => mappedEntities.Select(m => m.LogicalName).Contains(r.Type)).ToArray();
                ent.RelationshipsManyToMany = ent.RelationshipsManyToMany.ToList().Where(r => mappedEntities.Select(m => m.LogicalName).Contains(r.Type)).ToArray();
            }
        }
    }
}
