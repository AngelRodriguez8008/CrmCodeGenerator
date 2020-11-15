using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace CrmCodeGenerator.VSPackage.Xrm
{
    public static class OrganizationServiceExtensions
    {
        public static List<string> GetAllEntityNames(this IOrganizationService service, bool includeUnpublish = true)
        {
            var request = new RetrieveAllEntitiesRequest
            {
                EntityFilters = EntityFilters.Entity,
                RetrieveAsIfPublished = includeUnpublish,
            };
            var response = (RetrieveAllEntitiesResponse)service.Execute(request);
            List<string> result = response.EntityMetadata.Select(e => e.LogicalName).ToList();
            return result;
        }

        public static EntityMetadata[] GetEntitiesMetadata(this IOrganizationService service, IEnumerable<string> entities, bool includeUnpublish = true)
        {
            var results = new List<EntityMetadata>();
            foreach (var entity in entities)
            {
                EntityMetadata entityMetadata = service.GetEntityMetadata(entity, includeUnpublish);
                results.Add(entityMetadata);
            }
            return results.ToArray();
        }

        public static EntityMetadata GetEntityMetadata(this IOrganizationService service, string entity, bool includeUnpublish = true)
        {
            var req = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.All,
                LogicalName = entity,
                RetrieveAsIfPublished = includeUnpublish
            };
            var res = (RetrieveEntityResponse)service.Execute(req);
            var entityMetadata = res.EntityMetadata;
            return entityMetadata;
        }

        public static EntityMetadata[] GetAllEntitiesMetadata(this IOrganizationService service, bool includeUnpublish = true)
        {
            OrganizationRequest request = new RetrieveAllEntitiesRequest
            {
                EntityFilters = EntityFilters.All,
                RetrieveAsIfPublished = includeUnpublish
            };
            var results = (RetrieveAllEntitiesResponse)service.Execute(request);
            var entities = results.EntityMetadata;
            return entities;
        }

        public static bool CheckConnection(this IOrganizationService service)
        {
            var success = service.WhoAmI().IsNullOrEmpty() == false;
            return success;
        }

        public static Guid? WhoAmI(this IOrganizationService service)
        {
            if (service == null)
                return null;

            var request = new WhoAmIRequest();
            var response = (WhoAmIResponse)service.Execute(request);

            var result = response?.UserId;
            return result;
        }
    }
}
