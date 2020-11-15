using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Description;
using CrmCodeGenerator.VSPackage.Model;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;

namespace CrmCodeGenerator.VSPackage.Xrm
{
    public class QuickConnection
    {
        public static OrganizationServiceProxy Connect(Settings settings)
            => Connect(settings.DiscoveryUrl, settings.Domain, settings.Username, settings.Password, settings.CrmOrg);

        public static OrganizationServiceProxy Connect(string url, string domain, string username, string password, string organization)
        {
            var credentials = GetCredentials(url, domain, username, password);
            ClientCredentials deviceCredentials = null;
            if (url.IndexOf("dynamics.com", StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                deviceCredentials = DeviceIdManager.LoadOrRegisterDevice(new Guid());
            }

            using (var disco = new DiscoveryServiceProxy(new Uri(url), null, credentials, deviceCredentials))
            {
                OrganizationDetailCollection orgs = DiscoverOrganizations(disco);
                if (orgs.Count <= 0)
                    return null;

                var found = TryGetOrganizationByName(orgs, organization, a => a.UniqueName) ??
                            TryGetOrganizationByName(orgs, organization, a => a.UrlName) ??
                            TryGetOrganizationByName(orgs, organization, a => a.FriendlyName);
                if (found == null)
                    return null;
                
                var orgUri = new Uri(found.Endpoints[EndpointType.OrganizationService]);
                var service = new OrganizationServiceProxy(orgUri, null, credentials, deviceCredentials);
                return service;
            }
        }
        
        private static OrganizationDetail TryGetOrganizationByName(OrganizationDetailCollection orgs, string name, Func<OrganizationDetail, string> nameSelector) 
            => orgs.FirstOrDefault(a => nameSelector(a).Equals(name, StringComparison.InvariantCultureIgnoreCase));

        public static OrganizationDetail GetSelectedOrganization(Settings settings)
        {
            var organizations = GerOrganizationDetails(settings);
            var details = organizations.FirstOrDefault(d => d.UrlName == settings.CrmOrg);
            return details;
        }

        public static List<string> GetOrganizationNames(Settings settings)
        {
            var orgs = GerOrganizationDetails(settings);
            var result = orgs.Select(d => d.UrlName).ToList();
            return result;
        }

        public static OrganizationDetail[] GerOrganizationDetails(Settings settings) 
            => GerOrganizationDetails(settings.DiscoveryUrl, settings.Domain, settings.Username, settings.Password);
        public static OrganizationDetail[] GerOrganizationDetails(string url, string domain, string username, string password)
        {
            OrganizationDetail[] Empty() => new OrganizationDetail[0];  

            var credentials = GetCredentials(url, domain, username, password);
            var deviceCredentials = GetDeviceCredentials(url);

            using (var disco = new DiscoveryServiceProxy(new Uri(url), null, credentials, deviceCredentials))
            {
                OrganizationDetailCollection organizations = DiscoverOrganizations(disco);
                if (organizations.Count > 0)
                {
                    return organizations.ToArray();
                }
            }
            return Empty();
        }

        private static ClientCredentials GetDeviceCredentials(string url)
        {
            ClientCredentials deviceCredentials = null;
            if (url.IndexOf("dynamics.com", StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                deviceCredentials = DeviceIdManager.LoadOrRegisterDevice(new Guid()); // TODO this was failing with some online connections
            }
            return deviceCredentials;
        }

        private static OrganizationDetailCollection DiscoverOrganizations(DiscoveryServiceProxy service)
        {
            RetrieveOrganizationsRequest request = new RetrieveOrganizationsRequest();
            RetrieveOrganizationsResponse response = (RetrieveOrganizationsResponse)service.Execute(request);

            return response.Details;
        }
        

        private static ClientCredentials GetCredentials(string url, string domain, string username, string password)
        {
            ClientCredentials credentials = new ClientCredentials();

            var config = ServiceConfigurationFactory.CreateConfiguration<IDiscoveryService>(new Uri(url));

            if (config.AuthenticationType == AuthenticationProviderType.ActiveDirectory)
            {
                credentials.Windows.ClientCredential = new System.Net.NetworkCredential(username, password, domain);
            }
            else if (config.AuthenticationType == AuthenticationProviderType.Federation
                || config.AuthenticationType == AuthenticationProviderType.LiveId
                || config.AuthenticationType == AuthenticationProviderType.OnlineFederation)
            {
                credentials.UserName.UserName = username;
                credentials.UserName.Password = password;
            }
            else if (config.AuthenticationType == AuthenticationProviderType.None)
            {
            }

            return credentials;
        }
    }
}
