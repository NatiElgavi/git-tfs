using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace AuthorFileGenerator
{
    /// <summary>The entry point for the console application.</summary>
    public static class Program
    {
        /// <summary>Execute the author file generator.</summary>
        public static void Main()
        {
            
            new AuthorGenerator().Scan("http://il-air-dev01:8080/tfs");
        }
    }

    /// <summary>Generates a git-tfs author mapping for all users found on a TFS server.</summary>
    /// <remarks>
    /// This script is used to generate a git-tfs author mapping for migrating from TFS to git. It
    /// will generate a list like the following:
    /// <code>
    ///    DOMAIN\john.smith = John Smith &lt;john.smith@example.org&gt;
    ///    DOMAIN\roni.scheutz = Roni Scheutz &lt;roni.scheutz@example.org&gt;
    ///    DOMAIN\shanel.jones = Shanel Jones &lt;shanel.jones@example.org&gt;
    /// </code>
    /// 
    /// See https://github.com/git-tfs/git-tfs/wiki/Clone. This is based on a code sample
    /// by Roni Scheutz @ http://netrsc.blogspot.ca/2010/07/retrieve-list-with-all-your-tfs-users.html.
    /// </remarks>
    public class AuthorGenerator
    {
        /*********
		** Public methods
		*********/
        /// <summary>Connect to a TFS server and generate a git-tfs author file for all users found.</summary>
        /// <param name="url">The root TFS URL, like "http://team:8080/tfs".</param>
        public void Scan(string url)
        {
            // scan for TFS users
            List<Identity> identities = new List<Identity>();
            TfsConfigurationServer server = this.GetServer(url, CredentialCache.DefaultCredentials);
            Console.WriteLine("Scanning TFS for users...");
            foreach (TfsTeamProjectCollection collection in this.GetProjectCollections(server))
            {
                Console.WriteLine("  collection: {0}", collection.Name);
                foreach (Project project in this.GetProjects(collection))
                {
                    Console.WriteLine("    project: {0}", project.Name);
                    try
                    {
                        identities.AddRange(this.GetUsers(project, collection));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("\tThe Project: '{0}' throws an exception: {1} and will be ignored.", project.Name, ex.Message);
                    }
                }
            }

            // get unique users
            Identity[] users = (
                from Identity rawIdentity in identities
                group rawIdentity by rawIdentity.DistinguishedName into identityGroup
                orderby identityGroup.First().AccountName
                select identityGroup.First()
            ).ToArray();

            // display list
            Console.Write("\n\nDetected users\n--------------\n");
            List<string> lines = new List<string>();

            foreach (Identity identity in users)
            {
                string line =  string.Format(@"{0}\{1} = {2} <{3}>", identity.Domain, identity.AccountName, identity.DisplayName, identity.MailAddress);
                lines.Add(line);
                Console.WriteLine(line);
            }

            System.IO.File.WriteAllLines("Authors.txt", lines);
            Console.WriteLine("Check Authors.txt for a copy of the result.");
        }


        /*********
		** Protected methods
		*********/
        /// <summary>Get the TFS configuration server.</summary>
        /// <param name="url">The root TFS URL, like "http://team:8080/tfs".</param>
        /// <param name="credentials">The TFS credentials with which to access the server.</param>
        protected TfsConfigurationServer GetServer(string url, ICredentials credentials)
        {
            // get server
            TfsConfigurationServer server = new TfsConfigurationServer(new Uri(url));

            // authenticate
            server.EnsureAuthenticated();
            server.Authenticate();
            if (!server.HasAuthenticated)
                throw new InvalidOperationException("Authentication to TFS failed.");

            return server;
        }

        /// <summary>Get the TFS project collections.</summary>
        /// <param name="server">The TFS server from which to read the project collections.</param>
        protected IEnumerable<TfsTeamProjectCollection> GetProjectCollections(TfsConfigurationServer server)
        {
            ReadOnlyCollection<CatalogNode> projectNodes = server.CatalogNode.QueryChildren(new[] { CatalogResourceTypes.ProjectCollection }, false, CatalogQueryOptions.None);
            foreach (CatalogNode projectNode in projectNodes)
                yield return server.GetTeamProjectCollection(new Guid(projectNode.Resource.Properties["InstanceId"]));
        }

        /// <summary>Get the TFS projects in a collection.</summary>
        /// <param name="collection">The TFS project collection.</param>
        protected IEnumerable<Project> GetProjects(TfsTeamProjectCollection collection)
        {
            WorkItemStore workItemStore = collection.GetService<WorkItemStore>();
            return workItemStore.Projects.Cast<Project>();
        }

        /// <summary>Get the users in a TFS project.</summary>
        /// <param name="project">The TFS project.</param>
        /// <param name="collection">The TFS project collection.</param>
        protected IEnumerable<Identity> GetUsers(Project project, TfsTeamProjectCollection collection)
        {
            // get identity service
            VersionControlServer versionControl = collection.GetService<VersionControlServer>();
            TeamProject teamProject = versionControl.GetTeamProject(project.Name);
            IGroupSecurityService securityService = collection.GetService<IGroupSecurityService>();

            // get identities
            Identity[] appGroups = securityService.ListApplicationGroups(teamProject.ArtifactUri.AbsoluteUri);
            foreach (Identity group in appGroups)
            {
                Identity[] groupMembers = securityService.ReadIdentities(SearchFactor.Sid, new string[] { group.Sid }, QueryMembership.Expanded);
                foreach (Identity member in groupMembers)
                {
                    if (member.Members != null)
                    {
                        foreach (string memberSid in member.Members)
                        {
                            Identity memberInfo = securityService.ReadIdentity(SearchFactor.Sid, memberSid, QueryMembership.None);
                            if (memberInfo.Type == IdentityType.WindowsUser)
                                yield return memberInfo;
                        }
                    }
                }
            }
        }
    }
}