using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Core.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.Operations;
using Microsoft.TeamFoundation.TestManagement.Client;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Text;

using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using System.Collections;
using Microsoft.TeamFoundation.Framework.Client;

using Migration.Common;
using Migration.Common.Log;
using Migration.WIContract;

using VsWebApi = Microsoft.VisualStudio.Services.WebApi;
using WebApi = Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using WebModel = Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;

namespace WorkItemImport
{
    public class Agent
    {
        private readonly MigrationContext _context;
        public Settings Settings { get; private set; }

        public TfsTeamProjectCollection Collection
        {
            get; private set;
        }

        private WorkItemStore _store;
        public WorkItemStore Store
        {
            get
            {
                if (_store == null)
                    _store = new WorkItemStore(Collection, WorkItemStoreFlags.BypassRules);

                return _store;
            }
        }

        public VsWebApi.VssConnection RestConnection { get; private set; }
        public Dictionary<string, int> IterationCache { get; private set; } = new Dictionary<string, int>();
        public int RootIteration { get; private set; }
        public Dictionary<string, int> AreaCache { get; private set; } = new Dictionary<string, int>();
        public int RootArea { get; private set; }

        private WebApi.WorkItemTrackingHttpClient _wiClient;
        public WebApi.WorkItemTrackingHttpClient WiClient
        {
            get
            {
                if (_wiClient == null)
                    _wiClient = RestConnection.GetClient<WebApi.WorkItemTrackingHttpClient>();

                return _wiClient;
            }
        }

        private Agent(MigrationContext context, Settings settings, VsWebApi.VssConnection restConn, TfsTeamProjectCollection soapConnection)
        {
            _context = context;
            Settings = settings;
            RestConnection = restConn;
            Collection = soapConnection;
        }

        #region Static
        internal static Agent Initialize(MigrationContext context, Settings settings)
        {
            var restConnection = EstablishRestConnection(settings);
            if (restConnection == null)
                return null;

            var soapConnection = EstablishSoapConnection(settings);
            if (soapConnection == null)
                return null;

            var agent = new Agent(context, settings, restConnection, soapConnection);

            // check if projects exists, if not create it
            var project = agent.GetOrCreateProjectAsync().Result;
            if (project == null)
            {
                Logger.Log(LogLevel.Critical, "Could not establish connection to the remote Azure DevOps/TFS project.");
                return null;
            }

            (var iterationCache, int rootIteration) = agent.CreateClasificationCacheAsync(settings.Project, WebModel.TreeStructureGroup.Iterations).Result;
            if (iterationCache == null)
            {
                Logger.Log(LogLevel.Critical, "Could not build iteration cache.");
                return null;
            }

            agent.IterationCache = iterationCache;
            agent.RootIteration = rootIteration;

            (var areaCache, int rootArea) = agent.CreateClasificationCacheAsync(settings.Project, WebModel.TreeStructureGroup.Areas).Result;
            if (areaCache == null)
            {
                Logger.Log(LogLevel.Critical, "Could not build area cache.");
                return null;
            }

            agent.AreaCache = areaCache;
            agent.RootArea = rootArea;

            return agent;
        }

        internal Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem CreateWI(string type)
        {
            var project = Store.Projects[Settings.Project];
            var wiType = project.WorkItemTypes[type];
            return wiType.NewWorkItem();
        }

        internal Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem GetWorkItem(int wiId)
        {
            return Store.GetWorkItem(wiId);
        }

        private static VsWebApi.VssConnection EstablishRestConnection(Settings settings)
        {
            try
            {
                Logger.Log(LogLevel.Info, "Connecting to Azure DevOps/TFS...");
                var credentials = new VssBasicCredential("", settings.Pat);
                var uri = new Uri(settings.Account);
                return new VsWebApi.VssConnection(uri, credentials);
            }
            catch (Exception ex)
            {
                Logger.Log(ex, $"Cannot establish connection to Azure DevOps/TFS.", LogLevel.Critical);
                return null;
            }
        }

        private static TfsTeamProjectCollection EstablishSoapConnection(Settings settings)
        {
            NetworkCredential netCred = new NetworkCredential(string.Empty, settings.Pat);
            VssBasicCredential basicCred = new VssBasicCredential(netCred);
            VssCredentials tfsCred = new VssCredentials(basicCred);
            var collection = new TfsTeamProjectCollection(new Uri(settings.Account), tfsCred);
            collection.Authenticate();
            return collection;
        }

        #endregion

        #region Setup

        internal async Task<TeamProject> GetOrCreateProjectAsync()
        {
            ProjectHttpClient projectClient = RestConnection.GetClient<ProjectHttpClient>();
            Logger.Log(LogLevel.Info, "Retreiving project info from Azure DevOps/TFS...");
            TeamProject project = null;

            try
            {
                project = await projectClient.GetProject(Settings.Project);
            }
            catch (Exception ex)
            {
                Logger.Log(ex, $"Failed to get Azure DevOps/TFS project '{Settings.Project}'.");
            }

            if (project == null)
                project = await CreateProject(Settings.Project, $"{Settings.ProcessTemplate} project for Jira migration", Settings.ProcessTemplate);

            return project;
        }

        internal async Task<TeamProject> CreateProject(string projectName, string projectDescription = "", string processName = "Scrum")
        {
            Logger.Log(LogLevel.Warning, $"Project '{projectName}' does not exist.");
            Console.WriteLine("Would you like to create one? (Y/N)");
            var answer = Console.ReadKey();
            if (answer.KeyChar != 'Y' && answer.KeyChar != 'y')
                return null;

            Logger.Log(LogLevel.Info, $"Creating project '{projectName}'.");

            // Setup version control properties
            Dictionary<string, string> versionControlProperties = new Dictionary<string, string>
            {
                [TeamProjectCapabilitiesConstants.VersionControlCapabilityAttributeName] = SourceControlTypes.Git.ToString()
            };

            // Setup process properties       
            ProcessHttpClient processClient = RestConnection.GetClient<ProcessHttpClient>();
            Guid processId = processClient.GetProcessesAsync().Result.Find(process => { return process.Name.Equals(processName, StringComparison.InvariantCultureIgnoreCase); }).Id;

            Dictionary<string, string> processProperaties = new Dictionary<string, string>
            {
                [TeamProjectCapabilitiesConstants.ProcessTemplateCapabilityTemplateTypeIdAttributeName] = processId.ToString()
            };

            // Construct capabilities dictionary
            Dictionary<string, Dictionary<string, string>> capabilities = new Dictionary<string, Dictionary<string, string>>
            {
                [TeamProjectCapabilitiesConstants.VersionControlCapabilityName] = versionControlProperties,
                [TeamProjectCapabilitiesConstants.ProcessTemplateCapabilityName] = processProperaties
            };

            // Construct object containing properties needed for creating the project
            TeamProject projectCreateParameters = new TeamProject()
            {
                Name = projectName,
                Description = projectDescription,
                Capabilities = capabilities
            };

            // Get a client
            ProjectHttpClient projectClient = RestConnection.GetClient<ProjectHttpClient>();

            TeamProject project = null;
            try
            {
                Logger.Log(LogLevel.Info, "Queuing project creation...");

                // Queue the project creation operation 
                // This returns an operation object that can be used to check the status of the creation
                OperationReference operation = await projectClient.QueueCreateProject(projectCreateParameters);

                // Check the operation status every 5 seconds (for up to 30 seconds)
                Microsoft.VisualStudio.Services.Operations.Operation completedOperation = WaitForLongRunningOperation(operation.Id, 5, 30).Result;

                // Check if the operation succeeded (the project was created) or failed
                if (completedOperation.Status == OperationStatus.Succeeded)
                {
                    // Get the full details about the newly created project
                    project = projectClient.GetProject(
                        projectCreateParameters.Name,
                        includeCapabilities: true,
                        includeHistory: true).Result;

                    Logger.Log(LogLevel.Info, $"Project created (ID: {project.Id})");
                }
                else
                {
                    Logger.Log(LogLevel.Error, "Project creation operation failed: " + completedOperation.ResultMessage);
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex, "Exception during create project.", LogLevel.Critical);
            }

            return project;
        }

        private async Task<Microsoft.VisualStudio.Services.Operations.Operation> WaitForLongRunningOperation(Guid operationId, int interavalInSec = 5, int maxTimeInSeconds = 60, CancellationToken cancellationToken = default(CancellationToken))
        {
            OperationsHttpClient operationsClient = RestConnection.GetClient<OperationsHttpClient>();
            DateTime expiration = DateTime.Now.AddSeconds(maxTimeInSeconds);
            int checkCount = 0;

            while (true)
            {
                Logger.Log(LogLevel.Info, $" Checking status ({checkCount++})... ");

                Microsoft.VisualStudio.Services.Operations.Operation operation = await operationsClient.GetOperation(operationId, cancellationToken);

                if (!operation.Completed)
                {
                    Logger.Log(LogLevel.Info, $"   Pausing {interavalInSec} seconds...");

                    await Task.Delay(interavalInSec * 1000);

                    if (DateTime.Now > expiration)
                    {
                        Logger.Log(LogLevel.Error, $"Operation did not complete in {maxTimeInSeconds} seconds.");
                    }
                }
                else
                {
                    return operation;
                }
            }
        }

        private async Task<(Dictionary<string, int>, int)> CreateClasificationCacheAsync(string project, WebModel.TreeStructureGroup structureGroup)
        {
            try
            {
                Logger.Log(LogLevel.Info, $"Building {(structureGroup == WebModel.TreeStructureGroup.Iterations ? "iteration" : "area")} cache...");
                WebModel.WorkItemClassificationNode all = await WiClient.GetClassificationNodeAsync(project, structureGroup, null, 1000);

                var clasificationCache = new Dictionary<string, int>();

                if (all.Children != null && all.Children.Any())
                {
                    foreach (var iteration in all.Children)
                        CreateClasificationCacheRec(iteration, clasificationCache, "");
                }

                return (clasificationCache, all.Id);
            }
            catch (Exception ex)
            {
                Logger.Log(ex, $"Error while building {(structureGroup == WebModel.TreeStructureGroup.Iterations ? "iteration" : "area")} cache.");
                return (null, -1);
            }
        }

        private void CreateClasificationCacheRec(WebModel.WorkItemClassificationNode current, Dictionary<string, int> agg, string parentPath)
        {
            string fullName = !string.IsNullOrWhiteSpace(parentPath) ? parentPath + "/" + current.Name : current.Name;

            agg.Add(fullName, current.Id);
            Logger.Log(LogLevel.Debug, $"{(current.StructureType == WebModel.TreeNodeStructureType.Iteration ? "Iteration" : "Area")} '{fullName}' added to cache");
            if (current.Children != null)
            {
                foreach (var node in current.Children)
                    CreateClasificationCacheRec(node, agg, fullName);
            }
        }

        public int? EnsureClasification(string fullName, WebModel.TreeStructureGroup structureGroup = WebModel.TreeStructureGroup.Iterations)
        {
            if (string.IsNullOrWhiteSpace(fullName))
            {
                Logger.Log(LogLevel.Error, "Empty value provided for node name/path.");
                throw new ArgumentException("fullName");
            }

            var path = fullName.Split('/');
            var name = path.Last();
            var parent = string.Join("/", path.Take(path.Length - 1));

            if (!string.IsNullOrEmpty(parent))
                EnsureClasification(parent, structureGroup);

            var cache = structureGroup == WebModel.TreeStructureGroup.Iterations ? IterationCache : AreaCache;

            lock (cache)
            {
                if (cache.TryGetValue(fullName, out int id))
                    return id;

                WebModel.WorkItemClassificationNode node = null;

                try
                {
                    node = WiClient.CreateOrUpdateClassificationNodeAsync(
                        new WebModel.WorkItemClassificationNode() { Name = name, }, Settings.Project, structureGroup, parent).Result;
                }
                catch (Exception ex)
                {
                    Logger.Log(ex, $"Error while adding {(structureGroup == WebModel.TreeStructureGroup.Iterations ? "iteration" : "area")} '{fullName}' to Azure DevOps/TFS.", LogLevel.Critical);
                }

                if (node != null)
                {
                    Logger.Log(LogLevel.Debug, $"{(structureGroup == WebModel.TreeStructureGroup.Iterations ? "Iteration" : "Area")} '{fullName}' added to Azure DevOps/TFS.");
                    cache.Add(fullName, node.Id);
                    Store.RefreshCache();
                    return node.Id;
                }
            }
            return null;
        }

        #endregion

        #region Import Revision

        private bool UpdateWIFields(IEnumerable<WiField> fields, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi)
        {
            var success = true;

            if (!wi.IsOpen || !wi.IsPartialOpen)
                wi.PartialOpen();

            foreach (var fieldRev in fields)
            {
                try
                {
                    var fieldRef = fieldRev.ReferenceName;
                    var fieldValue = fieldRev.Value;


                    switch (fieldRef)
                    {
                        case var s when s.Equals(WiFieldReference.IterationPath, StringComparison.InvariantCultureIgnoreCase):

                            var iterationPath = Settings.BaseIterationPath;

                            if (!string.IsNullOrWhiteSpace((string)fieldValue))
                            {
                                var iterationAppendValue = (string)fieldValue;
                                iterationAppendValue = Regex.Replace(iterationAppendValue, "[\\/$?*:&<>#%|+]", "-");
                                //Logger.Log(LogLevel.Info, $"Mapped IterationPath '{wi.IterationPath}'.");
                                iterationAppendValue = Regex.Replace(iterationAppendValue, "\\\\\"", " -");
                                if (string.IsNullOrWhiteSpace(iterationPath))
                                    // iterationPath = (string)fieldValue;
                                    iterationPath = iterationAppendValue;
                                else
                                    // iterationPath = string.Join("/", iterationPath, (string)fieldValue);
                                    iterationPath = string.Join("/", iterationPath, iterationAppendValue);
                            }
                            //Logger.Log(LogLevel.Info, $"Mapped IterationPath '{wi.IterationPath}'.");
                            if (!string.IsNullOrWhiteSpace(iterationPath))
                            {
                                EnsureClasification(iterationPath, WebModel.TreeStructureGroup.Iterations);
                                wi.IterationPath = $@"{Settings.Project}\{iterationPath}".Replace("/", @"\");
                            }
                            else
                            {
                                wi.IterationPath = Settings.Project;
                            }
                            Logger.Log(LogLevel.Debug, $"Mapped IterationPath '{wi.IterationPath}'.");
                            break;

                        case var s when s.Equals(WiFieldReference.AreaPath, StringComparison.InvariantCultureIgnoreCase):

                            var areaPath = Settings.BaseAreaPath;

                            if (!string.IsNullOrWhiteSpace((string)fieldValue))
                            {
                                var areaAppendValue = (string)fieldValue;
                                areaAppendValue = Regex.Replace(areaAppendValue, "[\\/$?*:&<>#%|+]", "-");
                                areaAppendValue = Regex.Replace(areaAppendValue, "\\\\\"", "-");
                                if (string.IsNullOrWhiteSpace(areaPath))                                    
                                    // areaPath = (string)fieldValue;
                                    areaPath = areaAppendValue;
                                else
                                    // areaPath = string.Join("/", areaPath, (string)fieldValue);
                                    areaPath = string.Join("/", areaPath, areaAppendValue);
                            }

                            if (!string.IsNullOrWhiteSpace(areaPath))
                            {
                                EnsureClasification(areaPath, WebModel.TreeStructureGroup.Areas);
                                wi.AreaPath = $@"{Settings.Project}\{areaPath}".Replace("/", @"\");
                            }
                            else
                            {
                                wi.AreaPath = Settings.Project;
                            }

                            Logger.Log(LogLevel.Debug, $"Mapped AreaPath '{wi.AreaPath}'.");

                            break;

                        case var s when s.Equals(WiFieldReference.ActivatedDate, StringComparison.InvariantCultureIgnoreCase) && fieldValue == null ||
                             s.Equals(WiFieldReference.ActivatedBy, StringComparison.InvariantCultureIgnoreCase) && fieldValue == null ||
                            s.Equals(WiFieldReference.ClosedDate, StringComparison.InvariantCultureIgnoreCase) && fieldValue == null ||
                            s.Equals(WiFieldReference.ClosedBy, StringComparison.InvariantCultureIgnoreCase) && fieldValue == null ||
                            s.Equals(WiFieldReference.Tags, StringComparison.InvariantCultureIgnoreCase) && fieldValue == null:

                            SetFieldValue(wi, fieldRef, fieldValue);
                            break;
                        default:
                            if (fieldValue != null)
                            {
                                SetFieldValue(wi, fieldRef, fieldValue);
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(ex, $"Failed to update fields.");
                    success = false;
                }
            }

            return success;
        }

        private static void SetFieldValue(Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi, string fieldRef, object fieldValue)
        {
            try
            {
                //if (fieldRef != "WEF_3833D59EA86C405986E666B2297ACB95_Kanban.Column" && fieldRef != "WEF_3833D59EA86C405986E666B2297ACB95_Kanban.Lane")
                if (!fieldRef.EndsWith("Kanban.Column") && !fieldRef.EndsWith("Kanban.Lane"))
                {

                    var field = wi.Fields[fieldRef];

                    if (fieldRef == "System.Title")
                    {
                        int length = fieldValue.ToString().Length;
                        if (length > 155)
                        {
                            fieldValue = fieldValue.ToString().Substring(0, 154);
                        }
                    }

                field.Value = fieldValue;

                if (field.IsValid)
                    Logger.Log(LogLevel.Debug, $"Mapped '{fieldRef}' '{fieldValue}'.");
                else
                {
                    field.Value = null;
                    Logger.Log(LogLevel.Warning, $"Mapped empty value for '{fieldRef}', because value was not valid");
                }
                }

            }
            catch (ValidationException ex)
            {
                //Logger.Log(LogLevel.Error, ex.Message);
                Logger.Log(LogLevel.Error, $"WorkItem ID: '{wi.Id}' : Field Name: '{fieldRef}' : Field Value: '{fieldValue}' : '{ex.Message}'.");
            }


        }

        private bool ApplyAttachments(WiRevision rev, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi, Dictionary<string, Attachment> attachmentMap)
        {
            var success = true;

            if (!wi.IsOpen)
                wi.Open();

            foreach (var att in rev.Attachments)
            {
                try
                {
                    Logger.Log(LogLevel.Debug, $"Adding attachment '{att.ToString()}'.");
                    if (att.Change == ReferenceChangeType.Added)
                    {
                        var newAttachment = new Attachment(att.FilePath, att.Comment);
                        wi.Attachments.Add(newAttachment);

                        attachmentMap.Add(att.AttOriginId, newAttachment);
                    }
                    else if (att.Change == ReferenceChangeType.Removed)
                    {
                        Attachment existingAttachment = IdentifyAttachment(att, wi);
                        if (existingAttachment != null)
                        {
                            wi.Attachments.Remove(existingAttachment);
                        }
                        else
                        {
                            success = false;
                            Logger.Log(LogLevel.Error, $"Could not find migrated attachment '{att.ToString()}'.");
                        }
                    }
                }
                catch (AbortMigrationException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    Logger.Log(ex, $"Failed to apply attachments for '{wi.Id}'.");
                    success = false;
                }
            }

            if (rev.Attachments.Any(a => a.Change == ReferenceChangeType.Removed))
                wi.Fields[CoreField.History].Value = $"Removed attachments(s): { string.Join(";", rev.Attachments.Where(a => a.Change == ReferenceChangeType.Removed).Select(a => a.ToString()))}";

            return success;
        }

        private Attachment IdentifyAttachment(WiAttachment att, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi)
        {
            if (_context.Journal.IsAttachmentMigrated(att.AttOriginId, out int attWiId))
                return wi.Attachments.Cast<Attachment>().SingleOrDefault(a => a.Id == attWiId);
            return null;
        }

        private bool ApplyLinks(WiRevision rev, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi)
        {
            bool success = true;

            if (!wi.IsOpen)
                wi.Open();


            foreach (var link in rev.Links)
            {
                try
                {
                    int sourceWiId = _context.Journal.GetMigratedId(link.SourceOriginId);
                    int targetWiId = _context.Journal.GetMigratedId(link.TargetOriginId);

                    link.SourceWiId = sourceWiId;
                    link.TargetWiId = targetWiId;

                    if (link.TargetWiId == -1)
                    {
                        var errorLevel = Settings.IgnoreFailedLinks ? LogLevel.Warning : LogLevel.Error;
                        Logger.Log(errorLevel, $"'{link.ToString()}' - target work item for Jira '{link.TargetOriginId}' is not yet created in Azure DevOps/TFS.");
                        success = false;
                        continue;
                    }

                    if (link.Change == ReferenceChangeType.Added && !AddLink(link, wi))
                    {
                        success = false;
                    }
                    else if (link.Change == ReferenceChangeType.Removed && !RemoveLink(link, wi))
                    {
                        success = false;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(ex, $"Failed to apply links for '{wi.Id}'.");
                    success = false;
                }
            }

            if (rev.Links.Any(l => l.Change == ReferenceChangeType.Removed))
                wi.Fields[CoreField.History].Value = $"Removed link(s): { string.Join(";", rev.Links.Where(l => l.Change == ReferenceChangeType.Removed).Select(l => l.ToString()))}";
            else if (rev.Links.Any(l => l.Change == ReferenceChangeType.Added))
                wi.Fields[CoreField.History].Value = $"Added link(s): { string.Join(";", rev.Links.Where(l => l.Change == ReferenceChangeType.Added).Select(l => l.ToString()))}";

            return success;
        }

        private bool AddLink(WiLink link, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi)
        {
            var linkEnd = ParseLinkEnd(link, wi);

            if (linkEnd != null)
            {
                try
                {
                    var relatedLink = new RelatedLink(linkEnd, link.TargetWiId);
                    relatedLink = ResolveCiclycalLinks(relatedLink, wi);
                    if (!IsDuplicateWorkItemLink(wi.Links, relatedLink))
                    {
                        wi.Links.Add(relatedLink);
                        return true;
                    }
                    return false;
                }

                catch (Exception ex)
                {

                    Logger.Log(LogLevel.Error, ex.Message);
                    return false;
                }
            }
            else
                return false;

        }

        private bool IsDuplicateWorkItemLink(LinkCollection links, RelatedLink relatedLink)
        {
            var containsRelatedLink = links.Contains(relatedLink);
            var hasSameRelatedWorkItemId = links.OfType<RelatedLink>()
                .Any(l => l.RelatedWorkItemId == relatedLink.RelatedWorkItemId);

            if (!containsRelatedLink && !hasSameRelatedWorkItemId)
                return false;

            Logger.Log(LogLevel.Warning, $"Duplicate work item link detected to related workitem id: {relatedLink.RelatedWorkItemId}, Skipping link");
            return true;


        }

        private RelatedLink ResolveCiclycalLinks(RelatedLink link, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi)
        {
            if (link.LinkTypeEnd.LinkType.IsNonCircular && DetectCycle(wi, link))
                return new RelatedLink(link.LinkTypeEnd.OppositeEnd, link.RelatedWorkItemId);

            return link;
        }

        private bool DetectCycle(Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem startingWi, RelatedLink startingLink)
        {
            var nextWiLink = startingLink;
            do
            {
                var nextWi = Store.GetWorkItem(nextWiLink.RelatedWorkItemId);
                nextWiLink = nextWi.Links.OfType<RelatedLink>().FirstOrDefault(rl => rl.LinkTypeEnd.Id == startingLink.LinkTypeEnd.Id);

                if (nextWiLink != null && nextWiLink.RelatedWorkItemId == startingWi.Id)
                    return true;

            } while (nextWiLink != null);

            return false;
        }

        private WorkItemLinkTypeEnd ParseLinkEnd(WiLink link, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi)
        {
            var props = link.WiType?.Split('-');
            var linkType = wi.Project.Store.WorkItemLinkTypes.SingleOrDefault(lt => lt.ReferenceName == props?[0]);
            if (linkType == null)
            {
                Logger.Log(LogLevel.Error, $"'{link.ToString()}' - link type ({props?[0]}) does not exist in project");
                return null;
            }

            WorkItemLinkTypeEnd linkEnd = null;

            if (linkType.IsDirectional)
            {
                if (props?.Length > 1)
                    linkEnd = props[1] == "Forward" ? linkType.ForwardEnd : linkType.ReverseEnd;
                else
                    Logger.Log(LogLevel.Error, $"'{link.ToString()}' - link direction not provided for '{wi.Id}'.");
            }
            else
                linkEnd = linkType.ForwardEnd;

            return linkEnd;
        }

        private bool RemoveLink(WiLink link, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi)
        {
            var linkToRemove = wi.Links.OfType<RelatedLink>().SingleOrDefault(rl => rl.LinkTypeEnd.ImmutableName == link.WiType && rl.RelatedWorkItemId == link.TargetWiId);
            if (linkToRemove == null)
            {
                Logger.Log(LogLevel.Warning, $"{link.ToString()} - cannot identify link to remove for '{wi.Id}'.");
                return false;
            }
            wi.Links.Remove(linkToRemove);
            return true;
        }

        private void SaveWorkItem(WiRevision rev, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem newWorkItem)
        {
            if (!newWorkItem.IsValid())
            {
                var reasons = newWorkItem.Validate();
                foreach (Microsoft.TeamFoundation.WorkItemTracking.Client.Field reason in reasons)
                    Logger.Log(LogLevel.Info, $"Field: '{reason.Name}', Status: '{reason.Status}', Value: '{reason.Value}'");
            }
            try
            {
                newWorkItem.Save(SaveFlags.MergeAll);
            }
            catch (FileAttachmentException faex)
            {
                Logger.Log(faex,
                    $"[{faex.GetType().ToString()}] {faex.Message}. Attachment {faex.SourceAttachment.Name}({faex.SourceAttachment.Id}) in {rev.ToString()} will be skipped.");
                newWorkItem.Attachments.Remove(faex.SourceAttachment);
                SaveWorkItem(rev, newWorkItem);
            }
            catch (WorkItemLinkValidationException wilve)
            {
                Logger.Log(wilve, $"[{wilve.GetType()}] {wilve.Message}. Link Source: {wilve.LinkInfo.SourceId}, Target: {wilve.LinkInfo.TargetId} in {rev} will be skipped.");
                var exceedsLinkLimit = RemoveLinksFromWiThatExceedsLimit(newWorkItem);
                if (exceedsLinkLimit)
                    SaveWorkItem(rev, newWorkItem);
            }
        }

        private bool RemoveLinksFromWiThatExceedsLimit(Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem newWorkItem)
        {
            var links = newWorkItem.Links.OfType<RelatedLink>().ToList();
            var result = false;
            foreach (var link in links)
            {
                var relatedWorkItem = GetWorkItem(link.RelatedWorkItemId);
                var relatedLinkCount = relatedWorkItem.RelatedLinkCount;
                if (relatedLinkCount != 1000)
                    continue;

                newWorkItem.Links.Remove(link);
                result = true;
            }

            return result;
        }

        private void EnsureAuthorFields(WiRevision rev)
        {
            if (rev.Index == 0 && !rev.Fields.HasAnyByRefName(WiFieldReference.CreatedBy))
            {
                rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.CreatedBy, Value = rev.Author });
            }
            if (!rev.Fields.HasAnyByRefName(WiFieldReference.ChangedBy))
            {
                rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.ChangedBy, Value = rev.Author });
            }
        }

        private void EnsureAssigneeField(WiRevision rev, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi)
        {
            string assignedTo = wi.Fields[WiFieldReference.AssignedTo].Value.ToString();

            if (rev.Fields.HasAnyByRefName(WiFieldReference.AssignedTo))
            {
                var field = rev.Fields.First(f => f.ReferenceName.Equals(WiFieldReference.AssignedTo, StringComparison.InvariantCultureIgnoreCase));
                assignedTo = field.Value?.ToString() ?? string.Empty;
                rev.Fields.RemoveAll(f => f.ReferenceName.Equals(WiFieldReference.AssignedTo, StringComparison.InvariantCultureIgnoreCase));
            }
            rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.AssignedTo, Value = assignedTo });
        }

        private void EnsureDateFields(WiRevision rev, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi)
        {
            if (rev.Index == 0 && !rev.Fields.HasAnyByRefName(WiFieldReference.CreatedDate))
            {
                rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.CreatedDate, Value = rev.Time.ToString("o") });
            }
            if (!rev.Fields.HasAnyByRefName(WiFieldReference.ChangedDate))
            {
                if (wi.ChangedDate == rev.Time)
                {
                    rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.ChangedDate, Value = rev.Time.AddMilliseconds(1).ToString("o") });
                }
                else
                    rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.ChangedDate, Value = rev.Time.ToString("o") });
            }

        }


        private void EnsureFieldsOnStateChange(WiRevision rev, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi)
        {
            if (rev.Index != 0 && rev.Fields.HasAnyByRefName(WiFieldReference.State))
            {
                var wiState = wi.Fields[WiFieldReference.State]?.Value?.ToString() ?? string.Empty;
                var revState = rev.Fields.GetFieldValueOrDefault<string>(WiFieldReference.State) ?? string.Empty;
                if (wiState.Equals("Done", StringComparison.InvariantCultureIgnoreCase) && revState.Equals("New", StringComparison.InvariantCultureIgnoreCase))
                {
                    rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.ClosedDate, Value = null });
                    rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.ClosedBy, Value = null });

                }
                if (!wiState.Equals("New", StringComparison.InvariantCultureIgnoreCase) && revState.Equals("New", StringComparison.InvariantCultureIgnoreCase))
                {
                    rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.ActivatedDate, Value = null });
                    rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.ActivatedBy, Value = null });
                }

                if (revState.Equals("Done", StringComparison.InvariantCultureIgnoreCase) && !rev.Fields.HasAnyByRefName(WiFieldReference.ClosedBy))
                    rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.ClosedBy, Value = rev.Author });
            }
        }

        private bool CorrectDescription(Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi, WiItem wiItem, WiRevision rev)
        {
            string description = wi.Type.Name == "Bug" ? wi.Fields[WiFieldReference.ReproSteps].Value.ToString() : wi.Description;
            if (string.IsNullOrWhiteSpace(description))
                return false;

            bool descUpdated = false;

            CorrectImagePath(wi, wiItem, rev, ref description, ref descUpdated);

            if (descUpdated)
            {
                if (wi.Type.Name == "Bug")
                {
                    wi.Fields[WiFieldReference.ReproSteps].Value = description;
                }
                else
                {
                    wi.Fields[WiFieldReference.Description].Value = description;
                }
            }

            return descUpdated;
        }

        private void CorrectComment(Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi, WiItem wiItem, WiRevision rev)
        {
            var currentComment = wi.History;
            var commentUpdated = false;
            CorrectImagePath(wi, wiItem, rev, ref currentComment, ref commentUpdated);

            if (commentUpdated)
                wi.Fields[CoreField.History].Value = currentComment;
        }

        private void CorrectImagePath(Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi, WiItem wiItem, WiRevision rev, ref string textField, ref bool isUpdated)
        {
            foreach (var att in wiItem.Revisions.SelectMany(r => r.Attachments.Where(a => a.Change == ReferenceChangeType.Added)))
            {
                var fileName = att.FilePath.Split('\\')?.Last() ?? string.Empty;
                if (textField.Contains(fileName))
                {
                    var tfsAtt = IdentifyAttachment(att, wi);

                    if (tfsAtt != null)
                    {
                        string imageSrcPattern = $"src.*?=.*?\"([^\"])(?=.*{att.AttOriginId}).*?\"";
                        textField = Regex.Replace(textField, imageSrcPattern, $"src=\"{tfsAtt.Uri.AbsoluteUri}\"");
                        isUpdated = true;
                    }
                    else
                        Logger.Log(LogLevel.Warning, $"Attachment '{att.ToString()}' referenced in text but is missing from work item {wiItem.OriginId}/{wi.Id}.");
                }
            }
            if (isUpdated)
            {
                DateTime changedDate;
                if (wiItem.Revisions.Count > rev.Index + 1)
                    changedDate = RevisionUtility.NextValidDeltaRev(rev.Time, wiItem.Revisions[rev.Index + 1].Time);
                else
                    changedDate = RevisionUtility.NextValidDeltaRev(rev.Time);

                wi.Fields[WiFieldReference.ChangedDate].Value = changedDate;
                wi.Fields[WiFieldReference.ChangedBy].Value = rev.Author;
            }
        }

        // public bool ImportRevision(WiRevision rev, WorkItem wi)
        public bool ImportRevision(WiRevision rev, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wi, TfsTeamProjectCollection collection, int thisRevision, String jsonPath, Hashtable htParent, String urlValue)
        {
            var incomplete = false;
            try
            {
                if (rev.Index == 0)
                    EnsureClassificationFields(rev);

                EnsureDateFields(rev, wi);
                EnsureAuthorFields(rev);
                EnsureAssigneeField(rev, wi);
                EnsureFieldsOnStateChange(rev, wi);

                var attachmentMap = new Dictionary<string, Attachment>();
                if (rev.Attachments.Any() && !ApplyAttachments(rev, wi, attachmentMap))
                    incomplete = true;

                //var thisWorkitemID = wi.Id.ToString();
                var thisIssueID = rev.ParentOriginId;
                foreach (var fieldData in rev.Fields)
                {                    
                    String lateFieldName = "";
                    String lateFieldValue = "";
                    
                    if (fieldData.ReferenceName.EndsWith("Kanban.Column") || fieldData.ReferenceName.EndsWith("Kanban.Lane"))
                    {
                        //Logger.Log(LogLevel.Info, rev.ParentOriginId + " : " + htParent[thisIssueID]);
                        lateFieldName = fieldData.ReferenceName;
                        lateFieldValue = fieldData.Value.ToString();
                        if (htParent.ContainsKey(thisIssueID))
                        {
                            Hashtable htLateUpdates = (Hashtable)htParent[thisIssueID];
                            if (!htLateUpdates.ContainsKey(lateFieldName))
                            {
                                htLateUpdates.Add(lateFieldName, lateFieldValue);
                                htParent[thisIssueID] = htLateUpdates;
                            }
                            else
                                htLateUpdates[lateFieldName] = lateFieldValue;
                        }
                        else
                        {
                            Hashtable htLateUpdates = new Hashtable();
                            htLateUpdates.Add(lateFieldName, lateFieldValue);
                            htParent.Add(thisIssueID, htLateUpdates);
                        }                        
                        //Logger.Log(LogLevel.Info, lateFieldName + ":" + lateFieldValue);
                    }
                }

                if (rev.Fields.Any() && !UpdateWIFields(rev.Fields, wi))
                    incomplete = true;

                if (rev.Links.Any() && !ApplyLinks(rev, wi))
                    incomplete = true;

                if (incomplete)
                    Logger.Log(LogLevel.Warning, $"'{rev.ToString()}' - not all changes were saved.");

                if (rev.Attachments.All(a => a.Change != ReferenceChangeType.Added) && rev.AttachmentReferences)
                {
                    Logger.Log(LogLevel.Debug, $"Correcting description on '{rev.ToString()}'.");
                    CorrectDescription(wi, _context.GetItem(rev.ParentOriginId), rev);
                }
                if (!string.IsNullOrEmpty(wi.History))
                {
                    Logger.Log(LogLevel.Debug, $"Correcting comments on '{rev.ToString()}'.");
                    CorrectComment(wi, _context.GetItem(rev.ParentOriginId), rev);
                }

                SaveWorkItem(rev, wi);

                foreach (var wiAtt in rev.Attachments)
                {
                    if (attachmentMap.TryGetValue(wiAtt.AttOriginId, out Attachment tfsAtt) && tfsAtt.IsSaved)
                        _context.Journal.MarkAttachmentAsProcessed(wiAtt.AttOriginId, tfsAtt.Id);
                }

                if (rev.Attachments.Any(a => a.Change == ReferenceChangeType.Added) && rev.AttachmentReferences)
                {
                    Logger.Log(LogLevel.Debug, $"Correcting description on separate revision on '{rev.ToString()}'.");

                    try
                    {
                        if (CorrectDescription(wi, _context.GetItem(rev.ParentOriginId), rev))
                            SaveWorkItem(rev, wi);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log(ex, $"Failed to correct description for '{wi.Id}', rev '{rev.ToString()}'.");
                    }
                }

                _context.Journal.MarkRevProcessed(rev.ParentOriginId, wi.Id, rev.Index);

                Logger.Log(LogLevel.Debug, $"Imported revision.");

                //Logger.Log(LogLevel.Info, "***rev.ParentOriginId:" + rev.ParentOriginId);
                Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemType wit = wi.Type;
                String wiTypeName = wit.Name;
                //Logger.Log(LogLevel.Info, "**** " + wi.Project + "***** " + wi.Id + "***** " + rev.ToString());
                //String thisRevisionTemp = rev.ToString().Split(',').Last().Trim();
                //String thisRevision = thisRevisionTemp.ToString().Split(' ').Last().Trim();
                // Logger.Log(LogLevel.Info, "***thisRevision:" + thisRevision);
                // if (wiTypeName == "Test Case" && int.Parse(thisRevision) == totalRevisions-1)
                //if (wiTypeName == "Test Case" && int.Parse(thisRevision) == 0)
                if (wiTypeName == "Test Case")
                {
                    //string path = "C:\\migration\\UKBL\\data1\\" + rev.ParentOriginId + ".json";
                    string path = jsonPath + "\\" + rev.ParentOriginId + ".json";
                    string text = File.ReadAllText(path);
                    //JToken test = JArray.Parse(text).Last();
                    var jsonobject = JObject.Parse(text);
                    // Logger.Log(LogLevel.Info, "******* readjson: " + jsonobject);                    
                    JProperty allRevisions = jsonobject.Property("Revisions");
                    JToken lastrevision = allRevisions.Value.Last;
                    String lastindex = lastrevision["Index"].ToString();
                    //Logger.Log(LogLevel.Info, "******* index value: " + allRevisions.Value);
                    //Logger.Log(LogLevel.Info, "******* index value: " + allRevisions.Value.Last);
                    //Logger.Log(LogLevel.Info, "******* last index value: " + lastrevision["Index"].ToString());

                    if (int.Parse(lastindex) == thisRevision)
                    {
                        WebClient client = new WebClient();
                        String userName = "ranjan_jira_to_ado";
                        String passWord = "myjira";
                        string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(userName + ":" + passWord));
                        client.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                        String jiraURL = "https://jira.thomsonreuters.com/rest/api/2/issue/" + rev.ParentOriginId;
                        ITestManagementService service = (ITestManagementService)collection.GetService(typeof(ITestManagementService));
                        ITestManagementTeamProject testProject = service.GetTeamProject(wi.Project);
                        ITestCase testCase = testProject.TestCases.Find(wi.Id);
                        if (CheckURLNotValid(client, jiraURL))
                        {
                            Logger.Log(LogLevel.Info, "******* OLD JIRA ***********");
                            //Logger.Log(LogLevel.Info, "******* " + wi.Project + "************" + wi.Id);
                            ITestStep step = testCase.CreateTestStep();
                            // step.Title = "*****test test test**********";
                            // step.Description = "test test test";
                            // step.ExpectedResult = "";
                            testCase.Actions.Add(step);
                            //testCase.Save();
                        }
                        else
                        {
                            Logger.Log(LogLevel.Info, "******* NEW JIRA ***********");                            
                            
                            string jiraIssueText = client.DownloadString(jiraURL);
                            var jsonobject1 = JObject.Parse(jiraIssueText);
                            String issueID = jsonobject1.Property("id").Value.ToString();
                            //Logger.Log(LogLevel.Info, "***IssueID : " + issueID);
                            string teststepText = client.DownloadString("https://jira.thomsonreuters.com/rest/zapi/latest/teststep/" + issueID);
                            var jsonobject2 = JObject.Parse(teststepText);
                            //JProperty allTeststep = jsonobject2.Property("stepBeanCollection");
                            //var jsonobject3 = jsonobject2["stepBeanCollection"];
                            JProperty allTeststep = null;
                            allTeststep = jsonobject2.Property("stepBeanCollection");
                            //Logger.Log(LogLevel.Info, "***allTeststep.Count : " + allTeststep.Count);
                            //Logger.Log(LogLevel.Info, "***allTeststep.Value.ToList().Count : " + allTeststep.Value.ToList().Count);
                            if (allTeststep == null || allTeststep.Value.ToList().Count == 0)
                            {
                                ITestStep stepTemp = testCase.CreateTestStep();
                                testCase.Actions.Add(stepTemp);
                                //testCase.Save();
                            }
                            foreach (JToken x in allTeststep.Value)
                            {
                                String teststep = x["step"].ToString();
                                String data = x["data"].ToString();
                                String result = x["result"].ToString();
                                ITestStep step1 = testCase.CreateTestStep();
                                step1.Title = "Step: " + teststep + "\nData: " + data;
                                // step1.Description = "Step: " + teststep + "\nData: " + data;
                                step1.ExpectedResult = result;
                                // Logger.Log(LogLevel.Info, "***step1.Title : " + step1.Title);
                                // Logger.Log(LogLevel.Info, "***step1.ExpectedResult : " + step1.ExpectedResult);
                                testCase.Actions.Add(step1);
                                // testCase.Save();
                            } // end foreach allTeststep

                            string testExecutionText = client.DownloadString("https://jira.thomsonreuters.com/rest/zapi/latest/execution?issueId=" + issueID);
                            var jsonobject3 = JObject.Parse(testExecutionText);
                            //Logger.Log(LogLevel.Info, "***jsonobject3 : " + jsonobject3);
                            //JProperty allExecutions = jsonobject3.Property("executions");
                            //var jsonobject3 = jsonobject3["executions"];                               
                            JProperty allExecutions = jsonobject3.Property("executions");
                            JProperty allStatusText = jsonobject3.Property("status");
                            //Logger.Log(LogLevel.Info, "***allExecutions : " + allExecutions);
                            // int iterationId = 0;
                            foreach (JToken execution in allExecutions.Value)
                            {
                                string releaseStatus = "";
                                string cycleId = "";
                                string cycleName = "";
                                string folderId = "";
                                string folderName = "";
                                string versionId = "";
                                string versionName = "";  
                                string jiraIssueKey = "";
                                string comment = "";
                                string executedByDisplay = "";
                                string executedByID = "";
                                string executedOn = "";
                                string executedOnVal = "";
                                string executionStatusText = "";
                                //string executionOutcome = "";


                                //Logger.Log(LogLevel.Info, "***execution : " + execution);
                                //Logger.Log(LogLevel.Info, "***cycleId : " + execution["cycleId"]);
                                cycleId = execution["cycleId"].ToString();
                                //Logger.Log(LogLevel.Info, "***cycleId : " + cycleId);
                                cycleName = execution["cycleName"].ToString();
                                //Logger.Log(LogLevel.Info, "***cycleName : " + cycleName);
                                if (execution["folderId"] != null)
                                {
                                    folderId = execution["folderId"].ToString();
                                    //Logger.Log(LogLevel.Info, "***folderId : " + folderId);
                                    folderName = execution["folderName"].ToString();
                                    //Logger.Log(LogLevel.Info, "***folderName : " + folderName);
                                }
                                versionId = execution["versionId"].ToString();
                                //Logger.Log(LogLevel.Info, "***versionId : " + versionId);
                                versionName = execution["versionName"].ToString();
                                //Logger.Log(LogLevel.Info, "***versionName : " + versionName);
                                String projectId = execution["projectId"].ToString();
                                //Logger.Log(LogLevel.Info, "***projectId : " + projectId);
                                String projectKey = execution["projectKey"].ToString();
                                //Logger.Log(LogLevel.Info, "***projectKey : " + projectKey);
                                String executionStatus = execution["executionStatus"].ToString();
                                //Logger.Log(LogLevel.Info, "***executionStatus : " + executionStatus);
                                
                                jiraIssueKey = execution["issueKey"].ToString();
                                comment = execution["comment"].ToString();
                                if (execution["executedByDisplay"] != null)
                                {
                                    executedByDisplay = execution["executedByDisplay"].ToString();
                                }
                                if (execution["executedBy"] != null)
                                {
                                    executedByID = execution["executedBy"].ToString();
                                }
                                if (execution["executedOn"] != null)
                                {
                                    executedOn = execution["executedOn"].ToString();
                                }
                                if (execution["executedOnVal"] != null)
                                {
                                    executedOnVal = execution["executedOnVal"].ToString();
                                }

                                //Logger.Log(LogLevel.Info, "***before foreach allStatusText");
                                //Logger.Log(LogLevel.Info, "***allStatusText : " + allStatusText);
                                foreach (JProperty statusText in allStatusText.Value)
                                {
                                    //Logger.Log(LogLevel.Info, "***statusText : " + statusText);
                                    JToken statusTextValue = statusText.Value;
                                    //Logger.Log(LogLevel.Info, "***statusTextValue : " + statusText.Value);
                                    //Logger.Log(LogLevel.Info, "***statusTextID : " + statusText["id"].ToString());
                                    //Logger.Log(LogLevel.Info, "***statusTextName : " + statusText["name"].ToString());
                                    //Logger.Log(LogLevel.Info, "***statusTextId : " + statusTextValue["id"].ToString());
                                    if (statusTextValue["id"].ToString() == executionStatus)
                                    {
                                        executionStatusText = statusTextValue["name"].ToString();
                                        //Logger.Log(LogLevel.Info, "***executionStatusText : " + executionStatusText);
                                        break;
                                    }
                                }

                                string projectVersionText = client.DownloadString("https://jira.thomsonreuters.com/rest/api/2/project/" + int.Parse(projectId) + "/version");
                                var jsonobject4 = JObject.Parse(projectVersionText);
                                //Logger.Log(LogLevel.Info, "***jsonobject4 : " + jsonobject4);
                                JProperty allVersions = jsonobject4.Property("values");
                                //Logger.Log(LogLevel.Info, "***allVersions : " + allVersions);
                                foreach (JToken version in allVersions.Value)
                                {                                    
                                    if (version["id"].ToString() == versionId)
                                    {
                                        releaseStatus = version["released"].ToString();
                                        //Logger.Log(LogLevel.Info, "***releaseStatus : " + releaseStatus);
                                        break;
                                    }
                                }
                                if (versionId == "-1")
                                {
                                    releaseStatus = "False";
                                    //Logger.Log(LogLevel.Info, "***releaseStatus : " + releaseStatus);
                                }

                                ITestPlanCollection plans = null;

                                /*ITestPlan defaultPlan = testProject.TestPlans.Create();
                                defaultPlan.Name = "Default";
                                defaultPlan.Save();  */                              
                                plans = testProject.TestPlans.Query("Select * From TestPlan");
                                //Logger.Log(LogLevel.Info, "000111");
                                var check = plans.Where(pp => pp.Name == projectKey);
                                //Logger.Log(LogLevel.Info, "11111");
                                /*if (plans != null &&  check!=null && check.Any()) // at least one test plan already exists
                                {
                                    //Logger.Log(LogLevel.Info, "2222222");
                                    var p = plans.Where(pp => pp.Name == projectKey).First();
                                    //Logger.Log(LogLevel.Info, "333333");
                                    //Logger.Log(LogLevel.Info, "***Plan Name : " + p.Name);
                                }
                                else // create the non-existent test plan*/
                                if (plans == null || check == null || !check.Any())
                                {
                                    ITestPlan testPlan = testProject.TestPlans.Create();
                                    testPlan.Name = projectKey;
                                    testPlan.Save();
                                    //Logger.Log(LogLevel.Info, "***Test Plan Name : " + testPlan.Name);
                                }
                                ITestPlan plan = testProject.TestPlans.Query("Select * From TestPlan").Where(pp => pp.Name == projectKey).First();

                                int suiteId;
                                int lastSuiteId;

                                IEnumerable<ITestSuiteEntry> pSuites = plan.RootSuite.Entries.Where(s => s.Title == "All Releases");
                                if (pSuites == null || !pSuites.Any())
                                {
                                    IStaticTestSuite newSuite = testProject.TestSuites.CreateStatic();
                                    newSuite.Title = "All Releases";
                                    suiteId = newSuite.Id;
                                    plan.RootSuite.Entries.Add(newSuite);
                                    //IStaticTestSuite todelete = newSuite;
                                    //Logger.Log(LogLevel.Info, "-----ReferenceEquals 1----- " + Object.ReferenceEquals(newSuite, todelete));
                                    //Logger.Log(LogLevel.Info, "-----ID Comparison 1----- " + newSuite.Id + "-----ID Comparison 2----- " + todelete.Id);
                                    plan.Save();
                                }
                                //suiteId = plan.RootSuite.Entries.Where(s => s.Title == "All Releases").First().Id;
                                IStaticTestSuite pSuite = plan.RootSuite.Entries.Where(s => s.Title == "All Releases").First().TestSuite as IStaticTestSuite;
                                //Logger.Log(LogLevel.Info, "-----ReferenceEquals 2----- " + Object.ReferenceEquals(pSuite, plan.RootSuite.Entries.Where(s => s.Title == "All Releases").First().TestSuite as IStaticTestSuite));
                                //Logger.Log(LogLevel.Info, "-----ID Comparison 3----- " + pSuite.Id + "-----ID Comparison 4----- " + plan.RootSuite.Entries.Where(s => s.Title == "All Releases").First().TestSuite.Id);

                                if (!String.IsNullOrEmpty(releaseStatus))
                                {
                                    if (releaseStatus == "True")
                                    {
                                        releaseStatus = "Released";
                                    }
                                    else
                                    {
                                        releaseStatus = "Unreleased";
                                    }
                                    
                                    IEnumerable<ITestSuiteEntry> pSuitesL1 = pSuite.Entries.Where(s => s.Title == releaseStatus);
                                    if (pSuitesL1 == null || !pSuitesL1.Any())
                                    {
                                        IStaticTestSuite newSuite = testProject.TestSuites.CreateStatic();
                                        newSuite.Title = releaseStatus;
                                        lastSuiteId = newSuite.Id;
                                        pSuite.Entries.Add(newSuite);
                                        plan.Save();
                                    }
                                    //lastSuiteId = pSuite.Entries.Where(s => s.Title == releaseStatus).First().Id;                                  
                                    IStaticTestSuite pSuiteLast = pSuite.Entries.Where(s => s.Title == releaseStatus).First().TestSuite as IStaticTestSuite;
                                    //Logger.Log(LogLevel.Info, "-----ReferenceEquals 3----- " + Object.ReferenceEquals(pSuiteLast, pSuite.Entries.Where(s => s.Title == releaseStatus).First().TestSuite as IStaticTestSuite));
                                    //Logger.Log(LogLevel.Info, "-----ID Comparison 5----- " + pSuiteLast.Id + "-----ID Comparison 6----- " + pSuite.Entries.Where(s => s.Title == releaseStatus).First().TestSuite.Id);

                                    IStaticTestSuite versionSuite = null;

                                    if (!String.IsNullOrEmpty(versionName))
                                    {
                                        IEnumerable<ITestSuiteEntry> pSuitesL2 = pSuiteLast.Entries.Where(s => s.Title == versionName);
                                        if (pSuitesL2 == null || !pSuitesL2.Any())
                                        {
                                            IStaticTestSuite newSuite = testProject.TestSuites.CreateStatic();
                                            newSuite.Title = versionName;
                                            lastSuiteId = newSuite.Id;
                                            pSuiteLast.Entries.Add(newSuite);
                                            plan.Save();
                                        }
                                        //lastSuiteId = pSuiteLast.Entries.Where(s => s.Title == versionName).First().Id;
                                        //IStaticTestSuite temp = pSuiteLast.Entries.Where(s => s.Title == versionName).First().TestSuite as IStaticTestSuite;
                                        pSuiteLast = pSuiteLast.Entries.Where(s => s.Title == versionName).First().TestSuite as IStaticTestSuite;
                                        //Logger.Log(LogLevel.Info, "-----ReferenceEquals 4----- " + Object.ReferenceEquals(pSuiteLast, temp));
                                        //Logger.Log(LogLevel.Info, "-----ID Comparison 7----- " + pSuiteLast.Id + "-----ID Comparison 8----- " + temp.Id);

                                        versionSuite = pSuiteLast;
                                    }

                                    if (!String.IsNullOrEmpty(cycleName))
                                    {
                                        IEnumerable<ITestSuiteEntry> pSuitesL3 = pSuiteLast.Entries.Where(s => s.Title == cycleName);
                                        if (pSuitesL3 == null || !pSuitesL3.Any())
                                        {
                                            IStaticTestSuite newSuite = testProject.TestSuites.CreateStatic();
                                            newSuite.Title = cycleName;
                                            lastSuiteId = newSuite.Id;
                                            pSuiteLast.Entries.Add(newSuite);
                                            plan.Save();
                                        }
                                        //lastSuiteId = pSuiteLast.Entries.Where(s => s.Title == versionName).First().Id;
                                        //IStaticTestSuite temp = pSuiteLast.Entries.Where(s => s.Title == cycleName).First().TestSuite as IStaticTestSuite;
                                        pSuiteLast = pSuiteLast.Entries.Where(s => s.Title == cycleName).First().TestSuite as IStaticTestSuite;
                                        //Logger.Log(LogLevel.Info, "-----ReferenceEquals 5----- " + Object.ReferenceEquals(pSuiteLast, temp));
                                        //Logger.Log(LogLevel.Info, "-----ID Comparison 9----- " + pSuiteLast.Id + "-----ID Comparison 10----- " + temp.Id);

                                    }

                                    if (!String.IsNullOrEmpty(folderName))
                                    {
                                        IEnumerable<ITestSuiteEntry> pSuitesL4 = pSuiteLast.Entries.Where(s => s.Title == folderName);
                                        if (pSuitesL4 == null || !pSuitesL4.Any())
                                        {
                                            IStaticTestSuite newSuite = testProject.TestSuites.CreateStatic();
                                            newSuite.Title = folderName;
                                            lastSuiteId = newSuite.Id;
                                            pSuiteLast.Entries.Add(newSuite);
                                            plan.Save();
                                        }
                                        //lastSuiteId = pSuiteLast.Entries.Where(s => s.Title == versionName).First().Id;
                                        //IStaticTestSuite temp = pSuiteLast.Entries.Where(s => s.Title == folderName).First().TestSuite as IStaticTestSuite;
                                        pSuiteLast = pSuiteLast.Entries.Where(s => s.Title == folderName).First().TestSuite as IStaticTestSuite;
                                        //Logger.Log(LogLevel.Info, "-----ReferenceEquals 6----- " + Object.ReferenceEquals(pSuiteLast, temp));
                                        //Logger.Log(LogLevel.Info, "-----ID Comparison 11----- " + pSuiteLast.Id + "-----ID Comparison 12----- " + temp.Id);

                                    }

                                    //bool toAdd = true;
                                       
                                    if (pSuiteLast.AllTestCases.Count> 0)
                                    {
                                        foreach (var currentTestCase in pSuiteLast.AllTestCases)
                                        {
                                            if (currentTestCase.WorkItem.Title.Contains(jiraIssueKey))
                                            {
                                                //Logger.Log(LogLevel.Info, "This testcase already added to the TestSuite");
                                                // toAdd = false;
                                                pSuiteLast.Entries.Remove(currentTestCase);
                                                break;
                                            }
                                        }

                                    }
                                    //if (toAdd)
                                    //{                                        
                                        pSuiteLast.Entries.Add(testCase);
                                        plan.Save();
                                    //}

                                    ITestPointCollection testPointColl = plan.QueryTestPoints("SELECT * FROM TestPoint WHERE SuiteId = " + pSuiteLast.Id + " AND TestCaseId = " + testCase.Id);
                                    if (testPointColl.Count > 0)
                                    {
                                        ITestPoint testpoint = testPointColl.First();
                                        //var testRun = testProject.TestRuns.Find(122);
                                        //Logger.Log(LogLevel.Info, "testProject.TestRuns.Count " + testProject.TestRuns.Count);
                                        
                                        var testRuns = testProject.TestRuns.Query("select * From TestRun WHERE Title = 'TestSuite ID: " + pSuiteLast.Id + "'");
                                        //int testRunId = testpoint.MostRecentRunId;
                                        //Logger.Log(LogLevel.Info, "-----testRunId----- " + testRunId);
                                        ITestRun run = null;
                                        bool isNew = false;
                                        /*if (!testRuns.Any())
                                        //if (testRunId <= 0)
                                        {
                                            //Logger.Log(LogLevel.Info, "-----TestRun NOT present-----");
                                            run = plan.CreateTestRun(false);
                                            run.Title = " TestSuite ID: " + pSuiteLast.Id + "; TestPoint ID: " + testpoint.Id.ToString() + "; TestCase ID: " + testCase.Id;
                                            //run.Title = " TestSuite ID: " + pSuiteLast.Id;
                                            run.AddTestPoint(testpoint, null);
                                            run.Save();
                                            run.Refresh();
                                            isNew = true;
                                        }
                                        else
                                        {
                                            //var testRuns = testProject.TestRuns.Query("select * From TestRun WHERE Id = '" + testRunId + "'");
                                            //Logger.Log(LogLevel.Info, "-----TestRun already present-----");
                                            //run = testRuns.First();
                                            run = testRuns.First();                                             
                                        }*/
                                        run = plan.CreateTestRun(false);
                                        run.Title = " TestSuite ID: " + pSuiteLast.Id + "; TestPoint ID: " + testpoint.Id.ToString() + "; TestCase ID: " + testCase.Id;
                                        run.AddTestPoint(testpoint, null);
                                        run.Save();
                                        run.Refresh();


                                        var testResults = run.QueryResults();
                                        //ITestCaseResult testResult = testResults.Single(r => r.TestCaseId == testCase.Id);
                                        //run.AddTest(int.Parse(issueID), pSuiteLast.DefaultConfigurations[0].Id, pSuiteLast.Plan.Owner);
                                        //Logger.Log(LogLevel.Info, "testpoint.ConfigurationId is " + testpoint.ConfigurationId);
                                        //run.AddTest(int.Parse(issueID), pSuiteLast.DefaultConfigurations[0].Id, pSuiteLast.Plan.Owner);
                                        //run.AddTest(int.Parse(issueID), testpoint.ConfigurationId, pSuiteLast.Plan.Owner);
                                        //Logger.Log(LogLevel.Info, "Afer TestRun.AddTest Call");
                                        int countTestResults = 0;
                                        countTestResults = testResults.Count;
                                        //Logger.Log(LogLevel.Info, "countTestResults: " + countTestResults);
                                        ITestCaseResult testResult = null;
                                        if (isNew)
                                        {
                                            testResult = run.QueryResults()[0];
                                        }
                                        else
                                        {
                                            //testResult = run.QueryResults()[countTestResults];
                                            testResult = run.QueryResults()[0];
                                        }



                                        //ITestCaseResult testResult = run.QueryResults()[0]; // this works fine
                                        //ITestCaseResult testResult = run.QueryResults()[countTestResults-1]; 
                                        //ITestCaseResult testResult = run.QueryResults().Single(r => r.TestPointId == testpoint.Id);
                                        //ITestCaseResultCollection testResults = testProject.TestResults.ByTestId(testCase.Id);
                                        //iterationId = iterationId + 1;

                                        var iteration = testResult.CreateIteration(1);
                                        //var iteration = testResult.CreateIteration(iterationId);

                                        //testResult.DateStarted = result.DateStarted;
                                        if (!string.IsNullOrEmpty(executedOn))
                                        {
                                            testResult.DateCompleted = Convert.ToDateTime(executedOn);
                                            iteration.DateCompleted = Convert.ToDateTime(executedOn);                                            
                                            iteration.DateStarted = Convert.ToDateTime(executedOn);                                            
                                        }
                                            
                                        //testResult.Outcome = executionOutcome;
                                        if (executionStatusText == "PASS")
                                        {
                                            //testResult.Outcome = TestOutcome.Passed;
                                            iteration.Outcome = TestOutcome.Passed;
                                        }
                                        else if (executionStatusText == "FAIL")
                                        {
                                            //testResult.Outcome = TestOutcome.Failed;
                                            iteration.Outcome = TestOutcome.Failed;
                                        }
                                        else if (executionStatusText == "WIP")
                                        {
                                            //testResult.Outcome = TestOutcome.InProgress;
                                            iteration.Outcome = TestOutcome.InProgress;
                                        }
                                        else if (executionStatusText == "BLOCKED")
                                        {
                                            //testResult.Outcome = TestOutcome.Blocked;
                                            iteration.Outcome = TestOutcome.Blocked;
                                        }
                                        else if (executionStatusText == "UNEXECUTED")
                                        {
                                            //testResult.Outcome = TestOutcome.NotExecuted;
                                            iteration.Outcome = TestOutcome.NotExecuted;
                                        }
                                        else
                                        {
                                            //testResult.Outcome = TestOutcome.NotExecuted;
                                            iteration.Outcome = TestOutcome.NotExecuted;
                                        }
                                        //testResult.Comment = comment;
                                        iteration.Comment = comment;
                                        //var runByUser = $"{System.Environment.UserDomainName}\\{System.Environment.UserName}";
                                        //TeamFoundationIdentity runByUser = new TeamFoundationIdentity();
                                        //testResult.RunBy = executedByDisplay;
                                        //testResult.OwnerName = executedByDisplay;
                                        //testResult.State = TestResultState.Completed;
                                        testResult.Iterations.Add(iteration);
                                        testResult.State = TestResultState.Completed;
                                        testResult.Outcome = iteration.Outcome;

                                        //testResult.AssociateWorkItem(wi);

                                        //Logger.Log(LogLevel.Info, "for each wi.Links");
                                        foreach (var link in wi.Links.OfType<RelatedLink>().ToList())
                                        {
                                            var relatedWorkItem = GetWorkItem(link.RelatedWorkItemId);
                                            
                                            //Logger.Log(LogLevel.Info, "related work item is " + relatedWorkItem.Id);
                                        }

                                        //Logger.Log(LogLevel.Info, "Add Link");
                                        String linkName = "Microsoft.VSTS.Common.TestedBy-Forward";
                                        var props = linkName?.Split('-');
                                        var linkType = wi.Project.Store.WorkItemLinkTypes.SingleOrDefault(lt => lt.ReferenceName == props?[0]);
                                        if (linkType != null)
                                        {
                                            WorkItemLinkTypeEnd linkEnd = null;
                                            linkEnd = linkType.ForwardEnd;
                                            //Logger.Log(LogLevel.Info, "linkEnd" + linkEnd.ToString());
                                            if (linkEnd != null)
                                            {
                                                //var relatedLink = new RelatedLink(linkEnd, pSuiteLast.Id);
                                                var relatedLink = new RelatedLink(linkEnd, versionSuite.Id);
                                                if (!IsDuplicateWorkItemLink(wi.Links, relatedLink))
                                                {
                                                    wi.Links.Add(relatedLink);
                                                    wi.Save();
                                                    //Logger.Log(LogLevel.Info, "created link");
                                                }
                                                else
                                                {
                                                    //Logger.Log(LogLevel.Info, "Did not create Duplicate Link");
                                                }
                                            }
                                            
                                        } // linktype

                                        //Logger.Log(LogLevel.Info, "********" + run.Title);
                                        testResult.Save();
                                        testResult.Refresh();
                                        run.Save();
                                        run.Refresh();
                                        testpoint.Save();
                                        testpoint.Refresh();
                                        pSuiteLast.Refresh();
                                    }

                                } // end if releasestatus


                            } // end foreach allexecutions
                              
                        }

                        testCase.Save();     
                        
/*                        Logger.Log(LogLevel.Info, "************");
                        IStaticTestSuite newSuite = testProject.TestSuites.CreateStatic();
                        Logger.Log(LogLevel.Info, "+++++++++++++");
                        newSuite.Title = "All Releases";
                        testPlan.RootSuite.Entries.Add(newSuite);
                        IStaticTestSuite newSubSuite1 = testProject.TestSuites.CreateStatic();
                        newSubSuite1.Title = "Unreleased";
                        IStaticTestSuite newSubSuite2 = testProject.TestSuites.CreateStatic();
                        newSubSuite2.Title = "Released";
                        IStaticTestSuite pSuite = testPlan.RootSuite.Entries.Where(s => s.Title == "All Releases").First().TestSuite as IStaticTestSuite;
                        pSuite.Entries.Add(newSubSuite1);
                        pSuite.Entries.Add(newSubSuite2);
                        testPlan.Save();
                        IStaticTestSuite newSubSubSuite1 = testProject.TestSuites.CreateStatic();
                        newSubSubSuite1.Title = "RPC 5.0";
                        IStaticTestSuite pSubSuite = pSuite.Entries.Where(s => s.Title == "Unreleased").First().TestSuite as IStaticTestSuite;
                        pSubSuite.Entries.Add(newSubSubSuite1);
                        newSubSubSuite1.Entries.Add(testCase);
                        testPlan.Save();

                        // Create test configuration. You can reuse this instead of creating a new config everytime.
                        ITestConfiguration config = CreateTestConfiguration(testProject, string.Format("My test config {0}", DateTime.Now));

                        // Create test points. 
                        IList<ITestPoint> testPoints = CreateTestPoints(newSubSubSuite1,
                                                                        testPlan,                                                                        
                                                                        new IdAndName[] { new IdAndName(config.Id, config.Name) });

                        ITestRun testRun = testProject.TestRuns.Create();

                        testRun.DateStarted = DateTime.Now;
                        //testRun.AddTestPoint(testPoints, _currentIdentity);
                        // Create test run using test points.
                        ITestRun run = CreateTestRun(testProject, testPlan, testPoints);
                        //testRun.DateCompleted = DateTime.Now;
                        //testRun.Save(); // so results object is created

                        var result1 = testRun.QueryResults()[0];
                        //result.Owner = _currentIdentity;
                        //result.Owner = _currentIdentity;
                        //result.RunBy = _currentIdentity;
                        // result.RunBy = _currentIdentity;
                        result1.State = TestResultState.Completed;
                        result1.DateStarted = DateTime.Now;
                        result1.Duration = new TimeSpan(0L);
                        result1.DateCompleted = DateTime.Now.AddMinutes(0.0);

                        var iteration = result1.CreateIteration(1);
                        iteration.DateStarted = DateTime.Now;
                        iteration.DateCompleted = DateTime.Now;
                        iteration.Duration = new TimeSpan(0L);
                        //iteration.Comment = "Run from ADO Test Steps Editor by " + _currentIdentity.DisplayName;
                        iteration.Comment = "Run from ADO Test Steps Editor by " + "prashant.ranjan@thomsonreuters.com";

                        for (int actionIndex = 0; actionIndex < testCase.Actions.Count; actionIndex++)
                        {
                            var testAction = testCase.Actions[actionIndex];
                            if (testAction is ISharedStepReference)
                                continue;

                            //var userStep = _testEditInfo.SimpleSteps[actionIndex];

                            var stepResult = iteration.CreateStepResult(testAction.Id);
                            stepResult.ErrorMessage = String.Empty;
                            //stepResult.Outcome = userStep.Outcome;
                            stepResult.Outcome = TestOutcome.Passed;

                            /*foreach (var attachmentPath in userStep.AttachmentPaths)
                            {
                                var attachment = stepResult.CreateAttachment(attachmentPath);
                                stepResult.Attachments.Add(attachment);
                            }*/
/*
                            iteration.Actions.Add(stepResult);
                        }

                        var overallOutcome = TestOutcome.Passed;
*/                        /*var overallOutcome = _testEditInfo.SimpleSteps.Any(s => s.Outcome != TestOutcome.Passed)
                            ? TestOutcome.Failed
                            : TestOutcome.Passed;*/
/*
                        iteration.Outcome = overallOutcome;

                        result1.Iterations.Add(iteration);

                        result1.Outcome = overallOutcome;
                        result1.Save(false);

*/                    


                    } // if thisrevision

                } // if testcase

                string path1 = jsonPath + "\\" + rev.ParentOriginId + ".json";
                string text1 = File.ReadAllText(path1);
                //JToken test = JArray.Parse(text).Last();
                var jsonobject5 = JObject.Parse(text1);
                // Logger.Log(LogLevel.Info, "******* readjson: " + jsonobject);
                JProperty allRevisions1 = jsonobject5.Property("Revisions");
                JToken lastrevision1 = allRevisions1.Value.Last;
                String lastindex1 = lastrevision1["Index"].ToString();
                
                var workitemID = wi.Id;
                //var workitemID = 163064;
                var jiraIssue = rev.ParentOriginId;

                wi.Close();

                if (int.Parse(lastindex1) == thisRevision)
                {
                    //var baseUrl = "https://dev.azure.com/tr-content-platform/Test/_apis/wit/workitems/$User%20Story";
                    //var baseUrl = "https://dev.azure.com/tr-content-platform";
                    var baseUrl = urlValue;
                    var pat = "";
                    // var vssConnection = new VssConnection(new Uri(baseUrl), new VssBasicCredential(string.Empty, pat));
                    //Microsoft.TeamFoundation.WorkItemTracking.WebApi.WorkItemTrackingHttpClient _workItemTrackingHttpClient = vssConnection.GetClient<Microsoft.TeamFoundation.WorkItemTracking.WebApi.WorkItemTrackingHttpClient>();
                    var document = new Microsoft.VisualStudio.Services.WebApi.Patch.Json.JsonPatchDocument();

                    //if (htLateUpdates.Count > 0)
                    if (htParent.ContainsKey(jiraIssue))
                    {
                        Hashtable htValues = (Hashtable)htParent[jiraIssue];
                        //foreach (var key in htLateUpdates.Keys)
                        foreach (var key in htValues.Keys)
                        {
                            //Logger.Log(LogLevel.Info, htValues.Count + ":" + key + ":" + htValues[key]);
                            String fieldPath = "/fields/" + key;
                            document.Add(new Microsoft.VisualStudio.Services.WebApi.Patch.Json.JsonPatchOperation()
                            {
                                Operation = Microsoft.VisualStudio.Services.WebApi.Patch.Operation.Add,                                
                                Path = fieldPath,                                
                                Value = htValues[key]

                            });
                        }
                        var vssConnection = new VssConnection(new Uri(baseUrl), new VssBasicCredential(string.Empty, pat));
                        Microsoft.TeamFoundation.WorkItemTracking.WebApi.WorkItemTrackingHttpClient _workItemTrackingHttpClient = vssConnection.GetClient<Microsoft.TeamFoundation.WorkItemTracking.WebApi.WorkItemTrackingHttpClient>();
                        var workItem = _workItemTrackingHttpClient.UpdateWorkItemAsync(document, workitemID).Result;
                    }
                } // if for board values

                return true;
            }
            catch (AbortMigrationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.Log(ex, $"Failed to import revisions for '{wi.Id}'.");
                return false;
            }
        }

        private bool CheckURLNotValid(WebClient client, String jiraURL)
        {
            try
            {
                string responseText = client.DownloadString(jiraURL);
            }
            catch
            {
                return true;
            }
            return false;
        }

        private void EnsureClassificationFields(WiRevision rev)
        {
            if (!rev.Fields.HasAnyByRefName(WiFieldReference.AreaPath))
                rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.AreaPath, Value = "" });

            if (!rev.Fields.HasAnyByRefName(WiFieldReference.IterationPath))
                rev.Fields.Add(new WiField() { ReferenceName = WiFieldReference.IterationPath, Value = "" });
        }

        private ITestConfiguration CreateTestConfiguration(ITestManagementTeamProject project, string title)
        {
            ITestConfiguration configuration = project.TestConfigurations.Create();
            configuration.Name = title;
            configuration.Description = "DefaultConfig";
            //configuration.Values.Add(new KeyValuePair<string, string>("Browser", "IE"));
            configuration.Values.Add(new KeyValuePair<string, string>("Browser", "Chrome"));
            configuration.Save();
            return configuration;
        }

        public static IList<ITestPoint> CreateTestPoints(IStaticTestSuite testSuite,
                                                         ITestPlan testPlan,                                                         
                                                         IList<IdAndName> testConfigs)
        {
            testSuite.SetEntryConfigurations(testSuite.Entries, testConfigs);
            testPlan.Save();

            ITestPointCollection tpc = testPlan.QueryTestPoints("SELECT * FROM TestPoint WHERE SuiteId = " + testSuite.Id);
            return new List<ITestPoint>(tpc);
        }

        private static ITestRun CreateTestRun(ITestManagementTeamProject project,
                                             ITestPlan plan,
                                             IList<ITestPoint> points)
        {
            ITestRun run = plan.CreateTestRun(false);
            foreach (ITestPoint tp in points)
            {
                run.AddTestPoint(tp, null);
            }

            run.Save();
            return run;
        }

        #endregion
    }
}