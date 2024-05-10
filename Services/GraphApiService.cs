using Azure.Identity;
using Coginov.GraphApi.Library.Enums;
using Coginov.GraphApi.Library.Helpers;
using Coginov.GraphApi.Library.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using SystemFile = System.IO.File;
using System.Web;
using Microsoft.Graph.Models;
using Microsoft.Graph.Search.Query;
using Microsoft.Graph.Users.Item.SendMail;
using Microsoft.Kiota.Abstractions.Authentication;
using Azure.Core;
using Microsoft.Kiota.Http.HttpClientLibrary.Middleware;
using System.Text.RegularExpressions;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Http.HttpClientLibrary.Middleware.Options;
using DriveUpload = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using System.Reflection.Metadata;
using Microsoft.Graph.Models.Security;
using Microsoft.Kiota.Abstractions;

namespace Coginov.GraphApi.Library.Services
{
    public class GraphApiService : IGraphApiService, IDisposable
    {
        private readonly ILogger logger;
        private AuthenticationConfig authConfig;
        private HttpClient graphHttpClient;
        private GraphServiceClient graphServiceClient;
        private List<DriveConnectionInfo> drivesConnectionInfo = new List<DriveConnectionInfo>();
        private DriveConnectionType connectionType;
        private string userId;
        private string siteId;
        private AuthenticationToken authenticationToken;
        private string msalCachePath;
        private string msalCacheFileName;
        private MsalCacheHelper msalCacheHelper;
        private IPublicClientApplication pca;
        private IConfidentialClientApplication cca;
        private InteractiveBrowserCredential ibc;

        // SharePointOnline
        private string siteUrl;
        private string[] docLibraries;

        // OneDrive
        private string oneDriveUserAccount;

        // MsTeams
        private string[] teams;

        // This setting is to simulate GraphApi errors during development
        private bool useChaosHandler = false;

        public GraphApiService(ILogger logger)
        {
            this.logger = logger;
        }

        public void Dispose()
        {
            if (msalCacheHelper != null)
            {
                msalCacheHelper.UnregisterCache(pca.UserTokenCache);
                SystemFile.Delete(Path.Combine(msalCachePath, msalCacheFileName));
            }
        }

        public async Task<bool> InitializeGraphApi(AuthenticationConfig authenticationConfig, bool forceInit = true)
        {
            this.authConfig = authenticationConfig;
            return await IsInitialized(forceInit);
        }

        public async Task<bool> InitializeSharePointOnlineConnection(AuthenticationConfig authenticationConfig, string siteUrl, string[] docLibraries, bool forceInit = false)
        {
            this.authConfig = authenticationConfig;
            this.siteUrl = siteUrl; 
            this.docLibraries = docLibraries;

            if (!await IsInitialized(forceInit))
                return false;

            try
            {
                connectionType = DriveConnectionType.SharePoinConnection;
                siteId = await GetSiteId(siteUrl);
                if (string.IsNullOrWhiteSpace(siteId))
                    return false;

                drivesConnectionInfo?.Clear();
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorInitializingSPConnection}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return false;
            }

            return true;
        }

        public async Task<bool> InitializeOneDriveConnection(AuthenticationConfig authenticationConfig, string userAccount, bool forceInit = false)
        {
            this.authConfig = authenticationConfig;
            this.oneDriveUserAccount = userAccount;

            if (!await IsInitialized(forceInit))
                return false;

            try
            {
                connectionType = DriveConnectionType.OneDriveConnection;
                userId = await GetUserId(userAccount);
                if (string.IsNullOrWhiteSpace(userId))
                    return false;

                drivesConnectionInfo?.Clear();
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorInitializingODConnection}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return false;
            }

            return true;
        }

        public async Task<bool> InitializeMsTeamsConnection(AuthenticationConfig authenticationConfig, string[]? teams, bool forceInit = false)
        {
            this.authConfig = authenticationConfig;
            this.teams = teams;

            if (!await IsInitialized(forceInit))
                return false;

            try
            {
                connectionType = DriveConnectionType.MSTeamsConnection;
                drivesConnectionInfo?.Clear();
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorInitializingTeamsConnection}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return false;
            }

            return true;
        }

        public async Task<bool> InitializeExchangeConnection(AuthenticationConfig authenticationConfig, bool forceInit = false)
        {
            authConfig = authenticationConfig;
            if (!await IsInitialized(forceInit))
                return false;

            connectionType = DriveConnectionType.ExchangeConnection;
            return true;
        }

        public async Task<string> GetTokenApplicationPermissions(string tenantId, string clientId, string clientSecret, string[] scopes)
        {
            try
            {
                var options = new ClientSecretCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                };

                // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
                var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);
                var authResult = await clientSecretCredential.GetTokenAsync(new TokenRequestContext(scopes));

                return authResult.Token;
            }
            catch(Exception ex)
            {
                logger.LogError($"{Resource.CannotGetJwtToken}. {ex.Message}");
                return null;
            }
        }

        public async Task<string> GetTokenDelegatedPermissions(string tenantId, string clientId, string[] scopes)
        {
            try
            {
                // https://learn.microsoft.com/dotnet/api/azure.identity.interactivebrowsercredential
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = new Uri($"https://login.microsoftonline.com/{tenantId}/v2.0"),
                    RedirectUri = new Uri("http://localhost")
                };

                ibc ??= new InteractiveBrowserCredential(options);

                var context = new TokenRequestContext(scopes);
                var authResult = await ibc.GetTokenAsync(context);

                return authResult.Token;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.CannotGetJwtToken}. {ex.Message}");
                return null;
            }
        }

        public async Task<string> GetUserId(string user)
        {
            try
            {
                var userObject = await graphServiceClient.Users[user].GetAsync();
                return userObject?.Id;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingUserId}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return null;
            }
        }

        public async Task<string> GetSiteId(string siteUrl)
        {
            try
            {
                var uri = new Uri(siteUrl);
                var siteId = await graphServiceClient.Sites[$"{uri.Host}:{uri.PathAndQuery}"].GetAsync(requestConfiguration => 
                                        {
                                            requestConfiguration.QueryParameters.Select = new string[] { "id" };
                                        });                    
                    

                return siteId.Id;
            }
            catch(Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingSiteId}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return null;
            }
        }

        public async Task<List<DriveConnectionInfo>> GetSharePointOnlineDrives()
        {
            drivesConnectionInfo = new List<DriveConnectionInfo>();
            if (docLibraries != null)
                // Removing leading and trailing spaces
                docLibraries = docLibraries.Select(x => x.Trim()).ToArray();

            try
            {
                // Here 'docLibraries' could contain the Doc Libraries to process or null if we want to process all Doc Libraries on the site
                var siteDrives = await graphServiceClient.Sites[siteId].Drives.GetAsync();
                var selectedDrives = siteDrives.Value.Where(x => docLibraries == null || docLibraries.Contains(x.Name));

                if (docLibraries == null)
                    selectedDrives = siteDrives.Value;
                else
                {
                    foreach (var library in docLibraries)
                    {
                        // Show error if provided Document Library name doesn't exist
                        if (siteDrives.Value.FirstOrDefault(x => x.Name == library) == null)
                            logger.LogError($"{Resource.LibraryNotFound}: {library}");
                    }
                }

                if (!selectedDrives.Any())
                    // Show error if provided Document Libraries names don't exist
                    logger.LogError($"{Resource.LibrariesNotFound}: {string.Join(",", docLibraries)}");

                foreach (var drive in selectedDrives)
                {
                    var driveInfo = new DriveConnectionInfo
                    {
                        Id = drive.Id,
                        Root = siteUrl.GetFolderNameFromSpoUrl(),
                        Path = drive.WebUrl,
                        Name = drive.Name,
                        DownloadCompleted = false
                    };

                    drivesConnectionInfo.Add(driveInfo);
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingDocLibraries}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
            }

            return drivesConnectionInfo;
        }

        public async Task<List<DriveConnectionInfo>> GetOneDriveDrives()
        {
            drivesConnectionInfo = new List<DriveConnectionInfo>();

            try
            {
                var userDrives = await graphServiceClient.Users[userId].Drives.GetAsync();
                foreach (var drive in userDrives.Value)
                {
                    var driveInfo = new DriveConnectionInfo
                    {
                        Id = drive.Id,
                        Root = oneDriveUserAccount,
                        Path = drive.WebUrl,
                        Name = drive.Name,
                        DownloadCompleted = false
                    };

                    drivesConnectionInfo.Add(driveInfo);
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingDrives}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
            }

            return drivesConnectionInfo;
        }

        public async Task<List<DriveConnectionInfo>> GetMsTeamDrives()
        {
            drivesConnectionInfo = new List<DriveConnectionInfo>();
            if (teams != null)
                // Removing leading and trailing spaces
                teams = teams.Select(x => x.Trim()).ToArray();

            try
            {
                // Here 'teams' could contain a list of MsTeams to process or null if we want to process all MsTeams on the organization
                GroupCollectionResponse groups;
                if (teams == null)
                    groups = await graphServiceClient.Groups.GetAsync(requestConfiguration =>
                                        {
                                            requestConfiguration.QueryParameters.Filter = "resourceProvisioningOptions/Any(x:x eq 'Team')";
                                        });
                else
                {
                    var filter = string.Join(" or ", teams.Select(x => $"displayName eq '{x.Trim()}'"));
                    filter = $"({filter}) and resourceProvisioningOptions / Any(x: x eq 'Team')";
                    groups = await graphServiceClient.Groups.GetAsync(requestConfiguration =>
                                        {
                                            requestConfiguration.QueryParameters.Filter = filter;
                                        });

                    if (groups.Value.Count == 0)
                    {
                        // If no teams found log error
                        logger.LogError($"{Resource.ErrorRetrievingTeams}: {string.Join(",", teams)}");
                    } 
                    else if (groups.Value.Count < teams.Count())
                    {
                        // If any of the teams was not found log error
                        var foundTeams = groups.Value.Select(x => x.DisplayName).ToList();
                        var notFoundTeams = teams.Where(x => !foundTeams.Contains(x));
                        logger.LogError($"{Resource.ErrorRetrievingTeams}: {string.Join(",", notFoundTeams)}");
                    }
                }

                foreach (var group in groups.Value)
                {
                    Drive drive;
                    try
                    {
                        drive = await graphServiceClient.Groups[group.Id].Drive.GetAsync();
                    }
                    catch
                    {
                        logger.LogWarning($"{Resource.CannotAccessDrive}: {group.DisplayName}");
                        continue;
                    }

                    var driveInfo = new DriveConnectionInfo
                    {
                        Id = drive.Id,
                        Root = group.DisplayName,
                        Path = drive.WebUrl,
                        Name = drive.Name,
                        GroupId = group.Id,
                        DownloadCompleted = false
                    };

                    drivesConnectionInfo.Add(driveInfo);
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingTeams}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
            }

            return drivesConnectionInfo;
        }

        public async Task<List<string>> GetDocumentIds(string driveId, DateTime lastDate, int skip, int top)
        {
            var documentIds = new List<string>();

            try
            {
                var path = drivesConnectionInfo.FirstOrDefault(x => x.Id == driveId)?.Path;

                var requestBody = new QueryPostRequestBody
                {
                    Requests = new List<SearchRequest>
                    {
                        new SearchRequest
                        {
                            EntityTypes = new List<EntityType?> { EntityType.DriveItem },
                            Query = new SearchQuery { QueryString = $"LastModifiedtime>{lastDate} AND Path:{path} AND IsDocument:Yes" },
                            From = skip,
                            Size = top,
                            SortProperties = new List<SortProperty>()
                            {
                                new SortProperty
                                {
                                    Name = "LastModifiedDateTime",
                                    IsDescending = false
                                }
                            }
                        }
                    }
                };


                QueryResponse searchResults = await graphServiceClient.Search.Query.PostAsync(requestBody);
                foreach (var searchResult in searchResults.Value.First().HitsContainers)
                {
                    if (searchResult.Total == 0)
                        break;

                    foreach (var hit in searchResult.Hits)
                        documentIds.Add(hit.HitId);
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingDocumentIds}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
            }

            return documentIds;
        }

        public async Task<DriveItemSearchResult> GetDocumentIds(string driveId, DateTime lastDate, int top, string skipToken)
        {
            var retryCount = ConstantHelper.DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    var drive = await GetDriveRoot(driveId);

                    var rootDriveItem = await graphServiceClient.Drives[drive.Id].Root.GetAsync();

                    BaseDeltaFunctionResponse deltaResponse;

                    if (string.IsNullOrWhiteSpace(skipToken))
                    {
                        deltaResponse = await graphServiceClient.Drives[drive.Id]
                               .Items[rootDriveItem.Id]
                               .Delta
                               .GetAsync(requestConfiguration =>
                               {
                                   requestConfiguration.QueryParameters.Top = top;
                                   requestConfiguration.QueryParameters.Orderby = new string[] { "lastModifiedDateTime" };
                               });

                    }
                    else
                    {
                        deltaResponse = await graphServiceClient.Drives[drive.Id]
                               .Items[rootDriveItem.Id]
                               .DeltaWithToken(skipToken)
                               .GetAsync(requestConfiguration =>
                               {
                                   requestConfiguration.QueryParameters.Top = top;
                                   requestConfiguration.QueryParameters.Orderby = new string[] { "lastModifiedDateTime" };
                               });
                    }

                    var deltaLink = deltaResponse.OdataNextLink ?? deltaResponse.OdataDeltaLink;

                    if (deltaLink != null)
                    {
                        var tokenString = Regex.Match(deltaLink, @"token='?[a-zA-Z0-9_.-]*'?");
                        var newToken = tokenString.Value.Replace("token=", "").Replace("'", "");
                        skipToken = newToken ?? skipToken;
                    }

                    var deltaResults = deltaResponse.BackingStore?.Get<List<Microsoft.Graph.Models.DriveItem>?>("value");

                    var driveItemResult = new DriveItemSearchResult
                    {
                        DocumentIds = new List<DriveItem>(),
                        HasMoreResults = deltaResponse.OdataNextLink != null,
                        SkipToken = skipToken,
                        LastDate = deltaResults.LastOrDefault()?.LastModifiedDateTime?.DateTime ?? lastDate
                    };

                    if (deltaResponse == null)
                        return driveItemResult;

                    foreach (var searchResult in deltaResults)
                    {
                        if (searchResult.Folder != null || searchResult.Deleted != null || searchResult.File == null)
                            continue;

                        driveItemResult.DocumentIds.Add(searchResult);
                    }

                    return driveItemResult;
                }
                catch (ODataError ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorRetrievingDocumentIds}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, ConstantHelper.DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }

            return null;
        }

        public async Task<DriveItem> GetDriveItem(string driveId, string documentId)
        {
            var retryCount = ConstantHelper.DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    var driveRoot = await GetDriveRoot(driveId);

                    return await graphServiceClient.Drives[driveRoot.Id]
                        .Items[documentId]
                        .GetAsync();
                }
                catch (ODataError ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorRetrievingDriveItem}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, ConstantHelper.DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }
            return null;
        }

        public async Task<DriveItem> SaveDriveItemToFileSystem(string driveId, string documentId, string downloadLocation)
        {
            var retryCount = ConstantHelper.DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    var document = await GetDriveItem(driveId, documentId);
                    if (document == null || document.File == null)
                        return null;

                    var drive = await GetSharePointDriveConnectionInfo(driveId);
                    var documentPath = document.ParentReference.Path.Replace($"/drives/{driveId}/root:", string.Empty).TrimStart('/').Replace(@"/", @"\");

                    string path = Path.Combine(downloadLocation, drive.Root, drive.Name, documentPath, document.Name);
                    System.IO.Directory.CreateDirectory(Path.GetDirectoryName(path));

                    document.AdditionalData.TryGetValue("@microsoft.graph.downloadUrl", out var downloadUrl);
                    var documentSize = document.Size;
                    var readSize = ConstantHelper.DEFAULT_CHUNK_SIZE;

                    using (FileStream outputFileStream = new FileStream(path.GetFilePathWithTimestamp(), FileMode.Create))
                    {
                        long offset = 0;
                        while (offset < documentSize)
                        {
                            var chunkSize = documentSize - offset > readSize ? readSize : documentSize - offset;
                            var req = new HttpRequestMessage(HttpMethod.Get, downloadUrl.ToString());
                            req.Headers.Range = new RangeHeaderValue(offset, chunkSize + offset - 1);

                            try
                            {
                                var graphClientResponse = await graphHttpClient.SendAsync(req);
                                if (graphClientResponse.StatusCode != System.Net.HttpStatusCode.OK && graphClientResponse.StatusCode != System.Net.HttpStatusCode.PartialContent)
                                    throw new TaskCanceledException();

                                using (var rs = await graphClientResponse.Content.ReadAsStreamAsync())
                                    rs.CopyTo(outputFileStream);

                                offset += readSize;
                            }
                            catch (TaskCanceledException)
                            {
                                // We got a timeout, try with a smaller chunk size
                                readSize /= 2;
                                offset = 0;
                                outputFileStream.Seek(offset, SeekOrigin.Begin);
                            }
                        }
                        document.AdditionalData.Add("FilePath", outputFileStream.Name);
                    }

                    return document;
                }
                catch (ODataError ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorSavingDriveItem}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, ConstantHelper.DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }

            return null;
        }

        /// <summary>
        /// Uploads a file from the file system to a cloud drive
        /// https://learn.microsoft.com/en-us/graph/sdks/large-file-upload
        /// </summary>
        /// <param name="driveId">Location(drive) where the document is to be uploaded(SPO, OneDrive, Teams) e.g: b!8iWW4uSCgUivMIG9AG1qEeKEpuugBHBKluSqT2GoUxM_0VutFV5zQIqEiYUABvpu</param>
        /// <param name="filePath">FileName and Path where the document is stored in the local file system. e.g: c:\forlder\file.txt</param>
        /// <param name="fileName">Optional parameter to change the name of the uploaded file</param>
        /// <param name="folderPath">Optional location in the document library where the file will be uploaded. e.g: folder1/folder2</param>
        /// <param name="onConflict">Optional conflict resolution behaviour. Default: replace</param>
        /// <returns></returns>
        public async Task<bool> UploadDocumentToDrive(string driveId, string filePath, string fileName = null, string folderPath = "", string onConflict = "replace")
        {
            if (fileName == null)
            {
                fileName = Path.GetFileName(filePath);
            }

            // Use properties to specify the conflict behavior. Posible values for onConflict = "fail (default) | replace | rename"
            var uploadSessionRequestBody = new DriveUpload.CreateUploadSessionPostRequestBody
            {
                Item = new DriveItemUploadableProperties
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "@microsoft.graph.conflictBehavior", onConflict },
                    },
                },
            };

            try
            {
                using var fileStream = File.OpenRead(filePath);

                // Create the upload session
                var myDrive = await graphServiceClient.Drives[driveId].GetAsync();
                var uploadSession = await graphServiceClient.Drives[driveId]
                    .Items["root"]
                    .ItemWithPath($"{folderPath}/{fileName}")
                    .CreateUploadSession
                    .PostAsync(uploadSessionRequestBody);

                var fileUploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, ConstantHelper.DEFAULT_CHUNK_SIZE, graphServiceClient.RequestAdapter);

                var totalLength = fileStream.Length;
                // Create a callback that is invoked after each slice is uploaded
                IProgress<long> progress = new Progress<long>(prog =>
                {
                    logger.LogInformation(string.Format(Resource.DriveItemUploadProgress, prog, totalLength));
                });

                // Upload the file
                var uploadResult = await fileUploadTask.UploadAsync(progress);

                if (uploadResult.UploadSucceeded)
                {
                    logger.LogInformation($"{Resource.DriveItemUploadComplete}: {uploadResult.ItemResponse.Id}");
                }
                else
                {
                    logger.LogError(Resource.DriveItemUploadFailed);
                    return false;
                }

                return true;
            }
            catch (ODataError ex)
            {
                logger.LogError($"{Resource.DriveItemUploadFailed}: {ex.Error?.Message}");
                return false;
            }

        }

        /// <summary>
        /// Delete a document from a drive using its Id
        /// https://learn.microsoft.com/en-us/graph/api/driveitem-delete?view=graph-rest-1.0
        /// </summary>
        /// <param name="driveId">Location(drive) where the document is located</param>
        /// <param name="documentId">Id of the document</param>
        /// <returns></returns>
        public async Task<bool> DeleteDocumentById(string driveId, string documentId)
        {
            try
            {
                await graphServiceClient.Drives[driveId].Items[documentId].DeleteAsync();
                return true;
            }
            catch (ODataError ex)
            {
                logger.LogError($"{Resource.ErrorDeletingDriveItem}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return false;
            }
        }

        /// <summary>
        /// Delete a document froma drive using its relative path
        /// https://learn.microsoft.com/en-us/graph/api/driveitem-delete?view=graph-rest-1.0
        /// </summary>
        /// <param name="driveId">Location(drive) where the document is located</param>
        /// <param name="documentPath">Path to the document in the drive</param>
        /// <returns></returns>
        public async Task<bool> DeleteDocumentByPath(string driveId, string documentPath)
        {
            try
            {
                var driveItem = await graphServiceClient.Drives[driveId]
                                                        .Items["root"]
                                                        .ItemWithPath(documentPath)
                                                        .GetAsync();

                return await DeleteDocumentById(driveId, driveItem.Id);
            }
            catch (ODataError ex)
            {
                logger.LogError($"{Resource.ErrorDeletingDriveItem}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return false;
            }
        }

        /// <summary>
        /// Move a document to a different location within the same drive. Graph Api SDK does not allow moving to a different drive
        /// https://learn.microsoft.com/en-us/graph/api/driveitem-move?view=graph-rest-1.0
        /// </summary>
        /// <param name="driveId">Location(drive) where the document is located</param>
        /// <param name="documentId">Id of the document</param>
        /// <param name="destFolderId">Folder id where the document will be moved (Optional, will take precedence over destFolder)</param>
        /// <param name="destFolder">Path of the folder where the document will be moved (Optional)</param>
        /// <param name="docNewName">New name of the document when moved (Optional)</param>
        /// <returns></returns>
        public async Task<bool> MoveDocument(string driveId, string documentId, string destFolderId = null, string destFolder = null, string docNewName = null)
        {
            try
            {
                if (destFolderId == null)
                {
                    var folder = await graphServiceClient.Drives[driveId].Items["root"].ItemWithPath(destFolder ?? "//").GetAsync();

                    if (folder == null)
                    {
                        logger.LogError(Resource.DestinationFolderNotFound);
                        return false;
                    }

                    destFolderId = folder.Id;
                }

                var requestBody = new DriveItem
                {
                    ParentReference = new ItemReference
                    {
                        Id = destFolderId
                    }
                };

                if (docNewName != null)
                    requestBody.Name = docNewName;

                var result = await graphServiceClient.Drives[driveId].Items[documentId].PatchAsync(requestBody);
                return true;
            }
            catch (ODataError ex)
            {
                logger.LogError($"{Resource.ErrorMovingDriveItem}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return false;
            }
        }

        /// <summary>
        /// Get a list of Sharepoint sites in a tenant along with a list of document libraries
        /// </summary>
        /// <param name="excludePersonalSites">If true method will not return Sharepoint Online personal sites</param>
        /// <returns>A dictionary containing site Urls as the Key and a list of its respectives DocumentLibraries as the Value</returns>
        public async Task<Dictionary<string, List<string>>> GetSharepointSitesAndDocLibs(bool excludePersonalSites = false, bool excludeSystemDocLibs = false)
        {
            try
            {
                var requestInformation = graphServiceClient.Sites.ToGetRequestInformation(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new[] { "Id", "WebUrl", "Name", "DisplayName" };
                    requestConfiguration.QueryParameters.Top = 200;
                });
                requestInformation.UrlTemplate = requestInformation.UrlTemplate.Replace("%24search", "search");
                requestInformation.QueryParameters.Add("search", "*");
                var sitesReponse = await graphServiceClient.RequestAdapter.SendAsync<SiteCollectionResponse>(requestInformation, SiteCollectionResponse.CreateFromDiscriminatorValue);
                var sites = new List<Site>();
                var sitesIterator = PageIterator<Site, SiteCollectionResponse>.CreatePageIterator(graphServiceClient, sitesReponse, (site) => { sites.Add(site); return true; });
                await sitesIterator.IterateAsync();

                // If the Url provided is not a tenant sharepoint root Url we will exclude personal sites anyway
                var isTenantRoot = siteUrl.IsRootUrl();
                excludePersonalSites |= !isTenantRoot;

                if (!excludePersonalSites)
                {
                    var sitesResponse = await graphServiceClient.Sites.GetAllSites.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter = "IsPersonalSite eq true";
                    });

                    if (sitesResponse != null && sitesResponse.Value.Any())
                        sites.AddRange(sitesResponse.Value);

                    var nextLink = sitesResponse.OdataNextLink;
                    while (nextLink != null)
                    {
                        var nextSitesResponse = await graphServiceClient.RequestAdapter.SendAsync(new RequestInformation { UrlTemplate = nextLink }, (parseNode) => new SiteCollectionResponse());
                        if (nextSitesResponse != null && nextSitesResponse.Value.Any())
                        {
                            sites.AddRange(nextSitesResponse.Value);
                            nextLink = nextSitesResponse.OdataNextLink;
                        }
                    }
                }

                // If it is not the tenant root url filter out sites outside current site url
                if (!isTenantRoot)
                {
                    sites = sites.Where(x => x.WebUrl.StartsWith(siteUrl, StringComparison.InvariantCultureIgnoreCase)).ToList();
                }

                return await GetSiteAndDocLibsDictionary(sites.ToList(), excludeSystemDocLibs);

            }
            catch (ODataError ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingSpoSites}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return null;
            }
        }

        /// <summary>
        /// Search folders and document sets in a Sahrepoint Document libray that match a specific criteria.
        /// </summary>
        /// <param name="siteUrl">Url of the Sharepoint site. e.g: https://coginovportal.sharepoint.com/sites/Dev-JC</param>
        /// <param name="docLibrary">Display name of the document library. e.g: "Documenst Test"</param>
        /// <param name="searchField">Field to search for a specific value. Optional if searchFilter is specified</param>
        /// <param name="searchValue">Value to match in the specified field. Optional if searchFilter is specified</param>
        /// <param name="searchFilter">Optional search condition. Superseeds searchField/searchValue combination </param>
        /// <returns>List of field values for found folders</returns>
        public async Task<List<ListItem>> SearchSharepointOnlineFolders(string siteUrl, string docLibrary, string searchField = null, string searchValue = null, string searchFilter = null, int top = 200)
        {
            if (string.IsNullOrWhiteSpace(searchFilter))
            {
                if (string.IsNullOrWhiteSpace(searchField) || string.IsNullOrWhiteSpace(searchValue))
                {
                    logger.LogError(Resource.InvalidSearchParameters);
                    return null;
                }
            }

            try
            {
                var siteId = await GetSiteId(siteUrl);
                if (string.IsNullOrWhiteSpace(siteId))
                {
                    return null;
                }

                searchFilter ??= $"fields/{searchField} eq '{searchValue}'";

                var folders = await graphServiceClient.Sites[siteId].Lists[Uri.EscapeDataString(docLibrary)].Items.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Expand = new string[] { "fields", "driveItem" };
                    requestConfiguration.QueryParameters.Filter = $"(fields/ContentType eq 'Document Set' or fields/ContentType eq 'Folder') and ({searchFilter})";
                    requestConfiguration.QueryParameters.Select = new string[] { "sharepointIds" };
                    requestConfiguration.QueryParameters.Top = top;
                    requestConfiguration.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
                });

                folders.Value.ForEach(x =>
                {
                    x.Fields.AdditionalData["DriveId"] = x.GetDriveId();
                    x.Fields.AdditionalData["DriveItemId"] = x.GetDriveItemId();
                    x.Fields.AdditionalData["CreatedByName"] = x.GetCreatedByName();
                    x.Fields.AdditionalData["CreatedByEmail"] = x.GetCreatedByEmail();
                    x.Fields.AdditionalData["ModifiedByName"] = x.GetModifiedByName();
                    x.Fields.AdditionalData["ModifiedByEmail"] = x.GetModifiedByEmail();
                });

                return folders.Value;
            }
            catch (ODataError ex)
            {
                logger.LogError($"{Resource.ErrorSearchingFolders}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return null;
            }
        }

        /// <summary>
        /// Update one or more columns in a list of documents with new values
        /// </summary>
        /// <param name="items">List of items to be updated</param>
        /// <param name="columnKeyValues">Dictionary containing list of columns and values to be updated</param>
        /// <returns>A dictionary containing a folder as the key and an optional update error message as the value</returns>
        public async Task<Dictionary<ListItem, string>> UpdateSharePointOnlineItemFieldValue(List<ListItem> items, Dictionary<string, object> columnKeyValues)
        {
            if (columnKeyValues.Any(x => string.IsNullOrEmpty(x.Key)))
            {
                logger.LogError(Resource.InvalidUpdateParameters);
                return null;
            }

            try
            {
                var result = new Dictionary<ListItem, string>();
                var columnKeys = columnKeyValues.Keys;

                foreach (var item in items)
                {
                    var requestBody = new FieldValueSet
                    {
                        AdditionalData = columnKeyValues
                    };

                    try
                    {
                        var listItemResult = await graphServiceClient.Sites[siteId].Lists[item.SharepointIds.ListId].Items[item.SharepointIds.ListItemId].Fields.PatchAsync(requestBody);
                        result.Add(item, null);
                    }
                    catch (Exception ex)
                    {
                        result.Add(item, ex.Message);
                    }
                }

                return result;
            }
            catch (ODataError ex)
            {
                logger.LogError($"{Resource.ErrorUpdatingSharepointItems}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return null;
            }
        }

        /// <summary>
        /// Update one or more columns in a list of documents with new values
        /// </summary>
        /// <param name="items">List of items to be updated</param>
        /// <param name="columnKeyValues">Dictionary containing list of columns and values to be updated</param>
        /// <returns>A dictionary containing a folder as the key and an optional update error message as the value</returns>
        public async Task<Dictionary<ListItem, string>> UpdateSharePointOnlineItemFieldValue(List<DriveItemInfo> items, Dictionary<string, object> columnKeyValues)
        {
            if (columnKeyValues.Any(x => string.IsNullOrEmpty(x.Key)))
            {
                logger.LogError(Resource.ErrorUpdatingSharepointItems);
                return null;
            }

            try
            {
                var listItems = new List<ListItem>();

                foreach (var item in items)
                {
                    var listItemResult = await graphServiceClient.Drives[item.DriveId].Items[item.DriveItemId].ListItem.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Expand = new string[] { "fields", "driveItem" };
                        requestConfiguration.QueryParameters.Select = new string[] { "sharepointIds" };
                        requestConfiguration.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
                    });

                    listItems.Add(listItemResult);
                }

                return await UpdateSharePointOnlineItemFieldValue(listItems, columnKeyValues);
            }
            catch (ODataError ex)
            {
                logger.LogError($"{Resource.ErrorUpdatingSharepointItems}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return null;
            }
        }

        /// <summary>
        /// Get the list of files in a folder. Will return all files, retrieves items on batches of 'batchSize'
        /// </summary>
        /// <param name="driveItem">Object representing the folder that contains the files</param>
        /// <param name="batchSize">Number of files to download in each operation</param>
        /// <returns>List of driveitems representing the files in the folder</returns>
        public async Task<List<DriveItem>> GetListOfFilesInFolder(DriveItemInfo driveItem, DateTimeOffset? lastDate = null, int batchSize = 100)
        {
            if (driveItem == null)
            {
                logger.LogError("Invalid driveItem info");
                return null;
            }

            lastDate ??= DateTime.MinValue;

            try
            {
                var driveItemResult = await graphServiceClient.Drives[driveItem.DriveId].Items[driveItem.DriveItemId].Children.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Top = batchSize;
                    requestConfiguration.QueryParameters.Orderby = new string[] { "lastModifiedDateTime" };
                    requestConfiguration.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
                });

                var driveItemList = new List<DriveItem>();
                var pageIterator = PageIterator<DriveItem, DriveItemCollectionResponse>.CreatePageIterator(graphServiceClient, driveItemResult, (item) => 
                { 
                    if (item.Folder == null && item.LastModifiedDateTime > lastDate)
                        driveItemList.Add(item);
                    return true; 
                });

                await pageIterator.IterateAsync();
                
                return driveItemList;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingDocuments}: {ex.Message}");
                return null;
            }
        }

        #region Exchange Online methods

        public async Task<bool> SaveEmailToFileSystem(Message message, string downloadLocation, string userAccount, string fileName)
        {
            var retryCount = ConstantHelper.DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    string path = Path.Combine(downloadLocation, userAccount, fileName);
                    System.IO.Directory.CreateDirectory(Path.GetDirectoryName(path));

                    var email = await graphServiceClient.Users[userAccount].Messages[message.Id].Content.GetAsync();

                    using (FileStream outputFileStream = new FileStream(path, FileMode.Create))
                        email.CopyTo(outputFileStream);

                    return true;
                }
                catch (TaskCanceledException)
                {
                    // We got a timeout, ignore for now
                    logger.LogInformation($"{Resource.ErrorSavingExchangeMessage} File too big. Go to next");
                }
                catch (ODataError ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorSavingExchangeMessage}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, ConstantHelper.DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }

            return false;
        }

        public async Task<int?> GetInboxMessageCount(string userAccount)
        {
            try
            {
                return await graphServiceClient.Users[userAccount].Messages.Count.GetAsync();

            }
            catch (ODataError ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingExchangeMessagesCount}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
            }
            return null;
        }

        public async Task<MessageCollectionResponse> GetEmailsAfterDate(string userAccount, DateTime afterDate, int skipIndex = 0, int emailCount = 10, bool includeAttachments = false)
        {
            var retryCount = ConstantHelper.DEFAULT_RETRY_COUNT;
            var filter = $"createdDateTime gt {afterDate.ToString("s")}Z";

            while (retryCount-- > 0)
            {
                try
                {
                    return await graphServiceClient.Users[userAccount].Messages
                                                .GetAsync(requestConfiguration =>
                                                {
                                                    requestConfiguration.QueryParameters.Filter = filter;
                                                    requestConfiguration.QueryParameters.Skip = skipIndex;
                                                    requestConfiguration.QueryParameters.Top = emailCount;
                                                    requestConfiguration.QueryParameters.Orderby = new string[] { "createdDateTime" };
                                                    requestConfiguration.QueryParameters.Expand = new string[] { includeAttachments ? "attachments" : "" };
                                                });
                }
                catch (ODataError ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorRetrievingExchangeMessages}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, ConstantHelper.DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }

            return null;
        }

        public async Task<MessageCollectionResponse> GetEmailsFromFolderAfterDate(string userAccount, string folder, DateTime afterDate, int skipIndex = 0, int emailCount = 10, bool includeAttachments = false, bool preferText = false)
        {
            var retryCount = ConstantHelper.DEFAULT_RETRY_COUNT;
            var filter = $"createdDateTime gt {afterDate.ToString("s")}Z";

            while (retryCount-- > 0)
            {
                try
                {
                    var graphRequest = graphServiceClient.Users[userAccount].MailFolders[folder].Messages;

                    return await  graphRequest.GetAsync(requestConfiguration =>
                                        {
                                            if (preferText)
                                                requestConfiguration.Headers.Add("Prefer", "outlook.body-content-type=\"text\"");

                                            requestConfiguration.QueryParameters.Filter = filter;
                                            requestConfiguration.QueryParameters.Skip = skipIndex;
                                            requestConfiguration.QueryParameters.Top = emailCount;
                                            requestConfiguration.QueryParameters.Orderby = new string[] { "createdDateTime" };
                                            requestConfiguration.QueryParameters.Expand = new string[] { includeAttachments ? "attachments" : "" };
                                        });
                }
                catch (ODataError ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorRetrievingExchangeMessages}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, ConstantHelper.DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }

            return null;
        }

        public async Task<List<MailFolder>> GetEmailFolders(string userAccount)
        {
            try
            {
                var folderResult = await graphServiceClient.Users[userAccount].MailFolders.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "totalItemCount" };
                });

                var folderList = new List<MailFolder>();
                var pageIterator = PageIterator<MailFolder, MailFolderCollectionResponse>.CreatePageIterator(graphServiceClient, folderResult, (item) =>
                {
                    folderList.Add(item);
                    return true;
                });

                await pageIterator.IterateAsync();

                return folderList;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingDocuments}: {ex.Message}");
                return null;
            }
        }

        public async Task<MailFolder> GetEmailFolderById(string userAccount, string folderId)
        {
            var retryCount = ConstantHelper.DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    return await graphServiceClient.Users[userAccount].MailFolders[folderId].GetAsync();
                }
                catch (ODataError ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorRetrievingExchangeFolders}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, ConstantHelper.DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }

            return null;
        }

        public async Task<bool> ForwardEmail(string userAccount, string emailId, string forwardAccount)
        {
            var retryCount = ConstantHelper.DEFAULT_RETRY_COUNT;

            var requestBody = new Microsoft.Graph.Users.Item.Messages.Item.Forward.ForwardPostRequestBody
            {
                ToRecipients = new List<Recipient>{
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Name = forwardAccount,
                            Address = forwardAccount
                        }
                    }
                }
            };

            while (retryCount-- > 0)
            {
                try
                {
                    await graphServiceClient.Users[userAccount].Messages[emailId]
                            .Forward
                            .PostAsync(requestBody);

                    return true;
                }
                catch (ODataError ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorForwardingEmail}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, ConstantHelper.DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }

            return false;
        }

        public async Task<bool> SendEmail(string fromAccount, string toAccounts, string subject, string body, List<Attachment> attachments = null)
        {
            var attachmentsCollection = new List<Attachment>();
            attachments?.ForEach(x => attachmentsCollection.Add(x));

            var sendMailBody = new SendMailPostRequestBody
            {
                Message = new Message
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = body
                    },
                    ToRecipients = toAccounts.ParseRecipients(),
                    Attachments = attachmentsCollection
                }
            };

            var retryCount = ConstantHelper.DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    await graphServiceClient.Users[fromAccount]
                            .SendMail
                            .PostAsync(sendMailBody);

                    return true;
                }
                catch (ODataError ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorSendingEmail}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, ConstantHelper.DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }

            return false;
        }

        public async Task<bool> MoveEmailToFolder(string userAccount, string emailId, string newFolder)
        {
            var retryCount = ConstantHelper.DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    var folder = (await GetEmailFolders(userAccount)).FirstOrDefault(x => x.DisplayName.Equals(newFolder, StringComparison.InvariantCultureIgnoreCase));
                    if (folder == null)
                        return false;

                    await graphServiceClient.Users[userAccount].Messages[emailId]
                            .Move
                            .PostAsync(new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                            {
                                DestinationId = folder.Id
                            });

                    return true;
                }
                catch (ODataError ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorMovingEmail}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, ConstantHelper.DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }

            return false;
        }

        /// <summary>
        /// Delete email from user account
        /// https://learn.microsoft.com/en-us/graph/api/message-delete
        /// </summary>
        /// <param name="userAccount">Account(email address) containing the email to be deleted</param>
        /// <param name="emailId">Id of the email to be deleted</param>
        /// <returns></returns>
        public async Task<bool> RemoveEmail(string userAccount, string emailId)
        {
            try
            {
                await graphServiceClient.Users[userAccount].Messages[emailId].DeleteAsync();
                return true;
            }
            catch (ODataError ex)
            {
                logger.LogError($"{Resource.ErrorRemovingEmail}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return false;
            }
        }

        #endregion

        #region Private Methods

        private async Task<Dictionary<string,List<string>>> GetSiteAndDocLibsDictionary(List<Site> sites, bool excludeSystemDocLibs = false)
        {
            if (sites == null || !sites.Any()) { return null; }

            try
            {
                var batchSize = 20;
                var index = 0;
                var siteDocsDictionary = new Dictionary<string, List<string>>();

                var batch = sites.Skip(index * batchSize).Take(batchSize).ToList();
                while (batch.Any())
                {
                    var batchRequestContent = new BatchRequestContentCollection(graphServiceClient);
                    var requestList = new List<RequestInformation>();
                    var requestIdDictionary = new Dictionary<Site, string>();

                    foreach (var item in batch)
                    {
                        var request = graphServiceClient.Sites[item.Id].Drives.ToGetRequestInformation(requestConfiguration =>
                        {
                            requestConfiguration.QueryParameters.Select = excludeSystemDocLibs ? Array.Empty<string>() : new string[] { "name", "system", "weburl" };
                        });
                        requestList.Add(request);
                        requestIdDictionary.Add(item, await batchRequestContent.AddBatchRequestStepAsync(request));
                    }

                    var drivesResponse = await graphServiceClient.Batch.PostAsync(batchRequestContent);

                    foreach (var item in requestIdDictionary)
                    {
                        if (siteDocsDictionary.ContainsKey(item.Key.WebUrl))
                            continue;

                        try
                        {
                            var drivesResult = await drivesResponse.GetResponseByIdAsync<DriveCollectionResponse>(item.Value);
                            var drives = drivesResult.Value.DistinctBy(x => x.Name).ToList();
                            siteDocsDictionary.Add(item.Key.WebUrl, drives.Select(x => x.Name).ToList());
                        }
                        catch(Exception ex)
                        {
                            logger.LogError($"{Resource.ErrorRetrievingDocLibraries}: {item.Key.Name}. {ex.Message}. {ex.InnerException?.Message}");
                            continue;
                        }
                    }

                    batch = sites.Skip(++index * batchSize).Take(batchSize).ToList();
                };

                return siteDocsDictionary;
            }
            catch(Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingDocLibraries}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return null;
            }
        }

        private async Task<Drive> GetDriveRoot(string driveId)
        {
            switch(connectionType)
            {
                case DriveConnectionType.OneDriveConnection:
                    return await graphServiceClient.Users[userId].Drives[driveId].GetAsync();
                case DriveConnectionType.SharePoinConnection:
                    return await graphServiceClient.Sites[siteId].Drives[driveId].GetAsync();
                case DriveConnectionType.MSTeamsConnection:
                    var groupId = drivesConnectionInfo.FirstOrDefault(s => s.Id == driveId)?.GroupId;
                    return await graphServiceClient.Groups[groupId].Drives[driveId].GetAsync();
                default: 
                    return null;
            }
        }

        private async Task<bool> InitializeAppPermissions()
        {
            string[] scopes = new string[] { $"{authConfig.ApiUrl}.default" };

            try
            {
                if (UseClientSecret())
                {
                    // using Azure.Identity;
                    var options = new ClientSecretCredentialOptions
                    {
                        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    };

                    // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
                    var clientSecretCredential = new ClientSecretCredential(authConfig.Tenant, authConfig.ClientId, authConfig.ClientSecret, options);
                    // Try to get a token to make sure the credentials provided work
                    await clientSecretCredential.GetTokenAsync(new TokenRequestContext(scopes));
                    graphServiceClient = new GraphServiceClient(clientSecretCredential, scopes);
                }
                else
                {
                    // using Azure.Identity;
                    var options = new ClientCertificateCredentialOptions
                    {
                        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    };
                    var clientCertificate = new X509Certificate2(authConfig.CertificateName);
                    // https://learn.microsoft.com/dotnet/api/azure.identity.clientcertificatecredential
                    var clientCertCredential = new ClientCertificateCredential(authConfig.Tenant, authConfig.ClientId, clientCertificate, options);
                    // Try to get a token to make sure the credentials provided work
                    await clientCertCredential.GetTokenAsync(new TokenRequestContext(scopes));
                    graphServiceClient = new GraphServiceClient(clientCertCredential, scopes);
                }
            }
            catch (ODataError ex) when (ex.Message.Contains("AADSTS70011"))
            {
                logger.LogError($"{Resource.ErrorInitializingGraph}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return await Task.FromResult(false);
            }

            return true;
        }

        private async Task<bool> InitializeDelegatedPermissions()
        {
            try
            {
                // The permission scopes required
                var graphScopes = new string[] { 
                    "https://graph.microsoft.com/Files.Read.All",
                    "https://graph.microsoft.com/Group.Read.All",
                    "https://graph.microsoft.com/Sites.Read.All",
                    "https://graph.microsoft.com/User.Read.All",
                    "https://graph.microsoft.com/Mail.Read.Shared"
                };

                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = authConfig.Tenant,
                    ClientId = authConfig.ClientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                // https://learn.microsoft.com/dotnet/api/azure.identity.interactivebrowsercredential
                var interactiveCredential = new InteractiveBrowserCredential(options);
                var context = new TokenRequestContext(graphScopes);
                await interactiveCredential.GetTokenAsync(context);

                graphServiceClient = new GraphServiceClient(interactiveCredential, graphScopes);
            }
            catch(Exception ex)
            {
                logger.LogError($"{Resource.ErrorInitializingGraph}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return await Task.FromResult(false);
            }

            return true;
        }

        private async Task<bool> InitializeUsingAccessToken()
        {
            authenticationToken = InitializeFromTokenPath();
            if (authenticationToken == null) 
                return false;

            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(authenticationConfig: authConfig, authenticationToken: authenticationToken, logger: logger));
                graphServiceClient = new GraphServiceClient(authenticationProvider);
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorInitializingGraph}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return await Task.FromResult(false);
            }

            return true;
        }

        private async Task<bool> IsInitialized(bool forceInit = false)
        {
            if (!forceInit && graphServiceClient != null)
                // TODO: Check if graphServiceClient is still connected
                // Even when we could already have an instance of the client the connection may had been lost
                return true;

            bool connected;
            switch (authConfig.AuthenticationMethod)
            {
                case AuthMethod.OAuthAppPermissions:
                    connected = await InitializeAppPermissions();
                    break;
                case AuthMethod.OAuthDelegatedPermissions:
                    connected = await InitializeDelegatedPermissions();
                    break;
                case AuthMethod.OAuthJwtAccessToken:
                    connected = await InitializeUsingAccessToken();
                    break;
                default:
                    return false;
            }

            InitGraphHttpClient();

            return connected;
        }

        private void InitGraphHttpClient()
        {
            if (graphHttpClient != null)
                return;

            var handlers = GraphClientFactory.CreateDefaultHandlers();

            if (useChaosHandler)
            {
                // Remove the default Retry Handler
                var retryHandler = handlers.Where(h => h is RetryHandler).FirstOrDefault();
                handlers.Remove(retryHandler);

                // Add the Chaos Handler
                handlers.Add(new ChaosHandler(new ChaosHandlerOption()
                {
                    ChaosPercentLevel = 50
                }));
            }

            graphHttpClient = GraphClientFactory.Create(handlers);
            graphHttpClient.Timeout = TimeSpan.FromHours(3);
        }

        private bool UseClientSecret()
        {
            if (!string.IsNullOrWhiteSpace(authConfig.ClientSecret))
                return true;
            else if (!string.IsNullOrWhiteSpace(authConfig.CertificateName))
                return false;
            else
                throw new Exception(Resource.ChooseClientOrCertificate);
        }

        private static X509Certificate2 ReadCertificate(string certificateName)
        {
            if (string.IsNullOrWhiteSpace(certificateName))
                throw new ArgumentException(Resource.CertificateEmpty, nameof(certificateName));

            var store = new X509Store(StoreName.My, StoreLocation.LocalMachine);
            store.Open(OpenFlags.ReadOnly);
            var certCollection = store.Certificates;
            var currentCerts = certCollection.Find(X509FindType.FindBySubjectDistinguishedName, certificateName, false);
            return currentCerts.Count == 0 ? null : currentCerts[0];
        }

        private int GetRetryAfterSeconds(ODataError ex)
        {
            IEnumerable<string> retries;
            switch (ex.ResponseStatusCode)
            {
                case (int)System.Net.HttpStatusCode.TooManyRequests:
                case (int)System.Net.HttpStatusCode.ServiceUnavailable:
                case (int)System.Net.HttpStatusCode.GatewayTimeout:
                    return ex.ResponseHeaders.ContainsKey("Retry-After") ? int.Parse(ex.ResponseHeaders["Retry-After"].First()) : ConstantHelper.DEFAULT_RETRY_IN_SECONDS;
                default:
                    return 1;
            }
        }

        private AuthenticationToken InitializeFromTokenPath()
        {
            if (!SystemFile.Exists(authConfig.TokenPath)) 
                return null;

            string tokenString;
            AuthenticationToken token;
            try
            {
                tokenString = SystemFile.ReadAllText(authConfig.TokenPath);
                if (string.IsNullOrWhiteSpace(tokenString)) 
                    return null;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.CannotReadTokenFile}: {ex.Message}. {ex.InnerException?.Message}");
                return null;
            }

            try
            {
                var uncryptedToken = AesHelper.DecryptToString(tokenString);
                return GetTokenFromString(uncryptedToken);
            }
            catch (Exception ex)
            {
                logger.LogWarning($"{Resource.TokenUnencrypted}: {ex.Message}. {ex.InnerException?.Message}");
            }

            try
            {
                token = GetTokenFromString(tokenString);
                if (token != null)
                    SystemFile.WriteAllText(authConfig.TokenPath, AesHelper.EncryptToString(tokenString));

                return token;
            }
            catch (Exception ex)
            {
                logger.LogWarning($"{Resource.CannotSaveToken}: {ex.Message}. {ex.InnerException?.Message}");
                return null;
            }
        }

        private static AuthenticationToken GetTokenFromString(string token)
        {
            try
            {
                return JsonConvert.DeserializeObject<AuthenticationToken>(token);
            }
            catch
            {
                //Invalid token format
                return null;
            }
        }

        private async Task<DriveConnectionInfo> GetSharePointDriveConnectionInfo(string driveId)
        {
            var driveInfo = drivesConnectionInfo?.FirstOrDefault(x => x.Id == driveId);
            if (driveInfo != null)
                return driveInfo;

            var drive = await graphServiceClient.Drives[driveId].GetAsync();
            
            driveInfo = new DriveConnectionInfo
            {
                Id = drive.Id,
                Root = siteUrl.GetFolderNameFromSpoUrl(),
                Path = drive.WebUrl,
                Name = drive.Name,
                DownloadCompleted = false
            };

            if (drivesConnectionInfo == null)
                drivesConnectionInfo = new List<DriveConnectionInfo>();

            drivesConnectionInfo.Add(driveInfo);

            return driveInfo;
        }

        #endregion
    }
}