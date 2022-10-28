using Azure.Identity;
using Coginov.GraphApi.Library.Enums;
using Coginov.GraphApi.Library.Helpers;
using Coginov.GraphApi.Library.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
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
using Newtonsoft.Json;
using SystemFile = System.IO.File;

namespace Coginov.GraphApi.Library.Services
{
    public class GraphApiService : IGraphApiService, IDisposable
    {
        private const long DEFAULT_CHUNK_SIZE = 1024 * 1024 * 1024; // 1 GB
        private const int DEFAULT_RETRY_IN_SECONDS = 1;
        private const int DEFAULT_RETRY_COUNT = 5;
        private readonly ILogger logger;
        private AuthenticationConfig authConfig;
        private GraphServiceClient graphServiceClient;
        private List<DriveConnectionInfo> drivesConnectionInfo;
        private DriveConnectionType connectionType;
        private string userId;
        private string siteId;
        private AuthenticationToken authenticationToken;
        private string msalCachePath;
        private string msalCacheFileName;
        private MsalCacheHelper msalCacheHelper;
        private IPublicClientApplication pca;
        private IConfidentialClientApplication cca; 

        public GraphApiService(ILogger logger)
        {
            this.logger = logger;
        }

        public void Dispose()
        {
            if (msalCacheHelper != null)
            {
                msalCacheHelper.UnregisterCache(pca.UserTokenCache);
                System.IO.File.Delete(Path.Combine(msalCachePath, msalCacheFileName));
            }
        }

        public async Task<bool> InitializeSharePointConnection(AuthenticationConfig authenticationConfig, string siteRoot, params string[] docLibraries)
        {
            authConfig = authenticationConfig;

            if (!await IsInitialized())
            {
                return false;
            }

            try
            {
                connectionType = DriveConnectionType.SharePoinConnection;
                siteId = await GetSiteId(siteRoot);
                if (siteId == null)
                {
                    return false;
                }

                await GetDrives(docLibraries, siteRoot.GetFolderNameFromSpoUrl());
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorInitializingSPConnection}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return false;
            }

            return true;
        }

        public async Task<bool> InitializeOneDriveConnection(AuthenticationConfig authenticationConfig, string user)
        {
            authConfig = authenticationConfig;

            if (!await IsInitialized())
            {
                return false;
            }

            try
            {
                connectionType = DriveConnectionType.OneDriveConnection;
                userId = await GetUserId(user);
                await GetDrives(root: user);
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorInitializingODConnection}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return false;
            }

            return true;
        }

        public async Task<bool> InitializeMsTeamsConnection(AuthenticationConfig authenticationConfig, params string[]? teams)
        {
            authConfig = authenticationConfig;

            if (!await IsInitialized())
            {
                return false;
            }

            try
            {
                connectionType = DriveConnectionType.MSTeamsConnection;
                await GetDrives(teams);
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorInitializingTeamsConnection}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return false;
            }

            return true;
        }

        public async Task<bool> InitializeExchangeConnection(AuthenticationConfig authenticationConfig)
        {
            authConfig = authenticationConfig;

            if (!await IsInitialized())
            {
                return false;
            }

            connectionType = DriveConnectionType.ExchangeConnection;
            return true;
        }


        public async Task<string> GetUserId(string user)
        {
            try
            {
                var userObject = await graphServiceClient.Users[user].Request().GetAsync();
                return userObject?.Id;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingUserId}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return null;
            }
        }

        public async Task<string> GetSiteId(string site)
        {
            try
            {
                var uri = new Uri(site);
                var subsite = uri.PathAndQuery.TrimStart('/').Replace("sites/", "");
                var siteId = string.IsNullOrEmpty(subsite)
                    ? await graphServiceClient.Sites[$"{uri.Host}:"].Request().Select("id").GetAsync()
                    : await graphServiceClient.Sites[$"{uri.Host}:"].Sites[subsite].Request().Select("id").GetAsync();

                return siteId?.Id?.Split(",")[1];
            }
            catch(Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingSiteId}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return null;
            }
        }

        public async Task<List<DriveConnectionInfo>> GetDrives(string[]? drives = null, string root = "")
        {
            if (drivesConnectionInfo != null)
            {
                return drivesConnectionInfo;
            }

            if (drives != null)
            {
                // Removing leading and trailing spaces
                drives = drives.Select(x => x.Trim()).ToArray();
            }

            drivesConnectionInfo = new List<DriveConnectionInfo>();

            try
            {
                switch (connectionType)
                {
                    case DriveConnectionType.OneDriveConnection:
                        var userDrives = await graphServiceClient.Users[userId].Drives.Request().GetAsync();
                        foreach (var drive in userDrives)
                        {
                            var driveInfo = new DriveConnectionInfo
                            {
                                Id = drive.Id,
                                Root = root,
                                Path = drive.WebUrl,
                                Name = drive.Name,
                                DownloadCompleted = false
                            };

                            drivesConnectionInfo.Add(driveInfo);
                        }

                        break;
                    case DriveConnectionType.SharePoinConnection:
                        // Here 'drives' could contain the Doc Libraries to process or null if we want to process all Doc Libraries on the site
                        var siteDrives = await graphServiceClient.Sites[siteId].Drives.Request().GetAsync();
                        if (drives != null)
                        {
                            foreach (var library in drives)
                            {
                                // Show error if provided Document Library name doesn't exist
                                if (siteDrives.FirstOrDefault(x => x.Name == library) == null)
                                {
                                    logger.LogError($"{Resource.LibraryNotFound}: {library}");
                                }
                            }
                        }

                        var selectedDrives = siteDrives.Where(x => drives == null || drives.Contains(x.Name));
                        if (drives == null || !selectedDrives.Any())
                        {
                            logger.LogWarning(Resource.NoLibraryFound);
                            selectedDrives = siteDrives;
                        }

                        foreach (var drive in selectedDrives)
                        {
                            var driveInfo = new DriveConnectionInfo
                            {
                                Id = drive.Id,
                                Root = root,
                                Path = drive.WebUrl,
                                Name = drive.Name,
                                DownloadCompleted = false
                            };

                            drivesConnectionInfo.Add(driveInfo);
                        }

                        break;
                    case DriveConnectionType.MSTeamsConnection:
                        // Here 'drives' could contain a list of MsTeams to process or null if we want to process all MsTeams on the organization
                        IGraphServiceGroupsCollectionPage groups;
                        if (drives == null)
                        {
                            groups = await graphServiceClient.Groups.Request().Filter("resourceProvisioningOptions/Any(x:x eq 'Team')").GetAsync();
                        } 
                        else
                        {
                            var filter = string.Join(" or ", drives.Select(x => $"displayName eq '{x.Trim()}'"));
                            filter = $"{(filter)} and (resourceProvisioningOptions / Any(x: x eq 'Team'))";
                            groups = await graphServiceClient.Groups.Request().Filter(filter).GetAsync();
                            if (groups.Count == 0)
                            {
                                //if no teams found default to all
                                groups = await graphServiceClient.Groups.Request().Filter("resourceProvisioningOptions/Any(x:x eq 'Team')").GetAsync();
                            }
                        }
                        foreach(var group in groups)
                        {
                            Drive drive;
                            try
                            {
                                drive = await graphServiceClient.Groups[group.Id].Drive.Request().GetAsync();
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

                        break;
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingDrives}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
            }

            return drivesConnectionInfo;
        }

        // This method will only work with Delegated Permissions authentication
        public async Task<List<string>> GetDocumentIds(string driveId, DateTime lastDate, int skip, int top)
        {
            var documentIds = new List<string>();

            try
            {
                var path = drivesConnectionInfo.FirstOrDefault(x => x.Id == driveId)?.Path;
                var searchRequest = new List<SearchRequestObject>
                {
                    new SearchRequestObject
                    {
                        EntityTypes = new List<EntityType>() { EntityType.DriveItem },
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
                };

                var searchResults = await graphServiceClient.Search.Query(searchRequest).Request().PostAsync();
                foreach (var searchResult in searchResults.First().HitsContainers)
                {
                    if (searchResult.Total == 0)
                    {
                        break;
                    }

                    foreach (var hit in searchResult.Hits)
                    {
                        documentIds.Add(hit.HitId);
                    }
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
            var retryCount = DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    var queryOptions = new List<QueryOption> { new QueryOption("$skiptoken", skipToken) };
                    var searchCollection = await GetDriveRoot(driveId)
                        .Root
                        .Search("")
                        .Request(queryOptions)
                        .OrderBy("lastModifiedDateTime")
                        .Top(top)
                        .GetAsync();

                    var driveItemResult = new DriveItemSearchResult
                    {
                        DocumentIds = new List<string>(),
                        HasMoreResults = searchCollection.NextPageRequest != null,
                        SkipToken = searchCollection.NextPageRequest != null
                            ? searchCollection.NextPageRequest.QueryOptions.FirstOrDefault(x => x.Name == "$skiptoken").Value
                            : skipToken,
                        LastDate = searchCollection?.LastOrDefault()?.LastModifiedDateTime?.DateTime ?? lastDate
                    };

                    if (searchCollection == null)
                    {
                        return driveItemResult;
                    }

                    foreach (var searchResult in searchCollection)
                    {
                        if (searchResult.Folder != null ||
                            searchResult.LastModifiedDateTime == null ||
                            searchResult.LastModifiedDateTime.Value.DateTime < lastDate)
                        {
                            continue;
                        }

                        driveItemResult.DocumentIds.Add(searchResult.Id);
                    }

                    return driveItemResult;
                }
                catch (ServiceException ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorRetrievingDocumentIds}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(String.Format(Resource.GraphRetryAttempts, retryInSeconds, DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }
            return null;
        }

        public async Task<DriveItem> GetDriveItem(string driveId, string documentId)
        {
            var retryCount = DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    return await GetDriveRoot(driveId)
                        .Items[documentId]
                        .Request()
                        .GetAsync();
                }
                catch (ServiceException ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorRetrievingDriveItem}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(String.Format(Resource.GraphRetryAttempts, retryInSeconds, DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }
            return null;
        }

        public async Task<DriveItem> SaveDriveItemToFileSystem(string driveId, string documentId, string downloadLocation)
        {
            var retryCount = DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    var document = await GetDriveItem(driveId, documentId);
                    if (document == null || document.File == null)
                    {
                        return null;
                    }

                    var drive = drivesConnectionInfo.First(x => x.Id == driveId);
                    var documentPath = document.ParentReference.Path.Replace($"/drives/{driveId}/root:", string.Empty).TrimStart('/').Replace(@"/", @"\");

                    string path = Path.Combine(downloadLocation, drive.Root, drive.Name, documentPath, document.Name);
                    System.IO.Directory.CreateDirectory(Path.GetDirectoryName(path));

                    document.AdditionalData.TryGetValue("@microsoft.graph.downloadUrl", out var downloadUrl);
                    var documentSize = document.Size;
                    var readSize = DEFAULT_CHUNK_SIZE;

                    using (FileStream outputFileStream = new FileStream(path, FileMode.Create))
                    {
                        long offset = 0;
                        while (offset < documentSize)
                        {
                            var chunkSize = documentSize - offset > readSize ? readSize : documentSize - offset;
                            var req = new HttpRequestMessage(HttpMethod.Get, downloadUrl.ToString());
                            req.Headers.Range = new RangeHeaderValue(offset, chunkSize + offset - 1);

                            try
                            {
                                var graphClientResponse = await graphServiceClient.HttpProvider.SendAsync(req);
                                using (var rs = await graphClientResponse.Content.ReadAsStreamAsync())
                                {
                                    rs.CopyTo(outputFileStream);
                                }
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
                    }

                    document.AdditionalData.Add("FilePath", path);
                    return document;
                }
                catch (ServiceException ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorSavingDriveItem}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(String.Format(Resource.GraphRetryAttempts, retryInSeconds, DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }
            return null;
        }

        #region Exchange Online methods

        public async Task<bool> SaveEmailToFileSystem(Message message, string downloadLocation, string userAccount, string fileName)
        {
            var retryCount = DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    string path = Path.Combine(downloadLocation, userAccount, fileName);
                    System.IO.Directory.CreateDirectory(Path.GetDirectoryName(path));

                    var email = await graphServiceClient.Users[userAccount].Messages[message.Id].Content.Request().GetAsync();

                    using (FileStream outputFileStream = new FileStream(path, FileMode.Create))
                    {
                        email.CopyTo(outputFileStream);
                    }
                    return true;
                }
                catch (TaskCanceledException)
                {
                    // We got a timeout, ignore for now
                    logger.LogInformation($"{Resource.ErrorSavingExchangeMessage} File too big. Go to next");
                }
                catch (ServiceException ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorSavingExchangeMessage}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }
            return false;
        }

        public async Task<IUserMessagesCollectionPage> GetEmailsAfterDate(string userAccount, DateTime afterDate, int skipIndex = 0, int emailCount = 10, bool includeAttachments = false)
        {
            var retryCount = DEFAULT_RETRY_COUNT;
            var filter = $"createdDateTime gt {afterDate.ToString("s")}Z";

            while (retryCount-- > 0)
            {
                try
                {
                    return await graphServiceClient.Users[userAccount].Messages.Request()
                                                .Filter(filter).Skip(skipIndex).Top(emailCount)
                                                .OrderBy("createdDateTime").Expand(includeAttachments ? "attachments" : "")
                                                .GetAsync();
                }
                catch (ServiceException ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorRetrievingExchangeMessages}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }
            return null;
        }

        public async Task<IMailFolderMessagesCollectionPage> GetEmailsFromFolderAfterDate(string userAccount, string folder, DateTime afterDate, int skipIndex = 0, int emailCount = 10, bool includeAttachments = false, bool preferText = false)
        {
            var retryCount = DEFAULT_RETRY_COUNT;
            var filter = $"createdDateTime gt {afterDate.ToString("s")}Z";

            while (retryCount-- > 0)
            {
                try
                {
                    var graphRequest = graphServiceClient.Users[userAccount].MailFolders[folder].Messages.Request();
                    if (preferText)
                    {
                        graphRequest = graphRequest.Header("Prefer", "outlook.body-content-type=\"text\"");
                    }

                    return await  graphRequest.Filter(filter).Skip(skipIndex).Top(emailCount)
                                              .OrderBy("createdDateTime").Expand(includeAttachments ? "attachments" : "")
                                              .GetAsync();
                }
                catch (ServiceException ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorRetrievingExchangeMessages}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }
            return null;
        }

        public async Task<List<MailFolder>> GetEmailFolders(string userAccount)
        {
            var retryCount = DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    var foldersResult =  await graphServiceClient.Users[userAccount].MailFolders.Request()
                                                    .Select("id,displayName,totalItemCount").GetAsync();
                    return foldersResult.CurrentPage.ToList();
                }
                catch (ServiceException ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorRetrievingExchangeFolders}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }
            return null;
        }

        public async Task<MailFolder> GetEmailFolderById(string userAccount, string folderId)
        {
            var retryCount = DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    return await graphServiceClient.Users[userAccount].MailFolders[folderId].Request().GetAsync();
                }
                catch (ServiceException ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorRetrievingExchangeFolders}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }
            return null;
        }

        public async Task<bool> ForwardEmail(string userAccount, string emailId, string forwardAccount)
        {
            var retryCount = DEFAULT_RETRY_COUNT;
            var forwardToRecipient = new List<Recipient>{   
                new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Name = forwardAccount,
                        Address = forwardAccount
                    }
                }
            };

            while (retryCount-- > 0)
            {
                try
                {
                    await graphServiceClient.Users[userAccount].Messages[emailId]
                            .Forward(forwardToRecipient)
                            .Request()
                            .PostAsync();
                    return true;
                }
                catch (ServiceException ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorForwardingEmail}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }
            return false;
        }

        public async Task<bool> MoveEmailToFolder(string userAccount, string emailId, string newFolder)
        {
            var retryCount = DEFAULT_RETRY_COUNT;

            while (retryCount-- > 0)
            {
                try
                {
                    var folder = (await GetEmailFolders(userAccount)).FirstOrDefault(x => x.DisplayName.Equals(newFolder, StringComparison.InvariantCultureIgnoreCase));

                    if (folder == null)
                    {
                        return false;
                    }

                    await graphServiceClient.Users[userAccount].Messages[emailId]
                            .Move(folder.Id)
                            .Request()
                            .PostAsync();
                    return true;
                }
                catch (ServiceException ex)
                {
                    var retryInSeconds = GetRetryAfterSeconds(ex);
                    logger.LogError($"{Resource.ErrorMovingEmail}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                    logger.LogError(string.Format(Resource.GraphRetryAttempts, retryInSeconds, DEFAULT_RETRY_COUNT - retryCount));
                    Thread.Sleep(retryInSeconds * 1000);
                }
            }
            return false;
        }

        #endregion

        #region Private Methods

        private IDriveRequestBuilder GetDriveRoot(string driveId)
        {
            switch(connectionType)
            {
                case DriveConnectionType.OneDriveConnection:
                    return graphServiceClient.Users[userId].Drives[driveId];
                case DriveConnectionType.SharePoinConnection:
                    return graphServiceClient.Sites[siteId].Drives[driveId];
                case DriveConnectionType.MSTeamsConnection:
                    var groupId = drivesConnectionInfo.FirstOrDefault(s => s.Id == driveId)?.GroupId;
                    return graphServiceClient.Groups[groupId].Drives[driveId];
                default: 
                    return null;
            }
        }

        private async Task<bool> InitializeAppPermissions()
        {
            if (UseClientSecret())
            {
                cca = ConfidentialClientApplicationBuilder.Create(authConfig.ClientId)
                    .WithClientSecret(authConfig.ClientSecret)
                    .WithAuthority(new Uri(authConfig.Authority))
                    .Build();
            }
            else
            {
                X509Certificate2 certificate = ReadCertificate(authConfig.CertificateName);
                cca = ConfidentialClientApplicationBuilder.Create(authConfig.ClientId)
                    .WithCertificate(certificate)
                    .WithAuthority(new Uri(authConfig.Authority))
                    .Build();
            }

            // With client credentials flows the scopes is ALWAYS of the shape "resource/.default", as the 
            // application permissions need to be set statically (in the portal or by PowerShell), and then granted by
            // a tenant administrator. 
            string[] scopes = new string[] { $"{authConfig.ApiUrl}.default" };

            try
            {
                graphServiceClient =
                    new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) => {
                    // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                    var authResult = await cca.AcquireTokenForClient(scopes).ExecuteAsync();

                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                    }));

            }
            catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS70011"))
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
                // Configure the MSAL client to get tokens
                var pcaOptions = new PublicClientApplicationOptions
                {
                    ClientId = authConfig.ClientId,
                    TenantId = authConfig.Tenant,
                    RedirectUri = "http://localhost"
                };

                msalCachePath = $"{authConfig.TokenPath ?? "C:\\Temp"}\\MsalCache";
                msalCacheFileName = $"MSALCache-{Guid.NewGuid()}.plaintext";
                System.IO.Directory.CreateDirectory(Path.GetDirectoryName(msalCachePath));

                var storageProperties = new StorageCreationPropertiesBuilder(msalCacheFileName, msalCachePath)
                                            .WithUnprotectedFile()
                                            .Build();

                pca = PublicClientApplicationBuilder
                        .CreateWithApplicationOptions(pcaOptions)
                        .Build();

                msalCacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);
                msalCacheHelper.RegisterCache(pca.UserTokenCache);

                // The permission scopes required
                var graphScopes = new string[] { 
                    "https://graph.microsoft.com/Files.Read.All",
                    "https://graph.microsoft.com/Group.Read.All",
                    "https://graph.microsoft.com/Sites.Read.All",
                    "https://graph.microsoft.com/User.Read.All",
                };

                var authResult = await pca.AcquireTokenInteractive(graphScopes).ExecuteAsync();

                graphServiceClient =
                    new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                    {
                        if (IsTokenCloseToExpiration(authResult.AccessToken)) 
                        {
                            var accounts = await pca.GetAccountsAsync();
                            authResult = await pca.AcquireTokenSilent(graphScopes, accounts.FirstOrDefault())
                                                .WithForceRefresh(true)
                                                .ExecuteAsync();
                        }

                        // Add the access token in the Authorization header of the API request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                    }));
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
            if (authenticationToken == null) return false;

            try
            {
                graphServiceClient =
                    new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                    {
                        await Task.Run(async () =>
                        {
                            await GetValidToken();
                            if (authenticationToken != null)
                            {
                                // Add the access token in the Authorization header of the API request.
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authenticationToken.Access_Token);
                            }
                            else
                            {
                                logger.LogError(Resource.CannotGetJwtToken);
                            }
                        });
                    }));
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorInitializingGraph}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return await Task.FromResult(false);
            }

            return true;
        }

        // This other method can also be used to initialize the GraphServiceClient
        // With this method we can tap into the middleware pipeline
        // The ChaosHandler can be used to generate random errors and test recovery
        // See: https://camerondwyer.com/2021/09/23/how-to-use-the-microsoft-graph-sdk-chaos-handler-to-simulate-graph-api-errors/
        private bool InitGraphServiceClient()
        {
            var tokenCredential = new ClientSecretCredential(authConfig.Tenant, authConfig.ClientId, authConfig.ClientSecret);

            // Use the static GraphClientFactory to get the default pipeline
            var handlers = GraphClientFactory.CreateDefaultHandlers(new TokenCredentialAuthProvider(tokenCredential));

            // Remove the default Retry Handler
            var retryHandler = handlers.Where(h => h is RetryHandler).FirstOrDefault();
            handlers.Remove(retryHandler);

            // Add a custom middleware handler (the Chaos Handler) to the pipeline
            handlers.Add(new ChaosHandler(new ChaosHandlerOption()
            {
                ChaosPercentLevel = 25
            }));

            // Now we have an extra step of creating a HTTPClient passing in the customized pipeline
            var httpClient = GraphClientFactory.Create(handlers);

            // Then we construct the Graph Service Client using the HTTPClient
            graphServiceClient = new GraphServiceClient(httpClient);

            return true;
        }

        private async Task<bool> IsInitialized()
        {
            if (graphServiceClient != null)
            {
                // TODO: Check if graphServiceClient is still connected
                // Even when we could already have an instance of the client the connection may had been lost
                return true;
            }

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

            if (graphServiceClient != null)
            {
                //Increase GraphClient HttpClient timeout 
                graphServiceClient.HttpProvider.OverallTimeout = TimeSpan.FromHours(3);
            }
            return connected;
        }

        private bool UseClientSecret()
        {
            if (!string.IsNullOrWhiteSpace(authConfig.ClientSecret))
            {
                return true;
            }
            else if (!string.IsNullOrWhiteSpace(authConfig.CertificateName))
            {
                return false;
            }
            else
            {
                throw new Exception(Resource.ChooseClientOrCertificate);
            }
        }

        private static X509Certificate2 ReadCertificate(string certificateName)
        {
            if (string.IsNullOrWhiteSpace(certificateName))
            {
                throw new ArgumentException(Resource.CertificateEmpty, nameof(certificateName));
            }

            var store = new X509Store(StoreName.My, StoreLocation.LocalMachine);
            store.Open(OpenFlags.ReadOnly);
            var certCollection = store.Certificates;
            var currentCerts = certCollection.Find(X509FindType.FindBySubjectDistinguishedName, certificateName, false);
            return currentCerts.Count == 0 ? null : currentCerts[0];
        }

        private int GetRetryAfterSeconds(ServiceException ex)
        {
            IEnumerable<string> retries;
            switch (ex.StatusCode)
            {
                case System.Net.HttpStatusCode.TooManyRequests:
                case System.Net.HttpStatusCode.ServiceUnavailable:
                case System.Net.HttpStatusCode.GatewayTimeout:
                    return ex.ResponseHeaders.TryGetValues("Retry-After", out retries) ? int.Parse(retries.First()) : DEFAULT_RETRY_IN_SECONDS;
                default:
                    return 1;
            }
        }

        private AuthenticationToken InitializeFromTokenPath()
        {
            if (!SystemFile.Exists(authConfig.TokenPath)) return null;

            string tokenString;
            AuthenticationToken token;
            try
            {
                tokenString = SystemFile.ReadAllText(authConfig.TokenPath);
                if (string.IsNullOrWhiteSpace(tokenString)) return null;
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
                {
                    SystemFile.WriteAllText(authConfig.TokenPath, AesHelper.EncryptToString(tokenString));
                }
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

        private async Task GetValidToken(int timeInMinutes = 30)
        {
            var jwtToken = new JwtSecurityToken(authenticationToken.Access_Token);
            TimeSpan dateDiff = jwtToken.ValidTo - DateTime.UtcNow;
            if (dateDiff.TotalMinutes < timeInMinutes)
            {
                await RefreshToken();
            }
        }

        private async Task RefreshToken()
        {
            try
            {
                string refreshUrl = $"https://login.microsoftonline.com/{authConfig.Tenant}/oauth2/v2.0/token";
                Dictionary<string, string> data = new Dictionary<string, string>
                {
                    { "grant_type", "refresh_token" },
                    { "client_id", authConfig.ClientId },
                    { "refresh_token", authenticationToken.Refresh_Token }
                };

                using var httpClient = new HttpClient();
                var response = await httpClient.PostAsync(refreshUrl, new FormUrlEncodedContent(data));

                var result = await response.Content.ReadAsStringAsync();

                authenticationToken = JsonConvert.DeserializeObject<AuthenticationToken>(result);

                SystemFile.WriteAllText(authConfig.TokenPath, AesHelper.EncryptToString(result));
            }
            catch (Exception)
            {
                logger.LogError(Resource.CannotGetRefreshToken);
            }
        }

        public static bool IsTokenCloseToExpiration(string token, int timeInMinutes = 10)
        {
            var jwtToken = new JwtSecurityToken(token);
            TimeSpan dateDiff = jwtToken.ValidTo - DateTime.UtcNow;
            if (dateDiff.TotalMinutes < timeInMinutes)
            {
                return true;
            }
            return false;
        }

        #endregion
    }
}