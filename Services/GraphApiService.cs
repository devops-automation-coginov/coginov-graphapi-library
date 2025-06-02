using Azure.Identity;
using Coginov.GraphApi.Library.Enums;
using Coginov.GraphApi.Library.Helpers;
using Coginov.GraphApi.Library.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using SystemFile = System.IO.File;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using Microsoft.Kiota.Abstractions.Authentication;
using Azure.Core;
using Microsoft.Kiota.Http.HttpClientLibrary.Middleware;
using System.Text.RegularExpressions;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Http.HttpClientLibrary.Middleware.Options;
using DriveUpload = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Authentication.Azure;
using Microsoft.Graph.Drives.Item.Items.Item.Copy;
using System.Threading;

namespace Coginov.GraphApi.Library.Services
{
    public class GraphApiService : IGraphApiService
    {
        /// <summary>
        /// The logger instance used for logging messages within this service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Configuration settings for authentication with Microsoft Graph API.
        /// </summary>
        private AuthenticationConfig authConfig;

        /// <summary>
        /// HttpClient instance used for making HTTP requests to the Microsoft Graph API, primarily for large file downloads.
        /// </summary>
        private HttpClient graphHttpClient;

        /// <summary>
        /// The Microsoft Graph service client used for interacting with the Microsoft Graph API.
        /// </summary>
        private GraphServiceClient graphServiceClient;

        /// <summary>
        /// A list to store connection information for drives (document libraries), particularly for SharePoint and Teams.
        /// </summary>
        private List<DriveConnectionInfo> drivesConnectionInfo = new List<DriveConnectionInfo>();

        /// <summary>
        /// Specifies the type of connection being used with Microsoft Graph (e.g., OneDrive, SharePoint, Teams).
        /// </summary>
        private DriveConnectionType connectionType;

        /// <summary>
        /// The user principal name (UPN) or object ID of the user, used primarily for OneDrive connections.
        /// </summary>
        private string userId;

        /// <summary>
        /// The unique identifier of the SharePoint site.
        /// </summary>
        private string siteId;

        /// <summary>
        /// An object holding the authentication token retrieved for accessing Microsoft Graph.
        /// </summary>
        private AuthenticationToken authenticationToken;

        /// <summary>
        /// An instance of the InteractiveBrowserCredential used for delegated permissions authentication.
        /// </summary>
        private InteractiveBrowserCredential ibc;

        // SharePointOnline
        /// <summary>
        /// The URL of the SharePoint site.
        /// </summary>
        private string siteUrl;
        /// <summary>
        /// An array of document library names within the SharePoint site.
        /// </summary>
        private string[] docLibraries;

        // OneDrive
        /// <summary>
        /// The user principal name (UPN) or object ID of the user for OneDrive access.
        /// </summary>
        private string oneDriveUserAccount;

        // MsTeams
        /// <summary>
        /// An array of Microsoft Teams identifiers (e.g., group IDs).
        /// </summary>
        private string[] teams;

        /// <summary>
        /// A flag to enable or disable the Chaos Handler, which simulates random Graph API errors for development and testing purposes. Defaults to <c>false</c>.
        /// </summary>
        private bool useChaosHandler = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphApiService"/> class with the provided logger.
        /// </summary>
        /// <param name="logger">The logger instance to use for logging messages.</param>
        public GraphApiService(ILogger logger)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Initializes the Microsoft Graph API service using the provided authentication configuration.
        /// </summary>
        /// <param name="authenticationConfig">The <see cref="AuthenticationConfig"/> object containing the necessary authentication parameters.</param>
        /// <param name="forceInit">Optional. A boolean value indicating whether to force re-initialization even if the service is already initialized. Defaults to <c>true</c>.</param>
        /// <returns><c>true</c> if the initialization was successful; otherwise, <c>false</c>.</returns>
        public async Task<bool> InitializeGraphApi(AuthenticationConfig authenticationConfig, bool forceInit = true)
        {
            this.authConfig = authenticationConfig;
            return await IsInitialized(forceInit);
        }

        /// <summary>
        /// Initializes the connection to SharePoint Online.
        /// </summary>
        /// <param name="authenticationConfig">The authentication configuration for connecting to Microsoft Graph.</param>
        /// <param name="siteUrl">The URL of the SharePoint Online site.</param>
        /// <param name="docLibraries">An array of document library names within the site to interact with.</param>
        /// <param name="forceInit">A boolean value indicating whether to force re-initialization even if already initialized.</param>
        /// <returns>A boolean value indicating whether the initialization was successful.</returns>
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

        /// <summary>
        /// Initializes the connection to a user's OneDrive for Business.
        /// </summary>
        /// <param name="authenticationConfig">The authentication configuration for connecting to Microsoft Graph.</param>
        /// <param name="userAccount">The user's email address or principal name for accessing their OneDrive.</param>
        /// <param name="forceInit">A boolean value indicating whether to force re-initialization even if already initialized.</param>
        /// <returns>A boolean value indicating whether the initialization was successful.</returns>
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

        /// <summary>
        /// Initializes the connection to Microsoft Teams.
        /// </summary>
        /// <param name="authenticationConfig">The authentication configuration for connecting to Microsoft Graph.</param>
        /// <param name="teams">An optional array of Microsoft Teams IDs to interact with. If null, interacts with all accessible teams.</param>
        /// <param name="forceInit">A boolean value indicating whether to force re-initialization even if already initialized.</param>
        /// <returns>A boolean value indicating whether the initialization was successful.</returns>
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

        /// <summary>
        /// Initializes the connection to Exchange Online.
        /// </summary>
        /// <param name="authenticationConfig">The authentication configuration for connecting to Microsoft Graph.</param>
        /// <param name="forceInit">A boolean value indicating whether to force re-initialization even if already initialized.</param>
        /// <returns>A boolean value indicating whether the initialization was successful.</returns>
        public async Task<bool> InitializeExchangeConnection(AuthenticationConfig authenticationConfig, bool forceInit = false)
        {
            authConfig = authenticationConfig;
            if (!await IsInitialized(forceInit))
                return false;

            connectionType = DriveConnectionType.ExchangeConnection;
            return true;
        }

        /// <summary>
        /// Retrieves an access token using application permissions.
        /// </summary>
        /// <param name="tenantId">The ID of the Azure Active Directory tenant.</param>
        /// <param name="clientId">The client ID (application ID) of the Azure AD application.</param>
        /// <param name="clientSecret">The client secret of the Azure AD application.</param>
        /// <param name="scopes">An array of scopes required for the access token.</param>
        /// <returns>The access token as a string, or null if an error occurs.</returns>
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

        /// <summary>
        /// Retrieves an access token using delegated permissions, requiring user interaction through a browser.
        /// </summary>
        /// <param name="tenantId">The ID of the Azure Active Directory tenant.</param>
        /// <param name="clientId">The client ID (application ID) of the Azure AD application.</param>
        /// <param name="scopes">An array of scopes required for the access token.</param>
        /// <param name="cancellationToken">A cancellation token that can be used to cancel the operation.</param>
        /// <returns>The access token as a string, or null if an error occurs or the operation is cancelled.</returns>
        public async Task<string> GetTokenDelegatedPermissions(string tenantId, string clientId, string[] scopes, CancellationToken cancellationToken = default)
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

                if (cancellationToken == default)
                {
                    // If no calcellationToken is passed we call GetTokenAsync this way
                    var authResult = await ibc.GetTokenAsync(context);
                    return authResult.Token;
                }

                // If cancellationToken is passed we need to use this approach to be able to cancel
                // WARNING: GetTokenAsync is not handeling cancellationToken as expected here
                // var authResult = await ibc.GetTokenAsync(context, cancellationToken);
                // WORK AROUND: To cancel the task we wrap the GetTokenAsync call in a task that
                // can be canceled.
                // Although this will not truly cancel the underlying operation, it will allow
                // you to handle the cancellation logic in the client code that calls this method
                var tokenTask = Task.Run(async () => await ibc.GetTokenAsync(context, cancellationToken), cancellationToken);
                return (await tokenTask).Token;
            }
            catch (OperationCanceledException ex)
            {
                logger.LogError($"{Resource.TimeOutGettingJwtToken}. {ex.Message}. {ex.InnerException?.Message}");
                return null;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.CannotGetJwtToken}. {ex.Message}. {ex.InnerException?.Message}");
                return null;
            }
        }

        /// <summary>
        /// Retrieves the unique identifier (ID) of a user from Microsoft Graph.
        /// </summary>
        /// <param name="user">The user's principal name (UPN) or email address.</param>
        /// <returns>The user's ID as a string, or null if the user is not found or an error occurs.</returns>
        public async Task<string> GetUserId(string user)
        {
            try
            {
                var userObject = await graphServiceClient.Users[user].GetAsync();
                return userObject?.Id;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingUserId}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
        }

        /// <summary>
        /// Retrieves the unique identifier (ID) of a SharePoint Online site from its URL.
        /// </summary>
        /// <param name="siteUrl">The full URL of the SharePoint Online site.</param>
        /// <returns>The site's ID as a string, or null if the site is not found or an error occurs.</returns>
        public async Task<string> GetSiteId(string siteUrl)
        {
            try
            {
                var uri = new Uri(siteUrl);
                var siteId = await graphServiceClient.Sites[$"{uri.Host}:{uri.PathAndQuery}"]
                    .GetAsync(requestConfiguration => 
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

        /// <summary>
        /// Retrieves a list of document library drive information for the configured SharePoint Online site.
        /// </summary>
        /// <returns>A list of <see cref="DriveConnectionInfo"/> objects representing the document libraries, or an empty list if an error occurs.</returns>
        public async Task<List<DriveConnectionInfo>> GetSharePointOnlineDrives()
        {
            drivesConnectionInfo = new List<DriveConnectionInfo>();
            if (docLibraries != null)
                // Removing leading and trailing spaces
                docLibraries = docLibraries.Select(x => x.Trim()).ToArray();

            try
            {
                // Here 'docLibraries' could contain the Doc Libraries to process or null if we want to process all Doc Libraries on the site
                var siteDrives = await GetSiteDrives(siteId);
                var selectedDrives = siteDrives.Where(x => docLibraries == null || docLibraries.Contains(x.Name));

                if (docLibraries == null)
                    selectedDrives = siteDrives;
                else
                {
                    foreach (var library in docLibraries)
                    {
                        // Show error if provided Document Library name doesn't exist
                        if (siteDrives.FirstOrDefault(x => x.Name == library) == null)
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
                logger.LogError($"{Resource.ErrorRetrievingDocLibraries}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return drivesConnectionInfo;
        }

        /// <summary>
        /// Retrieves a list of drive information for the configured OneDrive for Business user.
        /// </summary>
        /// <returns>A list of <see cref="DriveConnectionInfo"> objects representing the user's OneDrive drives, or an empty list if an error occurs.</returns>
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
                logger.LogError($"{Resource.ErrorRetrievingDrives}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return drivesConnectionInfo;
        }

        /// <summary>
        /// Retrieves a list of drive information for the configured Microsoft Teams.
        /// </summary>
        /// <returns>A list of <see cref="DriveConnectionInfo"/> objects representing the Team drives, or an empty list if an error occurs.</returns>
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
                    groups = await graphServiceClient.Groups
                        .GetAsync(requestConfiguration =>
                        {
                            requestConfiguration.QueryParameters.Filter = "resourceProvisioningOptions/Any(x:x eq 'Team')";
                        });
                else
                {
                    var filter = string.Join(" or ", teams.Select(x => $"displayName eq '{x.Replace("'", "''").Trim()}'"));
                    filter = $"({filter}) and resourceProvisioningOptions / Any(x: x eq 'Team')";
                    groups = await graphServiceClient.Groups
                        .GetAsync(requestConfiguration =>
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
                logger.LogError($"{Resource.ErrorRetrievingTeams}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return drivesConnectionInfo;
        }

        /// <summary>
        /// Retrieves a list of document IDs and metadata from a specific drive, using delta queries for incremental changes.
        /// </summary>
        /// <param name="driveId">The ID of the drive to query.</param>
        /// <param name="lastDate">The last known date of changes. Used for initial delta query. Subsequent calls use the skip token.</param>
        /// <param name="top">The maximum number of items to retrieve in a single request (for pagination).</param>
        /// <param name="skipToken">A token used to retrieve the next page of delta results. Null for the initial request.</param>
        /// <returns>A <see cref="DriveItemSearchResult"/> object containing the list of document IDs, information about more results, the new 
        /// skip token, and the latest modified date, or null if an error occurs.
        /// </returns>
        public async Task<DriveItemSearchResult> GetDocumentIds(string driveId, DateTime lastDate, int top, string skipToken)
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
                        .GetAsDeltaGetResponseAsync(requestConfiguration =>
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
                        .GetAsDeltaWithTokenGetResponseAsync(requestConfiguration =>
                        {
                            requestConfiguration.QueryParameters.Top = top;
                            requestConfiguration.QueryParameters.Orderby = new string[] { "lastModifiedDateTime" };
                        });
                }

                var deltaLink = deltaResponse?.OdataNextLink ?? deltaResponse?.OdataDeltaLink;

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
                    HasMoreResults = deltaResponse?.OdataNextLink != null,
                    SkipToken = skipToken,
                    LastDate = deltaResults?.LastOrDefault()?.LastModifiedDateTime?.DateTime ?? lastDate
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
                logger.LogError($"{Resource.ErrorRetrievingDocumentIds}: {ex.Message}. {ex.InnerException?.Message}");
                if (ex.ResponseStatusCode == 403)
                {
                    var drive = drivesConnectionInfo.FirstOrDefault(x => x.Id == driveId);
                    var errorMessage = $"{Resource.ErrorAccessDeniedToDrive}: '{drive?.Name}'. {ex.Message}. {ex.InnerException?.Message}";
                    logger.LogError(errorMessage);
                    return new DriveItemSearchResult
                    {
                        DocumentIds = new List<DriveItem>(),
                        ErrorMessage = errorMessage
                    };
                }
            }
            catch(Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingDocumentIds}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
        }

        /// <summary>
        /// Retrieves the metadata of a specific item (file or folder) from a drive.
        /// </summary>
        /// <param name="driveId">The ID of the drive containing the item.</param>
        /// <param name="documentId">The ID of the item to retrieve.</param>
        /// <returns>A <see cref="DriveItem"/> object representing the item's metadata, or null if an error occurs.</returns>
        public async Task<DriveItem> GetDriveItem(string driveId, string documentId)
        {
            try
            {
                var driveRoot = await GetDriveRoot(driveId);
                return await graphServiceClient.Drives[driveRoot.Id]
                    .Items[documentId]
                    .GetAsync();
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingDriveItem}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
        }

        /// <summary>
        /// Downloads a specific item (file) from a drive and saves it to the local file system.
        /// </summary>
        /// <param name="driveId">The ID of the drive containing the item.</param>
        /// <param name="documentId">The ID of the item to download.</param>
        /// <param name="downloadLocation">The local path where the file should be saved.</param>
        /// <returns>A <see cref="DriveItem"/> object representing the item's metadata with added local file path, or null if an error occurs.</returns>
        public async Task<DriveItem> SaveDriveItemToFileSystem(string driveId, string documentId, string downloadLocation)
        {
            try
            {
                var document = await GetDriveItem(driveId, documentId);
                if (document == null || document.File == null)
                    return null;

                var drive = await GetSharePointDriveConnectionInfo(driveId);
                var documentPath = document.ParentReference.Path.Replace($"/drives/{driveId}/root:", string.Empty).TrimStart('/').Replace(@"/", @"\");

                string filePath = Path.Combine(downloadLocation, drive.Root, drive.Name, documentPath, document.Name).GetFilePathWithTimestamp();
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));

                document.AdditionalData.Add("FilePath", filePath);
                document.AdditionalData.Add("ParentUrl", $"{drive.Path}{document.ParentReference.Path.ExtractStringAfterRoot()}");

                try
                {
                    var driveRoot = await GetDriveRoot(driveId);
                    var documentStream = await graphServiceClient.Drives[driveRoot.Id].Items[documentId].Content.GetAsync();

                    using (FileStream outputFileStream = new FileStream(filePath, FileMode.Create))
                        documentStream.CopyTo(outputFileStream);

                    return document;
                }
                catch (Exception)
                {
                    // We got an error while saving document content. Let's try in chuncks in case it is too big
                    return await SaveDriveItemToFileSystem(document, filePath);
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorSavingDriveItem}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
        }

        /// <summary>
        /// Retrieves the metadata of a specific item (file) from a drive and adds the parent URL and an empty file path to its additional data.
        /// This method is intended for retrieving metadata without downloading the file content.
        /// </summary>
        /// <param name="driveId">The ID of the drive containing the item.</param>
        /// <param name="documentId">The ID of the item to retrieve metadata for.</param>
        /// <returns>A <see cref="DriveItem"/> object representing the item's metadata with added 'ParentUrl' and empty 'FilePath' in AdditionalData, or null if an error occurs.</returns>
        public async Task<DriveItem> GetDriveItemMetadata(string driveId, string documentId)
        {
            try
            {
                var document = await GetDriveItem(driveId, documentId);
                if (document == null || document.File == null)
                    return null;

                var drive = await GetSharePointDriveConnectionInfo(driveId);
                var documentPath = document.ParentReference.Path.Replace($"/drives/{driveId}/root:", string.Empty).TrimStart('/').Replace(@"/", @"\");

                document.AdditionalData.Add("FilePath", string.Empty);
                document.AdditionalData.Add("ParentUrl", $"{drive.Path}{document.ParentReference.Path.ExtractStringAfterRoot()}");

                return document;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorGettingDriveItemMetadata}: {ex.Message}. {ex.InnerException?.Message}");
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
        /// <param name="onConflict">Optional conflict resolution behaviour. Default: rename</param>
        /// <returns></returns>
        public async Task<bool> UploadDocumentToDrive(string driveId, string filePath, string fileName = null, string folderPath = "", string onConflict = "rename")
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
                using var fileStream = System.IO.File.OpenRead(filePath);

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
            catch (Exception ex)
            {
                logger.LogError($"{Resource.DriveItemUploadFailed}: {ex.Message}. {ex.InnerException?.Message}");
            }


            return false;
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
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorDeletingDriveItem}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return false;
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
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorDeletingDriveItem}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return false;
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
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorMovingDriveItem}: {ex.Message}. {ex.InnerException?.Message}");
                throw;
            }
        }

        /// <summary>
        /// Moves a document from one location to another within SharePoint Online.
        /// https://stackoverflow.com/questions/66478737/invalidrequest-moving-driveitem-between-drives
        /// </summary>
        /// <param name="driveId">The unique identifier of the drive (document library) where the document is currently located.</param>
        /// <param name="documentId">The unique identifier of the document to move.</param>
        /// <param name="destSite">The URL of the destination SharePoint site.</param>
        /// <param name="destDocLib">The name of the destination document library within the destination site.</param>
        /// <param name="destFolderId">Optional. The unique identifier of the destination folder within the destination document library. If not provided, <paramref name="destFolder"/> is used.</param>
        /// <param name="destFolder">Optional. The path to the destination folder within the destination document library (e.g., "Subfolder/AnotherFolder"). If not provided, the document will be moved to the root of the destination library.</param>
        /// <param name="docNewName">Optional. The new name for the document after it has been moved. If not provided, the original name is retained.</param>
        /// <returns><c>true</c> if the document was moved successfully; otherwise, <c>false</c>.</returns>
        public async Task<bool> MoveDocument(string driveId, string documentId, string destSite, string destDocLib, string destFolderId = null, string destFolder = null, string docNewName = null)
        {
            try
            {
                siteUrl = destSite;
                siteId = await GetSiteId(destSite);
                if (siteId == null)
                    return false;

                var drives = await GetSharePointOnlineDrives();
                var destDrive = drives.FirstOrDefault(x => x.Name == destDocLib);
                if (destDrive == null)
                    return false;

                if (destFolderId == null)
                {
                    var folder = await graphServiceClient.Drives[destDrive.Id].Items["root"].ItemWithPath(destFolder ?? "//").GetAsync();

                    if (folder == null)
                    {
                        logger.LogError(Resource.DestinationFolderNotFound);
                        return false;
                    }

                    destFolderId = folder.Id;
                }

                var parentReference = new ItemReference
                {
                    DriveId = destDrive.Id,
                    Id = destFolderId
                };

                var requestBody = new CopyPostRequestBody
                {
                    ParentReference = new ItemReference
                    {
                        DriveId = destDrive.Id,
                        Id = destFolderId,
                    }
                };

                if (docNewName != null)
                    requestBody.Name = docNewName;

                // TODO: We need to confirm if the Copy operation succeeded. The only way to validate
                // that is if we call another api and find the file in the destination. If not found
                // do not delete it from the source location
                await graphServiceClient.Drives[driveId].Items[documentId].Copy.PostAsync(requestBody);
                
                // TODO: Before we delete the item we need to confirm if the Copy operation succeeded.
                // We need to implement that verification before deleting the item from soruce location
                await graphServiceClient.Drives[driveId].Items[documentId].DeleteAsync();

                // TODO: Return true only if previous steps were successfull: Copy and Delete
                return true;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorMovingDriveItem}: {ex.Message}. {ex.InnerException?.Message}");
                throw;
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

                // There is a known bug in Graph API SDK when searching sites with * character: https://github.com/microsoftgraph/msgraph-sdk-dotnet/issues/1884
                // This is a workaround to overcome this: https://github.com/microsoftgraph/msgraph-sdk-dotnet/issues/1826#issuecomment-1531405983
                requestInformation.UrlTemplate = requestInformation.UrlTemplate.Replace("%24search", "search");
                requestInformation.QueryParameters.Add("search", "*");

                var sites = new List<Site>();

                // This is one of the ways to retrieve paged results from GraphAPI SDK. Using PageIterator
                // https://learn.microsoft.com/en-us/graph/sdks/paging
                var sitesReponse = await graphServiceClient.RequestAdapter.SendAsync(requestInformation, SiteCollectionResponse.CreateFromDiscriminatorValue);
                var sitesIterator = PageIterator<Site, SiteCollectionResponse>.CreatePageIterator(graphServiceClient, sitesReponse, (site) => { sites.Add(site); return true; });
                await sitesIterator.IterateAsync();

                // If the Url provided is not a tenant sharepoint root Url we will exclude personal sites anyway
                var isTenantRoot = siteUrl.IsRootUrl();
                excludePersonalSites |= !isTenantRoot;

                if (!excludePersonalSites)
                {
                    var personalSitesResponse = await graphServiceClient.Sites.GetAllSites.GetAsGetAllSitesGetResponseAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter = "IsPersonalSite eq true";
                    });

                    if (personalSitesResponse != null && personalSitesResponse.Value.Any())
                        sites.AddRange(personalSitesResponse.Value);

                    // This is the other way to retrieve paged results from GraphAPI SDK. Using OdataNextLink
                    // https://learn.microsoft.com/en-us/graph/sdks/paging
                    var nextLink = personalSitesResponse.OdataNextLink;
                    while (nextLink != null)
                    {
                        var nextSitesResponse = await graphServiceClient.Sites.GetAllSites.WithUrl(nextLink).GetAsGetAllSitesGetResponseAsync();
                        if (nextSitesResponse != null && nextSitesResponse.Value.Any())
                        {
                            sites.AddRange(nextSitesResponse.Value);
                            nextLink = nextSitesResponse.OdataNextLink;
                        }
                    }
                }

                // The sites list returned by the search query returns all sites in the tenant. We found no other way to retrieve all sites.
                // Since this is not a recurring process we can afford it. So, if we are not processing the tenant root url we must filter out sites outside current site url
                if (!isTenantRoot)
                {
                    sites = sites.Where(x => x.WebUrl.StartsWith(siteUrl, StringComparison.InvariantCultureIgnoreCase)).ToList();
                }

                var siteDocs = await GetSiteAndDocLibsDictionary(sites.ToList(), excludeSystemDocLibs);
                return siteDocs.ToDictionary(x => x.Key, x => x.Value.Select(x => x.Name).ToList());

            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingSpoSites}: {ex.Message}. {ex.InnerException?.Message}");
            }
            return null;
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

                searchFilter ??= $"fields/{searchField} eq '{searchValue.Replace("'", "''").Trim()}'";

                var folders = await graphServiceClient.Sites[siteId].Lists[Uri.EscapeDataString(docLibrary)].Items
                    .GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Expand = new string[] { "fields", "driveItem" };
                        requestConfiguration.QueryParameters.Filter = $"(fields/ContentType eq 'Document Set' or fields/ContentType eq 'Folder') and ({searchFilter})";
                        requestConfiguration.QueryParameters.Select = new string[] { "sharepointIds" };
                        requestConfiguration.QueryParameters.Top = top;
                        //The search for keywords is always done with a set http header "prefer" with the value "HonorNonIndexedQueriesWarningMayFailRandomly" 
                        //This way we are able to search over index terms that are not mapped in an index. The other option is to add an index to the filed to be filtered
                        //https://learn.microsoft.com/en-us/answers/questions/1255945/use-graph-api-get-items-on-a-sharepoint-list-with
                        requestConfiguration.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
                    });

                var foldersIterator = new PageIterator<ListItem, ListItemCollectionResponse>();

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
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorSearchingFolders}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
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
                    catch (ODataError ex)
                    {
                        result.Add(item, ex.Message);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorUpdatingSharepointItems}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
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
                    var listItemResult = await graphServiceClient.Drives[item.DriveId].Items[item.DriveItemId].ListItem
                        .GetAsync((requestConfiguration) =>
                        {
                            requestConfiguration.QueryParameters.Expand = new string[] { "fields", "driveItem" };
                            requestConfiguration.QueryParameters.Select = new string[] { "sharepointIds" };
                        });

                    listItems.Add(listItemResult);
                }

                return await UpdateSharePointOnlineItemFieldValue(listItems, columnKeyValues);
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorUpdatingSharepointItems}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
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
                var driveItemResult = await graphServiceClient.Drives[driveItem.DriveId].Items[driveItem.DriveItemId].Children
                    .GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Top = batchSize;
                        requestConfiguration.QueryParameters.Orderby = new string[] { "lastModifiedDateTime" };
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
                logger.LogError($"{Resource.ErrorRetrievingDocuments}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
        }

        #region Exchange Online methods

        /// <summary>
        /// Downloads the content of an email message and saves it to the local file system.
        /// </summary>
        /// <param name="message">The <see cref="Message"/> object representing the email to save.</param>
        /// <param name="downloadLocation">The local directory where the email should be saved.</param>
        /// <param name="userAccount">The user's email address or principal name.</param>
        /// <param name="fileName">The name to use for the saved email file.</param>
        /// <returns>A boolean value indicating whether the email was successfully saved.</returns>
        public async Task<bool> SaveEmailToFileSystem(Message message, string downloadLocation, string userAccount, string fileName)
        {
            // The current limit for email size in Office 365 is 15O MB. Tested with a big email and working as expected
            try
            {
                string path = Path.Combine(downloadLocation, userAccount, fileName);
                Directory.CreateDirectory(Path.GetDirectoryName(path));

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
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorSavingExchangeMessage}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return false;
        }

        /// <summary>
        /// Retrieves the number of messages in a user's inbox.
        /// </summary>
        /// <param name="userAccount">The user's email address or principal name.</param>
        /// <returns>The number of messages in the inbox, or null if an error occurs.</returns>
        public async Task<int?> GetInboxMessageCount(string userAccount)
        {
            try
            {
                return await graphServiceClient.Users[userAccount].Messages.Count.GetAsync();

            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingExchangeMessagesCount}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
        }

        /// <summary>
        /// Retrieves emails from a user's inbox that were created after a specified date.
        /// This method was implemented initially for QoreAudit, but it could be used by any application
        /// </summary>
        /// <param name="userAccount">The user's email address.</param>
        /// <param name="afterDate">The date and time to filter emails after.</param>
        /// <param name="skipIndex">The number of emails to skip in the result set (for pagination).</param>
        /// <param name="emailCount">The maximum number of emails to retrieve in a single request.</param>
        /// <param name="includeAttachments">A boolean value indicating whether to include attachments in the retrieved emails.</param>
        /// <param name="preferText">A boolean value indicating whether to prefer the text body over HTML if available.</param>
        /// <param name="filterOperator">
        /// The operator to use when filtering by date ("gt" for greater than, "ge" for greater than or equal to). Default is "gt". 
        /// When using "ge" to include messages with the same timestamp the the caller needs to de-duplicate and remove emails already processed
        /// </param>
        /// <returns>A <see cref="MessageCollectionResponse"/> containing the retrieved emails, or null if an error occurs.</returns>
        public async Task<MessageCollectionResponse> GetEmailsAfterDate(
            string userAccount, 
            DateTime afterDate, 
            int skipIndex = 0, 
            int emailCount = 10, 
            bool includeAttachments = false,
            bool preferText = false,
            string filterOperator = "gt")
        {
            filterOperator = filterOperator.Trim();
            if (string.IsNullOrEmpty(filterOperator))
            {
                // If no operator is received then we default to gt
                filterOperator = "gt";
            }

            // IMPORTANT: We need to create the filter this way to avoid UTC issues:
            // $"createdDateTime ge {afterDate.ToString("s")}Z";
            var filter = $"createdDateTime {filterOperator} {afterDate.ToString("s")}Z";

            try
            {
                return await graphServiceClient.Users[userAccount].Messages
                    .GetAsync(requestConfiguration =>
                    {
                        if (preferText)
                        {
                            requestConfiguration.Headers.Add("Prefer", "outlook.body-content-type=\"text\"");
                        }

                        requestConfiguration.QueryParameters.Filter = filter;
                        requestConfiguration.QueryParameters.Skip = skipIndex;
                        requestConfiguration.QueryParameters.Top = emailCount;
                        requestConfiguration.QueryParameters.Orderby = new string[] { "createdDateTime asc" };

                        // Select only the fields you need (recommended for performance)
                        requestConfiguration.QueryParameters.Select = GraphHelper.SelectedEmailFields;

                        // Include attachments only when needed
                        requestConfiguration.QueryParameters.Expand = new string[] { includeAttachments ? "attachments" : "" };
                    });
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingExchangeMessages}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
        }

        /// <summary>
        /// Retrieves emails from a specific folder in a user's mailbox that were created after a specified date.
        /// This method was implemented initially for QoreMail / Siemens, but it could be used by any application
        /// </summary>
        /// <param name="userAccount">The user's email address.</param>
        /// <param name="folder">The name of the mail folder to retrieve emails from.</param>
        /// <param name="afterDate">The date and time to filter emails after.</param>
        /// <param name="skipIndex">The number of emails to skip in the result set (for pagination).</param>
        /// <param name="emailCount">The maximum number of emails to retrieve in a single request.</param>
        /// <param name="includeAttachments">A boolean value indicating whether to include attachments in the retrieved emails.</param>
        /// <param name="preferText">A boolean value indicating whether to prefer the text body over HTML if available.</param>
        /// <param name="filterOperator">
        /// The operator to use when filtering by date ("gt" for greater than, "ge" for greater than or equal to). Default is "gt". 
        /// When using "ge" to include messages with the same timestamp the the caller needs to de-duplicate and remove emails already processed
        /// </param>
        /// <returns>A <see cref="MessageCollectionResponse"/> containing the retrieved emails, or null if an error occurs.</returns>
        public async Task<MessageCollectionResponse> GetEmailsFromFolderAfterDate(
            string userAccount, 
            string folder, 
            DateTime afterDate, 
            int skipIndex = 0, 
            int emailCount = 10, 
            bool includeAttachments = false, 
            bool preferText = false,
            string filterOperator = "gt")
        {
            filterOperator = filterOperator.Trim();
            if (string.IsNullOrEmpty(filterOperator))
            {
                // If no operator is received then we default to gt
                filterOperator = "gt";
            }

            // IMPORTANT: We need to create the filter this way to avoid UTC issues:
            // $"createdDateTime ge {afterDate.ToString("s")}Z";
            var filter = $"createdDateTime {filterOperator} {afterDate.ToString("s")}Z";

            try
            {
                var graphRequest = graphServiceClient.Users[userAccount].MailFolders[folder].Messages;
                return await  graphRequest
                    .GetAsync(requestConfiguration =>
                    {
                        if (preferText)
                        {
                            requestConfiguration.Headers.Add("Prefer", "outlook.body-content-type=\"text\"");
                        }

                        requestConfiguration.QueryParameters.Filter = filter;
                        requestConfiguration.QueryParameters.Skip = skipIndex;
                        requestConfiguration.QueryParameters.Top = emailCount;
                        requestConfiguration.QueryParameters.Orderby = new string[] { "createdDateTime asc" };

                        // Select only the fields you need (recommended for performance)
                        requestConfiguration.QueryParameters.Select = GraphHelper.SelectedEmailFields;

                        // Include attachments only when needed
                        requestConfiguration.QueryParameters.Expand = new string[] { includeAttachments ? "attachments" : "" };
                    });
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingExchangeMessages}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
        }

        /// <summary>
        /// Retrieves a page of emails from a user's inbox that were created on or after a specified date.
        /// This method is optimized for pagination and avoids UTC conversion issues.
        /// The caller is responsible for de-duplicating and removing already processed emails.
        /// This method was implemented initially for QoreMail.Core, but it could be used by any application
        /// </summary>
        /// <param name="userAccount">The user's email address.</param>
        /// <param name="afterDate">The date and time to filter emails on or after.</param>
        /// <param name="pageSize">The maximum number of emails to retrieve in a single page.</param>
        /// <param name="includeAttachments">A boolean value indicating whether to include attachments in the retrieved emails.</param>
        /// <param name="preferText">A boolean value indicating whether to prefer the text body over HTML if available.</param>
        /// <param name="filterOperator">
        /// The operator to use when filtering by date ("gt" for greater than, "ge" for greater than or equal to). Default is "gt". 
        /// When using "ge" to include messages with the same timestamp the the caller needs to de-duplicate and remove emails already processed
        /// </param>
        /// <returns>A <see cref="MessageCollectionResponse"/> containing the retrieved emails, or null if an error occurs.</returns>
        public async Task<MessageCollectionResponse> GetEmailsAfterDate(
            string userAccount,
            DateTime afterDate,
            int pageSize = 10,
            bool includeAttachments = false,
            bool preferText = false,
            string filterOperator = "ge")
        {
            filterOperator = filterOperator.Trim();
            if (string.IsNullOrEmpty(filterOperator))
            {
                // If no operator is received then we default to ge
                filterOperator = "ge";
            }

            // IMPORTANT: We need to create the filter this way to avoid UTC issues:
            // $"createdDateTime ge {afterDate.ToString("s")}Z";
            var filter = $"createdDateTime {filterOperator} {afterDate.ToString("s")}Z";

            try
            {
                return await graphServiceClient.Users[userAccount].Messages
                    .GetAsync(requestConfiguration =>
                {
                    if (preferText)
                    {
                        requestConfiguration.Headers.Add("Prefer", "outlook.body-content-type=\"text\"");
                    }

                    requestConfiguration.QueryParameters.Filter = filter;
                    requestConfiguration.QueryParameters.Top = pageSize;
                    requestConfiguration.QueryParameters.Orderby = new[] { "createdDateTime asc" };

                    // Select only the fields you need (recommended for performance)
                    requestConfiguration.QueryParameters.Select = GraphHelper.SelectedEmailFields;

                    // Include attachments only when needed
                    requestConfiguration.QueryParameters.Expand = new string[] { includeAttachments ? "attachments" : "" };
                });
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingExchangeMessages}: {ex.Message}. {ex.InnerException?.Message}");
                return null;
            }
        }

        /// <summary>
        /// Retrieves a page of emails from a specific folder in a user's mailbox that were created on or after a specified date.
        /// This method is optimized for pagination and avoids UTC conversion issues.
        /// The caller is responsible for de-duplicating and removing already processed emails.
        /// </summary>
        /// <param name="userAccount">The user's email address.</param>
        /// <param name="folder">The name of the mail folder to retrieve emails from.</param>
        /// <param name="afterDate">The date and time to filter emails on or after.</param>
        /// <param name="pageSize">The maximum number of emails to retrieve in a single page.</param>
        /// <param name="includeAttachments">A boolean value indicating whether to include attachments in the retrieved emails.</param>
        /// <param name="preferText">A boolean value indicating whether to prefer the text body over HTML if available.</param>
        /// <param name="filterOperator">
        /// The operator to use when filtering by date ("gt" for greater than, "ge" for greater than or equal to). Default is "gt". 
        /// When using "ge" to include messages with the same timestamp the the caller needs to de-duplicate and remove emails already processed
        /// </param>
        /// <returns>A <see cref="MessageCollectionResponse"/> containing the retrieved emails, or null if an error occurs.</returns>
        public async Task<MessageCollectionResponse> GetEmailsFromFolderAfterDate(
            string userAccount,
            string folder,
            DateTime afterDate,
            int pageSize = 10,
            bool includeAttachments = false,
            bool preferText = false,
            string filterOperator = "ge")
        {
            filterOperator = filterOperator.Trim();
            if (string.IsNullOrEmpty(filterOperator))
            {
                // If no operator is passed we default to ge
                filterOperator = "ge";
            }

            // IMPORTANT: We need to create the filter this way to avoid UTC issues:
            // $"createdDateTime ge {afterDate.ToString("s")}Z";
            var filter = $"createdDateTime {filterOperator} {afterDate.ToString("s")}Z";

            try
            {
                var graphRequest = graphServiceClient.Users[userAccount].MailFolders[folder].Messages;

                return await graphRequest.GetAsync(requestConfiguration =>
                {
                    if (preferText)
                    {
                        requestConfiguration.Headers.Add("Prefer", "outlook.body-content-type=\"text\"");
                    }

                    requestConfiguration.QueryParameters.Filter = filter;
                    requestConfiguration.QueryParameters.Top = pageSize;
                    requestConfiguration.QueryParameters.Orderby = new[] { "createdDateTime asc" };

                    // Select only the fields you need (recommended for performance)
                    requestConfiguration.QueryParameters.Select = GraphHelper.SelectedEmailFields;

                    // Include attachments only when needed
                    requestConfiguration.QueryParameters.Expand = new string[] { includeAttachments ? "attachments" : "" };
                });
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingExchangeMessages}: {ex.Message}. {ex.InnerException?.Message}");
                return null;
            }
        }

        /// <summary>
        /// Retrieves the next page of emails using the provided next link from a previous response.
        /// </summary>
        /// <param name="nextLink">The URL to the next page of emails.</param>
        /// <returns>A <see cref="MessageCollectionResponse"/> containing the next page of emails, or null if an error occurs.</returns>
        public async Task<MessageCollectionResponse> GetEmailsFromNextLink(string nextLink)
        {
            try
            {
                return await graphServiceClient.RequestAdapter.SendAsync(
                    new RequestInformation
                    {
                        HttpMethod = Method.GET,
                        UrlTemplate = nextLink,
                        PathParameters = new Dictionary<string, object>() // required even if empty
                    },
                    MessageCollectionResponse.CreateFromDiscriminatorValue);
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingExchangeMessagesFromNextLink}: {ex.Message}. {ex.InnerException?.Message}");
                return null;
            }
        }

        /// <summary>
        /// Retrieves a list of mail folders for a specified user account.
        /// </summary>
        /// <param name="userAccount">The user principal name (UPN) or object ID of the user.</param>
        /// <returns>A list of <see cref="MailFolder"/> objects, or null if an error occurs.</returns>
        public async Task<List<MailFolder>> GetEmailFolders(string userAccount)
        {
            try
            {
                var folderResult = await graphServiceClient.Users[userAccount].MailFolders
                    .GetAsync(requestConfiguration =>
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
                logger.LogError($"{Resource.ErrorRetrievingDocuments}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
        }

        /// <summary>
        /// Retrieves a specific mail folder by its ID for a given user account.
        /// </summary>
        /// <param name="userAccount">The user principal name (UPN) or object ID of the user.</param>
        /// <param name="folderId">The unique identifier of the mail folder.</param>
        /// <returns>A <see cref="MailFolder"/> object if found, or null if an error occurs.</returns>
        public async Task<MailFolder> GetEmailFolderById(string userAccount, string folderId)
        {
            try
            {
                return await graphServiceClient.Users[userAccount].MailFolders[folderId].GetAsync();
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingExchangeFolders}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
        }

        /// <summary>
        /// Forwards a specific email message to another recipient.
        /// </summary>
        /// <param name="userAccount">The user principal name (UPN) or object ID of the user who owns the email.</param>
        /// <param name="emailId">The unique identifier of the email message to forward.</param>
        /// <param name="forwardAccount">The email address of the recipient to forward the email to.</param>
        /// <returns><c>true</c> if the email was forwarded successfully; otherwise, <c>false</c>.</returns>
        public async Task<bool> ForwardEmail(string userAccount, string emailId, string forwardAccount)
        {
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

            try
            {
                await graphServiceClient.Users[userAccount].Messages[emailId]
                    .Forward
                    .PostAsync(requestBody);

                return true;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorForwardingEmail}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return false;
        }

        /// <summary>
        /// Sends an email message from a specified user account.
        /// </summary>
        /// <param name="fromAccount">The user principal name (UPN) or object ID of the user sending the email.</param>
        /// <param name="toAccounts">A comma-separated string of email addresses of the recipients.</param>
        /// <param name="subject">The subject line of the email message.</param>
        /// <param name="body">The content of the email message.</param>
        /// <param name="attachments">An optional list of attachments to include in the email.</param>
        /// <returns><c>true</c> if the email was sent successfully; otherwise, <c>false</c>.</returns>
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

            try
            {
                await graphServiceClient.Users[fromAccount]
                        .SendMail
                        .PostAsync(sendMailBody);

                return true;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorSendingEmail}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return false;
        }

        /// <summary>
        /// Sends an email message from a specified user account.
        /// </summary>
        /// <param name="fromAccount">The user principal name (UPN) or object ID of the user sending the email.</param>
        /// <param name="toAccounts">A comma-separated string of email addresses of the recipients.</param>
        /// <param name="subject">The subject line of the email message.</param>
        /// <param name="body">The content of the email message.</param>
        /// <param name="attachments">An optional list of attachments to include in the email.</param>
        /// <returns><c>true</c> if the email was sent successfully; otherwise, <c>false</c>.</returns>
        public async Task<bool> MoveEmailToFolder(string userAccount, string emailId, string newFolder)
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
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorMovingEmail}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return false;
        }

        /// <summary>
        /// Delete email from user account
        /// https://learn.microsoft.com/en-us/graph/api/message-delete
        /// </summary>
        /// <param name="userAccount">Account(email address) containing the email to be deleted</param>
        /// <param name="emailId">Id of the email to be deleted</param>
        /// <returns><c>true</c> if the email was removed successfully; otherwise, <c>false</c>.</returns>
        public async Task<bool> RemoveEmail(string userAccount, string emailId)
        {
            try
            {
                await graphServiceClient.Users[userAccount].Messages[emailId].DeleteAsync();
                return true;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRemovingEmail}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return false;
        }

        /// <summary>
        /// Retrieves a list of Azure Active Directory group names that the user, associated with the provided access token, is a member of.
        /// </summary>
        /// <param name="azureAccessToken">The Azure AD access token used for authentication.</param>
        /// <returns>A list of strings representing the names of the Azure AD groups the user belongs to. Returns an empty list if the user is not a member of any Azure AD groups or if an error occurs.</returns>
        public async Task<List<string>> GetAzureAdGroupsFromAccessToken(string azureAccessToken)
        {
            var groups = new List<string>();

            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(logger, token: azureAccessToken));
                var graphServiceClient = new GraphServiceClient(authenticationProvider);

                var memberOfGroups = await graphServiceClient.Me.MemberOf.GetAsync();
                foreach (var group in memberOfGroups.Value)
                {
                    if (group is not Microsoft.Graph.Models.Group)
                    {
                        // Skip Groups that are not Azure AD Groups
                        // This applies to Windows AD Groups on hybrid environments
                        continue;
                    }

                    var groupName = ((Microsoft.Graph.Models.Group)group).DisplayName;
                    if (string.IsNullOrWhiteSpace(groupName))
                    {
                        // Fall back mechanism in case that DisplayName is not available
                        groupName = ((Microsoft.Graph.Models.Group)group).Id;
                    }

                    groups.Add(groupName);
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorRetrievingAzureAdGroups}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return groups;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Retrieves a list of drives (document libraries) for a specified SharePoint site.
        /// </summary>
        /// <param name="siteId">The unique identifier of the SharePoint site.</param>
        /// <param name="excludeSystemDrives">Optional. A boolean value indicating whether to exclude system drives. Defaults to <c>false</c>.</param>
        /// <returns>A list of <see cref="Drive"/> objects associated with the site, or null if the site is not found or an error occurs.</returns>
        private async Task<List<Drive>> GetSiteDrives(string siteId, bool excludeSystemDrives = false)
        {
            if (authConfig.AuthenticationMethod != AuthMethod.OAuthAppPermissions)
                return (await graphServiceClient.Sites[siteId].Drives.GetAsync()).Value;

            var site = await graphServiceClient.Sites[siteId].GetAsync();
            if (site == null)
                return null;

            var siteDocs = await GetSiteAndDocLibsDictionary(new List<Site> { site }, excludeSystemDrives);
            return siteDocs.FirstOrDefault().Value;
        }

        /// <summary>
        /// Saves the content of a Microsoft Graph <see cref="DriveItem"/> to a file on the local file system.
        /// </summary>
        /// <param name="document">The <see cref="DriveItem"/> object representing the file to save.</param>
        /// <param name="filePath">The desired file path (without timestamp) on the local file system where the file will be saved. A timestamp will be appended to the filename to ensure uniqueness.</param>
        /// <returns>The original <see cref="DriveItem"/> object if the save operation was successful; otherwise, <c>null</c>.</returns>
        private async Task<DriveItem> SaveDriveItemToFileSystem(DriveItem document, string filePath)
        {
            try
            {
                document.AdditionalData.TryGetValue("@microsoft.graph.downloadUrl", out var downloadUrl);
                var documentSize = document.Size;
                var readSize = ConstantHelper.DEFAULT_CHUNK_SIZE;

                using (FileStream outputFileStream = new FileStream(filePath.GetFilePathWithTimestamp(), FileMode.Create))
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
                }

                return document;
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorSavingDriveItem}: {ex.Message}. {ex.InnerException?.Message}");
            }

            return null;
        }

        /// <summary>
        /// Retrieves a dictionary of document libraries (drives) for a list of SharePoint sites.
        /// </summary>
        /// <param name="sites">A list of <see cref="Site"/> objects to retrieve document libraries for.</param>
        /// <param name="excludeSystemDocLibs">Optional. A boolean value indicating whether to exclude system document libraries. Defaults to <c>false</c>.</param>
        /// <returns>A dictionary where the key is the site's web URL and the value is a list of <see cref="Drive"/> objects (document libraries) for that site. Returns null if the input list of sites is null or empty, or if an error occurs.</returns>
        private async Task<Dictionary<string,List<Drive>>> GetSiteAndDocLibsDictionary(List<Site> sites, bool excludeSystemDocLibs = false)
        {
            if (sites == null || !sites.Any()) { return null; }

            try
            {
                var batchSize = 20;
                var index = 0;
                var siteDocsDictionary = new Dictionary<string, List<Drive>>();

                var batch = sites.Skip(index * batchSize).Take(batchSize).ToList();
                while (batch.Any())
                {
                    // Using GraphAPI batch calls to retrieve document libraries for each of the sites:
                    // https://learn.microsoft.com/en-us/graph/sdks/batch-requests
                    // https://learn.microsoft.com/en-us/graph/json-batching
                    var batchRequestContent = new BatchRequestContentCollection(graphServiceClient);
                    var requestList = new List<RequestInformation>();
                    var requestIdDictionary = new Dictionary<Site, string>();

                    foreach (var item in batch)
                    {
                        var request = graphServiceClient.Sites[item.Id].Drives.ToGetRequestInformation(requestConfiguration =>
                        {
                            requestConfiguration.QueryParameters.Select = excludeSystemDocLibs ? Array.Empty<string>() : new string[] { "id", "name", "system", "weburl" };
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
                            siteDocsDictionary.Add(item.Key.WebUrl, drives);
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

        /// <summary>
        /// Retrieves the root <see cref="Drive"/> object based on the connection type and drive ID.
        /// </summary>
        /// <param name="driveId">The unique identifier of the drive.</param>
        /// <returns>The root <see cref="Drive"/> object, or null if the connection type is invalid or an error occurs.</returns>
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

        /// <summary>
        /// Initializes the Microsoft Graph service client for application permissions using either a client secret or a certificate.
        /// </summary>
        /// <returns><c>true</c> if the initialization was successful; otherwise, <c>false</c>.</returns>
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

                    if (authConfig.UseChaosHander)
                    {
                        return InitializeWithChaosHandler(clientSecretCredential, scopes);
                    }
                    else
                    {
                        graphServiceClient = new GraphServiceClient(clientSecretCredential, scopes);
                    }
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
                logger.LogError($"{Resource.ErrorInitializingGraph}: {ex.Message}. {ex.InnerException?.Message}");
                return await Task.FromResult(false);
            }
            catch (Exception ex)
            {
                logger.LogError($"{Resource.ErrorInitializingGraph}: {ex.Message}. {ex.InnerException?.Message}");
                return await Task.FromResult(false);
            }

            return true;
        }

        /// <summary>
        /// Initializes the Microsoft Graph service client for delegated permissions using interactive browser authentication.
        /// </summary>
        /// <returns><c>true</c> if the initialization was successful; otherwise, <c>false</c>.</returns>
        private async Task<bool> InitializeDelegatedPermissions()
        {
            try
            {
                // Require only the  permissions listed in the application registration
                // https://learn.microsoft.com/en-us/entra/identity-platform/scopes-oidc#the-default-scope
                var graphScopes = new string[] { $"{authConfig.ApiUrl}.default" };

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

                if (authConfig.UseChaosHander)
                {
                    return InitializeWithChaosHandler(interactiveCredential, graphScopes);
                }
                else
                {
                    graphServiceClient = new GraphServiceClient(interactiveCredential, graphScopes);
                }
            }
            catch(Exception ex)
            {
                logger.LogError($"{Resource.ErrorInitializingGraph}: {ex.Message}. {ex.InnerException?.Message ?? ""}");
                return await Task.FromResult(false);
            }

            return true;
        }

        /// <summary>
        /// Initializes the Microsoft Graph service client with a Chaos Handler for simulating random server failures.
        /// We can use ChaosHandler to simulate server failures
        /// https://learn.microsoft.com/en-us/graph/sdks/customize-client?tabs=csharp
        /// https://camerondwyer.com/2021/09/23/how-to-use-the-microsoft-graph-sdk-chaos-handler-to-simulate-graph-api-errors/
        /// </summary>
        /// <param name="credential">The <see cref="TokenCredential"/> to use for authentication.</param>
        /// <param name="scopes">An array of permission scopes required for the application.</param>
        /// <returns><c>true</c> if the initialization was successful; otherwise, <c>false</c>.</returns>
        private bool InitializeWithChaosHandler(TokenCredential credential, string[] scopes)
        {
            try
            {
                // tokenCredential is one of the credential classes from Azure.Identity
                // scopes is an array of permission scope strings
                var authProvider = new AzureIdentityAuthenticationProvider(credential, scopes: scopes);

                var handlers = GraphClientFactory.CreateDefaultHandlers();

                // Remove the default Retry Handler
                var retryHandler = handlers.Where(h => h is RetryHandler).FirstOrDefault();
                handlers.Remove(retryHandler);

                // Add a new one ChaosHandler simulates random server failures
                // Microsoft.Kiota.Http.HttpClientLibrary.Middleware.ChaosHandler
                handlers.Add(new ChaosHandler(new ChaosHandlerOption()
                {
                    ChaosPercentLevel = authConfig.ChaosHandlerPercent
                }));

                var httpClient = GraphClientFactory.Create(handlers);
                graphServiceClient = new GraphServiceClient(httpClient, authProvider);
                return true;
            }
            catch(Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Initializes the Microsoft Graph service client using an access token stored in a file.
        /// </summary>
        /// <returns><c>true</c> if the initialization was successful; otherwise, <c>false</c>.</returns>
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

        /// <summary>
        /// Checks if the Microsoft Graph service client is initialized. If not, it attempts to initialize it based on the configured authentication method.
        /// </summary>
        /// <param name="forceInit">Optional. A boolean value indicating whether to force re-initialization even if the client is already initialized. Defaults to <c>false</c>.</param>
        /// <returns><c>true</c> if the client is initialized or successfully initialized; otherwise, <c>false</c>.</returns>
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

        /// <summary>
        /// Initializes the HttpClient used for Microsoft Graph API calls, potentially adding a Chaos Handler for simulating random failures.
        /// </summary>
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

        /// <summary>
        /// Determines whether to use a client secret or a certificate for application authentication based on the configuration.
        /// </summary>
        /// <returns><c>true</c> if a client secret is configured; <c>false</c> if a certificate is configured.</returns>
        /// <exception cref="Exception">Thrown if neither a client secret nor a certificate name is configured.</exception>
        private bool UseClientSecret()
        {
            if (!string.IsNullOrWhiteSpace(authConfig.ClientSecret))
                return true;
            else if (!string.IsNullOrWhiteSpace(authConfig.CertificateName))
                return false;
            else
                throw new Exception(Resource.ChooseClientOrCertificate);
        }

        /// <summary>
        /// Initializes an <see cref="AuthenticationToken"/> object by reading and potentially decrypting a token from a configured file path.
        /// </summary>
        /// <returns>An <see cref="AuthenticationToken"/> object if the token is successfully read and processed; otherwise, <c>null</c>.</returns>
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

        /// <summary>
        /// Deserializes a JSON string into an <see cref="AuthenticationToken"/> object.
        /// </summary>
        /// <param name="token">The JSON string representing the authentication token.</param>
        /// <returns>An <see cref="AuthenticationToken"/> object if deserialization is successful; otherwise, <c>null</c>.</returns>
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

        /// <summary>
        /// Retrieves connection information for a SharePoint drive, either from a cached list or by querying the Microsoft Graph API.
        /// </summary>
        /// <param name="driveId">The unique identifier of the SharePoint drive.</param>
        /// <returns>A <see cref="DriveConnectionInfo"/> object containing information about the drive, or null if the drive is not found or an error occurs.</returns>
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