using Coginov.GraphApi.Library.Models;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Coginov.GraphApi.Library.Services
{
    public interface IGraphApiService
    {
        // Methods mostly used by QoreAudit
        Task<bool> InitializeSharePointOnlineConnection(AuthenticationConfig authenticationConfig, string siteUrl, string[] docLibraries, bool forceInit = false);
        Task<bool> InitializeOneDriveConnection(AuthenticationConfig authenticationConfig, string userAccount, bool forceInit = false);
        Task<bool> InitializeMsTeamsConnection(AuthenticationConfig authenticationConfig, string[]? teams, bool forceInit = false);
        Task<List<DriveConnectionInfo>> GetSharePointOnlineDrives();
        Task<List<DriveConnectionInfo>> GetOneDriveDrives();
        Task<List<DriveConnectionInfo>> GetMsTeamDrives();
        Task<DriveItemSearchResult> GetDocumentIds(string driveId, DateTime lastDate, int top, string skipToken);
        Task<DriveItem> SaveDriveItemToFileSystem(string driveId, string documentId, string downloadLocation);
        Task<DriveItem> GetDriveItemMetadata(string driveId, string documentId);
        Task<Dictionary<string, List<string>>> GetSharepointSitesAndDocLibs(bool excludePersonalSites = false, bool excludeSystemDocLibs = false);
        Task<List<string>> GetAzureAdGroupsFromAccessToken(string azureAccessToken);

        // Methods used by QoreAudit and QoreMail
        Task<bool> InitializeExchangeConnection(AuthenticationConfig authenticationConfig, bool forceInit = false);
        Task<MessageCollectionResponse> GetEmailsAfterDate(string userAccount, DateTime afterDate, int skipIndex = 0, int emailCount = 10, bool includeAttachments = false);
        Task<MessageCollectionResponse> GetEmailsFromFolderAfterDate(string userAccount, string Folder, DateTime afterDate, int skipIndex = 0, int emailCount = 10, bool includeAttachments = false, bool preferText = false);
        Task<bool> SaveEmailToFileSystem(Message message, string downloadLocation, string userAccount, string fileName);
        Task<MailFolder> GetEmailFolderById(string userAccount, string folderId);
        Task<List<MailFolder>> GetEmailFolders(string userAccount);


        // Methods not currently in use. Developed for future features and needs 
        Task<bool> ForwardEmail(string userAccount, string emailId, string forwardAccount);
        Task<bool> MoveEmailToFolder(string userAccount, string emailId, string newFolder);
        Task<bool> RemoveEmail(string userAccount, string emailId);
        Task<bool> SendEmail(string fromAccount, string toAccounts, string subject, string body, List<Attachment> attachments = null);
        Task<int?> GetInboxMessageCount(string userAccount);
        Task<bool> UploadDocumentToDrive(string driveId, string filePath, string fileName = null, string folderPath = "", string onConflict = "replace");
        Task<bool> DeleteDocumentById(string driveId, string documentId);
        Task<bool> DeleteDocumentByPath(string driveId, string documentPath);
        Task<bool> MoveDocument(string driveId, string documentId, string destFolderId = null, string destFolder = null, string docNewName = null);
        Task<bool> MoveDocument(string driveId, string documentId, string destSite, string destDocLib, string destFolderId = null, string destFolder = null, string docNewName = null);

        // Methods currently only used by the QoreImport tool. Could be used by other projects in the future
        Task<bool> InitializeGraphApi(AuthenticationConfig authenticationConfig, bool forceInit = true);
        Task<List<ListItem>> SearchSharepointOnlineFolders(string siteUrl, string docLibrary, string searchField = null, string searchValue = null, string searchFilter = null, int top = 200);
        Task<Dictionary<ListItem, string>> UpdateSharePointOnlineItemFieldValue(List<ListItem> items, Dictionary<string, object> columnKeyValues);
        Task<Dictionary<ListItem, string>> UpdateSharePointOnlineItemFieldValue(List<DriveItemInfo> items, Dictionary<string, object> columnKeyValues);
        Task<List<DriveItem>> GetListOfFilesInFolder(DriveItemInfo driveItem, DateTimeOffset? lastDate = null, int batchSize = 100);
        Task<string> GetTokenApplicationPermissions(string tenantId, string clientId, string clientSecret, string[] scopes);
        Task<string> GetTokenDelegatedPermissions(string tenantId, string clientId, string[] scopes, CancellationToken cancellationToken = default); // Also used by QoreAudit
    }
}