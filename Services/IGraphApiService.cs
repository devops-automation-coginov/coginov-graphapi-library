using Coginov.GraphApi.Library.Models;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Coginov.GraphApi.Library.Services
{
    public interface IGraphApiService
    {
        Task<bool> InitializeSharePointOnlineConnection(AuthenticationConfig authenticationConfig, string siteUrl, string[] docLibraries);
        Task<bool> InitializeOneDriveConnection(AuthenticationConfig authenticationConfig, string userAccount);
        Task<bool> InitializeMsTeamsConnection(AuthenticationConfig authenticationConfig, params string[]? teams);
        Task<bool> InitializeExchangeConnection(AuthenticationConfig authenticationConfig);
        Task<List<DriveConnectionInfo>> GetSharePointOnlineDrives();
        Task<List<DriveConnectionInfo>> GetOneDriveDrives();
        Task<List<DriveConnectionInfo>> GetMsTeamDrives();
        Task<List<string>> GetDocumentIds(string driveId, DateTime lastDate, int skip, int top);
        Task<DriveItemSearchResult> GetDocumentIds(string driveId, DateTime lastDate, int top, string skipToken);
        Task<DriveItem> SaveDriveItemToFileSystem(string driveId, string documentId, string downloadLocation);
        Task<IUserMessagesCollectionPage> GetEmailsAfterDate(string userAccount, DateTime afterDate, int skipIndex = 0, int emailCount = 10, bool includeAttachments = false);
        Task<IMailFolderMessagesCollectionPage> GetEmailsFromFolderAfterDate(string userAccount, string Folder, DateTime afterDate, int skipIndex = 0, int emailCount = 10, bool includeAttachments = false, bool preferText = false);
        Task<bool> SaveEmailToFileSystem(Message message, string downloadLocation, string userAccount, string fileName);
        Task<MailFolder> GetEmailFolderById(string userAccount, string folderId);
        Task<List<MailFolder>> GetEmailFolders(string userAccount);
        Task<bool> ForwardEmail(string userAccount, string emailId, string forwardAccount);
        Task<bool> MoveEmailToFolder(string userAccount, string emailId, string newFolder);
        Task<bool> SendEmail(string fromAccount, string toAccounts, string subject, string body, List<Attachment> attachments = null);
    }
}