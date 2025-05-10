using Coginov.GraphApi.Library.Models;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Coginov.GraphApi.Library.Helpers
{
    public static class GraphHelper
    {
        public static readonly string[] SelectedEmailFields = new[]
        {
            "id",                   // Unique message ID
            "internetMessageId",    // Unique internet message ID (for deduplication)
            "createdDateTime",      // Date and time the message was created
            "subject",              // Message subject
            "from",                 // Sender
            "toRecipients",         // Recipients
            "hasAttachments",       // Whether the message has attachments
            "body",                 // Full message body (if required)
            "ccRecipients",         // CC recipients (if needed)
            "bccRecipients",        // BCC recipients (if needed)
            "replyTo",              // Reply To (if needed)
            "importance",           // Importance of the message
            "receivedDateTime",     // Date and time message was received
            "sentDateTime",         // Date and time message was sent
            "parentFolderId"        // ID of the parent folder
        };

        public static Dictionary<string, object> GetFieldValues(this ListItem item)
        {
            return item.Fields.AdditionalData.ToDictionary(x => x.Key, x => x.Value);
        }

        public static string GetDriveItemId(this ListItem item)
        {
            return item.DriveItem.Id.Trim();// Fields.AdditionalData.ToDictionary(x => x.Key, x => x.Value);
        }

        public static string GetDriveId(this ListItem item)
        {
            return item.DriveItem.ParentReference.DriveId;// Fields.AdditionalData.ToDictionary(x => x.Key, x => x.Value);
        }

        public static DriveItemInfo GetDriveItemInfo(this ListItem item)
        {
            return new DriveItemInfo
            {
                DriveId = item.GetDriveId(),
                DriveItemId = item.GetDriveItemId()
            };
        }

        // Use this method to obtain the DriveItemInfo from a serialized list of fields (Folder)
        public static DriveItemInfo GetDriveItemInfoFromSerializedFields(this string serializedDictionary)
        {
            try
            {
                var fieldValues = JsonConvert.DeserializeObject<Dictionary<string, string>>(serializedDictionary);

                if (fieldValues.TryGetValue("DriveId", out var driveId) && fieldValues.TryGetValue("DriveItemId", out var driveItemId))
                {
                    return new DriveItemInfo
                    {
                        DriveId = driveId,
                        DriveItemId = driveItemId
                    };
                }
                return null;
            }
            catch(Exception) 
            {
                return null;
            }
        }
        
        // Use this method to obtain the DriveItemInfo from a serialized Graph DriveItem (File)
        public static DriveItemInfo GetDriveItemInfoFromSerializedDriveItem(this string serializedDriveItem)
        {
            try
            {
                var driveItem = JsonConvert.DeserializeObject<DriveItem>(serializedDriveItem);

                return new DriveItemInfo
                {
                    DriveId = driveItem.ParentReference.DriveId,
                    DriveItemId = driveItem.Id
                };
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static string GetCreatedByName(this ListItem item)
        {
            return item.DriveItem.CreatedBy.User.DisplayName;
        }

        public static string GetCreatedByEmail(this ListItem item)
        {
            return item.DriveItem.CreatedBy.User.AdditionalData.TryGetValue("email", out var createdByEmail) ? createdByEmail.ToString() : string.Empty;
        }

        public static string GetModifiedByName(this ListItem item)
        {
            return item.DriveItem.LastModifiedBy.User.DisplayName;
        }

        public static string GetModifiedByEmail(this ListItem item)
        {
            return item.DriveItem.LastModifiedBy.User.AdditionalData.TryGetValue("email", out var createdByEmail) ? createdByEmail.ToString() : string.Empty;
        }
    }
}
