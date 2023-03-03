using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Coginov.GraphApi.Library.Helpers
{
    public static class EmailHelper
    {
        public static string ToStringFormat(this EmailAddress email)
        {
            return $"{email.Name} <{email.Address}>";
        }

        public static string ToStringFormat(this IEnumerable<Recipient> emails)
        {
            return String.Join(';', emails.Select(x => x.EmailAddress.ToStringFormat()));
        }

        public static List<Recipient> ParseRecipients(this string recipients)
        {
            recipients.Replace(';', ',');
            try
            {
                return recipients.Split(',').Select(to =>
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = to.Trim()
                            }
                        }
                    ).ToList();

            }
            catch(Exception)
            {
                return new List<Recipient>();
            }
        }

        public static Attachment CreateAttachment(string fileName, byte[] fileBytes)
        {
            return new FileAttachment
            {
                ODataType = "#microsoft.graph.fileAttachment",
                Name = fileName,
                ContentBytes = fileBytes
            };
        }
    }
}
