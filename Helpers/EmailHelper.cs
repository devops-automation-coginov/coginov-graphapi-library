using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
