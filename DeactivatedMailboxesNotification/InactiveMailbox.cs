using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActivateMailBoxes
{
    internal class InactiveMailbox
    {
        internal Guid? MailboxId { get; set; }
        internal string EmailAddress { get; set; }
        internal string UserName { get; set; }
    }
}

