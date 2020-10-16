using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadSendProject
{
    struct EmailSenderParameters
    {
        /// <summary>
        /// Valid email address which will send the emails and be used for authentication.
        /// </summary>
        public string senderEmail;

        /// <summary>
        /// TODO: wrap it in a encrypted retrieval method?
        /// https://stackoverflow.com/questions/4155187/securing-a-password-in-source-code
        /// </summary>
        public string senderEmailPassword;

        /// <summary>
        /// List of all the people interested in the email.
        /// </summary>
        public List<string> receiverEmails;
    }
}
