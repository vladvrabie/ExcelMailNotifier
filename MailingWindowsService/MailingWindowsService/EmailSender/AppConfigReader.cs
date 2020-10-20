using MailingWindowsService.Logging;
using System;
using System.Configuration;
using System.Linq;

namespace MailingWindowsService.EmailSender
{
    class AppConfigReader
    {
        public ILogger logger;

        private static readonly string SENDER_EMAIL = "senderEmail";
        private static readonly string SENDER_PASSWORD = "senderEmailPassword";
        private static readonly string RECEIVER_EMAILS = "receiverEmails";


        public EmailSenderParameters GetEmailSenderParameters()
        {
            try
            {
                var appSettings = ConfigurationManager.AppSettings;
                if (appSettings.Count == 0)
                {
                    throw new ConfigurationErrorsException();
                }

                return new EmailSenderParameters()
                {
                    senderEmail = appSettings[SENDER_EMAIL],
                    senderEmailPassword = appSettings[SENDER_PASSWORD],
                    receiverEmails = appSettings[RECEIVER_EMAILS].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList(),
                };
            }
            catch (Exception ex)
            {
                logger?.LogE($"Error in retrieving email sender parameters; Message: {ex.Message}");
                return new EmailSenderParameters();
            }
        }
    }
}
