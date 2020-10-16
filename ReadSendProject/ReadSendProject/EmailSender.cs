using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using ReadSendProject.Logging;
using System;
using StringMatrix = System.Collections.Generic.List<System.Collections.Generic.List<string>>;

namespace ReadSendProject
{
    class EmailSender
    {
        public ILogger logger;
        private readonly EmailSenderParameters emailParams;

        public EmailSender(EmailSenderParameters parameters)
        {
            emailParams = parameters;
        }

        private bool TryBuildMessage(StringMatrix data, out MimeMessage message)
        {
            try
            {
                message = new MimeMessage();
                message.From.Add(MailboxAddress.Parse(emailParams.senderEmail));
                message.To.AddRange(emailParams.receiverEmails.ConvertAll((email) => MailboxAddress.Parse(email)));
                message.Subject = "Notificare expirare acte masini";

                var bodyBuilder = new BodyBuilder();

                bodyBuilder.TextBody = "\tBuna ziua,\n"
                    + "\n"
                    + "Urmatoarele masini au date de expirare in viitorul apropiat:\n"
                    + "\n"
                    + "\n"
                    + StringMatrixConverter.ToPlainTextTable(data)
                    + "\n"
                    + "\n"
                    + "O zi frumoasa!\n";

                bodyBuilder.HtmlBody = "<p>Buna ziua,<br>"
                    + "<br>"
                    + "<p>Urmatoarele masini au date de expirare in viitorul apropiat:<br>"
                    + "<br>"
                    + StringMatrixConverter.ToHtmlTable(data)
                    + "<br>"
                    + "<p>O zi frumoasa!<br>";

                message.Body = bodyBuilder.ToMessageBody();
                return true;
            }
            catch (ParseException ex)
            {
                logger?.LogE($"ParseException in TryBuildMessage: {ex.Message}");
                message = null;
                return false;
            }
            catch (Exception ex)
            {
                logger?.LogE($"Exception in TryBuildMessage: {ex.Message}");
                message = null;
                return false;
            }
        }

        private bool TrySmtpConnect(SmtpClient client)
        {
            try
            {
                var host = "smtp.gmail.com";
                var port = 465;
                client.Connect(host, port, SecureSocketOptions.SslOnConnect);
                logger?.LogI("Smtp Connected");
                return true;
            }
            catch (SmtpCommandException ex)
            {
                logger?.LogE($"SMPT Command error while trying to connect; Status code: {ex.StatusCode}\t\tMessage: {ex.Message}");
                return false;
            }
            catch (SmtpProtocolException ex)
            {
                logger?.LogE($"SMTP Protocol error while trying to connect; Message: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                logger?.LogE($"Exception in TrySmtpConnect: {ex.Message}");
                return false;
            }
        }

        private bool TrySmtpAuthenticate(SmtpClient client)
        {
            try
            {
                if (client.Capabilities.HasFlag(SmtpCapabilities.Authentication))
                {
                    // TODO: try to use some system.net.credentials or encryption for pass
                    var username = emailParams.senderEmail;
                    var password = emailParams.senderEmailPassword;
                    client.Authenticate(username, password);
                    logger?.LogI("Smtp authenticated.");
                }

                return true;
            }
            catch (AuthenticationException)
            {
                logger?.LogE("Authentication error: Invalid user name or password.");
                return false;
            }
            catch (SmtpCommandException ex)
            {
                logger?.LogE($"SMPT Command error while trying to authenticate; Status code: {ex.StatusCode}\t\tMessage: {ex.Message}");
                return false;
            }
            catch (SmtpProtocolException ex)
            {
                logger?.LogE($"SMTP Protocol error while trying to authenticate; Message: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                logger?.LogE($"Exception in TrySmtpAuthenticate: {ex.Message}");
                return false;
            }
        }

        private bool TrySmtpSend(SmtpClient client, MimeMessage message)
        {
            try
            {
                client.Send(message);
                logger?.LogI("Smtp email sent");
                return true;
            }
            catch (SmtpCommandException ex)
            {
                logger?.LogE($"SMPT Command error while sending message; Status code: {ex.StatusCode}");

                switch (ex.ErrorCode)
                {
                    case SmtpErrorCode.RecipientNotAccepted:
                        logger?.LogE("\tRecipient not accepted: {ex.Mailbox}");
                        break;
                    case SmtpErrorCode.SenderNotAccepted:
                        logger?.LogE("\tSender not accepted: {ex.Mailbox}");
                        break;
                    case SmtpErrorCode.MessageNotAccepted:
                        logger?.LogE("\tMessage not accepted.");
                        break;
                    case SmtpErrorCode.UnexpectedStatusCode:
                        logger?.LogE("\tUnexpected Status Code.");
                        break;
                    default:
                        logger?.LogE("\tOther Status Code ??");
                        break;
                }

                return false;
            }
            catch (SmtpProtocolException ex)
            {
                logger?.LogE($"SMTP Protocol error while sending message; Message: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                logger?.LogE($"Exception in TrySmtpSend: {ex.Message}");
                return false;
            }
        }

        private bool TrySmtpDisconnect(SmtpClient client)
        {
            try
            {
                client.Disconnect(quit: true);
                return true;
            }
            catch (Exception ex)
            {
                logger?.LogE($"Exception in TrySmtpDisconnect: {ex.Message}");
                return false;
            }
        }


        public bool TrySendEmail(StringMatrix data)
        {
            // https://blog.elmah.io/how-to-send-emails-from-csharp-net-the-definitive-tutorial/
            // https://github.com/jstedfast/MailKit/blob/master/Documentation/Examples/SmtpExamples.cs
            // https://github.com/jstedfast/MailKit/blob/master/Documentation/Examples/BodyBuilder.cs

            using (var client = new SmtpClient())
            {
                return TryBuildMessage(data, out MimeMessage message)
                    && TrySmtpConnect(client)
                    && TrySmtpAuthenticate(client)
                    && TrySmtpSend(client, message)
                    && TrySmtpDisconnect(client);
            }
        }
    }
}
