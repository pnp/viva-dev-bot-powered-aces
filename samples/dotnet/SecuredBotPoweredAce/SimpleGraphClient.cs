using System.Collections.Generic;
using System.Net.Mail;
using System.Threading.Tasks;
using System;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Authentication;
using System.Threading;

namespace SecuredBotPoweredAce
{
    // This class is a wrapper for the Microsoft Graph API
    // See: https://developer.microsoft.com/en-us/graph
    public class SimpleGraphClient
    {
        private readonly string _token;

        public SimpleGraphClient(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new ArgumentNullException(nameof(token));
            }

            _token = token;
        }

        // Sends an email on the users behalf using the Microsoft Graph API
        public async Task SendMailAsync(string toAddress, string subject, string content)
        {
            if (string.IsNullOrWhiteSpace(toAddress))
            {
                throw new ArgumentNullException(nameof(toAddress));
            }

            if (string.IsNullOrWhiteSpace(subject))
            {
                throw new ArgumentNullException(nameof(subject));
            }

            if (string.IsNullOrWhiteSpace(content))
            {
                throw new ArgumentNullException(nameof(content));
            }

            var graphClient = GetAuthenticatedClient();
            var recipients = new List<Recipient>
            {
                new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = toAddress,
                    },
                },
            };

            // Create the message.
            var email = new Message
            {
                Body = new ItemBody
                {
                    Content = content,
                    ContentType = BodyType.Text,
                },
                Subject = subject,
                ToRecipients = recipients,
            };

            // Send the message.
            await graphClient.Me.SendMail.PostAsync(new Microsoft.Graph.Me.SendMail.SendMailPostRequestBody {
                Message = email
            });
        }

        // Gets the top emails for the user using the Microsoft Graph API
        public async Task<Message[]> GetRecentMailAsync(int top)
        {
            var graphClient = GetAuthenticatedClient();
            var messages = await graphClient.Me.Messages.GetAsync(a => a.QueryParameters.Top = top);
            return messages.Value.ToArray();
        }

        // Get information about the user.
        public async Task<User> GetMeAsync()
        {
            var graphClient = GetAuthenticatedClient();
            var me = await graphClient.Me.GetAsync();
            return me;
        }

        // Get an Authenticated Microsoft Graph client using the token issued to the user.
        private GraphServiceClient GetAuthenticatedClient()
        {
            var graphClient = new GraphServiceClient(new BaseBearerTokenAuthenticationProvider(new TokenProvider(_token)));
            return graphClient;
        }
    }

    public class TokenProvider : IAccessTokenProvider
    {
        private string _token { get; set; }

        public TokenProvider(string token)
        {
            _token = token;
        }

        public AllowedHostsValidator AllowedHostsValidator => throw new NotImplementedException();

        public Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object>? additionalAuthenticationContext = null, CancellationToken cancellationToken = default)
        {
            return Task.FromResult(_token);
        }
    }
}
