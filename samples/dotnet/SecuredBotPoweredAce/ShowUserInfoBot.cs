// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using AdaptiveCards;
using AdaptiveCards.Templating;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Builder.SharePoint;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.SharePoint;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SecuredBotPoweredAce
{
    public class ShowUserInfoBot : SharePointActivityHandler
    {
        private static string adaptiveCardExtensionId = Guid.NewGuid().ToString();
        private readonly string _connectionName;

        private readonly IConfiguration _configuration;
        private readonly ILogger<ShowUserInfoBot> _logger;

        private static ConcurrentDictionary<string, CardViewResponse> cardViews = new ConcurrentDictionary<string, CardViewResponse>();
        private static ConcurrentDictionary<string, QuickViewResponse> quickViews = new ConcurrentDictionary<string, QuickViewResponse>();

        private static string HomeCardView_ID = "HOME_CARD_VIEW";
        private static string ErrorCardView_ID = "ERROR_CARD_VIEW";
        private static string SignInCardView_ID = "SIGN_IN_CARD_VIEW";
        private static string SignInQuickView_ID = "SIGN_IN_QUICK_VIEW";
        private static string UserEmailsQuickView_ID = "USER_EMAILS_QUICK_VIEW";

        public ShowUserInfoBot(
            IConfiguration configuration,
            ILogger<ShowUserInfoBot> logger) : base()
        {
            this._configuration = configuration;
            this._connectionName = configuration["ConnectionName"];

            this._logger = logger;

            // ************************************
            // Add the CardViews
            // ************************************

            // Prepare ACE data for all Card Views
            var aceData = new AceData()
            {
                Title = "Your emails!",
                CardSize = AceData.AceCardSize.Large,
                DataVersion = "1.0",
                Id = adaptiveCardExtensionId
            };

            // Home Card View (Primary Text Card View)
            CardViewResponse homeCardViewResponse = new CardViewResponse();
            homeCardViewResponse.AceData = aceData;
            homeCardViewResponse.CardViewParameters = CardViewParameters.PrimaryTextCardViewParameters(
                new CardBarComponent()
                {
                    Id = "HomeCardView",
                },
                new CardTextComponent()
                {
                    Text = "Welcome!"
                },
                new CardTextComponent()
                {
                    Text = "You are: <upn>"
                },
                new List<BaseCardComponent>()
                {
                    new CardButtonComponent()
                    {
                        Id = "UserEmails",
                        Title = "Show emails",
                        Action = new QuickViewAction()
                        {
                            Parameters = new QuickViewActionParameters()
                            {
                                View = UserEmailsQuickView_ID
                            }
                        }
                    },
                    new CardButtonComponent()
                    {
                        Id = "SignOut",
                        Title = "Sign out",
                        Action = new SubmitAction()
                    }

                });
            homeCardViewResponse.ViewId = HomeCardView_ID;

            homeCardViewResponse.OnCardSelection = new QuickViewAction()
            {
                Parameters = new QuickViewActionParameters()
                {
                    View = UserEmailsQuickView_ID
                }
            };

            cardViews.TryAdd(homeCardViewResponse.ViewId, homeCardViewResponse);

            // SignIn Card View (Primary Text Card View)
            CardViewResponse signInCardViewResponse = new CardViewResponse();
            signInCardViewResponse.AceData = aceData;
            signInCardViewResponse.CardViewParameters = CardViewParameters.SignInCardViewParameters(
                new CardBarComponent()
                {
                    Id = "SignInCardView",
                },
                new CardTextComponent()
                {
                    Text = "User's Sign in!"
                },
                new CardTextComponent()
                {
                    Text = "Please, sign in ..."
                },
                new CardButtonComponent()
                {
                    Id = "CompleteSignInButton",
                    Title = "Complete sign in",
                    Action = new QuickViewAction()
                    {
                        Parameters = new QuickViewActionParameters()
                        {
                            View = SignInQuickView_ID
                        }
                    }
                }
                );
            signInCardViewResponse.ViewId = SignInCardView_ID;

            cardViews.TryAdd(signInCardViewResponse.ViewId, signInCardViewResponse);

            // Home Card View (Primary Text Card View)
            CardViewResponse errorCardViewResponse = new CardViewResponse();
            errorCardViewResponse.AceData = aceData;
            errorCardViewResponse.CardViewParameters = CardViewParameters.BasicCardViewParameters(
                new CardBarComponent()
                {
                    Id = "ErrorCardView",
                },
                new CardTextComponent()
                {
                    Text = "An error occurred!"
                },
                new List<BaseCardComponent>()
                {
                    new CardButtonComponent()
                    {
                        Id = "OkError",
                        Title = "Ok",
                        Action = new SubmitAction()
                        {
                            Parameters = new Dictionary<string, object>()
                            {
                                { "viewToNavigateTo", HomeCardView_ID }
                            }
                        }
                    }
                });
            errorCardViewResponse.ViewId = ErrorCardView_ID;

            cardViews.TryAdd(errorCardViewResponse.ViewId, errorCardViewResponse);

            // ************************************
            // Add the Quick Views
            // ************************************

            // Sign In Quick View
            QuickViewResponse signInQuickViewResponse = new QuickViewResponse();
            signInQuickViewResponse.Title = "Sign In";
            signInQuickViewResponse.Template = new AdaptiveCard("1.5");

            AdaptiveContainer signInContainer = new AdaptiveContainer();
            signInContainer.Separator = true;

            AdaptiveTextBlock signInTitleText = new AdaptiveTextBlock
            {
                Text = "Complete Sign In",
                Color = AdaptiveTextColor.Dark,
                Weight = AdaptiveTextWeight.Bolder,
                Size = AdaptiveTextSize.Medium,
                Wrap = true,
                MaxLines = 1,
                Spacing = AdaptiveSpacing.None
            };
            signInContainer.Items.Add(signInTitleText);

            AdaptiveTextBlock signInDescriptionText = new AdaptiveTextBlock
            {
                Text = "Input the magic code from signing into Azure Active Directory in order to continue.",
                Color = AdaptiveTextColor.Dark,
                Size = AdaptiveTextSize.Default,
                Wrap = true,
                MaxLines = 6,
                Spacing = AdaptiveSpacing.None
            };
            signInContainer.Items.Add(signInDescriptionText);

            AdaptiveNumberInput signInMagicCodeInputField = new AdaptiveNumberInput
            {
                Placeholder = "Enter Magic Code",
                Id = "magicCode",
                IsRequired = true
            };
            signInContainer.Items.Add(signInMagicCodeInputField);

            AdaptiveSubmitAction signInSubmitAction = new AdaptiveSubmitAction
            {
                Title = "Submit",
                Id = "SubmitMagicCode"
            };
            signInQuickViewResponse.Template.Actions.Add(signInSubmitAction);

            signInQuickViewResponse.Template.Body.Add(signInContainer);
            signInQuickViewResponse.ViewId = SignInQuickView_ID;

            quickViews.TryAdd(signInQuickViewResponse.ViewId, signInQuickViewResponse);

            // User's Emails Quick View
            QuickViewResponse userEmailQuickViewResponse = new QuickViewResponse();
            userEmailQuickViewResponse.Title = "Your email messages";
            userEmailQuickViewResponse.Template = new AdaptiveCard("1.5");

            AdaptiveContainer userEmailContainer = new AdaptiveContainer();
            userEmailContainer.Separator = true;

            AdaptiveTextBlock titleText = new AdaptiveTextBlock();
            titleText.Text = "Please, sign in to see your recent emails ...";
            titleText.Color = AdaptiveTextColor.Dark;
            titleText.Weight = AdaptiveTextWeight.Bolder;
            titleText.Size = AdaptiveTextSize.Large;
            titleText.Wrap = true;
            titleText.MaxLines = 1;
            titleText.Spacing = AdaptiveSpacing.None;
            userEmailContainer.Items.Add(titleText);

            userEmailQuickViewResponse.Template.Body.Add(userEmailContainer);

            userEmailQuickViewResponse.ViewId = UserEmailsQuickView_ID;

            quickViews.TryAdd(userEmailQuickViewResponse.ViewId, userEmailQuickViewResponse);
        }

        protected async override Task<CardViewResponse> OnSharePointTaskGetCardViewAsync(ITurnContext<IInvokeActivity> turnContext, AceRequest aceRequest, CancellationToken cancellationToken)
        {
            // Check to see if the user has already signed in
            var user = await GetAuthenticatedUser(null, turnContext, cancellationToken);
            if (user != null)
            {
                var homeCardView = cardViews[HomeCardView_ID];
                if (homeCardView != null)
                {
                    ((homeCardView.CardViewParameters.Body.ToList())[0] as CardTextComponent).Text = $"You are: {user.UserPrincipalName}";
                    return homeCardView;
                }
            }
            else
            {
                var signInCardView = cardViews[SignInCardView_ID];
                if (signInCardView != null)
                {
                    var signInResource = await GetSignInResource(turnContext, cancellationToken);
                    var signInLink = signInResource != null ? new Uri(signInResource.SignInLink) : new Uri(string.Empty);

                    signInCardView.AceData.Properties = Newtonsoft.Json.Linq.JObject.FromObject(new Dictionary<string, object>() {
                        { "uri", signInLink },
                        { "connectionName", _connectionName}
                    });
                    return signInCardView;
                }
            }

            return cardViews[ErrorCardView_ID];
        }

        protected async override Task<QuickViewResponse> OnSharePointTaskGetQuickViewAsync(ITurnContext<IInvokeActivity> turnContext, AceRequest aceRequest, CancellationToken cancellationToken)
        {
            var nextQuickViewId = ((JObject)aceRequest.Data)["viewId"].ToString();
            if (nextQuickViewId == null || nextQuickViewId == UserEmailsQuickView_ID)
            {
                var emails = await GetAuthenticatedUserRecentEmails(null, turnContext, cancellationToken);
                if (emails != null)
                {
                    var result = quickViews[UserEmailsQuickView_ID];
                    result.Template = createDynamicAdaptiveCard(
                        readQuickViewJson("EmailsQuickView.json"),
                        new
                        {
                            Title = "Here are your recently received emails:",
                            Emails = (from e in emails
                                     select new { 
                                         Date = e.ReceivedDateTime.Value,
                                         From = e.From.EmailAddress.Name,
                                         e.Subject 
                                     }).ToArray()
                        });
                    return result;
                }
            }
            return quickViews[nextQuickViewId];
        }

        protected async override Task<BaseHandleActionResponse> OnSharePointTaskHandleActionAsync(ITurnContext<IInvokeActivity> turnContext, AceRequest aceRequest, CancellationToken cancellationToken)
        {
            if (turnContext != null)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }
            }
            JObject actionParameters = aceRequest.Data as JObject;

            if (actionParameters != null)
            {
                var actionId = actionParameters["id"].ToString();
                if (actionId == "SubmitMagicCode")
                {
                    var magicCode = actionParameters["data"]["magicCode"].ToString();
                    var user = await GetAuthenticatedUser(magicCode, turnContext, cancellationToken);

                    var homeCardView = cardViews[HomeCardView_ID];
                    if (homeCardView != null && user != null)
                    {
                        ((homeCardView.CardViewParameters.Body.ToList())[0] as CardTextComponent).Text = $"You are: {user.UserPrincipalName}";

                        return new CardViewHandleActionResponse
                        {
                            RenderArguments = homeCardView
                        };
                    }
                }
                else if (actionId == "SignOut")
                {
                    await SignOutUser(turnContext, cancellationToken);

                    var signInCardView = cardViews[SignInCardView_ID];
                    if (signInCardView != null)
                    {
                        var signInResource = await GetSignInResource(turnContext, cancellationToken);
                        var signInLink = signInResource != null ? new Uri(signInResource.SignInLink) : new Uri(string.Empty);

                        signInCardView.AceData.Properties = Newtonsoft.Json.Linq.JObject.FromObject(new Dictionary<string, object>() {
                            { "uri", signInLink },
                            { "connectionName", _connectionName}
                        });

                        return new CardViewHandleActionResponse
                        {
                            RenderArguments = signInCardView
                        };
                    }
                }
                else if (actionId == "OkError")
                {
                    return new CardViewHandleActionResponse
                    {
                        RenderArguments = cardViews[HomeCardView_ID]
                    };
                }
            }

            return new CardViewHandleActionResponse
            {
                RenderArguments = cardViews[ErrorCardView_ID]
            };
        }

        private async Task<Microsoft.Graph.Models.User> GetAuthenticatedUser(string magicCode, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                var response = await GetUserToken(magicCode, turnContext, cancellationToken).ConfigureAwait(false);
                if (response != null && !string.IsNullOrEmpty(response.Token))
                {
                    var client = new SimpleGraphClient(response.Token);
                    return await client.GetMeAsync().ConfigureAwait(false);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error while trying to retrieve current user's UPN");
            }

            return null;
        }

        private async Task<Microsoft.Graph.Models.Message[]> GetAuthenticatedUserRecentEmails(string magicCode, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                var response = await GetUserToken(magicCode, turnContext, cancellationToken).ConfigureAwait(false);
                if (response != null && !string.IsNullOrEmpty(response.Token))
                {
                    var client = new SimpleGraphClient(response.Token);
                    return await client.GetRecentMailAsync(10);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error while trying to retrieve current user's UPN");
            }

            return null;
        }


        private async Task<SignInResource> GetSignInResource(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            // Get the UserTokenClient service instance
            var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();

            // Retrieve the Sign In Resource from the UserTokenClient service instance
            var signInResource = await userTokenClient.GetSignInResourceAsync(_connectionName, (Microsoft.Bot.Schema.Activity)turnContext.Activity, null, cancellationToken).ConfigureAwait(false);
            return signInResource;
        }

        private async Task<TokenResponse> GetUserToken(string magicCode, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            // Get the UserTokenClient service instance
            var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();

            // Assuming the bot is already configured for SSO and the user has been authenticated
            return await userTokenClient.GetUserTokenAsync(
                turnContext.Activity.From.Id,
                _connectionName, // The name of your Azure AD connection
                turnContext.Activity.ChannelId,
                magicCode,
                cancellationToken).ConfigureAwait(false);
        }

        private async Task SignOutUser(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            // Get the UserTokenClient service instance
            var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();

            // Sign out the current user
            await userTokenClient.SignOutUserAsync(
                turnContext.Activity.From.Id,
                _connectionName, // The name of your Azure AD connection
                turnContext.Activity.ChannelId,
                cancellationToken).ConfigureAwait(false);
        }

        private string readQuickViewJson(string quickViewTemplateFileName)
        {
            string json = null;

            var templatesPath = _configuration["TemplatesPath"];
            var quickViewTemplatePath = Path.Combine(templatesPath, quickViewTemplateFileName);

            using (StreamReader sr = new StreamReader(quickViewTemplatePath))
            {
                json = sr.ReadToEnd();
            }

            return json;
        }

        private AdaptiveCard createDynamicAdaptiveCard(string cardJson, object dataSource)
        {
            AdaptiveCardTemplate template = new AdaptiveCardTemplate(cardJson);
            var cardJsonWithData = template.Expand(dataSource);

            // Deserialize the JSON string into an AdaptiveCard object
            AdaptiveCardParseResult parseResult = AdaptiveCard.FromJson(cardJsonWithData);

            // Check for errors during parsing
            if (parseResult.Warnings.Count > 0)
            {
                Trace.Write("Warnings during parsing:");
                foreach (var warning in parseResult.Warnings)
                {
                    Trace.Write(warning.Message);
                }
            }

            // Get the AdaptiveCard object
            AdaptiveCard card = parseResult.Card;

            return card;
        }
    }
}
