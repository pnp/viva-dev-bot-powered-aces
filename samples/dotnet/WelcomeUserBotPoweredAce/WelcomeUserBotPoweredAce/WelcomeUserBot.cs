// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.SharePoint;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.SharePoint;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Collections.Concurrent;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using AdaptiveCards;
using Newtonsoft.Json.Linq;
using System.Linq;
using Microsoft.Bot.Connector.Authentication;

namespace WelcomeUserBotPoweredAce
{
    public class WelcomeUserBot : SharePointActivityHandler
    {
        private static string adaptiveCardExtensionId = Guid.NewGuid().ToString();
        private readonly string _connectionName;

        private readonly IConfiguration _configuration;
        private readonly ILogger<WelcomeUserBot> _logger;

        private static ConcurrentDictionary<string, CardViewResponse> cardViews = new ConcurrentDictionary<string, CardViewResponse>();
        private static ConcurrentDictionary<string, QuickViewResponse> quickViews = new ConcurrentDictionary<string, QuickViewResponse>();

        private static string HomeCardView_ID = "HOME_CARD_VIEW";
        private static string ErrorCardView_ID = "ERROR_CARD_VIEW";
        private static string SignInCardView_ID = "SIGN_IN_CARD_VIEW";
        private static string SignedOutCardView_ID = "SIGNED_OUT_CARD_VIEW";
        private static string SignInQuickView_ID = "SIGN_IN_QUICK_VIEW";

        public WelcomeUserBot(
            IConfiguration configuration,
            ILogger<WelcomeUserBot> logger) : base()
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
                Title = "Welcome!",
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
                    Text = "Welcome <displayName>!"
                },
                new CardTextComponent()
                {
                    Text = "Your UPN is: <upn>"
                },
                new List<BaseCardComponent>()
                {
                    new CardButtonComponent()
                    {
                        Id = "SignOut",
                        Title = "Sign out",
                        Action = new SubmitAction()
                    }

                });
            homeCardViewResponse.ViewId = HomeCardView_ID;

            cardViews.TryAdd(homeCardViewResponse.ViewId, homeCardViewResponse);

            // SignIn Card View (Sign In Card View)
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
            signInCardViewResponse.CardViewParameters.CardViewType = "signIn";
            signInCardViewResponse.ViewId = SignInCardView_ID;

            cardViews.TryAdd(signInCardViewResponse.ViewId, signInCardViewResponse);

            // Signed out Card View (Basic Card View)
            CardViewResponse signedOutCardViewResponse = new CardViewResponse();
            signedOutCardViewResponse.AceData = aceData;
            signedOutCardViewResponse.CardViewParameters = CardViewParameters.BasicCardViewParameters(
                new CardBarComponent()
                {
                    Id = "SignedOutCardView",
                },
                new CardTextComponent()
                {
                    Text = "You are now signed out!"
                },
                new List<BaseCardComponent>()
                {
                    new CardButtonComponent()
                    {
                        Id = "OkSignedOut",
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
            signedOutCardViewResponse.ViewId = SignedOutCardView_ID;

            cardViews.TryAdd(signedOutCardViewResponse.ViewId, signedOutCardViewResponse);

            // Error Card View (Basic Card View)
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

            // Complete Sign In Quick View
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
                Text = "Input the magic code from Microsoft Entra ID to complete sign in.",
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
                Id = "magicCode"
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
        }

        protected async override Task<CardViewResponse> OnSharePointTaskGetCardViewAsync(ITurnContext<IInvokeActivity> turnContext, AceRequest aceRequest, CancellationToken cancellationToken)
        {
            JObject aceData = aceRequest.Data as JObject;
            string magicCode = null;
            if (aceData["magicCode"] != null)
            {
                magicCode = aceData["magicCode"].ToString();
            }
            // Check to see if the user has already signed in
            var (displayName, upn) = await GetAuthenticatedUser(magicCode, turnContext, cancellationToken);
            if (displayName != null && upn != null)
            {
                var homeCardView = cardViews[HomeCardView_ID];
                if (homeCardView != null)
                {
                    ((homeCardView.CardViewParameters.Header.ToList())[0] as CardTextComponent).Text = $"Welcome {displayName}!";
                    ((homeCardView.CardViewParameters.Body.ToList())[0] as CardTextComponent).Text = $"Your UPN is: {upn}";
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
                        { "connectionName", this._connectionName }
                    });
                    return signInCardView;
                }
            }

            return cardViews[ErrorCardView_ID];
        }

        protected async override Task<QuickViewResponse> OnSharePointTaskGetQuickViewAsync(ITurnContext<IInvokeActivity> turnContext, AceRequest aceRequest, CancellationToken cancellationToken)
        {
            var nextQuickViewId = ((JObject)aceRequest.Data)["viewId"].ToString();
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
                    var (displayName, upn) = await GetAuthenticatedUser(magicCode, turnContext, cancellationToken);

                    var homeCardView = cardViews[HomeCardView_ID];
                    if (homeCardView != null && displayName != null && upn != null)
                    {
                        ((homeCardView.CardViewParameters.Header.ToList())[0] as CardTextComponent).Text = $"Welcome {displayName}!";
                        ((homeCardView.CardViewParameters.Body.ToList())[0] as CardTextComponent).Text = $"Your UPN is: {upn}";

                        return new CardViewHandleActionResponse
                        {
                            RenderArguments = homeCardView
                        };
                    }
                }
                else if (actionId == "SignOut")
                {
                    await SignOutUser(turnContext, cancellationToken);

                    return new CardViewHandleActionResponse
                    {
                        RenderArguments = cardViews[SignedOutCardView_ID]
                    };
                }
                else if (actionId == "OkSignedOut")
                {
                    return new CardViewHandleActionResponse
                    {
                        RenderArguments = cardViews[SignInCardView_ID]
                    };
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

        private async Task<(string displayName, string upn)> GetAuthenticatedUser(string magicCode, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            string displayName = null;
            string upn = null;

            try
            {
                var response = await GetUserToken(magicCode, turnContext, cancellationToken).ConfigureAwait(false);
                if (response != null && !string.IsNullOrEmpty(response.Token))
                {
                    var token = new System.IdentityModel.Tokens.Jwt.JwtSecurityToken(response.Token);
                    displayName = token.Claims.FirstOrDefault(c => c.Type == System.IdentityModel.Tokens.Jwt.JwtRegisteredClaimNames.Name)?.Value;
                    upn = token.Claims.FirstOrDefault(c => c.Type == "upn")?.Value;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error while trying to retrieve current user's displayName and UPN!");
            }

            return (displayName, upn);
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
    }
}
