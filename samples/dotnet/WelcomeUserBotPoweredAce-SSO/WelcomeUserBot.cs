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
        private readonly IStorage _storage;

        private static ConcurrentDictionary<string, CardViewResponse> cardViews = new ConcurrentDictionary<string, CardViewResponse>();

        private static string HomeCardView_ID = "HOME_CARD_VIEW";
        private static string ErrorCardView_ID = "ERROR_CARD_VIEW";
        private static string SignInCardView_ID = "SIGN_IN_CARD_VIEW";
        private static string SignedOutCardView_ID = "SIGNED_OUT_CARD_VIEW";

        public WelcomeUserBot(
            IConfiguration configuration,
            IStorage storage,
            ILogger<WelcomeUserBot> logger) : base()
        {
            this._configuration = configuration;
            this._connectionName = configuration["ConnectionName"];

            this._logger = logger;
            this._storage = storage;

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
            );
            signInCardViewResponse.CardViewParameters.CardViewType = "signInSso";
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
        }

        protected async override Task<CardViewResponse> OnSharePointTaskGetCardViewAsync(ITurnContext<IInvokeActivity> turnContext, AceRequest aceRequest, CancellationToken cancellationToken)
        {
            // Check to see if the user has already signed in
            var (displayName, upn) = await GetAuthenticatedUser(magicCode: null, turnContext, cancellationToken);
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
                if (actionId == "SignOut")
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

        protected override Task OnSignInInvokeAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            SharePointSSOTokenExchangeMiddleware sso = new SharePointSSOTokenExchangeMiddleware(_storage, _connectionName);
            return sso.OnTurnAsync(turnContext, cancellationToken);
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
