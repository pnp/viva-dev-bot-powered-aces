import { 
  SharePointActivityHandler, TurnContext, AceRequest, AceData,
  CardViewResponse, QuickViewResponse, HandleActionResponse, CardViewHandleActionResponse,
  // TextInputCardView, 
  ImageCardView  
} from "botbuilder";
import * as AdaptiveCards from "adaptivecards";

export class CollectFeedbackBot extends SharePointActivityHandler {

  private readonly _botId: string = 'CollectFeedbackBot';
  private _cardViews: { [key: string]: CardViewResponse } = {};
  // private _quickViews: { [key: string]: QuickViewResponse } = {};

  private COLLECT_FEEDBACK_CARD_VIEW_ID: string = 'GET_FEEDBACK_CARD_VIEW';
  private OK_FEEDBACK_CARD_VIEW_ID: string = 'OK_FEEDBACK_CARD_VIEW';
  private SHOW_FEEDBACK_QUICK_VIEW_ID: string = 'FEEDBACK_QUICK_VIEW';

  /**
   * Build the ACE card views and quick views for the Bot Powered ACE.
   */
  constructor() {
    super();
    
    // Prepare the ACE data for all the card views and quick views.
    const aceData: AceData = {
      id: this._botId,
      title: 'Your feedback matters!',
      description: 'Please provide your feedback below.',
      cardSize: 'Large',
      iconProperty: 'Feedback',
      properties: {},
      dataVersion: '1.0',
    }; 

    // Collect Feedback Card View (Input Text Card View manual)
    const feedbackCardView: CardViewResponse = {
      aceData: aceData,
      cardViewParameters: {
        cardViewType: 'textInput',
        cardBar: [
          {
            componentName: 'cardBar',
            title: 'Feedback'
          }
        ],
        header: [
          {
            componentName: 'text',
            text: 'Please provide your feedback below.'
          }
        ],
        body: [
          {
            componentName: 'textInput',
            id: 'feedbackValue',
            placeholder: 'Your feedback ...'
          }
        ],
        footer: [
          {
            componentName: 'cardButton',
            id: 'SendFeedback',
            title: 'Submit',
            action: {
              type: 'Submit',
              parameters: {
                viewToNavigateTo: this.OK_FEEDBACK_CARD_VIEW_ID
              }
            }
          }
        ],
        image: {
          url: `https://${process.env.BOT_DOMAIN}/assets/Collect-Feedback.png`,
          altText: 'Feedback'
        }        
      },
      viewId: this.COLLECT_FEEDBACK_CARD_VIEW_ID,
      onCardSelection: {
        type: 'QuickView',
        parameters: {
          view: this.SHOW_FEEDBACK_QUICK_VIEW_ID
        }
      }
    };

    // #region Alternative syntax with function TextInputCardView

    // Collect Feedback Card View (Input Text Card View via TextInputCardView)
    // const feedbackCardView: CardViewResponse = {
    //   viewId: this.COLLECT_FEEDBACK_CARD_VIEW_ID,
    //   aceData: aceData,
    //   cardViewParameters: TextInputCardView(
    //     {
    //       componentName: 'cardBar',
    //       title: 'Feedback'
    //     },
    //     {
    //       componentName: 'text',
    //       text: 'Please provide your feedback below.'
    //     },
    //     {
    //       componentName: 'textInput',
    //       id: 'feedbackValue',
    //       placeholder: 'Your feedback ...'
    //     },
    //     [
    //       {
    //         componentName: 'cardButton',
    //         title: 'Submit',
    //         id: 'SendFeedback',
    //         action: {
    //           type: 'Submit',
    //           parameters: {
    //             viewToNavigateTo: this.OK_FEEDBACK_CARD_VIEW_ID
    //           }
    //         }
    //       }
    //     ]
    //   ),
    //   onCardSelection: {
    //     type: 'QuickView',
    //     parameters: {
    //       view: this.SHOW_FEEDBACK_QUICK_VIEW_ID
    //     }
    //   }
    // };

    // #endregion

    this._cardViews[this.COLLECT_FEEDBACK_CARD_VIEW_ID] = feedbackCardView;

    // OK Feedback Card View (Image Card View)
    const okFeedbackCardViewResponse: CardViewResponse = {
      aceData: aceData,
      cardViewParameters: ImageCardView(
        {
          componentName: 'cardBar',
          title: 'Feedback Collected'
        },
        {
          componentName: 'text',
          text: 'Here is your feedback \'<feedback>\' collected on \'<dateTimeFeedback>\''
        },
        {
          url: `https://${process.env.BOT_DOMAIN}/assets/Ok-Feedback.png`,
          altText: "Feedback collected"
        },
        [
          {
            componentName: 'cardButton',
            title: 'Ok',
            id: 'OkButton',
            action: {
              type: 'Submit',
              parameters: {
                viewToNavigateTo: this.COLLECT_FEEDBACK_CARD_VIEW_ID
              }
            }
          }
        ]
      ),
      viewId: this.COLLECT_FEEDBACK_CARD_VIEW_ID,
      onCardSelection: {
        type: 'QuickView',
        parameters: {
          view: this.SHOW_FEEDBACK_QUICK_VIEW_ID
        }
      }
    };
    this._cardViews[this.OK_FEEDBACK_CARD_VIEW_ID] = okFeedbackCardViewResponse;

  }

  protected override onSharePointTaskGetCardViewAsync(_context: TurnContext, _aceRequest: AceRequest): Promise<CardViewResponse> {
    return Promise.resolve(this._cardViews[this.COLLECT_FEEDBACK_CARD_VIEW_ID]);
  }

  protected override onSharePointTaskGetQuickViewAsync(_context: TurnContext, _aceRequest: AceRequest): Promise<QuickViewResponse> {

    // Prepare the AdaptiveCard for the Quick View
    const card = new AdaptiveCards.AdaptiveCard();
    card.version = new AdaptiveCards.Version(1, 5);
    const cardPayload = {
      type: 'AdaptiveCard',
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      body: [
          {
              type: 'TextBlock',
              text: 'Thanks for your feedback!',
              weight: 'Bolder',
              size: 'Large',
              wrap: true,
              maxLines: 1,
              spacing: 'None',
              color: 'Dark'
          },
          {
            type: 'TextBlock',
            text: 'We truly appreciate your effort in providing valuable feedback to us. Thanks!',
            weight: 'Normal',
            size: 'Medium',
            wrap: true,
            maxLines: 3,
            spacing: 'None',
            color: 'Dark'
        }
      ]
    };
    card.parse(cardPayload);

    // Add the Feedback QuickViews
    const feedbackQuickViewResponse: QuickViewResponse = {
      viewId: this.SHOW_FEEDBACK_QUICK_VIEW_ID,
      title: 'Your feedback',
      template: card,
      data: {},
      externalLink: null,
      focusParameters: null
    };

    return Promise.resolve(feedbackQuickViewResponse);
  }

  protected override onSharePointTaskHandleActionAsync(_context: TurnContext, _aceRequest: AceRequest): Promise<HandleActionResponse> {

    const requestData =_aceRequest.data;

    if (requestData.type === 'Submit' && requestData.id === 'SendFeedback') {

      const viewToNavigateTo = requestData.data.viewToNavigateTo;
      const feedbackValue = requestData.data.feedbackValue;
      const dateTimeFeedback = new Date();

      const nextCard = this._cardViews[viewToNavigateTo];
      const textContent = `Here is your feedback '${feedbackValue}' collected on '${dateTimeFeedback.toLocaleString()}'`;
      nextCard.cardViewParameters.header[0].text = textContent;

      const response: CardViewHandleActionResponse = {
        responseType: 'Card',
        renderArguments: nextCard,
      };

      return Promise.resolve(response);

    } else if (requestData.type === 'Submit' && requestData.id === 'OkButton') {

      const viewToNavigateTo = requestData.data.viewToNavigateTo;

      const response: CardViewHandleActionResponse = {
        responseType: 'Card',
        renderArguments: this._cardViews[viewToNavigateTo],
      };
      
      return Promise.resolve(response);
    }
  }
}
