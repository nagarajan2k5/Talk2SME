import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import * as bot from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import ProjectCards from "../ProjectCards"
import { GraphProvider } from "../../Graph/GraphProvider";

import { createMeetingService } from "../../Services/createMeetingService";
import { OnlineMeetingInput, OutlookEventInfo } from "../../Models/models";
import * as moment from 'moment';

import { conversationReferences } from "../talkToSmeBot/TalkToSmeBot"


// Initialize debug logging module
const log = debug("msteams");

const connectionName: string = process.env.OAuthConnectionName || '';

@PreventIframe("/projectsMessageExtension/config.html")
export default class ProjectsMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onQuery(context: bot.TurnContext, query: bot.MessagingExtensionQuery): Promise<bot.MessagingExtensionResult> {
        log("handler: onQuery");
        if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: ProjectCards
            } as bot.MessagingExtensionResult);
        } else {
            // the rest
            if (query.parameters && query.parameters[0]) {
                var queryString = (query.parameters[0].value || "").toLowerCase();
            }
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: await this.getProjectsandUsersCards(queryString)
            } as bot.MessagingExtensionResult);
        }
    }

    public async onCardButtonClicked(context: bot.TurnContext, value: any): Promise<void> {
        // Handle the Action.Submit action on the adaptive card
        log("handler: onCardButtonClicked");
        let requestor = context.activity.from.name;
        let cardInfo = context.activity.value;

        if (cardInfo) {
            switch (cardInfo.commandId) {
                case "Meeting":
                    // let i: IUserTokenProvider;
                    // i.getUserToken()
                    // Get access token for user.if already authenticated, we will get token.
                    // If user is not signed in, send sign in link in messaging extension.
                    let magicCode = "";//context.activity.value?.state || '';
                    //const conversationReference = bot.TurnContext.getConversationReference(activity);

                    let conversationReference;
                    for (conversationReference of Object.values(conversationReferences)) {
                        if (conversationReference.conversation.id === context.activity.conversation.id) {
                            magicCode = conversationReference.magicCode;
                            break;
                        }
                    }
                    log("Msg Extension magicCode: " + magicCode);
                    var tokenResponse = await (context.adapter as any).getUserToken(context, connectionName, magicCode);
                    if (!tokenResponse) {
                        log("Msg Extension: Sign in card");
                        const signInLink: any = await (context.adapter as any).getSignInLink(context, connectionName);
                        //send a sign in card!
                        const attachment = bot.CardFactory.signinCard("Sign in", signInLink, "Please sign in to schedule a meeting with SME");
                        const activity = bot.MessageFactory.attachment(attachment);
                        await context.sendActivity(activity);
                    }
                    else {
                        //send a meeting schedule confirm card!
                        await context.sendActivity("Thank you " + requestor + ", we have scheduled the meeting! Please check the calendar");
                        //Create a test meeting
                        const service = createMeetingService();
                        const startedAt = moment()
                        const meetingInput: OnlineMeetingInput = {
                            startDateTime: startedAt.add(30, 'm'),
                            endDateTime: startedAt.add(60, 'm'),
                            subject: "TalkToSME: Online meeting with SME - " + (value.parameters.Title ? value.parameters.Title : value.parameters.FullName),
                            smeEmailID: value.parameters.SMEContacts ? value.parameters.SMEContacts : value.parameters.EmailId
                        };

                        const meetingInfo: OutlookEventInfo = await service.createMeeting(meetingInput, tokenResponse.token);
                        //await context.sendActivity(JSON.stringify(meetingInfo));
                    }
                    break;
                default:
                    await context.sendActivity("Default");
                    break;
            }
        }
        return Promise.resolve();
    }

    public async onSubmitAction(context: bot.TurnContext, value: bot.MessagingExtensionAction): Promise<bot.MessagingExtensionResult> {
        log("handler: onSubmitAction");
        return Promise.resolve({} as bot.MessagingExtensionResult);
    }

    private async getProjectsandUsersCards(keyWord: string): Promise<any> {
        log("method: getProjectsandUsersCards");
        let result;
        try {
            if (keyWord) {
                result = ProjectCards.filter(p => p.content.body[0].text ?
                    (p.content.body[0].text.toLowerCase().includes(keyWord) ||
                        p.content.body[1].text.toLowerCase().includes(keyWord) ||
                        p.content.body[3].text.toLowerCase().includes(keyWord) ||
                        p.content.body[5].text.toLowerCase().includes(keyWord)) : false);

                //Merging User details
                let users = await this.getUsersCards(keyWord);
                users.forEach(user => {
                    result.push(user);
                });
            }
            else {
                result = ProjectCards;
            }
        } catch (error) {
            console.log("Error on getProjectsandUsers method");
            console.log(error);
        }
        return result;
    }

    private async getUsersCards(keyWord: string): Promise<any> {
        log("method: getUsersCards");
        let result = new Array<any>();
        try {
            const users = await GraphProvider.searchPeopleBySkills(keyWord);
            if (users) {
                users.forEach(data => {
                    const preview = {
                        contentType: "application/vnd.microsoft.card.thumbnail",
                        content: {
                            title: data.FullName,
                            text: data.Skills,
                            images: [
                                {
                                    url: `https://f1.pngfuel.com/png/323/743/633/icon-person-icon-design-symbol-avatar-silhouette-character-cartoon-head-png-clip-art.png`
                                }
                            ]
                        }
                    };
                    const card = {
                        "contentType": "application/vnd.microsoft.card.adaptive",
                        "content": {
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "version": "1.0",
                            "type": "AdaptiveCard",
                            "body": [

                                {
                                    "type": "TextBlock",
                                    "text": "Skills",
                                    "weight": "bolder",
                                    "isSubtle": false
                                },
                                {
                                    "type": "TextBlock",
                                    "text": data.Skills,
                                    "separator": true,
                                    "wrap": true
                                },
                                {
                                    "type": "TextBlock",
                                    "text": "SME Contact",
                                    "weight": "bolder",
                                    "isSubtle": false
                                },
                                {
                                    "type": "ColumnSet",
                                    "separator": true,
                                    "columns": [
                                        {
                                            "type": "Column",
                                            "width": "auto",
                                            "items": [
                                                {
                                                    "type": "Image",
                                                    "url": "https://f1.pngfuel.com/png/323/743/633/icon-person-icon-design-symbol-avatar-silhouette-character-cartoon-head-png-clip-art.png",
                                                    "size": "small",
                                                    "style": "person"
                                                }
                                            ]
                                        },
                                        {
                                            "type": "Column",
                                            "width": "stretch",
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "text": data.EmailId,
                                                    "weight": "bolder",
                                                    "wrap": true
                                                }
                                            ]
                                        }
                                    ]
                                }
                            ],
                            "actions": [
                                {
                                    "type": "Action.OpenUrl",
                                    "title": "Chat",
                                    "url": "https://teams.microsoft.com/l/chat/0/0?users=" + data.EmailId,
                                    "data": {
                                        "btnTalkToSME": data
                                    },
                                    "id": "Chat"
                                },
                                {
                                    "type": "Action.Submit",
                                    "title": "Meeting",
                                    "data": {
                                        "parameters": data,
                                        "msteams": {
                                            "type": "messageBack",
                                            "displayText": "Meeting",
                                            "text": "Meeting"
                                        },
                                        "commandId": "Meeting"
                                    },
                                    "id": "Meeting"
                                }
                            ]
                        }
                    };

                    result.push({ ...card, preview });
                });
            }
        }
        catch (error) {

        }
        return result;
    }






}
