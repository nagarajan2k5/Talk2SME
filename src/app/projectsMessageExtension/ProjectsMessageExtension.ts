import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import ProjectCards from "../ProjectCards"
import { GraphProvider } from "../../Graph/GraphProvider";

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/projectsMessageExtension/config.html")
export default class ProjectsMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {

        if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: ProjectCards
            } as MessagingExtensionResult);
        } else {
            // the rest
            if (query.parameters && query.parameters[0]) {
                var queryString = (query.parameters[0].value || "").toLowerCase();
            }
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: await this.getProjectsandUsersCards(queryString)
                // attachments: ProjectCards.filter(p => p.content.body[0].text ?
                //     (p.content.body[0].text.toLowerCase().includes(queryString) ||
                //         p.content.body[1].text.toLowerCase().includes(queryString) ||
                //         p.content.body[3].text.toLowerCase().includes(queryString) ||
                //         p.content.body[5].text.toLowerCase().includes(queryString)) : false
                // ),
            } as MessagingExtensionResult);
        }
    }

    public async onCardButtonClicked(context: TurnContext, value: any): Promise<void> {
        // Handle the Action.Submit action on the adaptive card
        log(`onCardButtonClicked, I got this ${value.id}`);
        // if (value.action === "moreDetails") {
        //     log(`I got this ${value.id}`);
        // }
        //log(JSON.stringify(context.activity));
        let requestor = context.activity.from.name;
        let cardInfo = context.activity.value;
        if (cardInfo) {
            await context.sendActivity("Thank you " + requestor + ", we will process the request soon..!");
        }

        return Promise.resolve();
    }

    private async getProjectsandUsersCards(keyWord: string): Promise<any> {
        let result;
        try {
            if (keyWord) {
                log("getProjectsandUsersCards");

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
        let result = new Array<any>();
        try {
            const users = await GraphProvider.searchPeopleBySkills(keyWord);
            console.log(JSON.stringify(users));
            log(JSON.stringify(users));
            if (users) {
                users.forEach(data => {
                    const preview = {
                        contentType: "application/vnd.microsoft.card.thumbnail",
                        content: {
                            title: data.FullName,
                            text: data.Skills,
                            images: [
                                {
                                    url: `https://www.vippng.com/png/detail/355-3555954_connect-icon-png-connected-icon.png`
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
                                    "type": "Action.Submit",
                                    "title": "Talk to Me",
                                    "data": {
                                        "btnTalkToSME": data
                                    },
                                    "id": "Talk to Me"
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
