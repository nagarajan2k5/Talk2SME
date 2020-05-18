import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import TechnologyCards from "../TechnologyCards";

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/projectsMessageExtension/config.html")
export default class ProjectsMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {
        
        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: "Headline"
                    },
                    {
                        type: "TextBlock",
                        text: "Description"
                    },
                    {
                        type: "Image",
                        url: `https://${process.env.HOSTNAME}/assets/icon.png`
                    }
                ],
                actions: [
                    {
                        type: "Action.Submit",
                        title: "More details",
                        data: {
                            action: "moreDetails",
                            id: "1234-5678"
                        }
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.0"
            });
        const preview = {
            contentType: "application/vnd.microsoft.card.thumbnail",
            content: {
                title: "Headline",
                text: "Description",
                images: [
                    {
                        url: `https://${process.env.HOSTNAME}/assets/icon.png`
                    }
                ]
            }
        };

        if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
            // initial run

            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: TechnologyCards
            } as MessagingExtensionResult);
        } else {
            // the rest
            if(query.parameters && query.parameters[0]){
            var queryString = query.parameters[0].value || "";
            }
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: TechnologyCards.filter(c => c.content.title ? c.content.title.toLowerCase().includes(queryString.toLowerCase()):true)
            } as MessagingExtensionResult);
        }
    }

    public async onCardButtonClicked(context: TurnContext, value: any): Promise<void> {
        // Handle the Action.Submit action on the adaptive card
        log(`I got this ${value.id}`);
        if (value.action === "moreDetails") {
            log(`I got this ${value.id}`);
        }
        return Promise.resolve();
    }






}
