import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import ProjectCards from "../ProjectCards"

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/projectsMessageExtension/config.html")
export default class ProjectsMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {
        
        if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
            // initial run

            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: ProjectCards
            } as MessagingExtensionResult);
        } else {
            // the rest
            if(query.parameters && query.parameters[0]){
            var queryString = (query.parameters[0].value || "").toLowerCase();
            }
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: ProjectCards.filter(p => p.content.body[0].text ? 
                    (p.content.body[0].text.toLowerCase().includes(queryString) || 
                        p.content.body[3].text.toLowerCase().includes(queryString)):false
                        ),
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
