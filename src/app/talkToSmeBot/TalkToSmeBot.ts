import { BotDeclaration, MessageExtensionDeclaration, PreventIframe } from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, TeamsActivityHandler, Activity } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import ProjectsMessageExtension from "../projectsMessageExtension/ProjectsMessageExtension";
import WelcomeCard from "./dialogs/WelcomeDialog";
import {CosmosDbPartitionedStorage} from "botbuilder-azure"

// Initialize debug logging module
const log = debug("msteams");

export const conversationReferences = [];

// Add memory storage.
//var storage = new MemoryStorage();
const storage = new CosmosDbPartitionedStorage({
    cosmosDbEndpoint: process.env.CosmosDbEndpoint || '',
    authKey: process.env.CosmosDbAuthKey || '',
    databaseId: process.env.CosmosDbDatabaseId || '',
    containerId: process.env.CosmosDbContainerId || '',
    compatibilityMode: false
});

/**
 * Implementation for TalkToSME Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)

export class TalkToSmeBot extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    /** Local property for ProjectsMessageExtension */
    @MessageExtensionDeclaration("projectsMessageExtension")
    private _projectsMessageExtension: ProjectsMessageExtension;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super();
        // Message extension ProjectsMessageExtension
        this._projectsMessageExtension = new ProjectsMessageExtension();


        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));

        // Set up the Activity processing

        this.onMessage(async (context: TurnContext): Promise<void> => {
           
            // TODO: add your own bot logic in here
            switch (context.activity.type) {
                case ActivityTypes.Message:
                    let text = TurnContext.removeRecipientMention(context.activity);
                    text = text.toLowerCase();
                    if (text.startsWith("hello")) {
                        await context.sendActivity("Oh, hello to you as well!");
                        return;
                    } else if (text.startsWith("help")) {
                        const dc = await this.dialogs.createContext(context);
                        await dc.beginDialog("help");
                    } else {
                        await context.sendActivity(`My training is under progress to answer all your queries!`);
                    }
                    break;
                default:
                    break;
            }
            // Save state changes
            return this.conversationState.saveChanges(context);
        });

        this.onConversationUpdate(async (context: TurnContext): Promise<void> => {
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                        addConversationReference(context.activity)
                        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                        await context.sendActivity({ attachments: [welcomeCard] });
                    }
                }
            }
        });

        this.onMessageReaction(async (context: TurnContext): Promise<void> => {
            const added = context.activity.reactionsAdded;
            if (added && added[0]) {
                await context.sendActivity({
                    textFormat: "xml",
                    text: `Thank you! That was an interesting reaction (<b>${added[0].type}</b>)`
                });
            }
        });;

        function addConversationReference(activity: Activity): void {
            const conversationReference = TurnContext.getConversationReference(activity);
            if(conversationReference.conversation){
            conversationReferences[conversationReference.conversation.id] = conversationReference;            
            }
        }
    }
}

// This function stores new user messages. Creates new utterance log if none exists.
async function logMessageText(storage, turnContext) {

    await turnContext.sendActivity("JSON Storage :" + JSON.stringify(storage));
    let utterance = turnContext.activity.text;
    // debugger;
    try {
        // Read from the storage.
        let storeItems = await storage.read(["UtteranceLogJS"])
        // Check the result.
        var UtteranceLogJS = storeItems["UtteranceLogJS"];
        if (typeof (UtteranceLogJS) != 'undefined') {
            // The log exists so we can write to it.
            storeItems["UtteranceLogJS"].turnNumber++;
            storeItems["UtteranceLogJS"].UtteranceList.push(utterance);
            // Gather info for user message.
            var storedString = storeItems.UtteranceLogJS.UtteranceList.toString();
            var numStored = storeItems.UtteranceLogJS.turnNumber;

            try {
                await storage.write(storeItems)
                await turnContext.sendActivity(`${numStored}: The list is now: ${storedString}`);
            } catch (err) {
                await turnContext.sendActivity(`Write failed of UtteranceLogJS: ${err}`);
            }
        }
        else{
            await turnContext.sendActivity(`Creating and saving new utterance log`);
            var turnNumber = 1;
            storeItems["UtteranceLogJS"] = { UtteranceList: [`${utterance}`], "eTag": "*", turnNumber }
            // Gather info for user message.
            var storedString = storeItems.UtteranceLogJS.UtteranceList.toString();
            var numStored = storeItems.UtteranceLogJS.turnNumber;

            try {
                await storage.write(storeItems)
                await turnContext.sendActivity(`${numStored}: The list is now: ${storedString}`);
            } catch (err) {
                await turnContext.sendActivity(`Write failed: ${err}`);
            }
        }
    }
    catch (err){
        await turnContext.sendActivity(`Read rejected. ${err}`);
    }
}
