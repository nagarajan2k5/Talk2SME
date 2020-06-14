import { BotDeclaration, MessageExtensionDeclaration, PreventIframe } from "express-msteams-host";
import * as debug from "debug";
import * as botDialogs from "botbuilder-dialogs";
import * as bot from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import ProjectsMessageExtension from "../projectsMessageExtension/ProjectsMessageExtension";
import WelcomeCard from "./dialogs/WelcomeDialog";
import { GraphProvider } from "../../Graph/GraphProvider";
import { WaterfallDialog } from "botbuilder-dialogs";
import { Activity } from "botbuilder";

// Initialize debug logging module
const log = debug("msteams");
const connectionName: string = process.env.OAuthConnectionName || '';

const MAIN_WATERFALL_DIALOG = "MainWaterfallDialog";
const OAUTH_PROMPT = "OAuthPrompt";
const SIGN_OUT = "SignOutDialog";

export const conversationReferences = [];

// Add memory storage.
var storage = new bot.MemoryStorage();
// const storage = new CosmosDbPartitionedStorage({
//     cosmosDbEndpoint: process.env.CosmosDbEndpoint || '',
//     authKey: process.env.CosmosDbAuthKey || '',
//     databaseId: process.env.CosmosDbDatabaseId || '',
//     containerId: process.env.CosmosDbContainerId || '',
//     compatibilityMode: false
// });

/**
 * Implementation for TalkToSME Bot
 */
@BotDeclaration(
    "/api/messages",
    new bot.MemoryStorage(),
    process.env.MicrosoftAppId,
    process.env.MicrosoftAppPassword)

export class TalkToSmeBot extends bot.TeamsActivityHandler {
    private readonly conversationState: bot.ConversationState;
    /** Local property for ProjectsMessageExtension */
    @MessageExtensionDeclaration("projectsMessageExtension")
    private _projectsMessageExtension: ProjectsMessageExtension;
    private readonly dialogs: botDialogs.DialogSet;
    private dialogState: bot.StatePropertyAccessor<botDialogs.DialogState>;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: bot.ConversationState) {
        super();
        log("TalkToSmeBot: constructor");
        // Message extension ProjectsMessageExtension
        this._projectsMessageExtension = new ProjectsMessageExtension();


        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new botDialogs.DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));

        this.dialogs.add(new botDialogs.OAuthPrompt(OAUTH_PROMPT, {
            connectionName,
            text: "Please sign in so I can show you your profile.",
            title: "Sign in",
            timeout: 300000
        }));

        //Meeting schedule dialog
        this.dialogs.add(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [

            this.getTokenStep.bind(this)//,
            //this.sendMeetingCard.bind(this),
            //this.showMeetingInfoCard.bind(this)
        ]))

        //Sign out dialog
        this.dialogs.add(new WaterfallDialog(SIGN_OUT, [
            this.userSignOut.bind(this)
        ]));

        // Set up the Activity processing

        this.onMessage(async (context: bot.TurnContext): Promise<void> => {

            log("handler: onMessage");
            await this.addConversationReference(context.activity);

            for (let conversationReference of Object.values(conversationReferences)) {
                log("Object: " + JSON.stringify(conversationReference))
            }

            // TODO: add your own bot logic in here
            switch (context.activity.type) {
                case bot.ActivityTypes.Message:
                    let text = bot.TurnContext.removeRecipientMention(context.activity);
                    text = text?.toLowerCase() || "";
                    console.log("Onmessage text: " + text);
                    if (text.startsWith("hello")) {
                        await context.sendActivity("Oh, hello to you as well!");
                        return;
                    } else if (text.startsWith("help")) {
                        const dc = await this.dialogs.createContext(context);
                        await dc.beginDialog("help");
                    } else if (text.startsWith("search user")) {
                        let result = await GraphProvider.searchPeopleBySkills(text.split(' ')[2]);
                        await context.sendActivity("Output: " + JSON.stringify(result));
                    }
                    else if (text.startsWith("list")) {
                        let result = await GraphProvider.getListItems(text.split(' ')[1]);
                        await context.sendActivity("Output: " + JSON.stringify(result));
                    }
                    else if (text.startsWith("update skill")) {
                        let result = await GraphProvider.updateSkillProficiency("nagarajan_s05@msnextlife.onmicrosoft.com", text.split(' ')[2]);
                        await context.sendActivity("Output: " + JSON.stringify(result));
                    }
                    else if (text.startsWith("search project")) {
                        let result = await GraphProvider.searchProjectsByKeyword(text.split(' ')[2]);
                        await context.sendActivity("Output: " + JSON.stringify(result));
                    } else if (text.startsWith("sign in")) {
                        log("Auth dialog begin");
                        const dc = await this.dialogs.createContext(context);
                        await dc.beginDialog(OAUTH_PROMPT);
                    }
                    else if (text.startsWith("sign out")) {
                        const dc = await this.dialogs.createContext(context);
                        await dc.beginDialog(SIGN_OUT);
                    }
                    else {
                        //await context.sendActivity(`My training is under progress to answer all your queries!`);
                        //const adapter: bot.IUserTokenProvider = context.adapter as bot.BotAdapter;  
                    }
                    break;
                default:
                    break;
            }
            // Save state changes
            return this.conversationState.saveChanges(context);
        });

        this.onConversationUpdate(async (context: bot.TurnContext): Promise<void> => {
            log("handler: onConversationUpdate");
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                        this.addConversationReference(context.activity);
                        const welcomeCard = bot.CardFactory.adaptiveCard(WelcomeCard);
                        await context.sendActivity({ attachments: [welcomeCard] });
                    }
                }
            }
        });

        this.onMessageReaction(async (context: bot.TurnContext): Promise<void> => {
            log("handler: onMessageReaction");
            const added = context.activity.reactionsAdded;
            if (added && added[0]) {
                await context.sendActivity({
                    textFormat: "xml",
                    text: `Thank you! That was an interesting reaction (<b>${added[0].type}</b>)`
                });
            }
        });
    }

    private async addConversationReference(activity: Activity, magicCode: string = ""): Promise<void> {
        log("magic code: " + magicCode);
        if (magicCode !== "") {
            log("function: addConversationReference");
            const conversationReference = bot.TurnContext.getConversationReference(activity);
            if (conversationReference.conversation) {
                conversationReferences[conversationReference.conversation.id] = { ...conversationReference, magicCode };
            }
        }
        return Promise.resolve();
    }

    public async run(context: bot.TurnContext) {
        log("handler: run");
        await super.run(context);

        // Save any state changes. The load happened during the execution of the Dialog.
        await this.conversationState.saveChanges(context, false);
    }

    protected async handleTeamsSigninVerifyState(context: bot.TurnContext, query: bot.SigninStateVerificationQuery): Promise<void> {
        log("handler: handleTeamsSigninVerifyState");
        await context.sendActivity("Well!, token received");
        log("context: " + JSON.stringify(context.activity));
        log("query: " + JSON.stringify(query));
        await this.addConversationReference(context.activity, query.state);
        await context.sendActivity(`You're now signed in.`);
        const dc = await this.dialogs.createContext(context);
        //await dc.beginDialog(MAIN_WATERFALL_DIALOG);        
        await dc.continueDialog();
    }

    protected async userSignOut(stepContext: botDialogs.WaterfallStepContext) {
        log("step: userSignOut  ");
        const adapter = stepContext.context.adapter as bot.BotFrameworkAdapter;
        await adapter.signOutUser(stepContext.context, connectionName);

        await stepContext.context.sendActivity(`You're now signed out from Bot.`);
        return await stepContext.endDialog();
    }

    protected async sendMeetingCard(stepContext: botDialogs.WaterfallStepContext) {
        log("step: sendMeetingCard  ");
        await stepContext.context.sendActivity("Meeting card");
        return await stepContext.endDialog();
    }

    private getTokenStep(stepContext: botDialogs.WaterfallStepContext) {
        log("step: getTokenStep");
        return stepContext.beginDialog(OAUTH_PROMPT);
        //stepContext.continueDialog();
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
        else {
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
    catch (err) {
        await turnContext.sendActivity(`Read rejected. ${err}`);
    }
}
