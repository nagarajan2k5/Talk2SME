// Technology cards
import { CardFactory, Attachment, CardAction, ContactRelationUpdateActionTypes } from "botbuilder";

const TechnologyCards = new Array<Attachment>();

export default TechnologyCards;

//TechnologyCards.splice(0,1);
TechnologyCards.push(CardFactory.thumbnailCard(
    "Infy Connect 2020 - Talk2SME",
    "Technologies: Teams App, Bot Framework, NodeJS, Messaging Extension.",
    ["https://picsum.photos/100?image=889"],
    ["Contact SME"]
));
TechnologyCards.push(CardFactory.thumbnailCard(
    "Infy Connect 2020 - Talk2SME 2.0",
    "Technologies: Teams App, Bot Framework, NodeJS, Messaging Extension.",
    ["https://picsum.photos/100?image=860"],
    ["Contact SME"]
));
TechnologyCards.push(CardFactory.thumbnailCard(
    "Infy Connect 2020 - NotificationFactory",
    "Technologies: SPFx, Teams App, Bot Framework, C#.Net, Messaging Extension",
    ["https://picsum.photos/100?image=851"],
    ["Contact SME"]

));


