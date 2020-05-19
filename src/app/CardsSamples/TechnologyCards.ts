// Technology cards
import { CardFactory, Attachment, CardAction, ContactRelationUpdateActionTypes } from "botbuilder";


import FlightCard = require("./FlightCard.json");

const TechnologyCards = new Array<any>();

export default TechnologyCards;

// Type of methods to add cards 

//TechnologyCards.splice(0,1);

//Method 1
TechnologyCards.push(CardFactory.thumbnailCard(
    "Infy Connect 2020 - Talk2SME",
    "**Technologies:** Teams App, Bot Framework, NodeJS, Messaging Extension.",
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

const card1 = CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    body: [
        {
            type: "TextBlock",
            size: "Large",
            text: "Adaptive Card Sample"
        },
        {
            type: "TextBlock",
            text: "A customizable card that can contain any combination of text, speech, images, buttons, and input fields.",
            wrap: true
        },
        {
            type: "Image",
            url: `https://picsum.photos/450/300?image=890`
        }
    ],
    actions: [
        {
            type: "Action.OpenUrl",
            title: "More Info",
            url: "https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-reference#adaptive-card"
        }
    ],
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.0"
});

const preview = {
    contentType: "application/vnd.microsoft.card.thumbnail",
    content: {
        title: "Adaptive Card Sample",
        text: "A customizable card that can contain any combination of text, speech, images, buttons, and input fields.",
        images: [
            {
                url: `https://picsum.photos/32/32?image=890`
            }
        ]
    }
};


//Method 2
TechnologyCards.push({...card1, preview }); 
TechnologyCards.push({...CardFactory.adaptiveCard({...FlightCard}),preview});

//Method 3
const tempCard =  JSON.parse(JSON.stringify({...FlightCard}));
tempCard.body[0].text = "Nagarajan";
TechnologyCards.push({...CardFactory.adaptiveCard({...tempCard}),preview});



//TechnologyCards.push(CardFactory.adaptiveCard(card1),);


