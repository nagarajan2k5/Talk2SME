// Project cards
import * as debug from "debug";
//import ProjectData = require("../app/CardsSamples/ProjectInfoData.json");

import { GraphProvider } from "../Graph/GraphProvider"; 
//logging modulec declared
const log = debug("msteams");

const ProjectCards = new Array<any>();

const ProjectData = GraphProvider.getListItems("nofilter");
log(JSON.stringify(ProjectData));
if (ProjectData) {
    ProjectData.then((res) =>
        res.forEach(data => {
            const preview = {
                contentType: "application/vnd.microsoft.card.thumbnail",
                content: {
                    title: data.Title,
                    text: data.Technology,
                    images: [
                        {
                            url: `https://banner2.cleanpng.com/20180528/byz/kisspng-case-study-computer-icons-blockchain-technology-re-5b0c44c685ce95.4336176315275306945481.jpg`
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
                            "text": data.Title,
                            "weight": "bolder",
                            "isSubtle": false
                        },
                        {
                            "type": "TextBlock",
                            "text": data.Abstract,
                            "separator": true,
                            "wrap": true

                        },
                        {
                            "type": "TextBlock",
                            "text": "Domain",
                            "weight": "bolder",
                            "isSubtle": false
                        },
                        {
                            "type": "TextBlock",
                            "text": data.BusinessScenario,
                            "separator": true,
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "Technology Stack",
                            "weight": "bolder",
                            "isSubtle": false
                        },
                        {
                            "type": "TextBlock",
                            "text": data.Technology,
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
                                            "text": data.SMEContacts,
                                            "weight": "bolder",
                                            "wrap": true
                                        },
                                        {
                                            "type": "TextBlock",
                                            "spacing": "none",
                                            "text": "Created {{DATE(" + data.CreatedOn + ", SHORT)}}",
                                            "isSubtle": true,
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
                            "url": "https://teams.microsoft.com/l/chat/0/0?users=" + data.SMEContacts,
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
                        },
                        {
                            "type": "Action.OpenUrl",
                            "title": "Case Study",
                            "url": data.CaseStudyURL
                        },
                    ]
                }
            };

            ProjectCards.push({ ...card, preview });
        }));
}

export default ProjectCards;





