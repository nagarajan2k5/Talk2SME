// Project cards
import * as debug from "debug";
import ProjectData = require("../app/CardsSamples/ProjectInfoData.json");

//logging modulec declared
const log = debug("msteams");

const ProjectCards = new Array<any>();

ProjectData.forEach(data => {
    const preview = {
        contentType: "application/vnd.microsoft.card.thumbnail",
        content: {
            title: data.ProjectName,
            text: data.TechnologyStack,
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
                    "text": data.ProjectName,
                    "weight": "bolder",
                    "isSubtle": false
                },
                {
                    "type": "TextBlock",
                    "text": data.Description,
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
                    "text": data.Domain,
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
                    "text": data.TechnologyStack,
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
                                    "text": data.SME,
                                    "weight": "bolder",
                                    "wrap": true
                                },
                                {
                                    "type": "TextBlock",
                                    "spacing": "none",
                                    "text": data.CreatedOn,
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
                    "type": "Action.Submit",
                    "title": "Talk to Me",
                    "data": {
                        "btnTalkToSME": data
                    },
                    "id": "Talk to Me"
                },
                {
                    "type": "Action.Submit",
                    "title": "Request for Case Study",
                    "data": {
                        "btnCaseStudy": data
                    },
                    "id": "Request for Case Study"
                }
            ]
        }
    };

    ProjectCards.push({ ...card, preview });
});

export default ProjectCards;





