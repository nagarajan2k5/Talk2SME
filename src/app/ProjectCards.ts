// Project cards
import ProjectData = require("../app/CardsSamples/ProjectInfoData.json");

const ProjectCards = new Array<any>();

ProjectData.forEach(data => {
    const preview = {
        contentType: "application/vnd.microsoft.card.thumbnail",
        content: {
            title: data.ProjectName,
            text: data.Description,
            images: [
                {
                    url: `https://picsum.photos/32/32?image=845`
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
                    "separator": true
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
                    "separator": true
                },
                {
                    "type": "TextBlock",
                    "text": "SME Contact",
                    "weight": "bolder",
                    "isSubtle": false
                },
                {
                    "type": "ColumnSet",
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
                        "x": "nagarajan"
                    }
                }
            ]
        }
    };

    ProjectCards.push({ ...card, preview });
});

export default ProjectCards;

//ProjectCards.filter(c => c.content.body[0].content.text ? c.content.body[0].content.text.toLowerCase().includes(queryString.toLowerCase()):true)





