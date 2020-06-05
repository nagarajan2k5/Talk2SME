import { SPFetchClient } from "@pnp/nodejs";
import { sp, SearchQueryBuilder } from "@pnp/sp/presets/all";
import { ISearchQuery, SearchResults } from "@pnp/sp/search";

import * as debug from "debug";
const log = debug("msteams");

export class PnPJsClient {

    private token: any;

    constructor() {

        const clientId = process.env.MicrosoftAppId || "";
        const clientSecret = process.env.MicrosoftAppPassword || "";
        const siteURL = process.env.SPO_SITE_URL || "";

        log("constructor: " + siteURL, clientId, clientSecret);

        sp.setup({
            sp: {
                fetchClientFactory: () => {
                    return new SPFetchClient(siteURL, clientId, clientSecret);
                },
            },
        });
    }

    /**
     * Check if a user exists
     * @param {string} skills Email address of the email's recipient.
     */
    public async searchPeopleBySkills(skills: string): Promise<SearchResults> {
        let res;
        try {
            log(skills);
            if (!skills) {
                throw new Error('Invalid `Skills` parameter received.');
            }
            console.log("SPO API Call");
            //const res = await sp.web.get();
            // URL: https://msnextlife.sharepoint.com/_api/search/query?querytext='Skills:*azure*'&sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'
            // Search specific path, URL: https://msnextlife.sharepoint.com/_api/search/query?querytext='(zela)'&querytemplate='{searchTerms} path:"https://msnextlife.sharepoint.com/Lists/ProjectRepo"'&EnableQueryRules=false

            res = await sp.search({
                Querytext: 'Skills:*' + skills + '*',
                SourceId: 'B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'
            });
            
            // const res1 = sp.profiles.setSingleValueProfileProperty("i:0#.f|membership|nagarajan_s05@msnextlife.onmicrosoft.com", "Skills", "PnP");
            // console.log(res1);
            // Get user profile info
            // const profile = await sp.profiles.userProfile;
            // console.log(profile);

            //console.log(JSON.stringify(res, null, 4));
            //console.log(JSON.stringify(res.RowCount));

        } catch (error) {
            console.log("searchPeopleBySkills method error");
            console.log(error);
        }
        return res;
    }

    /**
     * @param {string} keyword to search projects.
     */
    public async searchProjectsByKeyword(keyword: string): Promise<SearchResults> {
        let res;
        try {
            if (!keyword) {
                throw new Error('Invalid `keyword` parameter received.');
            }
            console.log("SPO API Call");
            // Search specific path, URL: https://msnextlife.sharepoint.com/_api/search/query?querytext='(zela)'&querytemplate='{searchTerms} path:"https://msnextlife.sharepoint.com/Lists/ProjectRepo"'&EnableQueryRules=false

            res = await sp.search({
                Querytext: keyword + '*',
                QueryTemplate: `{searchTerms} path:"https://msnextlife.sharepoint.com/Lists/ProjectRepo"`,
                EnableQueryRules: false
            });

            console.log(JSON.stringify(res, null, 4));
            console.log("Search count: " + JSON.stringify(res.RowCount));

        } catch (error) {
            console.log("searchPeopleBySkills method error");
            console.log(error);
        }
        return res;
    }
}