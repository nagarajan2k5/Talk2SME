import { Client } from '@microsoft/microsoft-graph-client';
import { User } from '@microsoft/microsoft-graph-types';

import * as debug from "debug";
const log = debug("msteams");

export class GraphClient {

    private token: string;
    private graphClient: Client;

    constructor(token: any) {
        if (!token || !token.trim()) {
            throw new Error('GraphClient: Invalid token received.');
        }

        this.token = token;

        this.graphClient = Client.init({
            authProvider: (done) => {
                done(null, this.token);
            }
        });
    }

    /**
     * Check if a user exists
     * @param {string} emailAddress Email address of the email's recipient.
     */
    public async userExists(emailAddress: string): Promise<boolean> {
        if (!emailAddress || !emailAddress.trim()) {
            throw new Error('Invalid `emailAddress` parameter received.');
        }
        try {
            const user: User = await this.graphClient.api(`/users/${emailAddress}`).get();
            return user ? true : false;
        } catch (error) {
            return false;
        }
    }

    /**
    * Check if a user exists
    * @param {string} keyWord Email address of the email's recipient.
    */
    public async getListItems(keyWord: string): Promise<any> {
        if (!keyWord || !keyWord.trim()) {
            throw new Error('Invalid `listName` parameter received.');
        }
        try {
            let apiURL = "/sites/" + (process.env.SPO_SITE_GUID || "") + "/lists/" + (process.env.SPO_LIST_GUID || "") + "/items/?expand=fields";
            console.log(apiURL);
            let res = await this.graphClient.api(apiURL).get();
            console.log("list found");
            return res;
        } catch (error) {
            console.log("list not found");
            console.log(error);
            return false;
        }
    }
    /**
    * Check if a user exists
    * @param {string} userId Email address of the email's recipient.
    * @param {string} skill Skill to update
    */
    public async updateSkillProficiency(userId: string, skill: string): Promise<any> {
        if (!userId || !userId.trim() || !skill || !skill.trim()) {
            throw new Error('Invalid `userId` or `skill` parameter received.');
        }
        try {
            let apiURL = "/me/profile/skills/"+userId;
            console.log(apiURL);

            const skillProficiency = {
                categories: [
                  "professional" //personal, professional, hobby
                ],
                displayName: skill, 
                proficiency: "expert", //elementary, limitedWorking, generalProfessional, advancedProfessional, expert, unknownFutureValue.
                webUrl: ""
              };

            let res = await this.graphClient.api(apiURL).version('beta').update(skillProficiency);
            console.log("skill updated");
            console.log(res);
            return res;
        } catch (error) {
            console.log("skill update failed");
            console.log(error);
            return false;
        }
    }


}