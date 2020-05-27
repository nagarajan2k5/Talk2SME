import { Client } from '@microsoft/microsoft-graph-client';
import { User } from '@microsoft/microsoft-graph-types';

import * as debug from "debug";
const log = debug("msteams");

export class GraphClient {

    private token: string;
    private graphClient: Client;

    constructor(token: any) {
        if (!token || !token.trim()) {
            throw new Error('SimpleGraphClient: Invalid token received.');
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
        console.log("client");
        if (!emailAddress || !emailAddress.trim()) {
            throw new Error('Invalid `emailAddress` parameter received.');
        }
        try {
            const user: User = await this.graphClient.api(`/users/${emailAddress}`).get();
            console.log("user found");
            return user ? true : false;
        } catch (error) {
            console.log("user not found");
            return false;
        }
    }

    /**
    * Check if a user exists
    * @param {string} listName Email address of the email's recipient.
    */
    public async getListItems(listName: string): Promise<any> {
        console.log("client");
        if (!listName || !listName.trim()) {
            throw new Error('Invalid `listName` parameter received.');
        }
        try {
            let apiURL = "/sites/" + (process.env.SPO_SITE_GUID || "") + "/lists/" + (process.env.SPO_LIST_GUID || "") + "/items/?expand=fields";
            log(apiURL);
            let res = await this.graphClient.api(apiURL).get();
            console.log("list found");
            return res;
        } catch (error) {
            console.log("list not found");
            return false;
        }
    }
}