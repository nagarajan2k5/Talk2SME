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
        log("client");
        if (!emailAddress || !emailAddress.trim()) {
            throw new Error('Invalid `emailAddress` parameter received.');
        }
        try {
            const user: User = await this.graphClient.api(`/users/${emailAddress}`).get();
            log("user found");
            return user ? true : false;
        } catch (error) {
            log("user not found");
            return false;
        }
    }

    /**
    * Check if a user exists
    * @param {string} listName Email address of the email's recipient.
    */
    public async getListItems(listName: string): Promise<any> {
        log("client");
        if (!listName || !listName.trim()) {
            throw new Error('Invalid `listName` parameter received.');
        }
        try {

            let res = await this.graphClient.api('/sites/1976f49c-7a98-4f8b-8da4-18f0a2cdce89/lists/00cfb8d5-9e98-4fbf-889d-5fc351c6c30f/items/?expand=fields').get();
            log("list found");
            return res;
        } catch (error) {
            log("list not found");
            return false;
        }
    }
}