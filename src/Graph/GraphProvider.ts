import { GraphClient } from './GraphClient';
import axios from "axios";
import qs = require('qs');
import { IListItem } from "./IListItem";

import * as debug from "debug";
const log = debug("msteams");

export class GraphProvider {

    public static accessToken: any;

    /**
     * @param {string} emailAddress The email address of the user.
     */
    public static async userExists(emailAddress: string): Promise<boolean> {
        let result;
        this.accessToken = await this.getAccessToken();
        if (this.accessToken && this.accessToken != "error") {
            const client = new GraphClient(this.accessToken);
            result = await client.userExists(emailAddress);
        }
        else {
            result = "not found";
        }
        return result;
    }
    /**
     * @param {string} listName The email address of the user.
     */
    public static async getListItems(listName: string): Promise<IListItem[]> {
        let result: any;
        try {
            this.accessToken = await this.getAccessToken();
            if (this.accessToken && this.accessToken != "error") {
                const client = new GraphClient(this.accessToken);
                result = await client.getListItems(listName);
                let lists: Array<IListItem> = new Array<IListItem>();
                if (result.value) {
                    result.value.map((item: any) => {
                        try {
                            lists.push({
                                Id: item.fields.id,
                                Title: item.fields.Title,
                                Abstract: item.fields.Abstract,
                                BusinessScenario: item.fields.BusinessScenario,
                                SolnHighlights: item.fields.SolnHighlights,
                                SMEContacts: item.fields.SMEContacts.map((elem: any) => {
                                    return elem.Email;
                                }).join(", "),
                                Technology: item.fields.Technology,
                                CaseStudyURL: item.fields.CaseStudyURL.Url,
                                CreatedOn: item.fields.Created
                            });
                        } catch (error) {
                            console.log("Error on mapping the item ID:" + item.fields.id);
                        }
                    });
                }
                result = lists;
            }
            else {
                //result = "not found";
                result = undefined;
            }
            return result;
        } catch (error) {
            console.log("List item mapping error : ");
            console.log(error);
            return result;

        }
    }

    public static async getAccessToken(): Promise<any> {
        try {
            let result: any;
            console.log("getAccessToken method");
            //const TOKEN_ENDPOINT = "https://login.microsoftonline.com/" + process.env.TENANT_ID + "/oauth2/v2.0/token";
            const TOKEN_ENDPOINT = (process.env.TOKEN_ENDPOINT || "").replace("tenatid", process.env.TENANT_ID || "");
            const postData = {
                client_id: process.env.MicrosoftAppId,
                scope: process.env.MS_GRAPH_SCOPE,
                client_secret: process.env.MicrosoftAppPassword,
                grant_type: 'client_credentials'
            };

            axios.defaults.headers.post['Content-Type'] =
                'application/x-www-form-urlencoded';
            result = await axios
                .post(TOKEN_ENDPOINT, qs.stringify(postData))
                .then(response => {
                    console.log("Token success");
                    return response.data.access_token;
                })
                .catch(error => {
                    console.log("Token error");
                    console.log(JSON.stringify(error));
                    return "error";
                });
            return result;
        } catch (error) {
            console.log("Fetch Token error : ");
            console.log(error);
            return "error on fetch Access Token";

        }
    }
}