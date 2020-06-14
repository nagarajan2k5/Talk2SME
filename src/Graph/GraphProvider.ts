import { GraphClient } from './GraphClient';
import { PnPJsClient } from "./PnPJsClient";
import axios from "axios";
import qs = require('qs');
import {IListItem, IUserInfo, IProjectInfo } from "../Models/models";

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
     * @param {string} keyWord The email address of the user.
     */
    public static async getListItems(keyWord: string): Promise<IListItem[]> {
        let result: any;
        try {
            this.accessToken = await this.getAccessToken();
            if (this.accessToken && this.accessToken != "error") {
                const client = new GraphClient(this.accessToken);
                result = await client.getListItems(keyWord);
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

    /**
     * @param {string} skill to search users.
     */
    public static async searchPeopleBySkills(skills: string): Promise<IUserInfo[]> {
        let users: Array<IUserInfo> = new Array<IUserInfo>();
        try {
            const client = new PnPJsClient();
            if (client) {
                let result = await client.searchPeopleBySkills(skills);
                //log("Result: " + JSON.stringify(result));
                if (result) {
                    result.PrimarySearchResults.map((item: any) => {
                        try {
                            users.push({
                                UserID: item.AccountName,
                                FullName: item.PreferredName,
                                EmailId: item.WorkEmail,
                                Skills: (item.Skills).replace(";", ", ")
                            });
                        } catch (error) {
                            console.log("SPO Search on mapping the item ID:" + item.fields.id);
                        }
                    });
                }
                log("Users: " + JSON.stringify(users));
                return users;
            }
        } catch (error) {
            console.log("SPO search item mapping error : ");
            console.log(error);
        }
        return users;
    }

    /**
    * @param {string} keyword to search projects.
    */
    public static async searchProjectsByKeyword(keyword: string): Promise<IProjectInfo[]> {
        let projs: Array<IProjectInfo> = new Array<IProjectInfo>();
        try {
            const client = new PnPJsClient();
            if (client) {
                console.log("Keyword: " + keyword);
                let result = await client.searchProjectsByKeyword(keyword);
                console.log("Result: " + JSON.stringify(result));
                if (result) {
                    result.PrimarySearchResults.map((item: any) => {
                        try {
                            projs.push({
                                Id: item.Title,
                                Title: item.Title,
                                Abstract: item.Title,
                                BusinessScenario: item.Title,
                                SolnHighlights: item.Title,
                                SMEContacts: item.Title,
                                Technology: item.Title,
                                CaseStudyURL: item.Title,
                                CreatedOn: item.Title
                            });
                        } catch (error) {
                            console.log("SPO Search on mapping the item ID:" + item.fields.id);
                        }
                    });
                }
                log("Projs: " + JSON.stringify(projs));
                return projs;
            }
        } catch (error) {
            console.log("SPO search item mapping error : ");
            console.log(error);
        }
        return projs;
    }

    /**
   * Check if a user exists
   * @param {string} userId Email address of the email's recipient.
   * @param {string} skill Skill to update
   */
    public static async updateSkillProficiency(userId: string, skill: string): Promise<any> {
        let result: any;
        try {
            this.accessToken = await this.getAccessToken();
            if (this.accessToken && this.accessToken != "error") {
                const client = new GraphClient(this.accessToken);
                result = await client.updateSkillProficiency(userId, skill);
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
            console.log("Url: " + TOKEN_ENDPOINT);
            axios.defaults.headers.post['Content-Type'] =
                'application/x-www-form-urlencoded';
            result = await axios
                .post(TOKEN_ENDPOINT, qs.stringify(postData))
                .then(response => {
                    console.log("Token success: " + response.data.access_token);
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