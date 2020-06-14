import { OnlineMeetingInput, OnlineMeeting, OutlookEventInfo } from '../Models/models';
//import { msalApp } from '../auth/msalApp';
import axios from 'axios';
import * as moment from 'moment';

import * as debug from "debug";
import { json } from 'express';
const log = debug("msteams");

export function createMeetingService() {
    return {
        async createMeeting(meeting: OnlineMeetingInput, accessToken: string) {
            /*
            const requestBody = {
                startDateTime: meeting.startDateTime?.toISOString(),
                endDateTime: meeting.endDateTime?.toISOString(),
                subject: meeting.subject
            };
            
            log("token: " + accessToken);
            log("Graph call: Me");
            const resMe = await axios.get('https://graph.microsoft.com/v1.0/me',
                {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                        'Content-type': 'application/json'
                    }
                });
            log("Response: " + JSON.stringify(resMe.data));

            log("Graph call: onlineMeetings");
            const response = await axios.post(
                'https://graph.microsoft.com/v1.0/me/onlineMeetings',
                requestBody,
                {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                        'Content-type': 'application/json'
                    }
                }
            );
            log("Response: " + JSON.stringify(response.data));
             // const preview = decodeURIComponent(
            //     (response.data.joinInformation.content?.split(',')?.[1] ?? '').replace(
            //         /\+/g,
            //         '%20'
            //     )
            // );

            // const createdMeeting = {
            //     id: response.data.id,
            //     creationDateTime: moment(response.data.creationDateTime),
            //     subject: response.data.subject,
            //     joinUrl: response.data.joinUrl,
            //     joinWebUrl: response.data.joinWebUrl,
            //     startDateTime: moment(response.data.startDateTime),
            //     endDateTime: moment(response.data.endDateTime),
            //     conferenceId: response.data.audioConferencing?.conferenceId || '',
            //     tollNumber: response.data.audioConferencing?.tollNumber || '',
            //     tollFreeNumber: response.data.audioConferencing?.tollFreeNumber || '',
            //     dialinUrl: response.data.audioConferencing?.dialinUrl || '',
            //     videoTeleconferenceId: response.data.videoTeleconferenceId,
            //     preview
            // } as OnlineMeeting;

            */

            log("Graph call: Create Outlook Event");

            log(JSON.stringify(meeting));

            const eventBody = {
                subject: meeting.subject,
                body: {
                    contentType: "HTML",
                    content: "Does this time work for you?"
                },
                start: {
                    dateTime: meeting.startDateTime?.toISOString(),
                    timeZone: "India Standard Time"
                },
                end: {
                    dateTime: meeting.endDateTime?.toISOString(),
                    timeZone: "India Standard Time"
                },
                location: {
                    displayName: "Microsoft Teams"
                },
                attendees: [
                    {
                        emailAddress: {
                            address: meeting.smeEmailID//,
                            //name: "User2 Reader"
                        },
                        type: "required"
                    }
                ],
                allowNewTimeProposals: true,
                isOnlineMeeting: true,
                onlineMeetingProvider: "teamsForBusiness"
            };
            const response = await axios.post(
                'https://graph.microsoft.com/v1.0/me/events',
                eventBody,
                {
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                        'Content-type': 'application/json'
                    }
                }
            );
            //log("Response: " + JSON.stringify(response.data));

            const preview = decodeURIComponent(
                response.data.onlineMeeting.joinUrl.replace(
                    /\+/g,
                    '%20'
                ));
            

            const createdMeeting = {
                id: response.data.id,
                creationDateTime: moment(response.data.creationDateTime),
                subject: response.data.subject,
                joinWebUrl: response.data.onlineMeeting.joinUrl,
                startDateTime: moment(response.data.start.dateTime),
                endDateTime: moment(response.data.end.dateTime),
                attendees: response.data.attendees[0]?.emailAddress.address || '',
                preview
            } as OutlookEventInfo;

            return createdMeeting;
        }
    };
}
