import * as moment from 'moment';
import { Moment } from 'moment';

export interface OnlineMeetingInput {
  subject?: string;
  startDateTime: Moment;
  endDateTime: Moment;
  smeEmailID: string;
}

export interface OnlineMeeting {
  id: string;
  joinWebUrl: string;
  subject: string;
  videoTeleconferenceId: string;
  creationDateTime: Moment;
  startDateTime: Moment;
  endDateTime: Moment;
  dialinUrl: string;
  conferenceId: string;
  tollNumber: string;
  tollFreeNumber: string;
  preview: string;
}

export interface OutlookEventInfo {
  id: string;
  joinWebUrl: string;
  subject: string;
  creationDateTime: Moment;
  startDateTime: Moment;
  endDateTime: Moment;
  attendees: string;
  preview: string;
}

export function createDefaultMeetingInput(): OnlineMeetingInput {
  return {
    subject: '',
    startDateTime: moment()
      .startOf('hour')
      .add(1, 'hour'),
    endDateTime: moment()
      .startOf('hour')
      .add(2, 'hour'),
    smeEmailID:"nagarajan_s05@msnextlife.com"
  };
}

export interface IListItem {
  Id: string;
  Title: string;
  Abstract: string;
  BusinessScenario: string;
  SolnHighlights: string;
  SMEContacts: string;
  Technology: string;
  CaseStudyURL: string;
  CreatedOn: string;
}

export interface IUserInfo {
  UserID: string;
  FullName: string;
  EmailId: string;
  Skills: string;
}

export interface IProjectInfo {
  Id: string;
  Title: string;
  Abstract: string;
  BusinessScenario: string;
  SolnHighlights: string;
  SMEContacts: string;
  Technology: string;
  CaseStudyURL: string;
  CreatedOn: string;
}
