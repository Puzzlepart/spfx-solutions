import { MessageBarType } from "office-ui-fabric-react/lib/MessageBar";

export interface Announcement {
    id: string;
    title: string;
    severity: string;
    forceDialog: boolean;
    content: string;
    consequence: string;
    affectedSystems: string;
    startDate: string;
    endDate: string;
    responsible: string;
    responsibleMail: string;
    customBgColor: string;
    infolink: {
        Description: string;
        Url: string;
    };
    
    getMessageBarType():MessageBarType;
}
export interface ServiceAnnouncementState {
    AnnouncementsToBeShown?: Announcement[];
    isloading?: boolean;
    modalShouldRender?: boolean;
    modalAnnouncement?: Announcement;
}

export interface User {
    name: string;
}