import { Alignment } from "../../TextAlignment";

export default interface IServiceAnnouncementProps {
    serverRelativeWebUrl: string;
    serviceAnnouncementListUrl: string;
    discardForSessionOnly: boolean;
    isMobile: boolean;
    textAlignment: Alignment;
    boldText: boolean;
    announcementLevels: string;
}