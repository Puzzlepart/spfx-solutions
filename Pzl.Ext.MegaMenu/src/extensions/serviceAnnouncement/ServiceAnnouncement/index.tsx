import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { Dialog } from 'office-ui-fabric-react/lib/Dialog';
import { Web } from '@pnp/sp';
import { loadStyles } from '@microsoft/load-themed-styles';
import ServiceAnnouncementProps from './IServiceAnnouncementProps';
import { ServiceAnnouncementState, Announcement } from './IServiceAnnouncementState';
import * as strings from 'GlobalNavigationApplicationCustomizerStrings';
import styles from './ServiceAnnouncement.module.scss';
import { Alignment } from '../../TextAlignment';

export default class ServiceAnnouncement extends React.Component<ServiceAnnouncementProps, ServiceAnnouncementState> {

    constructor(props) {
        super(props);
        this.state = {
            isloading: true,
            modalShouldRender: false
        };
    }
    public componentWillMount() {
        this.fetchData();
    }
    public render() {
        let announcementModal = this.state.modalShouldRender && !this.props.isMobile ?
            <Dialog isOpen={this.state.modalShouldRender} title={this.state.modalAnnouncement.title} isBlocking={false} onDismiss={() => this.setState({ modalShouldRender: false })}>
                <div hidden={!this.state.modalAnnouncement.affectedSystems || this.state.modalAnnouncement.affectedSystems.length === 0}>
                    <h4>{strings.Field_AffectedSystems_Title}</h4>
                    <p>{this.state.modalAnnouncement.affectedSystems}</p>
                </div>
                <div hidden={!this.state.modalAnnouncement.content || this.state.modalAnnouncement.content.length === 0}>
                    <h4>{strings.Field_Description_Title}</h4>
                    <p dangerouslySetInnerHTML={{ __html: this.state.modalAnnouncement.content }}></p>
                </div>
                <div hidden={!this.state.modalAnnouncement.consequence || this.state.modalAnnouncement.consequence.length === 0}>
                    <h4>{strings.Field_Consequence_Title}</h4>
                    <p>{this.state.modalAnnouncement.consequence}</p>
                </div>
                <div hidden={!this.state.modalAnnouncement.responsible || this.state.modalAnnouncement.responsible.length === 0}>
                    <h4>{strings.Field_Responsible_Title}</h4>
                    <p>
                        <Persona
                            text={this.state.modalAnnouncement.responsible}
                            size={PersonaSize.size40}
                            secondaryText={strings.Dialog_Contact_SecondaryText}
                            className={styles.responsiblePersona}
                            title={strings.Responsible_Hover_Title}
                            onClick={() => window.location.href = `mailto:${this.state.modalAnnouncement.responsibleMail}?subject=${this.state.modalAnnouncement.title}`} />
                    </p>
                </div>
            </Dialog>
            : null;
        if (this.state.isloading || !this.state.AnnouncementsToBeShown) {
            return null;
        } else {
            let textClass = '';
            if (this.props.textAlignment == Alignment.Center) {
                textClass = styles.announcementMessageMiddle;
            }
            if (this.props.textAlignment == Alignment.Right) {
                textClass = styles.announcementMessageRight;
            }
            if (this.props.boldText) {
                textClass += " " + styles.announcementMessageBold;
            }


            let messageBars = this.state.AnnouncementsToBeShown.map((announcement, idx) => {
                // Set background with on-demand class as style={{}} doesn't work, and styles={{}} is not available in ouifr until v6
                const className = `${styles.announcementMessage}-${idx}`;
                if (announcement.customBgColor && announcement.customBgColor.length > 0) {
                    let style = `.${className} {background-color: ${announcement.customBgColor};}`;
                    loadStyles(style);
                }

                return <MessageBar className={`${styles.announcementMessage} ${className}`} messageBarType={announcement.getMessageBarType()} isMultiline={true} onDismiss={() => this.registerAnnouncementRead(announcement.id)}>
                    <div className={textClass} onClick={() => {
                        this.setState({ modalShouldRender: true, modalAnnouncement: announcement });
                        if (this.props.isMobile) {
                            this.renderMobileAnnouncementAlert(announcement);
                        }
                    }}>{announcement.title}</div>
                </MessageBar>;
            });

            return (
                <div>
                    {messageBars}
                    {announcementModal}
                </div>
            );
        }
    }

    /**
     * Workaround, as of 20.07.2018, there are issues with office ui fabric modal dialogs on mobile. 
     */
    private renderMobileAnnouncementAlert(announcement) {
        var affectedSystems = this.cleanHtmlFromTextString(announcement.affectedSystems);
        var content = this.cleanHtmlFromTextString(announcement.content);
        var consequence = this.cleanHtmlFromTextString(announcement.consequence);

        var alertContent = `${announcement.title}\n\n\n\n${strings.Field_AffectedSystems_Title}\n\n${affectedSystems}\n\n\n\n${strings.Field_Description_Title}\n\n${content}\n\n\n\n${strings.Field_Consequence_Title}\n\n${consequence}`;
        window.alert(alertContent);
    }

    private cleanHtmlFromTextString(fieldValue) {
        var htmlCleaningDomElement = document.createElement("span");
        htmlCleaningDomElement.innerHTML = fieldValue && fieldValue.length > 0 ? fieldValue : "";
        return htmlCleaningDomElement.textContent || htmlCleaningDomElement.innerText;
    }
    /**
     * Gets a JSON-object of IDs of seen announcements. Uses session storage if discardForSessionOnly is enabled
     */
    private getAnnouncementReadStorage() {
        if (this.props.discardForSessionOnly) {
            return sessionStorage.getItem('seenAnnouncements');
        }
        return localStorage.getItem('seenAnnouncements');
    }

    /**
     * Sets a JSON-object of IDs of seen announcements. Uses session storage if discardForSessionOnly is enabled
     */
    private setAnnouncementReadStorage(seenAnnouncements: string) {
        if (this.props.discardForSessionOnly) {
            sessionStorage.setItem('seenAnnouncements', seenAnnouncements);
        } else {
            localStorage.setItem('seenAnnouncements', seenAnnouncements);
        }
    }

    /**
     * Registers that an announcement has been read/discarded. Updated the local/session storage
     */
    private registerAnnouncementRead(announcementId: string) {
        let seenAnnouncements = JSON.parse(this.getAnnouncementReadStorage());
        if (!seenAnnouncements) {
            seenAnnouncements = {};
        }
        seenAnnouncements[announcementId] = true;
        this.setAnnouncementReadStorage(JSON.stringify(seenAnnouncements));

        let newAnnouncementList = this.state.AnnouncementsToBeShown;
        newAnnouncementList = newAnnouncementList.filter((item: Announcement) => {
            return !(item.id === announcementId);
        });
        this.setState({
            AnnouncementsToBeShown: newAnnouncementList,
        });
    }

    private async fetchData() {
        let now = new Date();
        let spWeb = new Web(`${document.location.protocol}//${document.location.hostname}${this.props.serverRelativeWebUrl}`);
        const severityFilter = this.props.announcementLevels ? " and (" + this.props.announcementLevels.split(",").map(level => { return "PzlSeverity eq '" + level + "'" }).join(" or ") + ")" : "";
        let announcements: any[] = await spWeb.getList(`${this.props.serverRelativeWebUrl.replace(/\/$/, "")}/${this.props.serviceAnnouncementListUrl}`)
            .items.select("ID",
                "Title",
                "PzlResponsible/Title",
                "PzlResponsible/EMail",
                "PzlSeverity",
                "PzlContent",
                "PzlConsequences",
                "PzlAffectedSystems",
                "PzlForceAnnouncement",
                "PzlStartDate",
                "PzlEndDate")
            .filter("(PzlStartDate le datetime'" + now.toISOString() + "') and (PzlEndDate ge datetime'" + now.toISOString() + "')" + severityFilter)
            .expand("PzlResponsible").usingCaching().get();

        let seenAnnouncements = JSON.parse(this.getAnnouncementReadStorage());
        let relevantAnnouncements: Announcement[] = announcements.filter((item) => {
            let previouslySeen = false;
            if (seenAnnouncements) {
                previouslySeen = seenAnnouncements[item.ID] ? true : false;
            }
            return !previouslySeen;
        }).map((item): Announcement => {
            // support having "Warning (#rgb)" for custom colors
            let colorMatch = item.PzlSeverity.match(/(.*?)\((.*?)\)/);
            let bgColor = '';
            if (colorMatch && colorMatch.length === 3) {
                item.PzlSeverity = colorMatch[1].trim();
                bgColor = colorMatch[2];
            }

            return {
                id: item.ID,
                title: item.Title,
                severity: item.PzlSeverity,
                forceDialog: item.PzlForceAnnouncement,
                content: item.PzlContent,
                consequence: item.PzlConsequences,
                affectedSystems: item.PzlAffectedSystems,
                startDate: item.PzlStartDate,
                endDate: item.PzlEndDate,
                responsible: item.PzlResponsible ? item.PzlResponsible.Title : "",
                responsibleMail: item.PzlResponsible ? item.PzlResponsible.EMail : "",
                customBgColor: bgColor,
                getMessageBarType: function () {
                    switch (this.severity) {
                        case "Information":
                        case "Informasjon":
                            return MessageBarType.info;
                        case "Warning":
                        case "Advarsel":
                            return MessageBarType.warning;
                        case "Alert":
                        case "Varsel":
                            return MessageBarType.severeWarning;
                        case "Normal":
                            return MessageBarType.success;
                        default:
                            return MessageBarType.info;
                    }
                }
            };
        });
        this.setState({ AnnouncementsToBeShown: relevantAnnouncements, isloading: false });
        
        if (this.state.AnnouncementsToBeShown) {
            this.state.AnnouncementsToBeShown.forEach((msg, idx) => {
                if (msg.forceDialog) {
                    this.setState({ modalShouldRender: true, modalAnnouncement: msg });
                }
            });
        }
    }
}
