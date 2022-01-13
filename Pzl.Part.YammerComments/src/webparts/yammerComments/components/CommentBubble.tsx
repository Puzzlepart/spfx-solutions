import * as React from 'react';
import * as dayjs from 'dayjs';
import * as relativeTime from 'dayjs/plugin/relativeTime';
import { Icon, Persona, PersonaSize } from 'office-ui-fabric-react';
import { ServiceScope } from "@microsoft/sp-core-library";
import { LivePersona } from "@pnp/spfx-controls-react/lib/controls/LivePersona";
import IComment from '../interfaces/IComment';
import styles from './CommentBubble.module.scss';

export interface ICommentBubbleProps {
    serviceScope: ServiceScope;
    comments: IComment[];
}

export const CommentBubble: React.FunctionComponent<ICommentBubbleProps> = (props) => {

    dayjs.extend(relativeTime);

    const createThread = (comment: IComment, level: number) => {

        const style = { paddingLeft: 32 * level };

        return (
            <>
                <div className={styles.CommentBubble} style={style}>
                    <div className={styles.CommentUser}>
                        <LivePersona upn={comment.user.email} serviceScope={props.serviceScope} disableHover={false} template={
                            <Persona
                                imageUrl={`https://${window.location.hostname}/_layouts/15/userphoto.aspx?size=L&accountname=${comment.user.email}`}
                                size={PersonaSize.size32}
                            />
                        } />
                    </div>
                    <div className={styles.CommentGap}></div>
                    <div className={styles.CommentColumn}>
                        <div className={styles.CommentArea}>
                            <div className={styles.NameAndDate}>
                                <div className={styles.Name}>
                                    <LivePersona upn={comment.user.email} serviceScope={props.serviceScope} disableHover={false} template={
                                        <span className={styles.NameSpan}>{comment.user.name}</span>
                                    } />
                                </div>
                                <div>{dayjs(comment.created,).fromNow()}</div>
                            </div>
                            <div><span className={styles.CommentSpan}>{comment.text}</span></div>
                        </div>
                        <div className={styles.CommentActions} >

                            <Icon iconName='Like' className={styles.Icon}></Icon>
                            <Icon iconName='Comment' className={styles.Icon}></Icon>
                            <Icon iconName='More' className={styles.Icon}></Icon>

                        </div>
                    </div>
                </div>
                {
                    comment.replies.map((reply: IComment) => {
                        return createThread(reply, level + 1);
                    })
                }
            </>
        );
    };

    return (
        <div>
            {props.comments.map(comment => (
                createThread(comment, 0)
            ))}
        </div>
    );
};