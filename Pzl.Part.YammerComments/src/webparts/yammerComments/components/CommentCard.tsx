import * as React from 'react';
import { useState } from 'react';
import { Persona, PersonaSize, PrimaryButton, TextField } from 'office-ui-fabric-react';
import IUser from '../interfaces/IUser';
import styles from './CommentCard.module.scss';
import IComment from '../interfaces/IComment';
import * as strings from 'YammerCommentsWebPartStrings';
import { IYammerService } from '../services/YammerService';

export interface ICommentCardProps {
    user: IUser;
    comment: IComment;
    onNewComment(comment: IComment): void;
}

export const CommentCard: React.FunctionComponent<ICommentCardProps> = (props) => {

    const [comment, setComment] = useState<string>('');
    const [isPosting, setIsPosting] = useState<boolean>(false);

    const onCommentChange = React.useCallback(
        (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
            setComment(newValue || '');
        },
        [],
    );

    const postComment = async (): Promise<void> => {
        setIsPosting(true);
        props.onNewComment({
            text: comment,
            groupId: props.comment.groupId,
            replyToId: props.comment.replyToId
        });
        setComment('');
        setIsPosting(false);
    };

    return (
        <>{props.user &&
            <div className={styles.CommentCard}>
                <div className={styles.CommentUser}>
                    <Persona
                        imageUrl={`https://${window.location.hostname}/_layouts/15/userphoto.aspx?size=L&accountname=${props.user.email}`}
                        size={PersonaSize.size32}
                    />
                </div>
                <div className={styles.CommentGap}></div>
                <div className={styles.CommentArea}>
                    <TextField placeholder='Add a comment. Type @ to mention someone' onChange={onCommentChange} multiline autoAdjustHeight />
                </div>
                <div className={styles.CommentGap}></div>
                <div className={styles.CommentButton}>
                    <PrimaryButton text={isPosting ? strings.Posting : strings.Post} disabled={isPosting || comment.length < 1} onClick={postComment} />
                </div>
            </div>
        }
        </>
    );
};
