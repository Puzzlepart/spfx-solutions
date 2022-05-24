import * as React from 'react';
import { useState, useEffect } from 'react';
import { IPropertyPaneAccessor } from '@microsoft/sp-webpart-base';
import { ServiceScope } from "@microsoft/sp-core-library";
import { Icon, MessageBar, MessageBarType, PrimaryButton } from 'office-ui-fabric-react';
import { IYammerService } from "../services/YammerService";
import * as strings from 'YammerCommentsWebPartStrings';
import styles from './YammerComments.module.scss';
import IUser from '../interfaces/IUser';
import IComment from '../interfaces/IComment';
import { CommentCard } from './CommentCard';
import { CommentBubble } from './CommentBubble';
import { CommentShimmer } from './CommentShimmer';

export interface IYammerCommentsProps {
  serviceScope: ServiceScope;
  propertyPane: IPropertyPaneAccessor;
  yammerService: IYammerService;
  community: any;
}

export const YammerComments: React.FunctionComponent<IYammerCommentsProps> = (props) => {

  const [timestamp, setTimestamp] = useState<number>(Date.now());
  const [comments, setComments] = useState<IComment[]>();
  const [community, setCommunity] = useState<string>(props.community);
  const [isLoading, setIsLoading] = useState<boolean>(props.community !== null);
  const [messageBarStatus, setMessageBarStatus] = useState({
    type: MessageBarType.info,
    message: <span></span>,
    show: false
  });
  const [user, setUser] = useState<IUser>();

  useEffect(() => {
    setCommunity(props.community);
  }, [props]);


  useEffect(() => {

    const fetchData = async () => {
      try {

        const currentUser = await props.yammerService.getCurrentUser();
        setUser(currentUser);

        const result = await props.yammerService.getOpenGraphObjects();
        if (result && result.ogos) {

          const arrayOfThreadIds: string[][] = await Promise.all(result.ogos.map(async ogo => { return await props.yammerService.getWebLinkMessages(ogo.id); }));
          let threadIds: string[] = [];
          arrayOfThreadIds.forEach(nestedArray => {
            threadIds = threadIds.concat(nestedArray);
          });
          const messageThreads = await Promise.all(threadIds.map(async threadId => { return await props.yammerService.getMessagesInThread(threadId); }));

          const messages: IComment[] = [];
          messageThreads.forEach(thread => {
            const tree = props.yammerService.buildHierarchy(thread);
            messages.push(tree);
          });

          setComments(messages);
          setIsLoading(false);
        }
      } catch (error) {
        if ('InteractionRequiredAuthError' === error.name) {
          setMessageBarStatus({
            type: MessageBarType.error,
            message: <span>{error.message}</span>,
            show: true
          });
        }
      }
    };

    fetchData();
  }, [timestamp]);

  function openPropertyPanel() {
    props.propertyPane.open();
  }

  const onNewComment = async (comment: IComment): Promise<void> => {
    let newComment = await props.yammerService.postComment(comment);
    setTimestamp(Date.now());
  };

  return (
    <>
      <div>
        {
          (messageBarStatus.show) &&
          <MessageBar messageBarType={messageBarStatus.type}>{
            messageBarStatus.message
          }</MessageBar>
        }
      </div>
      <div>
        {community &&
          <aside aria-label={strings.WebPartTitle}>
            <div className={styles.row}>
              <h2>{strings.WebPartTitle}</h2>
            </div>
            <div className={styles.row}>
              <div className={styles.column}>
                <CommentCard user={user} comment={{ text: '', groupId: props.community }} onNewComment={onNewComment} />
                {!isLoading && <CommentBubble serviceScope={props.serviceScope} comments={comments} />}
                {isLoading && <div>
                  <CommentShimmer isReply={false} />
                  <div>&nbsp;</div>
                  <CommentShimmer isReply={true} />
                </div>
                }
              </div>
            </div>
          </aside>
        }
      </div>
      {!community &&
        <div className={styles.YammerPlaceholder}>
          <div className={styles.container}>
            <div className={styles.head}>
              <div className={styles.headContainer}>
                <Icon iconName='YammerLogo' className={styles.icon}></Icon>
                <span className={`${styles.text} ${styles.headerFluent}`}>{strings.WebPartTitle}</span>
              </div>
            </div>
            <div className={styles.description}>
              <span className={`${styles.descriptionText} ${styles.textFluent}`}>{strings.WebPartDescription}</span>
            </div>
            <div className={styles.description}>
              <PrimaryButton text={strings.WebPartSetUp} onClick={openPropertyPanel} />
            </div>
          </div>
        </div>
      }
    </>
  );

};
