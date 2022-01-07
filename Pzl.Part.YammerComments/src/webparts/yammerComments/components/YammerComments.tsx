import * as React from 'react';
import { useState, useEffect } from 'react';

import styles from './YammerComments.module.scss';
import { IYammerCommentsProps } from './IYammerCommentsProps';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';


const YammerComments: React.FunctionComponent<IYammerCommentsProps> = (props) => {

  const [community, setCommunity] = useState<string>(props.communityId);
  const [comment, setComment] = React.useState("");
  const [messageBarStatus, setMessageBarStatus] = React.useState({
    type: MessageBarType.info,
    message: <span></span>,
    show: false
  });

  useEffect(() => {
    setCommunity(props.communityId);
  }, [props]);

  useEffect(() => {
    (async () => {
      try {
        let webLink = await props.yammerService.getWebLink();
      } catch (error) {
        if ('InteractionRequiredAuthError' === error.name) {
          setMessageBarStatus({
            type: MessageBarType.error,
            message: <span>{error.message}</span>,
            show: true
          });
        }
      }
    })();

  }, []);

  return (
    <div className={styles.YammerComments}>
      <div>
        {
          (messageBarStatus.show) &&
          <MessageBar messageBarType={messageBarStatus.type}>{
            messageBarStatus.message 
          }</MessageBar>
        }
      </div>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.column}>
            <span className={styles.title}>Welcome to SharePoint!</span>
            <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
            <p className={styles.description}>{props.communityId}</p>
          </div>
        </div>
      </div>
    </div>
  );

};

export default YammerComments;
