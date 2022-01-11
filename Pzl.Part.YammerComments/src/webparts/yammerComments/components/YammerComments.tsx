import * as React from 'react';
import { useState, useEffect } from 'react';
import { IPropertyPaneAccessor } from '@microsoft/sp-webpart-base';
import { Icon, MessageBar, MessageBarType, PrimaryButton } from 'office-ui-fabric-react';
import { IYammerService } from "../services/YammerService";
import styles from './YammerComments.module.scss';
import * as strings from 'YammerCommentsWebPartStrings';

export interface IYammerCommentsProps {
  propertyPane: IPropertyPaneAccessor;
  yammerService: IYammerService;
  community: any;
}

export const YammerComments: React.FunctionComponent<IYammerCommentsProps> = (props) => {

  const [community, setCommunity] = useState<string>(props.community);

  const [comment, setComment] = React.useState("");

  const [messageBarStatus, setMessageBarStatus] = React.useState({
    type: MessageBarType.info,
    message: <span></span>,
    show: false
  });

  useEffect(() => {
    console.log('Community:' + props.community);
    setCommunity(props.community);
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

  function openPropertyPanel() {
    props.propertyPane.open();
  }

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
                <span className={styles.title}>Welcome to SharePoint!</span>
                <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
                <p className={styles.description}>{props.community}</p>
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
