import * as React from 'react';
import { useCallback, useState } from 'react';
import * as strings from 'PageExpiredWebPartStrings';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

export interface IPageExpiredProps {
  modified: Date;
  expireAfter: number;
  isEditor: boolean;
  verify(ev: unknown): void;
}

function addDays(date: Date, days: number): Date {
  date.setDate(date.getDate() + days);
  return date;
}

function format(days: number): string {
  if (18 * 30 < days) {
    const years = ~~(days / 365);
    return `${years} ${strings.YearsAgo}`
  } else if (11 * 30 < days) {
    return strings.AYearAgo;
  }
  else if (46 < days) {
    const months = ~~(days / 30);
    return `${months} ${strings.MonthsAgo}`
  }
  else if (26 < days) {
    return strings.AMonthAgo;
  }
  return `${days} ${strings.DaysAgo}`;
}

export const PageExpired: React.FunctionComponent<IPageExpiredProps> = (props) => {

  const today = new Date();
  const expiryDate = addDays(props.modified, props.expireAfter);
  const daysSinceModified = ~~((today.getTime() - props.modified.getTime()) / (1000 * 60 * 60 * 24));

  const sessionKey = `Ignore.${window.location.pathname}`;
  const [ignored, setIgnored] = useState<boolean>(sessionStorage.getItem(sessionKey) !== null);

  const ignore = useCallback((ev) => {
    sessionStorage.setItem(sessionKey, "Ignore");
    setIgnored(true);
  }, []);

  return (!ignored && expiryDate < today ? props.isEditor ? <>
    <MessageBar
      messageBarType={MessageBarType.warning}
      isMultiline={true}
      actions={
        <div>
          <PrimaryButton onClick={(ev) => props.verify(ev)}>{strings.Verify}</PrimaryButton>
          <DefaultButton onClick={ignore}>{strings.Ignore}</DefaultButton>
        </div>}>
      <div>
        <strong>{strings.PageWasPublished} {format(daysSinceModified)}</strong>
        <p>{strings.ExpirationMessage}</p>
      </div>
    </MessageBar>
  </> : <MessageBar>{strings.PageWasPublished} {format(daysSinceModified)}</MessageBar> : <></>);
};



