import * as React from 'react';
import styles from './BirthdayWebPart.module.scss';
import { IBirthdayWebPartProps } from './IBirthdayWebPartProps';
import * as moment from 'moment';
import { IBirthdayState } from './IBirthdayState';
import { sp, SearchQuery, SortDirection, ISearchQueryBuilder, SearchQueryBuilder, SearchResults, Item } from '@pnp/sp';
import { WebPartTitle} from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { escape, times } from '@microsoft/sp-lodash-subset';
import { authentication } from '@microsoft/teams-js';
import { SharePointPageContextDataProvider } from '@microsoft/sp-page-context';
import { IUser } from './IUser';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { dataFile } from './data/dataFile';
import * as strings from 'BirthdayWebPartWebPartStrings';

export default class BirthdayWebPart extends React.Component<IBirthdayWebPartProps, IBirthdayState> {
  private _users: IUser[] = [];
  constructor(props: IBirthdayWebPartProps) {
    super(props);
    this.state = { 
      items: []
    };
  }
  public async componentDidMount() {
    await this.fetchBirthdayData();
  }
  public render(): React.ReactElement<IBirthdayWebPartProps> {
    return (
      <div className={styles.birthdayWebPart}>
        <span className={styles.title}>
          <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} />        
        </span>
        <table>
          {(this.state.items) ? this.state.items.map(({ userEmail, userName, jobDescription, department, day, daysleft, years }, index) => (
            <tr key={`${index}`}>
              <td><img src={`/_layouts/15/userphoto.aspx?size=S&username=${userEmail}`} alt={userName} className={styles.userImage} /></td>
              <td className={styles.userInfo}><span className={styles.userName}>{userName}</span><br />
              <span className={styles.userBirthdayInfo}>{day}{(years<100)? `, ${strings.TextBecomes} ${years} ${strings.TextYears}`: ``}</span>
              </td>
              <td className={styles.userJob}>{jobDescription}<br />{department}</td>
              <td className={styles.daysleft}>
              {
                (daysleft === 0) ?
                  <Icon className={styles.icon} iconName={strings.CakeIconName} />  
                :               
                `${daysleft} ${strings.TextDays}`
              }          
              </td>
            </tr>
          )) : null}
        </table>
      </div>
    );
  }

  private async search() {
    const _searchQuerySettings: SearchQuery = {
      TrimDuplicates: false,
      RowLimit: 100,
      SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
      SelectProperties: ['WorkEmail', 'PreferredName', 'Department', 'JobTitle', 'Birthday', 'Birthyear'],
    };
    const builder: ISearchQueryBuilder = SearchQueryBuilder('* Birthday>=1/1/2000', _searchQuerySettings);
    let result = await sp.search(builder);
    return result.PrimarySearchResults;
  }

  private async fetchBirthdayData() {
    let people = await this.search();
    if (people && people.length > 0) {
      const today = moment.utc(new Date).format("YYYY-MM-DD");
      const thisYear: any = moment(today).format("YYYY");

      people.forEach((person: any) => {
        let birthday = moment.utc(person.Birthday).year(thisYear);
        let birthdayAsString = moment(birthday).format("YYYY-MM-DD");
        let nextBirthdayYear:number = Number(thisYear);

        // People having birthday next year, for sorting and numbers
        if (moment(birthdayAsString).isBefore(today)){
          birthday =  moment(birthdayAsString, "YYYY-MM-DD").add('years', 1);
          birthdayAsString = moment(birthday).format("YYYY-MM-DD");
          nextBirthdayYear = nextBirthdayYear + 1;
        }
        const daysbetween = moment(birthdayAsString).diff(moment(today), 'days');
        const old:number = nextBirthdayYear - Number(person.Birthyear);      
        this._users.push({ key: person.PreferredName, userEmail: person.WorkEmail ,userName: person.PreferredName, jobDescription: person.JobTitle, department: person.Department, daysleft: daysbetween, day: moment(birthday).format("DD.MM"), years: old , birthday: birthday.local().format() });
        this._users = this.sortBirthdays(this._users);
      });
      const items = this._users.splice(0, this.props.itemsCount);
      this.setState(
        { items }
      );
    }
  }

  // Sort Array of Birthdays
  private sortBirthdays(users: IUser[]) {
    return users.sort( (a, b) => {
      if (a.birthday > b.birthday) {
        return 1;
      }
      if (a.birthday < b.birthday) {
        return -1;
      }
      return 0;
    });
  }
}