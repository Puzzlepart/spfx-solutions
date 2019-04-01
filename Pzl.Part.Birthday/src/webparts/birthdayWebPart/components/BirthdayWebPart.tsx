import * as React from 'react';
import styles from './BirthdayWebPart.module.scss';
import { IBirthdayWebPartProps } from './IBirthdayWebPartProps';
import * as moment from 'moment';
import { IBirthdayState } from './IBirthdayState';
import { sp, SearchQuery, SortDirection, ISearchQueryBuilder, SearchQueryBuilder, SearchResults, Item } from '@pnp/sp';
import { escape, times } from '@microsoft/sp-lodash-subset';
import { authentication } from '@microsoft/teams-js';
import { SharePointPageContextDataProvider } from '@microsoft/sp-page-context';
import { IUser } from './IUser';

export default class BirthdayWebPart extends React.Component<IBirthdayWebPartProps, IBirthdayState> {
  private _users: IUser[] = [];
  constructor(props: IBirthdayWebPartProps) {
    super(props);
    this.state = { 
      items: [],
      happyBirthday: false 
    };
  }

  public async componentDidMount() {
    await this.fetchBirthdayData();
    console.log(this.state.items);
  }

  public componentWillUpdate(nextProps) {
    if (nextProps != this.props) {
      this.fetchBirthdayData();
    }
  }

  private async search() {
    const _searchQuerySettings: SearchQuery = {
      TrimDuplicates: false,
      RowLimit: 100,
      SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
      SelectProperties: ['PictureURL', 'PreferredName', 'Department', 'JobTitle', 'Birthday', 'Birthyear'],
    };
    const builder: ISearchQueryBuilder = SearchQueryBuilder('* Birthday>=1/1/2000', _searchQuerySettings);
    let result = await sp.search(builder);
    return result.PrimarySearchResults;
  }

  private async fetchBirthdayData() {
    let _birthdays: IUser[], _desemberBirthdays: IUser[];
    let items = await this.search();
    console.log(items);
    if (items && items.length > 0) {
      _birthdays = [], _desemberBirthdays = [];
      let today = moment.utc(new Date);
      let todayAsString = moment(today).format("YYYY-MM-DD");
      let thisYear: any = moment(today).format("YYYY");

      items.forEach((item: any) => {
        let birthday = moment.utc(item.Birthday);
        //let birthday = moment('2000-04-01', 'YYYY-MM-DD');

        // Finding days between today and birthday
        let birthdayThisYear = moment(birthday).year(thisYear);
        let bdayThisYearAsString = moment(birthdayThisYear).format("YYYY-MM-DD");
        if (moment(bdayThisYearAsString).isBefore(todayAsString)){
          birthdayThisYear =  moment(bdayThisYearAsString, "YYYY-MM-DD").add('years', 1);
          bdayThisYearAsString = moment(birthdayThisYear).format("YYYY-MM-DD");
        }
        let daysbetween = moment(bdayThisYearAsString).diff(moment(todayAsString), 'days');
        // if birthsday is today
        if (daysbetween === 0){
          this.setState({happyBirthday: true});
        }
       
        // Checking age of person, returns emty string if none
        const old:number = Number(thisYear) - Number(item.birthyear);
        const oldHtml = old > 0 ? ', fyller ' + old + ' år' : "";
        const  bdayDayMonth:any = moment(birthday).format("DD.MM");
        
        // Pushing person to array
        this._users.push({ key: item.PreferredName, userImage: item.PictureURL ,userName: item.PreferredName, jobDescription: item.JobTitle, department: item.Department, daysleft: (daysbetween + '\n dager').toString(), day: bdayDayMonth, birthyear: oldHtml.toString() , birthday: moment.utc(item.Birthday).local().format() });

        // Filter and sorting user array
        if (moment().format('MM') === '12'){
          _desemberBirthdays = this._users.filter((s) => {
            let _thisMonth = moment(s.birthday, ["MM-DD-YYYY", "YYYY-MM-DD", "DD/MM/YYYY", "MM/DD/YYYY"]).format('MM');
            return (_thisMonth === '12');
          });
          _desemberBirthdays = this.sortBirthdays(_desemberBirthdays);

          _birthdays = this._users.filter((s) => {
            let _thisMonth = moment(s.birthday, ["MM-DD-YYYY", "YYYY-MM-DD", "DD/MM/YYYY", "MM/DD/YYYY"]).format('MM');
            return (_thisMonth !== '12');
          });
          _birthdays = this.sortBirthdays(_birthdays);

          this._users = _desemberBirthdays.concat(_birthdays);

        } else {
          this._users = this.sortBirthdays(this._users);
        }
      });
      this.setState(
        { items: this._users }
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

  public render(): React.ReactElement<IBirthdayWebPartProps> {
    const icon:any = require('./img/cake_birthday.png');
    return (
      <div className={styles.birthdayWebPart}>
        <span className={styles.title}>Bursdager</span>
        <table>
          {(this.state.items) ? this.state.items.map(({ userImage, userName, jobDescription, department, day, daysleft, birthyear }, index) => (
            <tr key={`${index}`}>
              <td><img src={userImage} alt={userName} className={styles.userImage} /></td>
              <td className={styles.userInfo}><span className={styles.userName}>{userName}</span><br />
              <span className={styles.userBirthdayInfo}>{day}{birthyear}</span>
              </td>
              <td className={styles.userJob}>{jobDescription}<br />{department}</td>
              <td className={styles.daysleft}>
              {
                !this.state.happyBirthday ?
                daysleft
                :               
                <img src={icon} className={styles.icon} alt='' />     
              }          
             </td>
            </tr>
          )) : null}
        </table>
      </div>
    );
  }
}