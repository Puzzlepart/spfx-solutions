## pzl-part-birthday

If you need a birthday webpart to display upcomming birthdays, this is what you are looking for. You have option to set web part header and the number of persons to list.

![Demo](./Preview.png "Demo")


### Prerequisites

The web part is using the Local People results query for fetching data. Getting the properties for: 'WorkEmail', 'PreferredName', 'Department', 'JobTitle'.

For the web part to work you must add the clawled value for SPS-Birthday to a managed property 'RefinableDateXX' (You can use any RefinableDate. For exampel RefinableDate00). This can be set in the Search Schema settings.

IMPORTANT!
Set the Alias in RefinableDate00 (in the used RefinableDateXX) property to 'Birthday'.

OPTIONAL
By adding a year of birth property to the user profile you can get the age display for every person in the web part.
You must then in the Search Schema settings add the crawled value of year of birth to a managed property of 'RefinableStringXX' and set the alias to 'Birthyear'.


### Building the code

```bash
git clone the repo
npm i
gulp bundle --ship
gulp package-solution --ship
```

This package produces the following:

* sharepoint/solution/pzl-part-birthday.sppkg - package to install in the App Catalog