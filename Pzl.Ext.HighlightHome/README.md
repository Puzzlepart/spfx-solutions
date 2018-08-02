# Hightlight Home Page by Puzzlepart

This is a field customizer which will highlight which file in `Site Pages` is the actual home page for the site.

## Installing the solution

- Upload the file `pzl-ext-highlight-home.sppkg` to either your tenant app catalog or to a site app catalog.
- Add the application named *Hightlight Home Page by Puzzlepart* to your site, and wait for install to finish.
- Make sure your user is a site administrator or owner.
- Navigate to the `Site Pages` library, and you should see the current home page being high lighted in the list of all pages in the site.

## Troubleshooting
If you don't see the page highlighted right away, refresh your page. Also make sure you access the site as a site administrator or owner in order for the field customizer to be installed.

## Technical details
The solution adds a field customizer to the field `LinkFileName` which is the backing field for the `Name` column. It's not possible to modify this field using `elements.xml` directly, so the application also includes a run-once application customizer which will install the field customizer, and then remove itself.

## Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```