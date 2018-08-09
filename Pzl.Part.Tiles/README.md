# Modern Tiles 

## Summary

This webpart makes it possible to create Tiles in your modern intranet. Just connect the webpart to a list using the webpart property pane and specify which fields the webpart should use to render the following values **Description**, **Background image**, **Open in new tab**, **Url** and **Order**. Also set a fallback image if no background image is set. You can find these settings under **More settings** in the webpart property pane.

If you want to have multiple Tile webparts in your site, but only wish to persist it in one list - just create a choice column in your source list and set the **Choice field** and **Tile type** properties in your webpart property pane. Try it out!

### Building the code

```bash
git clone https://github.com/Puzzlepart/spfx-solutions.git
npm i
gulp --ship
gulp package-solution --ship
```

###

### Installing
* Copy `pzl-part-tiles.sppkg` from `sharepoint\solution` and install it in your tenant.
* Install the app in your site
* Enjoy the tiles!

