# Modern Tiles 

## Summary

This webpart makes it possible to create tiles in your modern intranet. Just connect the webpart to a list using the webpart property pane. Also specify which fields the webpart should use to render; Description, Background image, Open in new tab, Url and order. Also set a fallback image if no background image is set.

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

