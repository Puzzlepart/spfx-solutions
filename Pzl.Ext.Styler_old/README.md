# Pzl Styler
This Application customizer lets you override any and all CSS on SharePoint pages across the site collection. Place your CSS file named `PzlStyler.css` in a document library called "Styling" placed under /sites/CDN/ 

# Prereqs 
node version: 16.13.0
If you are using nvm, you can switch to the correct node version by running `nvm use`

# Building the code

```bash
npm install
```

```bash
gulp bundle --ship
gulp package-solution --ship
```

# Watching the code 
```bash
gulp serve
```


