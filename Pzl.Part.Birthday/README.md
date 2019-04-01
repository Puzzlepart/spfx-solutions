## pzl-part-birthday

This is where you include your WebPart documentation.

### Prerequisites
For at bursdag skal fanges opp av appen må SPS-Birthday legges til i RefinableDate-- (samme hvilken) og alias settes til "Birthday".

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
