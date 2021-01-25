'use strict';

const gulp = require('gulp');
const path = require('path');
const build = require('@microsoft/sp-build-web');

let copyIcons = build.subTask('copy-icons', function(gulp, buildOptions, done) {
    gulp.src('./*.svg')
        .pipe(gulp.dest('./temp/deploy'));
    done();
});
build.rig.addPostBuildTask(copyIcons);

const bundleAnalyzer = require('webpack-bundle-analyzer');

build.configureWebpack.mergeConfig({
    additionalConfiguration: (generatedConfiguration) => {
        const lastDirName = path.basename(__dirname);
        const dropPath = path.join(__dirname, 'temp', 'stats');
        generatedConfiguration.plugins.push(new bundleAnalyzer.BundleAnalyzerPlugin({
            openAnalyzer: false,
            analyzerMode: 'static',
            reportFilename: path.join(dropPath, `${lastDirName}.stats.html`),
            generateStatsFile: false,
            logLevel: 'error'
        }));

        return generatedConfiguration;
    }
});


build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.initialize(gulp);