
const webpack = require('webpack');

const path = require("path");


var config = {

    "entry" : {

        "content" : "./src/entry/content/index.ts" ,

        "background" : "./src/entry/background/index.ts",

        "popup" : "./src/entry/popup/index.ts",
        "proxy" : "./src/entry/proxy/index.ts",
    } ,

    "output": {

        "path": path.resolve(__dirname, "dist"),

        "filename" : "[name].js",
    } ,

    "module" : {

        "rules" : [
            {
                "test": /\.ts$/,

                "use": [
                    {
                        "loader": "ts-loader",
                    }
                ],
            },

            {
                "test" : /\.css$/,

                "use": [

                    {
                        "loader": "style-loader",
                    },

                    {
                        "loader": "css-loader",
                        "options": {
                            "modules" : true
                        }
                    }
                ]
            }
        ]
    } ,

    "resolve" : {

        "extensions" : [

            // load .ts files as modules
            ".ts",
        ]
    },

    "optimization" : {

        "minimizer" : []
    },

    "plugins" : [
        new (class {
      apply(compiler) {
 compiler.hooks.emit.tap("copy-static", (compilation) => {
          const fs = require("fs");
          const addFile = (abs, rel) => {
            compilation.assets[rel] = {
              source: () => fs.readFileSync(abs),
              size: () => fs.statSync(abs).size,
            };
          };
          addFile(path.join(__dirname, "manifest.json"), "manifest.json");
          addFile(path.join(__dirname, "src/entry/popup/popup.html"), "popup.html");
        });
      }
    })(),

    ]
}

module.exports = config