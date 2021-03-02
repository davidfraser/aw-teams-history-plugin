const WebpackUserscript = require('webpack-userscript')

module.exports = {
  plugins: [
    new WebpackUserscript({
        headers: {
            match: "https://teams.microsoft.com/*"
        }
    })
  ]
}