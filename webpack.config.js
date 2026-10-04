const { UserscriptPlugin } = require('webpack-userscript')
const path = require('path')

module.exports = env => {
    const isProduction = process.env.NODE_ENV === 'production'
    return {
        plugins: [
            new UserscriptPlugin({
                headers: {
                    match: [
                        "https://teams.microsoft.com/*",
                        "https://teams.cloud.microsoft/*",
                        // The new Teams client embeds the calendar as an Outlook web app
                        // in a cross-origin iframe; we inject there to read its event
                        // cache. Scoped to /hosted/ so we don't run on standalone Outlook.
                        "https://outlook.office.com/hosted/*",
                    ],
                    grant: ["GM_xmlhttpRequest", "GM_registerMenuCommand"],
                    connect: "*"
                }
            })
        ],
        resolve: {
            alias: {
                'aw-config$': path.resolve(__dirname, 'src', isProduction ? 'aw-prod-config.js' : 'aw-dev-config.js'),
            }
        },
        devtool: 'cheap-source-map',
    }
}