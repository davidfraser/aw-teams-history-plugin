const WebpackUserscript = require('webpack-userscript')
const path = require('path')

module.exports = env => {
    const isProduction = process.env.NODE_ENV === 'production'
    return {
        plugins: [
            new WebpackUserscript({
                headers: {
                    match: "https://teams.microsoft.com/*",
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