Krak Bot
========

Krak Bot an intelligent chat bot for [Microsoft Teams (free)](https://teams.live.com/) using Bing AI (Sydney) via [node-chatgpt-api](https://github.com/waylaidwanderer/node-chatgpt-api). You can ask any questions and get accurate answers straight from Microsoft Teams.

## Usage

1. Go to Krak Bot directory and install its dependencies: `npm install`
2. Register a free Microsoft account for Krak Bot (e.g. krakbot@example.com) and log in to [Microsoft Teams (free)](https://teams.live.com/)
3. Get an access token from Microsoft Teams by opening the Console (F12) and pasting: `JSON.parse(localStorage['ts.' + JSON.parse(localStorage['ts.userInfo']).userId + '.cache.token.service::api.fl.spaces.skype.com::MBI_SSL']).token`
4. Set the `ACCESS_TOKEN` environment variable with this token and run index.js: `ACCESS_TOKEN=... node index.js`
5. As a different user, open MS Teams and start a new chat with Krak Bot account (e.g. krakbot@example.com)

## License

&copy; 2023 Jerzy Głowacki under Apache 2.0 License.