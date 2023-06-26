import {BingAIClient} from '@waylaidwanderer/chatgpt-api';
import MarkdownIt from 'markdown-it';
import MarkdownItFootnote from 'markdown-it-footnote';
import MarkdownItSup from 'markdown-it-sup';
import fetch from 'node-fetch'
import {randomUUID} from 'crypto';

if (!process.env.ACCESS_TOKEN) {
    console.error('ACCESS_TOKEN environment variable not set');
    process.exit(1);
}

const authorizationHeader = "Bearer " + process.env.ACCESS_TOKEN;
const bingOptions = {
    userToken: '',
    cookies: '',
    proxy: '',
    debug: false,
};
const defaultHeaders = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "pl,en;q=0.9,en-GB;q=0.8,en-US;q=0.7",
    "clientinfo": "os=windows; osVer=10; proc=x86; lcid=pl-pl; deviceType=1; country=pl; clientName=skypeteams; clientVer=1415/1.0.0.2023050101; utcOffset=+02:00; timezone=Europe/Warsaw",
    "sec-ch-ua": "\"Not.A/Brand\";v=\"8\", \"Chromium\";v=\"114\", \"Microsoft Edge\";v=\"114\"",
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": "\"Windows\"",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "x-client-ui-language": "pl-pl",
    "x-ms-client-cpm": "ApplicationLaunch",
    "x-ms-client-env": "pds-prodl-c1-euwe-02",
    "x-ms-client-type": "web",
    "x-ms-client-version": "1415/1.0.0.2023050101",
    "x-ms-user-type": "null",
    "Referer": "https://teams.live.com/_",
    "Referrer-Policy": "strict-origin-when-cross-origin"
};
const md = new MarkdownIt({html: true}).use(MarkdownItFootnote).use(MarkdownItSup);
const uuid = randomUUID();
const strings = {
    botName: 'Krak Bot',
    newThread: 'Zaczynamy od nowa. W czym mogę Ci teraz pomóc?',
    learnMore: 'Dowiedz się więcej:'
};
let res, json, bingAIClient = new BingAIClient(bingOptions), bingResponse = {}, msgIds = {};

res = await fetch("https://teams.live.com/api/auth/v1.0/authz/consumer", {
    headers: {
        ...defaultHeaders,
        authorizationHeader,
        "claimschallengecapable": "true",
        "ms-teams-authz-type": "ExplicitLogin"
    },
    body: null,
    method: "POST"
});
json = await res.json();
const skypetoken = json.skypeToken?.skypetoken;
const skypeid = json.skypeToken?.skypeid;
if (!skypetoken) {
    console.error('Cannot authenticate:', json);
    process.exit(2);
}

res = await fetch("https://msgapi.teams.live.com/v2/users/ME/endpoints/" + uuid, {
    headers: {
        ...defaultHeaders,
        "authentication": "skypetoken=" + skypetoken,
        "content-type": "application/json"
    },
    body: "{\"startingTimeSpan\":0,\"endpointFeatures\":\"Agent,Presence2015,MessageProperties,CustomUserProperties,NotificationStream,SupportsSkipRosterFromThreads\",\"subscriptions\":[{\"channelType\":\"HttpLongPoll\",\"interestedResources\":[\"/v1/users/ME/conversations/ALL/properties\",\"/v1/users/ME/conversations/ALL/messages\",\"/v1/threads/ALL\"]}]}",
    method: "PUT"
});
json = await res.json();
let longPollUrl = json.subscriptions[0].longPollUrl;
console.log('Waiting for messages...');

while (res.ok) {
    res = await fetch(longPollUrl, {
        headers: {
            ...defaultHeaders,
            "authentication": "skypetoken=" + skypetoken
        },
        body: null,
        method: "GET"
    });
    json = await res.json();
    longPollUrl = json.next;
    json.eventMessages?.forEach(async m => {
        if (m.resourceType === 'NewMessage' && m.resource?.messagetype === 'RichText/Html') {
            if (m.resource.content && m.resource.imdisplayname !== strings.botName) {
                console.log('New message from:', m.resource.imdisplayname);
                const conversationId = m.resource.to;
                const clientMsgId = Math.floor(Math.random() * 9e18) + 1e18;
                const query = m.resource.content.replace(/(<[^>]+>)/g, '');
                if (query === '--') {
                    newThread(clientMsgId, conversationId);
                } else {
                    getReply(query, clientMsgId, conversationId);
                }
            } else if (m.resource.imdisplayname === strings.botName) {
                msgIds[m.resource.clientmessageid] = m.resource.id;
            }
        }
    });
}

function newThread(clientMsgId, conversationId) {
    clearContext();
    sendReply(strings.newThread, clientMsgId, conversationId);
}

async function getReply(query, clientMsgId, conversationId) {
    console.log('Querying Bing AI for:', query);
    sendTyping(conversationId);
    try {
        bingResponse = await getBingResponse(query, clientMsgId, conversationId);
    } catch (e) {
        console.error(e);
        clearContext();
        if (e.message === 'No message was generated.')
            return;
        console.log('Retrying');
        bingAIClient = new BingAIClient(bingOptions);
        bingResponse = await getBingResponse(query, clientMsgId, conversationId);
    }
    if (bingResponse.details?.adaptiveCards) {
        const reply = bingResponse.details.adaptiveCards[0].body.filter(b => b.type === 'TextBlock').map(b => b.size === 'small' ? '<span style="font-size:small;">​' + b.text + '</span>' : b.text).join('\n\n---\n\n');
        console.log('Reply:', reply);
        sendReply(md.render(reply), clientMsgId, conversationId);
    }
    setTimeout(clearContext, 15 * 60000);
}

function clearContext() {
    console.log('Clearing context');
    bingResponse = {};
}

function getBingResponse(query, clientMsgId, conversationId) {
    let draft = '', lastDraftLength = 0;
    return bingAIClient.sendMessage(query, {
        //conversationSignature: bingResponse.conversationSignature,
        //conversationId: bingResponse.conversationId,
        //clientId: bingResponse.clientId,
        //invocationId: bingResponse.invocationId,
        jailbreakConversationId: bingResponse.jailbreakConversationId || true,
        parentMessageId: bingResponse.messageId,
        onProgress: (token) => {
            process.stdout.write(token);
            draft += token;
            if (draft.length - lastDraftLength >= 25 && (token.includes('.') || token.includes('\n'))) {
                sendReply(md.render(draft.replace(/\[\^[0-9]+\^\]/g, '') + '…'), clientMsgId, conversationId);
                lastDraftLength = draft.length;
            }
        }
    });
}

function filterReply(reply) {
    return reply.replace(/\b(Bing|Sydney)\b/, 'Krak').replace('Learn more:', strings.learnMore).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

function sendReply(reply, clientMsgId, conversationId) {
    const body = '{"content":"' + filterReply(reply) + '","messagetype":"RichText/Html","contenttype":"text","amsreferences":[],"clientmessageid":"' + clientMsgId + '","imdisplayname":"' + strings.botName + '","properties":{"importance":"","subject":""}}';
    let url = 'https://msgapi.teams.live.com/v1/users/ME/conversations/' + conversationId + '/messages';
    let method = 'POST';
    if (msgIds[clientMsgId]) {
        url += '/' + msgIds[clientMsgId];
        method = 'PUT';
    }
    fetch(url, {
        headers: {
            ...defaultHeaders,
            "authentication": "skypetoken=" + skypetoken,
            "content-type": "application/json"
        },
        body,
        method
    });
}

function sendTyping(conversationId) {
    fetch("https://msgapi.teams.live.com/v1/users/ME/conversations/" + conversationId + "/messages", {
        headers: {
            ...defaultHeaders,
            "authentication": "skypetoken=" + skypetoken,
            "content-type": "application/json"
        },
        body: "{\"messagetype\":\"Control/Typing\",\"contenttype\":\"Application/Message\",\"content\":\"\"}",
        method: "POST"
    });
}
