const childProcess = require('child_process');
const http = require('http');
const https = require('https');
const WebSocket = require('ws');
const { EventEmitter } = require('events');

const events = {
  /*
  Triggered when a meeting is started. Technically, events for a meeting being started are not sent
  via the websocket. This event is really triggered by updates made to the channel (thread) that the
  meeting takes place in.
  {
    id: string          ID of the meeting
    title: string       Title of the meeting
    joinUrl: string     URL to join the meeting
    startedBy: string   ID of the user who started the meeting
    channel: {
      id: string      ID of the channel the meeting is in
    }
  }
  */
  NEW_MEETING: 10,
  /*
  Triggered when a meeting ends, which is when there is no one in the meeting anymore. Like the NEW_MEETING
  event, this is based on events that occur in the channel, not a specific call-ending event.
  {
    id: string          ID of the meeting
    title: string       Title of the meeting
    channel: {
      id: string      ID of the channel the meeting is in
    }
  }
  */
  MEETING_ENDED: 11,
  /*
  Triggered when a new message is received in a channel. Chats are also considered channels.
  {
    id: string          ID of the message
    content: string     Conent of the message. This is typically in HTML
    user: {
      name: string    Name of the user who sent the message
      id: string      ID of the user who sent the message
    }
    channel: {
      id: string      ID of the channel the message was sent in
      type: string    Either 'chat' or 'topic'. A topic is a conversation in a team channel
    }
  }
  */
  NEW_MESSAGE: 20,
  /*
  Triggered when a message is edited. Use this event to detect reactions, you would need to record the
  message by capturing the NEW_MESSAGE event, then detecting changes to the reactions on the message.
  {
    id: string          ID of the message
    content: string     Conent of the message. This is typically in HTML
    reactions: {
      [reaction: string]: users: string[]     An array of the users who have this reaction to the message.
    }
    user: {
      name: string    Name of the user who sent the message
      id: string      ID of the user who sent the message
    }
    channel: {
      id: string      ID of the channel the message was sent in
      type: string    Either 'chat' or 'topic'. A topic is a conversation in a team channel
    }
  }
  */
  MESSAGE_EDITED: 21,
  /*
  Triggered when a message is deleted.
  {
    id: string          ID of the message
    user: {
      name: string    Name of the user who sent the message
      id: string      ID of the user who sent the message
    }
    channel: {
      id: string      ID of the channel the message was sent in
      type: string    Either 'chat' or 'topic'. A topic is a conversation in a team channel
    }
  }
  */
  MESSAGE_DELETED: 22,
  /*
  Triggered when someone is typing in chat. There is no 'done typing' event, the client seems to just show
  the user is typing for 20s unless a message from them is received in the chat before that.
  {
    userId: string      ID of the user who is typing
    channel: {
      title: string   Title of the chat the user is typing in
      id: string      ID of the chat the user is typign in
    }
  }
  */
  CHAT_USER_TYPING: 23
};

const reactions = {
  YES: 'yes', LIKE: 'yes',                                // 👍
  HEART: 'heart',                                         // ❤️
  LAUGH: 'laugh',                                         // 😆
  SURPRISED: 'surprised',                                 // 😮
  GRINNING_FACE_BIG_EYES: 'grinningfacewithbigeyes',      // 😃
};

class TeamsClient extends EventEmitter {
  targetWsUrl;
  socket;

  _tokens = {
    skype: null,
    teams: null
  };

  _cache = {
    meetings: {}
  };

  REQ_ID_ENABLE_NET = 1;
  REQ_ID_GET_TEAMS_TOKEN = 2;
  REQ_ID_GET_SKYPE_TOKEN = 3;

  constructor(targetWsUrl) {
    super();
    this.targetWsUrl = targetWsUrl;
    this._connectSocket();
  }

  processMessage(message) {
    if (message.resourceType === 'NewMessage' && message.resource.messagetype === 'Event/Call') {
      // Cache this new call to be processed when the meeting data comes through. The reason for caching it
      // instead of just checking for the update is to ensure that the meeting is new and not just an update
      // to a meeting started some time ago.
      this._cache.meetings[message.resource.id] = (new Date(message.time)).getTime();
    } else if (
      message.resourceType === 'MessageUpdate'
      && message.resource.messagetype === 'Event/Call'
      && this._cache.meetings.hasOwnProperty(message.resource.id)
    ) {
      const timeAgo = Date.now() - this._cache.meetings[message.resource.id];
      // If the meeting message was more than 1min ago, delete from the cache and stop processing.
      if (timeAgo > 60 * 1000) {
        delete this._cache.meetings[message.resource.id];
        return;
      }

      // Check for the update that has the rest of the meeting data then emit the new meeting event.

      if (!('meeting' in message.resource.properties))
        return;

      const meetingData = JSON.parse(message.resource.properties.meeting);

      if (!('meetingJoinUrl' in meetingData))
        return;

      delete this._cache.meetings[message.resource.id];

      const meeting = {
        id: message.resource.skypeguid,
        title: meetingData.meetingtitle,
        joinUrl: meetingData.meetingJoinUrl,
        startedBy: meetingData.organizerId,
        channel: {
          id: message.resource.to
        }
      };

      // Cache this meeting so a MEETING_ENDED event can be generated later.
      this._cache.meetings[meeting.id] = meeting;

      this.emit(events.NEW_MEETING, meeting);
    }
  }

  /**
   * Send a message to a Teams channel (chat or team channel).
   * 
   * @param {string} channelId The ID of the channel to send the message to.
   * @param {string} message The content of the message to send. Should be in HTML.
   * 
   * @returns {Promise}
   */
  sendMessage(channelId, message) {
    const payload = {
      content: message,
      messagetype: "RichText/Html",
      contenttype: "text",
      amsreferences: [], // don't know what this is
      clientmessageid: `1337${Date.now()}`, // this is an ID for the client, the server has its own IDs
      imdisplayname: "Class Bot", // this seems to be ignored.
      properties: {
        importance: "", // you can specify 'high', and 'urgent'
        subject: "" // I think this has something to do with quoting previous messages.
      }
    };

    const url = `https://apac.ng.msg.teams.microsoft.com/v1/users/ME/conversations/${channelId}/messages`;
    const data = JSON.stringify(payload);

    return new Promise((resolve, reject) => {
      const req = https.request(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Content-Length': data.length,
          'Authentication': `skypetoken=${this._tokens.skype}`
        },
        timeout: 4000
      }, (res) => {
        if (res.statusCode >= 300) {
          return reject('sendMessage request failed.');
        }

        resolve();
      });

      req.on('error', reject);
      req.on('timeout', () => reject('timed out'));
      req.write(data);
      req.end();
    });
  }

  // Connect to the target frame devtools server.
  _connectSocket() {
    this.socket = new WebSocket(this.targetWsUrl, {perMessageDeflate: false});

    this.socket.on('open', () => {
      this.socket.on('message', (data) => {
        const msg = JSON.parse(data.toString());
  
        if ('id' in msg) {
          if (msg.id === this.REQ_ID_GET_TEAMS_TOKEN || msg.id === this.REQ_ID_GET_SKYPE_TOKEN) {
            this._processTokenResponse(msg.id, msg.result);
          }
        } else if ('method' in msg && msg.method === 'Network.webSocketFrameReceived') {
          try {
            let colonMatches = 0, i = 0;

            while (colonMatches < 3 && i < msg.params.response.payloadData.length) {
              if (msg.params.response.payloadData[i] === ':') {
                colonMatches++;
              }

              i++;
            }

            const rawPayload = msg.params.response.payloadData.substring(i);

            if (rawPayload[0] === "{") {
              const payload = JSON.parse(rawPayload);

              if (!('body' in payload))
                return;

              const body = JSON.parse(payload.body);
              this.processMessage(body);
            }
          } catch (error) {
            console.error(error);
          }
        }
      });
  
      // Tell Teams to send network events to us.
      this.socket.send(JSON.stringify({
        id: this.REQ_ID_ENABLE_NET,
        method: 'Network.enable'
      }));

      // Give time for the frame to load before requesting tokens.
      setTimeout(this._fetchTokens.bind(this), 10000);
    });

    // Reconnect the socket if it gets disconnected.
    this.socket.on('close', this._connectSocket);
  }

  // Extract tokens from Teams. These can be used to send messages among other uses.
  // This executes the code below in the Teams frame and will return the cached tokens from Teams.
  _fetchTokens() {
    const send = (message) => this.socket.send(JSON.stringify(message));

    send({
      id: this.REQ_ID_GET_TEAMS_TOKEN,
      method: 'Runtime.evaluate',
      params: {
        expression: "workerServer._stateAndRequestHandlers.get('graphql').requestHandler.contextValue.discoverService.aad._authTokenCache._cache.get('https://chatsvcagg.teams.microsoft.com');",
        returnByValue: true
      }
    });

    send({
      id: this.REQ_ID_GET_SKYPE_TOKEN,
      method: 'Runtime.evaluate',
      params: {
        expression: "workerServer._stateAndRequestHandlers.get('graphql').requestHandler.contextValue.discoverService._skypeToken.value;",
        returnByValue: true
      }
    });
  }

  _processTokenResponse(messageId, tokenResponse) {
    if (messageId === this.REQ_ID_GET_TEAMS_TOKEN) {
      this._tokens.teams = tokenResponse.result.value.token;
    } else {
      this._tokens.skype = tokenResponse.result.value;
    }
  }

}

/**
 * Launch teams with devtools enabled. Use devtools protocol to extract auth token and receive
 * events.
 * 
 * @param {string} exePath Path to the Teams executable.
 * @param {number} port    Port where devtools is served.
 * 
 * @returns {Promise<TeamsClient>}
 */
function connect(exePath, port) {
  // Start Teams with DevTools protocol enabled.
  const teamsProc = childProcess.execFile(
    exePath,
    [`--remote-debugging-port=${port}`],
    (error, stdout, stderr) => {
      console.log(error, stdout, stderr);
    }
  );

  const spawned = new Promise((resolve, reject) => {
    teamsProc.once('spawn', () => resolve());
    teamsProc.on('error', reject);
  });

  return spawned
    // Wait 5s to give Teams time to start, it's electron garbage after all.
    .then(() => new Promise((resolve) => setTimeout(resolve, 5000)))
    // Make a request to the devtools server requesting a list of debug targets.
    .then(() => new Promise((resolve) => {
      http.get(
        `http://localhost:${port}/json/list`,
        {
          headers: {
            'Accept': 'application/json'
          },
          timeout: 2000
        },
        (response) => {
          response.setEncoding('utf-8');

          let raw = '';
          response.on('data', (chunk) => {
            raw += chunk;
          });

          response.on('end', () => {
            resolve(JSON.parse(raw));
          });
        }
      );
    }))
    .then((targets) => {
      // Find the precompiled-shared-worker frame.
      // This frame is useful because it creates the WebSocket connection which receives Teams events like
      // new messages and new meetings. It is also possible to extract a useful token which can be used to
      // get data from the Teams API (teams, channels, chats, messages).
      const target = targets.find((candidate) => candidate.type == 'shared_worker' && candidate.url.indexOf('precompiled') >= 0);

      if (!target)
        throw "Unable to find precompiled-shared-worker";

      return new TeamsClient(target.webSocketDebuggerUrl);
    });
}

module.exports = { events, reactions, connect };