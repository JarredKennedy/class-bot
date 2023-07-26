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
    participants: [
      {
        id: string      ID of the user who participated in the meeting
        name: string    Name of the user who participated in the meeting
      }
    ]
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
  _devtoolsWsUrl;
  _socket;

  _skypeAuth;

  _tokens = {
    skype: null,
    skypeRefresh:  null,
    teams: null
  };

  _cache = {
    meetings: {}
  };

  REQ_ID_ENABLE_NET = 1;
  REQ_ID_GET_TEAMS_TOKEN = 2;
  REQ_ID_GET_SKYPE_TOKEN = 3;

  constructor(devtoolsWsUrl) {
    super();
    this._devtoolsWsUrl = devtoolsWsUrl;
    this._connectSockets();
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

    return this.skypeApiCall(`/users/ME/conversations/${channelId}/messages`, 'POST', payload);
  }

  /**
   * Makes a call to the Skype API. This method automatically handles authentication. Returns
   * a promise which will resolve with the request response if successful, otherwise it will
   * reject with some error.
   * 
   * @returns {Promise<object>}
   */
  skypeApiCall(endpoint, method, data) {
    if (this._skypeAuth)
      return this._skypeAuth.then(() => this.skypeApiCall(endpoint, method, data));

    const now = Date.now() / 1000;
    this._skypeAuth = (this._tokens.skype && this._tokens.skype.token && now < (this._tokens.skype.expires - 60))
      ? Promise.resolve()
      : (
          (this._tokens.skypeRefresh && this._tokens.skypeRefresh.token && now < (this._tokens.skypeRefresh.expires - 60))
            ? Promise.resolve()
            : this._fetchSkypeRefreshToken()
        ).then(this._refreshSkypeToken.bind(this));

    this._skypeAuth.finally(() => this._skypeAuth = undefined);

    return this._skypeAuth
      .then(() => {
        const headers = {'Authentication': `skypetoken=${this._tokens.skype.token}`};
    
        if (data)
          headers['Content-Type'] = 'application/json';

        return simpleRequest(`https://apac.ng.msg.teams.microsoft.com/v1${endpoint}`, {method, headers, timeout: 4000}, data);
      });
  }

  /**
   * Makes a call to the Teams API. This method handles authentication automatically. Returns
   * a promise which resolves with the response or rejects with an error message.
   * 
   * @returns {Promise<object>}
   */
  teamsApiCall(endpoint, method, data) {
    // Stub
  }

  /**
   * Handle a WebSocket message from the worker frame. This function examines the structure of the message
   * to see if it is for an event we care about and emits an event if it is.
   * 
   * @param {object} message The message sent via the WebSocket.
   */
  _processMessage(message) {
    if (message.resourceType === 'NewMessage' && message.resource.messagetype === 'Event/Call') {
      if (message.resource.content.indexOf('<ended/>') >= 0) {
        // This message indicates that a meeting has ended.
        // Participants list is in XML. Use minimal regex to get the participant names and IDs from the markup.
        const participants = [...message.resource.content.matchAll(/<part [^>]*>/g)].reduce((ranges, match) => {
          if (ranges.length)
            ranges[ranges.length-1].push(match.index);

          ranges.push([match.index]);
          return ranges;
        }, [])
        .map((range) => {
          const partStr = message.resource.content.substring.apply(message.resource.content, range);
          const idMatch = partStr.match(/identity="([^"]+)/);
          const id = idMatch[1];

          if (id.startsWith('28'))
            return null;

          const nameMatch = partStr.match(/<displayName>([^<]+)/);
          const name = nameMatch[1];
          
          return { id, name };
        })
        .filter((participant) => !!participant);

        const meeting = {
          id: message.resource.skypeguid,
          participants,
          channel: {
            id: message.resource.to
          }
        };

        this.emit(events.MEETING_ENDED, meeting);
      } else {
        // Cache this new call to be processed when the meeting data comes through. The reason for caching it
        // instead of just checking for the update is to ensure that the meeting is new and not just an update
        // to a meeting started some time ago.
        this._cache.meetings[message.resource.id] = (new Date(message.time)).getTime();
      }
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

  // Connect to the target frame devtools server.
  _connectSockets() {
    const socket = new WebSocket(this._devtoolsWsUrl, {perMessageDeflate: false});

    socket.on('open', () => {
      socket.on('message', this._handleSocketData.bind(this));
      this._socket = socket;

      // Tell Teams to send network events to us.
      socket.send(JSON.stringify({ id: this.REQ_ID_ENABLE_NET, method: 'Network.enable' }));
    });
  }

  /**
   * Handle a devtools message sent by a Teams frame.
   * 
   * @param {string} data The message the frame sent.
   */
  _handleSocketData(data) {
    const msg = JSON.parse(data.toString());

    if ('id' in msg) {
      switch (msg.id) {
        case this.REQ_ID_GET_SKYPE_TOKEN:
          return this.emit('tokenupdate.skype', msg.result);
        case this.REQ_ID_GET_TEAMS_TOKEN:
          return this.emit('tokenupdate.teams', msg.result);
      }
    } else if ('method' in msg && msg.method === 'Network.webSocketFrameReceived') {
      try {
        // WebSocket messages from Teams have some weird formatting, this removes that to extract the JSON
        // payload.

        // Messages we care about start with 3.
        if (msg.params.response.payloadData[0] != '3')
          return;

        let colonMatches = 0, i = 0;
        while (colonMatches < 3 && i < msg.params.response.payloadData.length) {
          if (msg.params.response.payloadData[i] === ':')
            colonMatches++;

          i++;
        }

        const rawPayload = msg.params.response.payloadData.substring(i);

        if (rawPayload[0] !== "{")
          return;

        const payload = JSON.parse(rawPayload);

        if (!('body' in payload))
          return;

        // The message will be inspected and an event will potentially be emitted.
        this._processMessage(JSON.parse(payload.body));
      } catch (error) {
        console.error(error);
      }
    }
  }

  /**
   * Fetch the current refresh token from Teams. This executes a line of code in the shared worker which should return
   * the Skype refresh token.
   * 
   * @returns {Promise}
   */
  _fetchSkypeRefreshToken() {
    const fetchToken = new Promise((resolve, reject) => {
      // Execute this JS code in Teams to extract the Skype refresh token.
      this._socket.send(JSON.stringify({
        id: this.REQ_ID_GET_SKYPE_TOKEN,
        method: 'Runtime.evaluate',
        params: {
          expression: "workerServer._stateAndRequestHandlers.get('graphql').requestHandler.contextValue.discoverService.aad.acquireTokenV2('https://api.spaces.skype.com');",
          returnByValue: true,
          awaitPromise: true
        }
      }));

      // There can only be one message-handler function for the WebSocket, so _handleSocketData will receive the result
      // of the above code execution. _handleSocketData will then emit an event with the response of the execution which
      // will provide the refresh token.
      this.once('tokenupdate.skype', (response) => {
        if (
          typeof response === 'object'
          && 'result' in response
          && typeof response.result === 'object'
          && 'value' in response.result
          && typeof response.result.value === 'object'
        ) {
          // Update the refresh token.
          this._tokens.skypeRefresh = {
            token: response.result.value.token,
            expires: response.result.value.expires
          };
          return resolve();
        }

        reject('fetch.response');
      });
    });
    // Allow 5s for the response to be received after sending the execution request.
    const timeout = new Promise((_, reject) => setTimeout(() => reject('fetch.timeout'), 5000));

    return Promise.race([fetchToken, timeout]);
  }

  /**
   * Make a request to the Teams API to generate a new Skype API token. Like all other token refresh/fetch functions, the
   * method returns a Promise which indicates only the success or failure of the operation, the token data is stored on
   * the instance.
   * 
   * @returns {Promise}
   */
  _refreshSkypeToken() {
    return simpleRequest('https://teams.microsoft.com/api/authsvc/v1.0/authz', {
      method: 'POST',
      headers: {
        'Content-Length': 0,
        'Authorization': `Bearer ${this._tokens.skypeRefresh.token}`
      },
      timeout: 4000
    }).then((response) => {
      this._tokens.skype = {
        token: response.tokens.skypeToken,
        expires: parseInt(Date.now() / 1000) + response.tokens.expiresIn
      };
    });
  }

}

/**
 * Make a simple request to a URL.
 */
function simpleRequest(url, options, data) {
  const protocol = url.startsWith('https') ? https : http;

  return new Promise((resolve, reject) => {
    let payload = '';
    if (data) {
      payload = (typeof data === 'string') ? data : JSON.stringify(data);
      options.headers = Object.assign(options.headers ?? {}, {'Content-Length': payload.length});
    }

    const request = protocol.request(url, options, (response) => {
      if (response.statusCode >= 300) {
        if (response.statusCode == 401 || response.statusCode == 403)
          return reject('unauthorized');

        return reject('request_failed');
      }

      response.setEncoding('utf-8');

      let raw = '';
      response.on('data', (chunk) => {
        raw += chunk;
      });

      response.on('end', () => {
        if (response.headers['content-type'].indexOf('application/json') >= 0) {
          return resolve(JSON.parse(raw));
        }

        resolve(raw);
      });
    });

    request.on('error', reject);
    request.on('timeout', () => reject('request_timed_out'));

    if (data) {
      request.write(payload);
    }

    request.end();
  });
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
    .then(() => simpleRequest(`http://localhost:${port}/json/list`, {
      headers: {'Accept': 'application/json'},
      timeout: 2000
    }))
    .then((targets) => {
      // Find the precompiled shared worker frame.
      const target = targets.find((candidate) => candidate.type == 'shared_worker' && candidate.url.indexOf('precompiled') >= 0);

      if (!target)
        throw "Unable to find precompiled-shared-worker";

      return new TeamsClient(target.webSocketDebuggerUrl);
    });
}

module.exports = { events, reactions, connect };