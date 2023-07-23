const childProcess = require('child_process');
const http = require('http');
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
      title: string   Title of the channel the meeting is in
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
      title: string   Title of the channel the meeting is in
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

  constructor(targetWsUrl) {
    this.targetWsUrl = targetWsUrl;
    this.connectSocket();
  }

  // Connect to the target frame devtools server.
  connectSocket() {
    this.socket = new WebSocket(targetWsUrl, {perMessageDeflate: false});

    this.socket.on('message', (data) => {
      const msg = JSON.parse(data.toString());

      if ('method' in msg && msg.method === 'Network.webSocketFrameReceived') {
        
      }
    });

    // Tell Teams to send network events to us.
    this.socket.send(JSON.stringify({
      id: 1,
      method: 'Network.enable'
    }));
  }

}

/**
 * Launch teams with devtools enabled. Use devtools protocol to extract auth token and receive
 * events.
 * 
 * @param {string} exePath Path to the Teams executable.
 * @param {number} port    Port where devtools is served.
 * 
 * @return {Promise<TeamsClient>}
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