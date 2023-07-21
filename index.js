const os = require('os');
const path = require('path');
const childProcess = require('child_process');
const http = require('http');
const WebSocket = require('ws');

const DEVTOOLS_PORT = 8315;

const exePath = path.join(os.homedir(), 'AppData\\Local\\Microsoft\\Teams\\current\\Teams.exe');
// Start Teams with DevTools protocol enabled.
const teamsProc = childProcess.execFile(
  exePath,
  [`--remote-debugging-port=${DEVTOOLS_PORT}`],
  (error, stdout, stderr) => {
    console.log(error, stdout, stderr);
  }
);

const spawned = new Promise((resolve, reject) => {
  teamsProc.once('spawn', () => resolve());
  teamsProc.on('error', reject);
});

spawned
  // Wait 5s to give Teams time to start, it's electron garbage after all.
  .then(() => new Promise((resolve) => setTimeout(resolve, 5000)))
  // Make a request to the devtools server requesting a list of debug targets.
  .then(() => new Promise((resolve) => {
    http.get(
      `http://localhost:${DEVTOOLS_PORT}/json/list`,
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

    // Now Teams is running with devtools enabled, and the target frame has been found, a websocket can be
    // created to communicate with the Teams shared worker frame.
    const ws = new WebSocket(target.webSocketDebuggerUrl, {perMessageDeflate: false});
    const connected = new Promise((resolve) => ws.once('open', resolve));
    connected
      .then(() => {
        ws.on('message', (message, isBinary) => console.log(message.toString()));

        ws.send(JSON.stringify({
          id: 1,
          method: 'Runtime.evaluate',
          params: {
            expression: "workerServer._stateAndRequestHandlers.get('graphql').requestHandler.contextValue.discoverService.aad._authTokenCache._cache.get('https://chatsvcagg.teams.microsoft.com').token"
          }
        }));
      })
  })
  .catch(console.error);