const os = require('os');
const path = require('path');
const config = require('./config.json');

const exePath = path.isAbsolute(config.teams.app) ? config.teams.app : path.join(os.homedir(), config.teams.app);

teamsClient = await teams.connect(exePath, config.teams.debugPort);

teamsClient.on(teams.events.NEW_MEETING, (meeting) => {
  // Ignore the meeting if it wasn't created by a teacher.
  if (!config.teachers.some(teacher => meeting.startedBy == teacher))
    return;

  // Find the class which runs in the channel the meeting was started in.
  channelClass = config.classes.find(tafeClass => tafeClass.channels.indexOf(meeting.channel.id) >= 0);

  // Ignore the meeting if there is no class in this channel.
  if (!channelClass)
    return;

  const now = new Date();
  // Ignore the meeting if the class isn't scheduled to run today.
  if (channelClass.start.day !== now.getDay())
    return;


  const classStarts = new Date(now);
  // No classes near midnight so idc about day.
  classStarts.setHours(channelClass.start.hour)
  classStarts.setMinutes(channelClass.start.mins);
  classStarts.setSeconds(0);
  classStarts.setMilliseconds(0);

  // If the class doesn't start 30mins before/after now, ignore meeting.
  if (Math.abs(now.getTime() - classStarts.getTime()) > 30 * 60 * 1000)
    return;

  // If all checks passed, tell group chat that the class has started.
  let message = `<h1>${channelClass.title} class is now meeting</h1>`;
  message += `<p style="font-size:x-large;"><a href="${meeting.joinUrl}">Join Here</a></p>`;
  message += '<hr>';
  message += `<em>This message was written by <a href="${config.repoUrl}">a bot</a></em>`;
  teamsClient.sendMessage(config.groupChatChannel, message);
});
