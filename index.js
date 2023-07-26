const os = require('os');
const path = require('path');
const config = require('./config.json');
const teams = require("./teams");

const exePath = path.isAbsolute(config.teams.app) ? config.teams.app : path.join(os.homedir(), config.teams.app);

(async () => {
  teamsClient = await teams.connect(exePath, config.teams.debugPort);

  const meetingMessages = {};

  // When a teacher starts a class, post a message in chat with the meeting link.
  teamsClient.on(teams.events.NEW_MEETING, async (meeting) => {
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
    message += getBotByline(config.repoUrl);

    try {
      const messageData = await teamsClient.sendMessage(config.groupChatChannel, message);

      // Record the meeting details, the class details, and the clientMessageId.
      meetingMessages[messageData.clientMessageId] = {
        class: channelClass,
        meeting
      };
    } catch (error) {
      console.error("Sending meeting message failed", error);
    }
  });

  // Record messages sent by the bot.
  teamsClient.on(teams.events.NEW_MESSAGE, (message) => {
    if (message.clientMessageId in meetingMessages) {
      const data = meetingMessages[message.clientMessageId];
      delete meetingMessages[message.clientMessageId];
      data.message = message;
      meetingMessages[data.meeting.id] = data;
    }
  });

  // When a class ends, update the meeting message with the details of the meeting.
  teamsClient.on(teams.events.MEETING_ENDED, async (meeting) => {
    if (meeting.id in meetingMessages) {
      // Retrieve data stored previously.
      const data = meetingMessages[meeting.id], now = new Date();
      const day = now.toLocaleDateString('en-AU', {day: '2-digit'});
      const month = now.toLocaleDateString('en-AU', {month: '2-digit'});
      const chatUrl = `https://teams.microsoft.com/l/message/${meeting.channel.id}/${data.meeting.messageId}`;
      const maxDuration = meeting.participants.reduce((max, participant) => Math.max(max, participant.duration), 0);
      const meetingHrs = Math.floor(maxDuration / 3600);
      const meetingMins = Math.floor((maxDuration % 3600) / 60);

      const divider = '<span style="font-size:inherit;"> · </span>';
      let message = `<h1><code class="skipProofing">WEEK${getCurrentCourseWeek(config.courseStart)}</code> ${data.class.title} Meeting</h1>`;
      // Meeting date.
      message += '<p><span contenteditable="false" title="Calendar" type="(1f4c5_calendar)" class="animated-emoticon-20-1f4c5_calendar">';
      message += '<img itemscope itemtype="http://schema.skype.com/Emoji" itemid="1f4c5_calendar" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/1f4c5_calendar/default/20_f.png" title="Calendar" style="width:20px;height:20px;"></span>';
      message += `<span style="font-size:inherit;"><strong>${day}</strong></span><span style="font-size:xx-small;"><strong>/${month}</strong></span>`;
      message += divider;
      // Meeting chat thread link.
      message += `<a href="${chatUrl}"><span style="font-size:inherit;"><strong>Chat Thread</strong></span></a>`;
      message += divider;
      // Meeting number of participants.
      message += '<span contenteditable="false" title="Student" type="(student)" class="animated-emoticon-20-student">';
      message += '<img itemscope itemtype="http://schema.skype.com/Emoji" itemid="student" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/student/default/20_f.png" title="Student" style="width:20px;height:20px;"></span>';
      message += `<span style="font-size:inherit;"><strong>${meeting.participants.length}</strong></span>`;

      // Meeting duration.
      if (meetingHrs || meetingMins) {
        message += divider;
        message += '<span contenteditable="false" title="Twelve oclock" type="(1f55b_twelveoclock)" class="animated-emoticon-20-1f55b_twelveoclock">';
        message += '<img itemscope itemtype="http://schema.skype.com/Emoji" itemid="1f55b_twelveoclock" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/1f55b_twelveoclock/default/20_f.png" title="Twelve oclock" style="width:20px;height:20px;"></span>';

        if (meetingHrs)
          message += `<strong>${meetingHrs}</strong><span style="font-size:xx-small;"><strong>HR</strong></span>`;
        if (meetingMins)
          message += `<strong>${meetingMins}</strong><span style="font-size:xx-small;"><strong>MIN</strong></span>`;
      }

      message += '</p>';
      message += getBotByline(config.repoUrl);

      try {
        await teamsClient.editMessage(data.message.channel.id, data.message.id, data.message.clientMessageId, message);
      } catch (error) {
        console.error("Updating meeting message failed", error);
      }

      delete meetingMessages[meeting.id];
    }
  });

})();

function getCurrentCourseWeek(start) {
  const now = parseInt(Date.now() / 1000);
  return Math.ceil((now - start) / (7 * 24 * 60 * 60));
}

function getBotByline(botUrl) {
  let byline = '<hr><p><span contenteditable="false" title="Cool robot" type="(coolrobot)" class="animated-emoticon-20-coolrobot">';
  byline += '<img itemscope itemtype="http://schema.skype.com/Emoji" itemid="coolrobot" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/coolrobot/default/20_f.png" title="Cool robot" style="width:20px;height:20px;"></span>';
  byline += `<a href="${botUrl}" rel="noreferrer noopener" title="${botUrl}" target="_blank"><i>CLASS BOT</i></a></p>`;
  return byline;
}