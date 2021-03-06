const CONSTANTS = {
  columnStart: 'C',
  emailColumn: 'A',
  nameColumn: 'B',
  rowEnd: 18,
  columnEnd: 19,
  marker: 'X',
  leader: '#548135'
};

const doPanelCheck = () => {
  const reminderDate = findReminderDate();
  if (!reminderDate) return;
  const { column, date } = reminderDate;
  const panelists = getPanelists(column);
  const leaders = findAndRemoveLeaders(panelists, 'isLeading', true);
  console.log(panelists);
  console.log(leaders);
  emailLeaders(leaders, panelists, date);
  emailPanelists(leaders, panelists, date);
}

const findReminderDate = () => {
  const sheet = SpreadsheetApp.getActiveSheet();
  const now = new Date();
  let found;
  let column = CONSTANTS.columnStart;
  for (let i = 0; i < CONSTANTS.rowEnd; i++) {
    const a1Notation = `${column}1`;
    const value = sheet.getRange(a1Notation).getValue();
    console.log(a1Notation, value);
    console.log(now.getDate(), value.getDate());
    const isFourDaysBefore = nowIsFourDaysBefore(now, value);
    console.log(isFourDaysBefore)
    if (isFourDaysBefore) {
      found = { column, a1Notation, date: value };
      break;
    }
    column = nextCharacter(column);
  }
  return found;
}    

const getPanelists = (column) => {
  const sheet = SpreadsheetApp.getActiveSheet();
  const rowOffset = 2;
  const panelists = [];
  for (let i = rowOffset; i < CONSTANTS.columnEnd + rowOffset; i++) {
    const a1Notation = `${column}${i}`;
    const isOnPanel = sheet.getRange(a1Notation).getValue() === CONSTANTS.marker;
    const isLeading = sheet.getRange(a1Notation).getBackground() === CONSTANTS.leader;
    if (isOnPanel) {
      const name = sheet.getRange(`${CONSTANTS.nameColumn}${i}`).getValue();
      const email = sheet.getRange(`${CONSTANTS.emailColumn}${i}`).getValue();
      const panelist = { isOnPanel, isLeading, name, email };
      panelists.push(panelist);
    }
  }
  return panelists;
}
                        
const nextCharacter = (c) => { 
  return String.fromCharCode(c.charCodeAt(0) + 1); 
}

const previousCharacter = (c) => { 
  return String.fromCharCode(c.charCodeAt(0) - 1); 
}

const nowIsFourDaysBefore = (now, then) => {
  if (now > then) {
    return false;
  }
  const oneDay = 24 * 60 * 60 * 1000; // h * m * s * ms
  const diffDays = Math.ceil(Math.abs((now - then) / oneDay));
  return diffDays === 4;
}
                        
const findAndRemoveLeaders = (array, attribute, value) => {
  const removed = [];
  for (let i = 0; i < array.length; i++) {
    if (array[i][attribute] === value) {
      removed.push(array.splice(i, 1)[0]);
    }
  }
  return removed;
}
                        
const formatLeaderEmail = (leader, panelists, date) => {
  const names = panelists.map(panelist => panelist.name);
  const tomorrow = `${date.getMonth()}\/${date.getDate()}`;
  return `
    Hi ${leader.name},<br/>
    <br/>
    This is a friendly reminder that you are leading the bible study panel on Tuesday, ${tomorrow} at 7:30 PM. The ${panelists.length} panelists who
    will be joining you are: ${names.join(', ')}.
    <br/><br/>
    Click <a href='https://www.kitchenergospelhall.com/study-schedule'>here<\/a> to see where we are starting.<br/>
    <br/>
    <br/>
    <p style="font-family:'Courier New'">
      This is an automated email; pretty please don't respond. It is coming from my personal email address but will soon come from a more suitable email address.
      If you are a nerd and would like to see how this code works, you can visit the GitHub repository <a href='https://github.com/GarethSharpe/google-scripts-shedule/blob/master/Code.js'>here<\/a>.
    </p>
  `;
}

const formatPanelistEmail = (panelist, leaders, panelists, date) => {
  const panelistNames = panelists.map(panelist => panelist.name);
  const leaderNames = leaders.map(leader => leader.name);
  const tomorrow = `${date.getMonth()}\/${date.getDate()}`;
  return `
    Hi ${panelist.name},<br/>
    <br/>
    This is a friendly reminder that you are joining the bible study panel on Tuesday, ${tomorrow} at 7:30 PM. The ${panelists.length} panelists who
    will be joining you are: ${panelistNames.join(', ')}. ${leaderNames.join(', ')} will be leading the study.<br/>
    <br/>
    Click <a href='https://www.kitchenergospelhall.com/study-schedule'>here<\/a> to see where we are starting.<br/>
    <br/>
    <br/>
    <p style="font-family:'Courier New'">
      This is an automated email; pretty please don't respond. It is coming from my personal email address but will soon come from a more suitable email address.
      If you are a nerd and would like to see how this code works, you can visit the GitHub repository <a href='https://github.com/GarethSharpe/google-scripts-shedule/blob/master/Code.js'>here<\/a>.
    </p>
  `;
}

const emailLeaders = (leaders, panelists, date) => {
  for (leader of leaders) {
    const leaderEmailTemplate = formatLeaderEmail(leader, panelists, date);
    MailApp.sendEmail({
      to: leader.email,
      subject: `${leader.name}, You're Leading!`,
      htmlBody: leaderEmailTemplate,
    });
  }
}

const emailPanelists = (leaders, panelists, date) => {
  for (panelist of panelists) {
    const panelistEmailTemplate = formatPanelistEmail(panelist, leaders, panelists, date);
    MailApp.sendEmail({
      to: panelist.email,
      subject: `${panelist.name}, You're on the Panel!`,
      htmlBody: panelistEmailTemplate,
    });
  }
}
