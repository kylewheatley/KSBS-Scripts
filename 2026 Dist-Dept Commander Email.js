//This program is for sending automated emails to the distict and department commanders
//The function sendEmails() sends all emails to district commanders, SAL commanders, and department/detachment commanders

//First Time Usage, Make sure OVERRIDE_EMAIL_TOGGLE is set to false if you are NOT testing and want to send emails to the commanders

//here is the link to the template document for the two legion automated emails:
// https://docs.google.com/document/d/1sBTx2dJaiO4lj7cr1-WarVT6LWXpYQZ045cEh0mvHsQ/edit?usp=sharing

//2026 Master Data Spreadsheet
const spreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1YVSvIIkK1J_-i9oNEVqWIqZYD9zNsezlWZ-FS68tQ8Y/edit?gid=658976025#gid=658976025';
const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);

//Setting to switch whether the emails are sent to the district commander email addresses. If set to true, emails will NOT be sent to the commanders, and will instead be sent to the override email
const OVERRIDE_EMAIL_TOGGLE = false;
const overrideEmail = "jackson.keeler@ksbstate.org";

const districtCommanderInfo = [
  {commanderName:"Ed Gunnels", districtEmail:"districtcommander1@kansaslegion.org", districtNumber:1, salCommanderName:"Kevin Funk", salCommanderEmail:"kfunk2407@gmail.com"},
  {commanderName:"Myra Jowers", districtEmail:"districtcommander2@kansaslegion.org", districtNumber:2, salCommanderName:"Rick Miller", salCommanderEmail:"salsquadron153olathe@outlook.com"},
  {commanderName:"Mark Stwalley", districtEmail:"districtcommander3@kansaslegion.org", districtNumber:3, salCommanderName:"Ronnie Rommel", salCommanderEmail:"ronnie.rommel72@gmail.com"},
  {commanderName:"Danny Roush", districtEmail:"districtcommander4@kansaslegion.org", districtNumber:4, salCommanderName:"", salCommanderEmail:""},
  {commanderName:"Joe Hulse", districtEmail:"districtcommander5@kansaslegion.org", districtNumber:5, salCommanderName:"Dan McDowell", salCommanderEmail:"djmcdowell84@gmail.com"},
  {commanderName:"Jimmy Strachan", districtEmail:"districtcommander6@kansaslegion.org", districtNumber:6, salCommanderName:"Kelly Pflughoeft", salCommanderEmail:"kamep5@yahoo.com"},
  {commanderName:"Bill Sykes", districtEmail:"districtcommander7@kansaslegion.org", districtNumber:7, salCommanderName:"Rick Munsch", salCommanderEmail:"rdmunsch@gmail.com"},
  {commanderName:"Henk Rijfkogel", districtEmail:"districtcommander8@kansaslegion.org", districtNumber:8, salCommanderName:"Michael Weber", salCommanderEmail:"spidercreek220@gmail.com"},
  {commanderName:"Sean Hankin", districtEmail:"districtcommander9@kansaslegion.org", districtNumber:9, salCommanderName:"Monte Lewis", salCommanderEmail:"montemegg57@gmail.com"},
  {commanderName:"Larry Rokey", districtEmail:"districtcommander10@kansaslegion.org", districtNumber:10, salCommanderName:"Mark Wessel", salCommanderEmail:"mwessel.alr21@gmail.com"},
  {commanderName:"Alan Zeitvogel", districtEmail:"districtcommander11@kansaslegion.org", districtNumber:11, salCommanderName:"Randy Porter", salCommanderEmail:"bightmeagain2@yahoo.com"},
];

const departmentAndDetachmentInfo ={
  departmentCommanderName:"Department Commander Evans", departmentCommanderEmail:"departmentcommander@kansaslegion.org", detachmentCommanderName:"Detachment Commander Gallentine", detachmentCommanderEmail:"jim.gallentine@gmail.com"
  };

function sendEmails(){
  //send an email to each district commander
  // districtCommanderInfo.forEach(district => {
  //   emailDistrictCommander(district);
  // })

  for(var x = 5; x <= 11; x++){
    sendSingleEmail(x);
  }

  //Send an email to the department commander and detachment commander
  emailDepartmentCommander();
}

//Manually select and send a single email. For testing purposes
function sendSingleEmail(district){
  // var district = 4;

  emailDistrictCommander(districtCommanderInfo[district-1]);
}

function readMasterSheetData(targetDistrict){
  //Grabbing the 2026 Master Data delegate lists
  const sheet = spreadsheet.getSheetByName('2026 Delegate List');

  //List of all columns used for the email tables
  const firstNameColumn = 'I';
  const lastNameColumn = 'K';
  const delegateSchoolColumn = 'T';
  const delegateParentColumn = 'AP';
  const parentPhoneColumn = 'AR';
  const totalPaidColumn = 'BB'; //Make sure to account for null values

  const legionDistrictColumn = 'V'; //Not put into table but used to filter by legion district

  //All columns to grab from the spreadseet put into the table
  const columns = [firstNameColumn, lastNameColumn, delegateSchoolColumn, delegateParentColumn, parentPhoneColumn, totalPaidColumn];

  const lastRow = sheet.getLastRow();

  //Creating an array containing all table data
  var valueData = []
  for(var i = 0; i < columns.length; i++){
    let range = sheet.getRange(columns[i] + '2:' + columns[i] + lastRow);
    valueData.push(range.getValues());
  }

  var districtData = sheet.getRange(legionDistrictColumn + '2:' + legionDistrictColumn + lastRow).getValues();
  
  for(var i = 0; i < valueData.length; i++){
    valueData[i] = valueData[i].filter((element, index) => districtData[index][0] == targetDistrict);
  }

  return valueData;
}

function getNumDelegates(district = "Grand Total"){
  //Grab the sheet containing district delegate totals
  const spreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1YVSvIIkK1J_-i9oNEVqWIqZYD9zNsezlWZ-FS68tQ8Y/edit?gid=658976025#gid=658976025';
  const spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  const sheet = spreadsheet.getSheetByName('2026 Delegate Count By Legion Dist');


  const lastRow = sheet.getLastRow();
  var targetValue = 0;
  try{
    const targetRow = sheet.getRange('A1:A' + lastRow).createTextFinder(district).findNext().getRow();
    targetValue = sheet.getRange(targetRow, 2).getValue();
  }
  catch{
    console.log("No data located for district: " + district + " for sheet '2026 Delegate Count By Legion Dist'");
  }
  
  return targetValue;
}

function districtCommanderTemplate(tableData, districtCount, districtInfo){
  //Get current date
  const currentDate = new Date();
  const timeZone = Session.getScriptTimeZone(); 
  const formattedDate = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy");

  

  //Change the greeting if there is no SAL district commander
  let greeting = "District Commander " + districtInfo.commanderName + " and SAL District Commander " + districtInfo.salCommanderName;
  if(districtInfo.salCommanderName == ''){
    greeting = districtInfo.commanderName;
  }

  return`
    <html>
      <head>
        <style>

          #delegates {
            border-collapse: collapse;
            width: 100%;
          }

          #delegates td, #delegates th {
            border: 1px solid #ddd;
            padding: 8px;
          }

          #delegates tr:nth-child(even){background-color: #f2f2f2;}

          #delegates tr:hover {background-color: #ddd;}

          #delegates th {
            padding-top: 12px;
            padding-bottom: 12px;
            text-align: left;
            background-color: black;
            color: white;
          }
        </style>

      </head>
      <body>
        <p>${greeting},<p>

        <p>Greetings from <i>your</i> Kansas Boys State! As of <strong>${formattedDate.toString()}</strong>, we currently have ${districtCount} signed up to attend the 2026 Session of Boys State from your district. </p>

        <p>As a reminder:<p>
        <ul>
          <li>There is no maximum number of of delegates a post or school can send</li>
          <li>The cost to attend is $375, with $325 coming from a sponsor or Post</li>
          <li><strong>Who</strong> – Young men who are Juniors (preferred), or Sophomores in the 25-26 school year</li>
          <li><strong>When</strong> – Sunday, May 31st through Saturday, June 6th, 2026</li>
          <li><strong>Where</strong> – Kansas State University in Manhattan</li>
          <li><strong>Why</strong> – To inculcate a sense of individual obligation to community, state and nation</li>
          <li><strong>How</strong> – Apply at www.ksbstate.org</li>
          <li><strong><i><u>Bonus</u></i></strong> – College Credit from K-State and Scouting Merit Badges can now be earned for Boys Staters!</li>
        </ul>

        <p>Below is a table of the current delegates from your district:</p>

        <table id="delegates">
          <tr>
            <th>Delegate First Name</th>
            <th>Delegate Last Name</th>
            <th>Delegate School</th>
            <th>Delegate Parent</th>
            <th>Delegate Parent Phone</th>
            <th>Total Paid</th>
          </tr>
          ${tableData}
        </table>

        <p>Our goal for this year is 350 delegates, that is 32 delegates per district! If your district has the most delegates, you will get our plaque with your district’s number on it for the next year!</p>

        <p>If you have any questions, please contact <a href="mailto:info@ksbstate.org">info@ksbstate.org</a>.</p>

        <p>Thank you for the support of your Kansas Boys State program!<p>

        <p>Kyle Wheatley<br>
        Executive Director<br>
        The American Legion Boys State of Kansas</p>
      </body>
    </html>
  `;
}

function emailDistrictCommander(district) {


  var tableData = readMasterSheetData(district.districtNumber);

  var rowLength = tableData.length;
  var colLength = tableData[0].length;

  var htmlTableData = "";
  for(var i = 0; i < colLength; i++){
    htmlTableData += "<tr>";
    for(var j = 0; j < rowLength; j++){
      htmlTableData += "<td>" + tableData[j][i] + "</td>"
    }
    htmlTableData += "</tr>";
  }

  const delegateCount = getNumDelegates(district.districtNumber);

  emailTemplate = districtCommanderTemplate(htmlTableData, delegateCount, district);

  if(OVERRIDE_EMAIL_TOGGLE){
    GmailApp.sendEmail(overrideEmail, "Delegate Count for District " + district.districtNumber, "", {  
      htmlBody: emailTemplate,  
      from: "info@ksbstate.org",  
      name: "Kansas Boys State"  
    });  
  }
  else{
    console.log("Sending email to district: " + district.districtNumber);
    GmailApp.sendEmail(district.districtEmail, "Delegate Count for District " + district.districtNumber, "", {  
      htmlBody: emailTemplate,  
      from: "info@ksbstate.org",  
      name: "Kansas Boys State"  
    });
    GmailApp.sendEmail(district.salCommanderEmail, "Delegate Count for District " + district.districtNumber, "", {  
      htmlBody: emailTemplate,  
      from: "info@ksbstate.org",  
      name: "Kansas Boys State"  
    }); 
  }
  
}

function emailDepartmentCommander(){
  const sheet = spreadsheet.getSheetByName('2026 Delegate Count By Legion Dist');
  const lastRow = sheet.getLastRow();

  const range = sheet.getRange('A2:B' + lastRow);
  const values = range.getValues();
  
  var table = "";
  for(var i = 0; i < values.length; i++){
    table += '<tr>';
    table += '<td>' + values[i][0] + '</td>';
    table += '<td>' + values[i][1] + '</td>';
    table += '</tr>';
  }

  var emailTemplate = `
    <html>
      <head>
        <style>

          #delegates {
            border-collapse: collapse;
            max-width: 500 px;
          }

          #delegates td, #delegates th {
            border: 1px solid #ddd;
            padding: 8px;
          }

          #delegates tr:nth-child(even){background-color: #f2f2f2;}

          #delegates tr:hover {background-color: #ddd;}

          #delegates th {
            padding-top: 12px;
            padding-bottom: 12px;
            text-align: left;
            background-color: black;
            color: white;
          }
        </style>

      </head>
      <body>
        <p>${departmentAndDetachmentInfo.departmentCommanderName} and ${departmentAndDetachmentInfo.detachmentCommanderName}, </p>

        <p>Greetings from <i>your</i> Kansas Boys State! Below you will find a breakdown of the number of delegates per District for the 2026 Session of Boys State.</p>

        <table id="delegates">
          <tr>
            <th>Legion District</th>
            <th># of Delegates</th>
          </tr>
          ${table}
        </table>

        <p>As a reminder:<p>
        <ul>
          <li>There is no maximum number of delegates a post or school can send</li>
          <li>The cost to attend is $375, with $325 coming from a sponsor or Post</li>
          <li><strong>Who</strong> – Young men who are Juniors (preferred), or Sophomores in the 25-26 school year</li>
          <li><strong>When</strong> – Sunday, May 31st through Saturday, June 6th, 2026</li>
          <li><strong>Where</strong> – Kansas State University in Manhattan</li>
          <li><strong>Why</strong> – To inculcate a sense of individual obligation to community, state and nation</li>
          <li><strong>How</strong> – Apply at www.ksbstate.org</li>
          <li><strong><i><u>Bonus</u></i></strong> – College Credit from K-State and Scouting Merit Badges can now be earned for Boys Staters!</li>
        </ul>

        <p>Our goal for this year is 350 delegates, that is 32 delegates per district! This year we are awarding a plaque to the district that has the most number of delegates to have for the next year and we are looking for additional ways to reward the district and post with the most delegates.</p>

        <p>If you have any questions, please contact <a href="mailto:info@ksbstate.org">info@ksbstate.org</a>.</p>

        <p>Thank you for the support of your Kansas Boys State program!<p>

        <p>Kyle Wheatley<br>
        Executive Director<br>
        The American Legion Boys State of Kansas</p>
      </body>
    </html>
  `;

  if(OVERRIDE_EMAIL_TOGGLE){
    GmailApp.sendEmail(overrideEmail, "Boys State of Kansas Department Totals", "", {  
      htmlBody: emailTemplate,  
      from: "info@ksbstate.org",  
      name: "Kansas Boys State"  
    });  
  }
  else{
    GmailApp.sendEmail(departmentAndDetachmentInfo.departmentCommanderEmail, "Boys State of Kansas Department Totals", "", {  
      htmlBody: emailTemplate,  
      from: "info@ksbstate.org",  
      name: "Kansas Boys State"  
    }); 
    GmailApp.sendEmail(departmentAndDetachmentInfo.detachmentCommanderEmail, "Boys State of Kansas Department Totals", "", {  
      htmlBody: emailTemplate,  
      from: "info@ksbstate.org",  
      name: "Kansas Boys State"  
    });
  }
}

function getAliases() {
  var aliases = GmailApp.getAliases();
  Logger.log(aliases);
}
