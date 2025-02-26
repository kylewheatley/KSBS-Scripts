function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var subject = "Kansas Boys State wants YOU!";
  var emailCount = 0;
  var maxEmails = 450; // Limit to 450 emails per run
  var totalEmails = 1980; // Total emails in the list (update if needed)
  var lastRow = Math.min(data.length, totalEmails + 1); // Prevent out-of-bounds errors

  var fileId = "1WB2gvsqfb8gs7PI5zPY2Bviif-CILs5i"; // Replace with your actual file ID
  var pdfFile = DriveApp.getFileById(fileId); // Get the PDF file from Google Drive

  var template = `
    <p><strong><firstname>,</strong></p>
    
    <p>My name is <strong>Kyle Wheatley</strong>, and I am the Executive Director of the American Legion Boys State of Kansas. I would love to invite you to apply to attend <strong>Kansas Boys State</strong> this summer from <strong>June 1 - June 7</strong> on the campus of <strong>Kansas State University</strong>.</p>

    <p>Kansas Boys State is a fast-paced, fun leadership and teamwork experience that fosters self-identity, mutual respect, and civic responsibility. Through simulated elections, political parties, and government at various levels, you develop skills in leadership, public speaking, conflict resolution, and networking.</p>

    <p><strong>Kansas Boys State is <em>not</em> a politics or debate camp</strong>—it is a leadership experience that has something in it for every student. From the football field, to the stage, to the classroom and beyond, Kansas Boys State has something to offer you <strong>no matter your background</strong>.</p>

    <h3>Highlights and opportunities include:</h3>
    <ul>
      <li><strong>Earn 3 hours of college credit</strong> from Kansas State University at a deeply discounted price</li>
      <li>Attend a <strong>college and career fair</strong> with over <strong>40 colleges, universities, and trade schools</strong></li>
      <li>Complete up to <strong>7 Scouting Merit Badges</strong></li>
      <li>Join a <strong>band and choir</strong> directed by Kansas State University faculty members</li>
      <li>Apply for a <strong>$1,000 director's scholarship</strong> and other <strong>Kansas American Legion Scholarships</strong></li>
      <li>Apply for the <strong>$1,250 Kansas Samsung Scholarship</strong> (eligible for $5,000 regional & $10,000 national scholarships)</li>
      <li><strong>Earn guaranteed scholarships</strong> at local colleges and universities</li>
      <li>Attend <strong>morning strength & conditioning workouts</strong> for Fall sports requirements</li>
    </ul>

    <p>The cost to attend is just <strong>$375</strong>, which covers <strong>food, housing, and programming for the week</strong>. Typically, you pay <strong>$50</strong>, and the remaining <strong>$325</strong> is covered by an <strong>American Legion Post or a community organization</strong> (e.g., Lions Club, Kiwanis Club, Rotary Club, Optimist Club, or even a church). If you need sponsorship help, <strong>apply anyway</strong>—we can help you find a sponsor!</p>

    <p>We would love to see you at Boys State this summer, where you'll join <strong>thousands of Kansans</strong> who have attended before you! A copy of our brochure is attached for you to see more about our program or visit our website to see a video about what a week at Boys State is like.</p>

    <p>🔗 <strong>Visit our website:</strong> <a href="https://ksbstate.org" target="_blank">https://ksbstate.org</a></p>
    <p>📸 <strong>Follow us on Instagram:</strong> <a href="https://www.instagram.com/kansasboysstate" target="_blank">@kansasboysstate</a></p>

    <p><strong>Onward!</strong></p>

    <p><strong>Kyle Wheatley</strong><br>
    Executive Director<br>
    The American Legion Boys State of Kansas</p>
  `;

  for (var i = 1; i < lastRow; i++) { // Loop only up to the last valid row
    var email = data[i][12]; // Column M (index 12)
    var firstName = data[i][0]; // Column A (index 0)
    var status = data[i][15]; // Column P (index 15)

    // Send only if email exists and status is blank
    if (email && status !== "Yes") {
      var message = template.replace("<firstname>", firstName);
      
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: message,
        attachments: [pdfFile.getAs(MimeType.PDF)] // Attach PDF
      });

      sheet.getRange(i + 1, 16).setValue("Yes"); // Update column P

      emailCount++;
      if (emailCount >= maxEmails) {
        Logger.log("Reached daily email limit of 450. Stopping.");
        break; // Stop execution after 450 emails
      }
    }
  }

  if (emailCount === 0) {
    Logger.log("No more emails left to send. All done!");
  }
}