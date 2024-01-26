// Valuetales is a Google Apps Script project intended to be ran on the Google Apps Script platform: https://script.google.com/
// The variables below will need to be filled out prior to running this script.

let startDate = new Date('01/01/2000'); // Start date in the format 'MM/DD/YYYY'. Set this as the same date the script will first be triggered. This will be used to keep track of the current week.
let senderEmail = ''; // Email address the sender (you).
let senderFirstName = ''; // First name of the sender (you).
let senderLastName = ''; // Last name of the sender (you). 
let recipients = [ // List of recipients. If there is only one recipient just add one value to this array.
  {
    firstName: '', // First name of the recipient.
    lastName: '', // Last name of the recipient.
    email: '', // Email address of the recipient.
    coverPhotoWeekOverride: -1, // Set this to a value between 1-50 to override the random image for the cover photo and take the image from the corresponding week. This can be set to -1 to keep the random image.
    googleDocLinkOverride: 'auto', // Set this to 'auto' to have the book auto-generated and sent without a review. Set this to 'skip' to skip sending the book on week 52. Set this to a Google Doc link to have your edited link sent instead.
  }
];
let questions = [ // List of 50 questions. Use these questions or make up your own.
  'What is the best vacation you\'ve ever been on and why?',
  'What are your favorite musicians, bands or albums?',
  'Describe one of your most memorable birthdays.',
  'What is some of the best advice you\'ve ever received in your life?',
  'What was your wedding like?',
  'Where did you grow up?',
  'What is one fad/trend from your younger years that you would like to see brought back? Are there any fads/trends that should never be brought back?',
  'What was your first job? Describe what it was like.',
  'Tell me something about your parents that I probably wouldn\'t know',
  'If you could relive a year of your life, what year would it be and why?',
  'If you could do any job in the world and it would pay you a comfortable salary what would you do and why?',
  'What are some personality traits of yours that you see in your children?',
  'Describe your best friend or friends.',
  'If you could choose your last meal what would it be?',
  'What about being an American makes you most proud?',
  'What was the first concert you went to?',
  'What\'s something you really disagreed with your parents about?',
  'Where were you and what can you tell me about when we landed on the moon, the events of 9/11, the Covid-19 pandemic, and any other historic events in your lifetime?',
  'What is the best deal you\'ve ever come across in your life?',
  'Do you have an oldest memory?',
  'What is one of the greatest physical challenges you have ever had to go through? What gave you the strength?',
  'What is one of your go-to stories, one you like telling over and over?',
  'What political issues do you consider most important?',
  'What traits do you share with your father and mother?',
  'What were your grandparents like?',
  'What weaknesses do you struggle with the most?',
  'What inventions have had the biggest impact on your day-to-day life?',
  'What was one of the most difficult things to overcome from your childhood? How did you do it?',
  'If you could pick someone to narrate your life who would it be and why?',
  'When was the first time you had an alcoholic beverage?',
  'What are some of your favorite quotes?',
  'What are some of your favorite films and TV Shows?',
  'What are your spiritual/religious beliefs?',
  'What accomplishments are you most proud of?',
  'Can you tell me a little about your siblings and growing up with them?',
  'What did you do when you were bored as a child?',
  'When you were little, what did you answer to the question: “What do you want to be when you grow up?”',
  'Are there any family recipes or meals that you would like to share?',
  'Where do you consider home? Did you ever have any desire to move away from where you consider home?',
  'What is the history of all the cars you and your family have owned?',
  'What are all the positions and job responsibilities you held throughout your career?',
  'Have you ever experienced a moment of pure awe or wonder? Have you ever experienced anything you would deem supernatural?',
  'What are some of your favorite feelings? (leisurely strolls on the beach perhaps)',
  'Tell me something I don\'t know about you.',
  'What was the best gift you\'ve ever received? What\'s the best gift you\'ve ever given?',
  'What are some of your favorite books?',
  'What are some of your happiest childhood memories?',
  'If you were to go on a separate trip with just you and each of your children where would you go for each child?',
  'If you had to create a family motto, what would it be?',
  'Are there any questions you wish you had asked your own parents? If so, what would be your answer to those same questions?'
];

// Google Docs values.
let pointsInInch = 72; // Google Docs default value.
let pixelsInInch = 96; // Google Docs default value.
let pageHeightInchFull = 8.5; // Page height in inches. This project is currently only setup to build a 5.5 inch x 8.5 inch book with 0.5 inch margins. Changing this value will result in unknown behavior.
let pageWidthInchFull = 5.5; // Page width in inches. This project is currently only setup to build a 5.5 inch x 8.5 inch book with 0.5 inch margins. Changing this value will result in unknown behavior.
let pageMarginInch = 0.5; // Page margins around the whole document in inches. This project is currently only setup to build a 5.5 inch x 8.5 inch book with 0.5 inch margins. Changing this value will result in unknown behavior.
let pageHeightInch = pageHeightInchFull - ( 2 * pageMarginInch ); // Page height of the editable area in inches.
let pageHeightUnit = 1800; // This 'unit' is specific to this project. This represents the height of the editable area. Changing this value result in unknown behavior.
let pageWidthUnit = 1080; // This 'unit' is specific to this project. This represents the width of the editable area. Changing this value result in unknown behavior.
const conversions = { // Conversions between all Google Docs units used in this project.
  unit_inch: pageHeightInch / pageHeightUnit,
  unit_pixel: (pageHeightInch / pageHeightUnit) * pixelsInInch,
  unit_point: (pageHeightInch / pageHeightUnit) * pointsInInch,
  inch_unit: pageHeightUnit / pageHeightInch,
  inch_pixel: pixelsInInch,
  inch_point: pointsInInch,
  pixel_unit: (1 / pixelsInInch) * (pageHeightUnit / pageHeightInch),
  pixel_inch: 1 / pixelsInInch,
  pixel_point: (1 / pixelsInInch) * pointsInInch,
  point_point: (1 / pointsInInch) * (pageHeightUnit / pageHeightInch),
  point_inch: 1 / pointsInInch,
  point_pixel: (1 / pointsInInch) * pixelsInInch
};

// Handles unit conversion between the Google Docs units: 'unit', 'inch', 'pixel', and 'point'.
function convertUnit(value, fromUnit, toUnit) {
  const factor = conversions[`${fromUnit}_${toUnit}`];
  if (!factor) {
    console.log('Invalid unit conversion. Using default value. May result in errors.');
    return value;
  }
  return value * factor;
}

// Validates a Google Doc link.
function validateGoogleDocLink(link){
  try {
    DocumentApp.openByUrl(link);
    return true;
  } catch (e) {
    console.log('Unable to open the Google Doc link: ' + link);
    return false;
  }
}

// Handles sending an email.
function sendEmail(data){
  let template = HtmlService.createTemplateFromFile('template'); // Creates template from file 'template.html'
  for (const [key, value] of Object.entries(data.template)) {
    template[key] = value;
  }
  let attachments = []; // Default is no attachments.
  if(data.hasOwnProperty('pdf')){
    attachments = [data.pdf];
  }
  let body = template.evaluate().getContent();
  GmailApp.sendEmail(
    data.email,
    data.subject,
    body, {
    htmlBody: body,
    inlineImages: {logo: data.logo},
    attachments: attachments
  });
}

// Handles asking the weekly questions. This should be setup as a weekly trigger in Google Apps Script.
function askQuestions(){
  let currentDate = new Date();
  if(startDate.getFullYear() <= 2023) { // Checking to see if the user (you) updated the startDate variable.
    console.error('Invalid startDate. Skipping askQuestions script. Please enter a valid startDate.');
    return 0;
  }
  let days = Math.floor((currentDate - startDate) / (24 * 60 * 60 * 1000));
  let weekNumber = Math.floor(days / 7);
  let base64logo = 'data:image/png;base64,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';
  let logoImageBlob = Utilities.newBlob(Utilities.base64Decode(base64logo.slice(22)),'image/png','logoImageBlob'); // Creates blob of 'valuetales' logo to use in email.

  if(weekNumber === 0){ // Special first week welcome message.
    for(let i = 0; i < recipients.length; i++){
      sendEmail({
        subject: 'Welcome to valuetales ' + recipients[i].firstName + '!',
        email: recipients[i].email,
        template: {
          paragraph: 'welcome',
          recipientFirstName: recipients[i].firstName,
          senderFirstName: senderFirstName,
          senderLastName: senderLastName,
        },
        logo: logoImageBlob
      });
    }
  }
  else if(weekNumber === 51){ // Special last week completion message.
    for(let i = 0; i < recipients.length; i++){
      // Sending thank you message to recipients.
      sendEmail({
        subject: 'A thank you from valuetales!',
        email: recipients[i].email,
        template: {
          paragraph: 'completed',
          recipientFirstName: recipients[i].firstName,
          senderFirstName: senderFirstName,
          senderLastName: senderLastName,
        },
        logo: logoImageBlob
      });
      // Sending message to the sender (you) to remind them that they have one week until the finished book is sent and to help out each recipient.
      sendEmail({
        subject: 'Last week to fix valuetales book for ' + recipients[i].firstName + ' ' + recipients[i].lastName,
        email: senderEmail,
        template: {
          paragraph: 'reminder',
          recipientFirstName: recipients[i].firstName,
          recipientLastName: recipients[i].lastName,
          senderFirstName: senderFirstName,
          senderLastName: senderLastName,
          googleDocLink: buildBook({
              firstName: recipients[i].firstName,
              lastName: recipients[i].lastName,
              email: recipients[i].email,
              coverPhotoWeekOverride: recipients[i].coverPhotoWeekOverride
            })
        },
        logo: logoImageBlob
      });
    }
  }
  else if(weekNumber === 52){ // Special completed book message.
    for(let i = 0; i < recipients.length; i++){
      let override = recipients[i].googleDocLinkOverride;
      if(override === 'skip'){
        // Send message to the sender (you) that they've skipped sending the book to the recipient.
        sendEmail({
          subject: 'Skipped sending valuetales book to ' + recipients[i].firstName + ' ' + recipients[i].lastName,
          email: senderEmail,
          template: {
            paragraph: 'skip',
            recipientFirstName: recipients[i].firstName,
            recipientLastName: recipients[i].lastName,
            senderFirstName: senderFirstName,
            senderLastName: senderLastName
          },
          logo: logoImageBlob
        });
      }
      else if(override === 'auto' || validateGoogleDocLink(override)){
        let link = override;
        if (override === 'auto'){
          link = buildBook({
            firstName: recipients[i].firstName,
            lastName: recipients[i].lastName,
            email: recipients[i].email,
            coverPhotoWeekOverride: recipients[i].coverPhotoWeekOverride
          })
        }
        let googleDocument = DocumentApp.openByUrl(link);
        googleDocument.saveAndClose();
        let pdfBlob = googleDocument.getAs('application/pdf');
        // Send completed book PDF message to the recipient.
        sendEmail({
          subject: 'Your valuetales book is ready to view!',
          email: recipients[i].email,
          template: {
            paragraph: 'book',
            recipientFirstName: recipients[i].firstName,
            senderFirstName: senderFirstName,
            senderLastName: senderLastName
          },
          logo: logoImageBlob,
          pdf: pdfBlob
        });
        // Send message to the sender (you) to let them know that the recipient has received the PDF.
        sendEmail({
          subject: 'Your valuetales book has been sent to ' + recipients[i].firstName + ' ' + recipients[i].lastName,
          email: senderEmail,
          template: {
            paragraph: 'buy',
            recipientFirstName: recipients[i].firstName,
            recipientLastName: recipients[i].lastName,
            senderFirstName: senderFirstName,
            senderLastName: senderLastName
          },
          logo: logoImageBlob,
          pdf: pdfBlob
        });
      }
      else{
        // Send message to the sender (you) that they selected an invalid value for googleDocLinkOverride and their book was not sent to the recipient.
        sendEmail({
          subject: 'Invalid valuetales googleDocLinkOverride value for ' + recipients[i].firstName + ' ' + recipients[i].lastName,
          email: senderEmail,
          template: {
            paragraph: 'invalid',
            recipientFirstName: recipients[i].firstName,
            recipientLastName: recipients[i].lastName,
            senderFirstName: senderFirstName,
            senderLastName: senderLastName
          },
          logo: logoImageBlob
        });
      }
    }
  }
  else if(weekNumber >= 52){ // Special message to sender letting them know to disable this service.
    sendEmail({
      subject: 'Reminder to disable the valuetales service',
      email: senderEmail,
      template: {
        paragraph: 'disable',
        senderFirstName: senderFirstName,
        senderLastName: senderLastName
      },
      logo: logoImageBlob
    });
  }
  else{ // Weekly question message.
    for(let i = 0; i < recipients.length; i++){
      sendEmail({
        subject: recipients[i].firstName + '\'s valuetales question week ' + String(weekNumber),
        email: recipients[i].email,
        template: {
          paragraph: 'question',
          recipientFirstName: recipients[i].firstName,
          senderFirstName: senderFirstName,
          senderLastName: senderLastName,
          question: questions[weekNumber - 1]
        },
        logo: logoImageBlob
      });
    }
  }
}

// Function to get the estimated paragraph height for a given paragraph and font size based on the total page height of 1800 units.
// The total page height of 1800 units was derived by using the monospaced font Courier in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins.
// The font sizes 10pt, 15pt, 18pt, and 30pt fit the document an exact number of times down and across. This is why the font Courier was chosen.
// 1800 units is chosen as the page height due to 450 being the lowest common multiple (LCM) of 45, 30, 25, and 15 which are the total number of lines in the document a 10pt, 15pt, 18pt, and 30pt can fit in respectively.
// The LCM of 450 was then multiplied by 4 to get 1800 units for the height to make the page height divisible by more values for more flexibility.
// There is no method in the Apps Script service for Google Docs to get a Paragraph's height making this function necessary to be able to keep track of the vertical alignment of items in the Document.
// The best workaround for this problem is to determine the number of lines in a Paragraph by approximating the Google Docs line wrap algorithm.
// After finding the number of lines in a Paragraph this value can then be multiplied by the line height for the specific font size of the Paragraph.
// The downside to this workaround is that there may be inconsistencies in this approximated line wrap algorithm and the Google Docs line wrap algorithm resulting in incorrect Paragraph line counts.
function getParagraphHeight(paragraph, fontSize, indented) {
  const lineWidths = {
    10: 54, // 10pt Courier font can fit 54 characters per line in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins.
    15: 36, // 15pt Courier font can fit 36 characters per line in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins.
    18: 30, // 18pt Courier font can fit 30 characters per line in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins.
    30: 18, // 30pt Courier font can fit 18 characters per line in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins.
  };
  const lineHeights = {
    10: 40, // 10pt Courier font can fit 45 lines in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins. 1800/45 = 40
    15: 60, // 15pt Courier font can fit 30 lines in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins. 1800/30 = 60
    18: 72, // 18pt Courier font can fit 25 lines in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins. 1800/25 = 72
    30: 120, // 30pt Courier font can fit 15 lines in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins. 1800/15 = 120
  };
  const indentCharacterCount = {
    10: 3, // 10pt Courier font can fit 3 characters in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins with a 0.25 inch tab.
    15: 2, // 15pt Courier font can fit 2 characters in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins with a 0.25 inch tab.
    18: 2, // 18pt Courier font can fit 2 characters in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins with a 0.25 inch tab. (This is the only font size where the tab is not an exact character count)
    30: 1, // 30pt Courier font can fit 1 characters in a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins with a 0.25 inch tab.
  };

  let lineWidth = lineWidths[fontSize]; // The maximum amount of characters that can fit in one line.
  let indentCount = indentCharacterCount[fontSize]; // The amount of characters in an indent.
  let indentLineWidth = lineWidth - indentCount; // The line width for the first line with an indent.
  let spaceLeft = lineWidth; // The amount of characters left in the current line.
  let words = paragraph.split(' '); // Words in the paragraph.
  let lineCount = 1; // Number of lines for the paragraph in Google Doc.
  let wordIterStart = 0;  // Array index to start iterating over words in paragraph.

  // Handling indented paragraphs.
  if(indented){
    if(words.length && words[0].length > indentLineWidth){ // Special case where word is longer than indented line.
      let indentAndWord = words[0].length + indentCount;
      spaceLeft = lineWidth - (indentAndWord % lineWidth)
      lineCount = lineCount + Math.floor(indentAndWord / lineWidth);
      wordIterStart = 1;
    }
    else{
      spaceLeft = lineWidth - indentCount;
    }
  }

  // Handling line wrap of paragraphs.
  for(let i = wordIterStart; i < words.length; i++){
    if(spaceLeft === lineWidth){
      if(words[i].length > lineWidth){
        spaceLeft = lineWidth - (words[i].length % lineWidth);
        lineCount = lineCount + Math.floor(words[i].length / lineWidth);
      }
      else{
        spaceLeft = lineWidth - words[i].length;
      }
    }
    else if(1 + words[i].length > spaceLeft){
      if(words[i].length > lineWidth){
        spaceLeft = lineWidth - (words[i].length % lineWidth);
        lineCount = lineCount + 1 + Math.floor(words[i].length / lineWidth);
      }
      else{
        spaceLeft = lineWidth - words[i].length;
        lineCount = lineCount + 1;
      }
    }
    else{
      spaceLeft = spaceLeft - (1 + words[i].length);
    }
  }
  return lineCount * lineHeights[fontSize];
}

// Used for testing the function buildBook with a recipient.
function buildBookTest(){
  buildBook(recipients[0])
}

// Builds the Google Docs book.
function buildBook(recipient) {
  // Gathering the emails.
  let subjectLine = recipient.firstName + '\'s valuetales question week';
  let threads = GmailApp.search('subject: ' + subjectLine);
  threads.sort(function(a, b){ // Sorts email threads by the question number.
    if (a.getFirstMessageSubject() < b.getFirstMessageSubject()){ return -1; }
    if (a.getFirstMessageSubject() > b.getFirstMessageSubject()){ return 1; }
    return 0;
  })

  // Parsing the email responses and images.
  let tales = []; // Array of valid 'tales'.
  for(let i = 0; i < threads.length; i++){
    let weekNumber = parseInt(threads[i].getFirstMessageSubject().split(' ').pop());
    let tale = {
      weekNumber: weekNumber,
      question: questions[weekNumber - 1],
      responseParagraphs: [],
      image: null
    };
    let imageFound = false; // Flag for an image attachment in the thread.
    let responseFound = false; // Flag for a suitable response in the thread.
    let messages = threads[i].getMessages();
    if(messages.length >= 2){
      for(let j = messages.length - 1; j >= 1 && (!imageFound || !responseFound); j--){ // Iterating through the thread's messages until a suitable image and response are found.
        // Gathering and formatting the email body into paragraphs.
        if(!responseFound){
          let messageBody = messages[j].getPlainBody();
          let regexList = [ // List of regexes used to clean email body by replacing them with an empty string. This is where you will need to remove items such as email signatures or previous messages in the thread.
            /On(\r\n|\s)*(Mon|Tue|Wed|Thu|Fri|Sat|Sun),(\r\n|\s)*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\r\n|\s)*([1-9]|[12][0-9]|3[01]),(\r\n|\s)*2[0-9]{3}(\r\n|\s)*at(\r\n|\s)*([0-1]?[0-9]|2[0-3]):[0-5][0-9](\r\n|\s)*(AM|PM)(\r\n|\s)*.*<.*?>(\r\n|\s)*wrote:(.|\n|\r)*/gm,
            /\[image: Logo\] <https:\/\/www\.my-company.com\/>(.|\n|\r)*/gm
          ]
          for (let k = 0; k < regexList.length; k++) {
            messageBody = messageBody.replace(regexList[k], '');
          }
          let paragraphs = messageBody.split(/(?:\r\n){2,}/gm); // Split paragraphs up by occurrences of consecutive '\r' and '\n' characters.
          let paragraphsFiltered = paragraphs.filter(function(a){return a !== ''}); // Removing any empty paragraphs.
          for (let k = 0; k < paragraphsFiltered.length; k++) {
            paragraphsFiltered[k] = paragraphsFiltered[k].replaceAll('\r\n', ' '); // Removing any extra new line characters.
          }
          if(paragraphsFiltered.length){
            tale.responseParagraphs = paragraphsFiltered;
            responseFound = true;
          }
        }
        // Checking for an image in the thread.
        if(!imageFound){
          let attachments = messages[j].getAttachments();
          if(attachments.length){
            for(let k = 0; k < attachments.length && !imageFound; k++){
              let contentType = attachments[k].getContentType();
              if(contentType.substring(0,5) === 'image'){
                tale.image = attachments[k].getAs(contentType);
                imageFound = true;
              }
            }
          }
        }
      }
    }
    if (tale.responseParagraphs.length){ // If there is a suitable paragraph add the 'tale'.
      tales.push(tale);
    }
  }

  // Setting up the Google Doc.
  let date = new Date();
  let doc = DocumentApp.create('valuetales-' + recipient.firstName + '-' + recipient.lastName + '-' + date.toUTCString());
  let body = doc.getBody();
  body.setPageHeight(pageHeightInchFull * pointsInInch);
  body.setPageWidth(pageWidthInchFull * pointsInInch);
  body.setMarginTop(pageMarginInch * pointsInInch);
  body.setMarginBottom(pageMarginInch * pointsInInch);
  body.setMarginLeft(pageMarginInch * pointsInInch);
  body.setMarginRight(pageMarginInch * pointsInInch);
  let paragraphNumber = 0;

  // Handles building paragraphs.
  let buildParagraph = (paragraph, fontSize, textStyle, alignment) => {
    let documentParagraph = body.insertParagraph(paragraphNumber,paragraph);
    paragraphNumber = paragraphNumber + 1;
    let style = {};
    style[DocumentApp.Attribute.FONT_FAMILY] = 'Courier';
    style[DocumentApp.Attribute.FONT_SIZE] = fontSize;
    style[DocumentApp.Attribute.LINE_SPACING] = 1;
    if(alignment === 'center'){
      style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    }
    else if(alignment === 'left'){
      style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
      documentParagraph.setIndentFirstLine(18); // Set in points (0.25 inch)
    }
    if(textStyle === 'normal'){
      style[DocumentApp.Attribute.BOLD] = false;
      style[DocumentApp.Attribute.ITALIC] = false;
    }
    else if(textStyle === 'italic'){
      style[DocumentApp.Attribute.BOLD] = false;
      style[DocumentApp.Attribute.ITALIC] = true;
    }
    else if(textStyle === 'bold'){
      style[DocumentApp.Attribute.BOLD] = true;
      style[DocumentApp.Attribute.ITALIC] = false;
    }
    documentParagraph.setAttributes(style);
    return {
      object: documentParagraph,
      height: getParagraphHeight(paragraph,fontSize,alignment === 'left')
    };
  };

  // Handles adding an image and scaling it to fit in the given height on the page.
  // Much of the logic in this function is a workaround for limitations in the Apps Script service for Google Docs.
  // A PositionedImage is required to be anchored to a Paragraph, but this results in an unwanted line after the image that can not be removed.
  // The best workaround for this is to add an empty line above the image in addition to this line to make the spacing look more intentional and even.
  // There is also no way to set the 'margins from text' on an image through the Apps Script service for Google Docs.
  // The default 'margins from text' on an image is 0.02 inches so this value will need to be subtracted from the height of the image to keep the units consistent.
  // The 'margins from text' on an image doesn't apply to the paragraph before the image so the paragraph before the image will need to add spacing after the paragraph to compensate.
  // The 'unit' 60, which is the line height of the font size 15pt, is used because when this 'unit' value is converted to inches the converted value will be in increments of 0.25 inches. 
  // Google Docs limits the width and height values in inches of images to two decimal places. Increments of 0.25 inches will keep the size of images accurate.
  let addImage = (image, height) =>  {
    let heightMarginsAndLine = 60 - (2 * convertUnit(0.02, 'inch', 'unit'));
    let trimHeight = height - (60 * 2) - (2 * convertUnit(0.02, 'inch', 'unit')); // Removing the height of the lines above and below and the margins.
    if(trimHeight < heightMarginsAndLine) { // Minimum height check.
      console.log('Invalid image height. Skipping adding image. May result in errors.');
      return 0;
    }
    else {
      let emptyParagraph = buildParagraph('',15, 'normal', 'center'); // This paragraph is for the space above an image.
      emptyParagraph.object.setSpacingAfter(convertUnit(0.02, 'inch', 'point')); // This is to counteract the margins from text on the image.
      let imageParagraph = buildParagraph('', 15, 'normal', 'center'); // This paragraph is necessary to contain the image but also places an unwanted line below the image.
      let positionedImage = imageParagraph.object.addPositionedImage(image);
      positionedImage.setLayout(DocumentApp.PositionedLayout.BREAK_BOTH).setTopOffset(-1.44); // -1.44 results in an offset of 0 (Apps Script Bug).
      // Scaling the image to fit in the box.
      let aspectRatioImage = positionedImage.getWidth() / positionedImage.getHeight();
      let aspectRatioBox = pageWidthUnit / trimHeight;
      let heightAdjusted = aspectRatioImage < aspectRatioBox ? trimHeight : (pageWidthUnit / aspectRatioImage); // Determining if the height needs to be adjusted.
      let imageHeight = heightMarginsAndLine + (Math.floor((heightAdjusted - heightMarginsAndLine) / 60) * 60); // Scaling Image to be multiple of 60.
      let imageWidth = imageHeight * aspectRatioImage;
      let leftOffset = convertUnit((pageWidthUnit - imageWidth)/2, 'unit', 'point');
      positionedImage.setLeftOffset(leftOffset - 1.44); // 1.44 is added to offsets (Apps Script Bug).
      positionedImage.setHeight(convertUnit(imageHeight, 'unit', 'pixel'));
      positionedImage.setWidth(convertUnit(imageWidth, 'unit', 'pixel'));
      return {
        object: positionedImage,
        height: imageHeight + (2 * convertUnit(0.02, 'inch', 'unit')) + (60 * 2)
      };
    }
  }
  
  // Adding the recipient's name to the title page.
  let titlePageNameText = recipient.firstName + ' ' + recipient.lastName;
  let titlePageNameParagraph = buildParagraph(titlePageNameText, 30, 'normal', 'center');
  
  // Adding image to the title page.
  let titlePageImage = { 
    object: null, 
    height: 0 
  };
  let coverTaleOverride = tales.find((tale) => tale.weekNumber === recipient.coverPhotoWeekOverride && tale.image !== null);
  if (recipient.coverPhotoWeekOverride >= 1 && recipient.coverPhotoWeekOverride <= 50 && coverTaleOverride) {
    titlePageImage = addImage(coverTaleOverride.image, pageHeightUnit * 0.625);
  } 
  else {
    let talesWithImages = tales.filter(function(a){return a.image !== null});
    if (talesWithImages.length){
      let randomTitleImage = talesWithImages[Math.floor(Math.random() * talesWithImages.length)].image
      titlePageImage = addImage(randomTitleImage, pageHeightUnit * 0.625);
    }
  }

  // Adding description to the title page.
  let descriptionYear = `${startDate.getFullYear()}`;
  if(startDate.getMonth() > 3 && startDate.getMonth() < 9){
    descriptionYear = `${startDate.getFullYear()}/${startDate.getFullYear() + 1}`
  }
  else if(startDate.getMonth() >= 9) {
    descriptionYear = `${startDate.getFullYear() + 1}`
  }
  let titlePageDescriptionText = `This book contains answers to weekly questions that ${senderFirstName} ${senderLastName} asked ${recipient.firstName} over the course of the year in ${descriptionYear}.`;
  if(getParagraphHeight(titlePageDescriptionText, 10, true) + titlePageImage.height + titlePageNameParagraph.height > pageHeightUnit){
    buildParagraph('', 10, 'normal', 'center').object.appendPageBreak();
  }
  buildParagraph(titlePageDescriptionText, 10, 'normal', 'left').object.appendPageBreak();

  // Adding in the question pages.
  for(let i = 0; i < tales.length; i++){
    // Adding the question.
    let questionParagraph = buildParagraph(tales[i].question, 18, 'italic', 'center');
    let questionPageHeight = questionParagraph.height;

    // Adding the image.
    if(tales[i].image){
      let questionImage = addImage(tales[i].image, pageHeightUnit * 0.4);
      questionPageHeight = questionPageHeight + questionImage.height;
    }
    else {
      let questionParagraphSpace = buildParagraph('', 15, 'normal', 'center');
      questionPageHeight = questionPageHeight + questionParagraphSpace.height;
    }

    // Adding the response.
    let responseParagraph = {
      object: null,
      height: 0
    };
    for(let j = 0; j < tales[i].responseParagraphs.length; j++){
      responseParagraph = buildParagraph(tales[i].responseParagraphs[j], 10, 'normal', 'left');
      if(j !== tales[i].responseParagraphs.length - 1){ // Adding a space between paragraphs.
        buildParagraph('', 10, 'normal', 'left');
      }
    }
    if(i !== tales.length - 1 && responseParagraph.object){ // Adding a page break between page questions.
      responseParagraph.object.appendPageBreak();
    }
  }

  // Returns link to Google Doc
  return doc.getUrl();
}