/**
 * Valuetales is a Google Apps Script project intended to be ran on the Google Apps Script platform.
 * You can access the Google Apps Script platform at: https://script.google.com/.
 * The variables below will need to be filled out prior to running this script.
 */

// Set startDate as the same date the script will first be triggered. This is used to keep track of the week number.
const startDate = new Date("01/01/1970"); // Start date in the format 'MM/DD/YYYY'.
const sender = {
  firstName: "", // First name of the sender (you).
  lastName: "", // Last name of the sender (you).
  email: "", // Email address the sender (you).
};

/**
 * List of recipients. If there is only one recipient just add one value to this array.
 *
 * coverPhoto: Handles selecting the cover photo for the Google Doc. Setting this to a value between 1 - 50 will use an
 * image from the corresponding week as the cover photo. Setting this value to null will use a random image for the
 * cover photo. If the recipient hasn't responded with images then there won't be an image used on the cover page.
 *
 * googleDocLinkOverride: Handles overriding the Google Doc sent at the end of the year. If a valid Google Doc link is
 * supplied, this Google Doc will be used instead of the auto generated one. This is helpful if you want to manually
 * make any edits before sending the Google Doc to your recipient. To use the auto generated option set this value to
 * null.
 */
const recipients = [
  {
    firstName: "", // First name of the recipient.
    lastName: "", // Last name of the recipient.
    email: "", // Email address of the recipient.
    coverPhoto: null, // Handles selecting the cover photo for the Google Doc.
    googleDocLinkOverride: null, // Handles overriding the Google Doc sent at the end of the year.
  },
];

// List of 50 questions. Use these questions or make up your own.
const questions = [
  "What is the best vacation you've ever been on and why?",
  "Who are some of your favorite musicians? What about bands or specific albums?",
  "Can you tell me about your most memorable birthday?",
  "What is the best advice you've ever received in your life?",
  "What was your wedding like?",
  "Where did you grow up?",
  "What is one fad/trend from your younger years that you would like to see brought back? What fads/trends should never be brought back?",
  "What was your first job? What it was like?",
  "What can you tell me about your parents that I probably wouldn't know?",
  "If you could relive a year of your life, what year would it be and why?",
  "If you could do any job in the world and it would pay you a comfortable salary what would you do and why?",
  "What are some personality traits of yours that you see in your children?",
  "Who is your best friend and/or friends?",
  "If you could choose your last meal what would it be?",
  "What is a lesson you've learned the hard way?",
  "What was the first concert you went to?",
  "What's something you really disagreed with your parents about?",
  "Where were you and what can you tell me about large historic events that happened in your lifetime?",
  "What is the best deal you've ever come across in your life?",
  "Do you have an oldest memory?",
  "What is one of the greatest physical challenges you have ever had to go through?",
  "What is one of your go-to stories, one you like telling over and over?",
  "What political issues do you consider most important?",
  "What traits do you share with your father and mother?",
  "What were your grandparents like?",
  "What weaknesses do you struggle with the most?",
  "What inventions have had the biggest impact on your day-to-day life?",
  "What was one of the most difficult things to overcome from your childhood? How did you do it?",
  "If you could pick someone to narrate your life who would it be and why?",
  "When was the first time you had an alcoholic beverage?",
  "What are some of your favorite quotes?",
  "What are some of your favorite films and TV Shows?",
  "What are your spiritual/religious beliefs?",
  "What accomplishments are you most proud of?",
  "Can you tell me a little about your siblings and growing up with them?",
  "What did you do when you were bored as a child?",
  "When you were little, what did you answer to the question: “What do you want to be when you grow up?”",
  "Are there any family recipes or meals that you would like to share?",
  "Where do you consider home? Did you ever have any desire to move away from where you consider home?",
  "What is the history of all the cars you and your family have owned?",
  "What are all the positions and job responsibilities you have had throughout your career?",
  "Have you ever experienced a moment of pure awe or wonder?",
  "Have you ever experienced anything you would deem supernatural?",
  "What is something I don't know about you?",
  "What was the best gift you've ever received? What's the best gift you've ever given?",
  "What are some of your favorite books?",
  "What are some of your happiest childhood memories?",
  "If you were to go on a separate trip with just you and each of your children where would you go for each child?",
  "If you had to create a family motto, what would it be?",
  "Are there any questions you wish you had asked your own parents? If so, what would be your answer to those same questions?",
];

// Google Docs values.
// This script is configured to build a 5.5 inch x 8.5 inch Google Doc with 0.5 inch margins using the Courier font.
const POINTS_IN_INCH = 72; // Google Docs default value.
const PIXELS_IN_INCH = 96; // Google Docs default value.
const PAGE_FULL_HEIGHT_INCH = 8.5; // Page height in inches.
const PAGE_FULL_WIDTH_INCH = 5.5; // Page width in inches.
const PAGE_MARGIN_INCH = 0.5; // Page margins around the whole document in inches.
const PAGE_HEIGHT_INCH = PAGE_FULL_HEIGHT_INCH - 2 * PAGE_MARGIN_INCH; // Page height of the editable area in inches.
const PAGE_WIDTH_INCH = PAGE_FULL_WIDTH_INCH - 2 * PAGE_MARGIN_INCH; // Page width of the editable area in inches.
const IMAGE_HEIGHT_INCH = 3; // Ideal height of images in inches.
const PARAGRAPH_INDENT_INCH = 0.25; // Paragraph indent size in inches.
const PARAGRAPH_OFFSET_POINTS = 1.44; // This offset value is added to paragraphs by default.
const DEFAULT_PARAGRAPH_STYLE = {
  // Default paragraph style options for every paragraph.
  [DocumentApp.Attribute.FONT_FAMILY]: "Courier",
  [DocumentApp.Attribute.LINE_SPACING]: 1,
};
const FONTS = {
  // Font sizes used in this project.
  small: 10,
  medium: 15,
  large: 18,
  huge: 30,
};

// Valuetales logo.
const BASE_64_LOGO =
  "data:image/png;base64,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";
const LOGO_IMAGE_BLOB = Utilities.newBlob(Utilities.base64Decode(BASE_64_LOGO.slice(22)), "image/png", "logoImageBlob");

// Weeks to send special emails instead of weekly question emails.
const FIRST_WEEK = 0;
const WRAP_UP_WEEK = 51;
const BOOK_DELIVERY_WEEK = 52;

// Forms titles.
const FORMS_TEMPLATE_NAME = "Valuetales Google Forms Template";
const FORMS_TITLE_PREFIX = "Valuetales Question Week";

// Max PDF size in bytes.
const MAX_BYTES_PDF = 15 * 1024 * 1024;

// Creates individual Google Forms for each question based on a Google Forms template.
const createForms = () => {
  const templateFormDriveFile = getFormDriveFile(FORMS_TEMPLATE_NAME);
  if (!templateFormDriveFile) {
    return;
  }
  const parentFolder = templateFormDriveFile.getParents().next();
  const formsFolder = parentFolder.createFolder("valuetales-forms");
  console.log("Setting up all valuetales questions as individual Google Forms. This may take some time...");
  questions.forEach((question, index) => {
    const formTitle = `${FORMS_TITLE_PREFIX} ${index + 1}`;
    console.log(`Creating ${formTitle}. ${question}`);
    const newFormDriveFile = templateFormDriveFile.makeCopy(formTitle, formsFolder);
    const newFormDriveFileID = newFormDriveFile.getId();
    const form = FormApp.openById(newFormDriveFileID);
    form.setTitle(formTitle);
    form.addParagraphTextItem().setTitle(question).setRequired(true);
    form.moveItem(form.getItems().length - 1, 0);
    recipients.forEach((recipient) => {
      form.addPublishedReader(recipient.email);
    });
  });
  console.log("All valuetales Google Forms have been created.");
  console.log("Make sure to open up each Google Form and restore the missing file upload folder.");
};

// Searches Google Drive for a Google Form matching a file name.
const getFormDriveFile = (fileName) => {
  const namedFormFileIterator = DriveApp.getFilesByName(`${fileName}`);
  while (namedFormFileIterator.hasNext()) {
    const file = namedFormFileIterator.next();
    if (file.getMimeType() === "application/vnd.google-apps.form") {
      console.log(`Google Form with the name "${fileName}" found.`);
      return file;
    }
  }
  console.error(`Google Form with the name "${fileName}" missing.`);
  return null;
};

// Handles sending weekly question emails. This should be setup as a weekly trigger in Google Apps Script.
const weeklyEmail = () => {
  const days = Math.floor((new Date() - startDate) / (24 * 60 * 60 * 1000));
  const week = Math.floor(days / 7);
  recipients.forEach((recipient) => {
    const recipientFirstName = recipient.firstName;
    const recipientLastName = recipient.lastName;
    const recipientFullName = `${recipientFirstName} ${recipientLastName}`;
    if (week === FIRST_WEEK) {
      // Special first week welcome message.
      sendEmail({
        recipient,
        subject: `Welcome to valuetales ${recipientFirstName}!`,
        paragraph: "welcome",
      });
    } else if (week === WRAP_UP_WEEK) {
      // Special last week wrap up completion messages.
      // Sending thank you message to recipients.
      sendEmail({
        recipient,
        subject: "A thank you from valuetales!",
        paragraph: "completed",
      });
      // Sending message to the sender (you) to remind them that they have one week until the finished book is sent to recipient.
      sendEmail({
        recipient,
        subject: `Last week to fix valuetales book for ${recipientFullName}`,
        paragraph: "reminder",
        googleDocLink: buildBook(recipient),
        toSender: true,
      });
    } else if (week === BOOK_DELIVERY_WEEK) {
      // Special completed book message.
      const bookPDF = getBookPDF(recipient, true);
      if (bookPDF) {
        const isLink = typeof bookPDF === "string";
        // Send completed book PDF message to the recipient.
        sendEmail({
          recipient,
          subject: "Your valuetales book is ready to view!",
          paragraph: "book",
          bookLink: isLink ? bookPDF : null,
          attachments: isLink ? [] : [bookPDF],
        });
        // Send message to the sender (you) to let them know that the recipient has received the book PDF.
        sendEmail({
          recipient,
          subject: `Your valuetales book has been sent to ${recipientFullName}`,
          paragraph: "buy",
          bookLink: isLink ? bookPDF : null,
          attachments: isLink ? [] : [bookPDF],
          toSender: true,
        });
      } else {
        // Send message to the sender (you) that they selected an invalid value for googleDocLinkOverride. Book was not sent to the recipient.
        sendEmail({
          recipient,
          subject: `Invalid valuetales googleDocLinkOverride value for ${recipientFullName}`,
          paragraph: "invalid",
          toSender: true,
        });
      }
    } else if (week >= BOOK_DELIVERY_WEEK) {
      // Special message to the sender (you) letting them know to disable this service.
      sendEmail({
        recipient,
        subject: "Reminder to disable the valuetales service",
        paragraph: "disable",
        toSender: true,
      });
    } else {
      // Weekly question message.
      sendEmail({
        recipient,
        subject: `${recipientFirstName}'s valuetales question week ${week}`,
        paragraph: "question",
        question: questions[week - 1],
        googleFormLink: getFormResponderLink(week),
      });
    }
  });
};

// Handles sending an email using an HTML template.
const sendEmail = ({
  recipient,
  subject,
  paragraph,
  bookLink = null,
  googleDocLink = null,
  googleFormLink = null,
  question = null,
  attachments = [],
  toSender = false,
}) => {
  const templateHTML = HtmlService.createTemplateFromFile("template");
  templateHTML.paragraph = paragraph;
  templateHTML.senderFirstName = sender.firstName;
  templateHTML.senderLastName = sender.lastName;
  templateHTML.recipientFirstName = recipient.firstName;
  templateHTML.recipientLastName = recipient.lastName;
  templateHTML.bookLink = bookLink;
  templateHTML.googleDocLink = googleDocLink;
  templateHTML.googleFormLink = googleFormLink;
  templateHTML.question = question;
  const body = templateHTML.evaluate().getContent();
  const email = toSender ? sender.email : recipient.email;
  GmailApp.sendEmail(email, subject, body, {
    htmlBody: body,
    inlineImages: { logo: LOGO_IMAGE_BLOB },
    attachments: attachments,
  });
};

// Used for testing buildBook for a recipient. Enter a recipients email to test building the book.
const validateBuildBook = (email = "") => {
  const recipient = recipients.find((recipient) => recipient.email === email);
  if (!recipient) {
    console.error(`Recipient with email "${email}" not found.`);
    return;
  }
  console.log(buildBook(recipient));
};

// Builds the Google Docs book.
const buildBook = (recipient) => {
  const tales = gatherTales(recipient); // Getting all responses and images from Google Forms.
  const { doc, body } = setupDoc(recipient); // Setting up the Google Doc.
  createTitlePage(body, recipient, tales); // Creating title page.
  createQuestionPages(body, tales); // Adding question pages.
  return doc.getUrl(); // Returns link to Google Doc
};

// Gathers all Google Forms responses and images from a recipient.
const gatherTales = (recipient) => {
  const tales = [];
  questions.forEach((question, index) => {
    const formName = `${FORMS_TITLE_PREFIX} ${index + 1}`;
    const formDriveFile = getFormDriveFile(formName);
    if (formDriveFile) {
      const formDriveFileID = formDriveFile.getId();
      const form = FormApp.openById(formDriveFileID);
      const submissions = form.getResponses();
      submissions.forEach((submission) => {
        if (submission.getRespondentEmail() === recipient.email) {
          const itemResponses = submission.getItemResponses();
          let driveImage = null;
          if (itemResponses.length >= 2) {
            const imageID = itemResponses[1].getResponse();
            const { height, width, rotation } = Drive.Files.get(imageID).imageMediaMetadata;
            driveImage = {
              file: DriveApp.getFileById(imageID),
              height: rotation % 2 !== 0 ? width : height,
              width: rotation % 2 !== 0 ? height : width,
            };
          }
          tales.push({
            week: index + 1,
            question,
            response: itemResponses[0].getResponse(),
            image: driveImage,
          });
        }
      });
    }
  });
  return tales;
};

// Sets up a new Google Doc for a recipient.
const setupDoc = (recipient) => {
  const doc = DocumentApp.create(`valuetales-${recipient.firstName}-${recipient.lastName}-${new Date().toUTCString()}`);
  const body = doc.getBody();
  body.setPageHeight(PAGE_FULL_HEIGHT_INCH * POINTS_IN_INCH);
  body.setPageWidth(PAGE_FULL_WIDTH_INCH * POINTS_IN_INCH);
  body.setMarginTop(PAGE_MARGIN_INCH * POINTS_IN_INCH);
  body.setMarginBottom(PAGE_MARGIN_INCH * POINTS_IN_INCH);
  body.setMarginLeft(PAGE_MARGIN_INCH * POINTS_IN_INCH);
  body.setMarginRight(PAGE_MARGIN_INCH * POINTS_IN_INCH);
  return { doc, body };
};

// Handles creating the title page for the Google Doc book.
const createTitlePage = (body, recipient, tales) => {
  // Adding the recipient's name to the title page.
  const titlePageNameText = `${recipient.firstName} ${recipient.lastName}`;
  buildParagraph(body, titlePageNameText, FONTS.huge, "normal", "center");
  // Adding image to the title page.
  const talesWithImages = tales.filter((tale) => tale.image !== null);
  const coverPhoto = talesWithImages.find((tale) => tale.week === recipient.coverPhoto)?.image;
  const randomImage =
    talesWithImages.length > 0 ? talesWithImages[Math.floor(Math.random() * talesWithImages.length)]?.image : null;
  if (coverPhoto) {
    addImage(body, coverPhoto);
  } else if (randomImage) {
    addImage(body, randomImage);
  }
  // Adding description to the title page.
  let descriptionYear = `${startDate.getFullYear()}`;
  if (startDate.getMonth() > 3 && startDate.getMonth() < 9) {
    descriptionYear = `${startDate.getFullYear()}/${startDate.getFullYear() + 1}`;
  } else if (startDate.getMonth() >= 9) {
    descriptionYear = `${startDate.getFullYear() + 1}`;
  }
  const titlePageDescriptionText = `This book is a collection of answers to weekly questions that ${sender.firstName} ${sender.lastName} asked ${recipient.firstName} over the course of the year in ${descriptionYear}.`;
  buildParagraph(body, titlePageDescriptionText, FONTS.small, "normal", "left").appendPageBreak();
};

// Appends a paragraph to a Google Doc body.
const buildParagraph = (body, text, fontSize, textStyle, alignment) => {
  const documentParagraph = body.appendParagraph(text);
  const style = DEFAULT_PARAGRAPH_STYLE;
  // Setting font size.
  style[DocumentApp.Attribute.FONT_SIZE] = fontSize;
  // Setting paragraph alignment.
  if (alignment === "center") {
    style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  } else if (alignment === "left") {
    style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
    documentParagraph.setIndentFirstLine(PARAGRAPH_INDENT_INCH * POINTS_IN_INCH);
  }
  // Setting bold and italics.
  style[DocumentApp.Attribute.BOLD] = textStyle === "bold" ? true : false;
  style[DocumentApp.Attribute.ITALIC] = textStyle === "italic" ? true : false;
  // Applying styles.
  documentParagraph.setAttributes(style);
  // Removes starter paragraph. There needs to be at least one paragraph in the document at all times.
  if (body.getNumChildren() === 2 && body.getChild(0).getText() === "") {
    body.getChild(0).removeFromParent();
  }
  return documentParagraph;
};

// Handles adding an image to a Google Doc body.
const addImage = (body, image) => {
  buildParagraph(body, "", FONTS.medium, "normal", "center");
  const imageParagraph = buildParagraph(body, "", FONTS.medium, "normal", "center");
  buildParagraph(body, "", FONTS.medium, "normal", "center"); // Need a paragraph after image here (Apps Script Bug).
  const positionedImage = imageParagraph.addPositionedImage(image.file);
  const landscape = image.width > image.height;
  const aspectRatioImage = image.width / image.height;
  const imageWidth = landscape ? PAGE_WIDTH_INCH : IMAGE_HEIGHT_INCH * aspectRatioImage;
  const imageHeight = landscape ? PAGE_WIDTH_INCH / aspectRatioImage : IMAGE_HEIGHT_INCH;
  const leftOffset = ((PAGE_WIDTH_INCH - imageWidth) / 2) * POINTS_IN_INCH;
  positionedImage.setLayout(DocumentApp.PositionedLayout.BREAK_BOTH).setTopOffset(-PARAGRAPH_OFFSET_POINTS); // Paragraph offset (Apps Script Bug).
  positionedImage.setLeftOffset(leftOffset - PARAGRAPH_OFFSET_POINTS); // Paragraph offset (Apps Script Bug).
  positionedImage.setHeight(imageHeight * PIXELS_IN_INCH);
  positionedImage.setWidth(imageWidth * PIXELS_IN_INCH);
};

// Handles creating all the question and response pages.
const createQuestionPages = (body, tales) => {
  tales.forEach((tale) => {
    // Adding the question.
    buildParagraph(body, tale.question, FONTS.large, "italic", "center");
    // Adding the image.
    if (tale.image) {
      addImage(body, tale.image);
    } else {
      buildParagraph(body, "", FONTS.medium, "normal", "center");
    }
    // Adding the response.
    tale.response.split("\n").forEach((responseParagraph) => {
      buildParagraph(body, responseParagraph, FONTS.small, "normal", "left");
    });
    buildParagraph(body, "", FONTS.medium, "normal", "center").appendPageBreak();
  });
};

// Builds the final book PDF for a recipient.
const getBookPDF = (recipient, share = false) => {
  const override = recipient.googleDocLinkOverride;
  if (override === null || validateGoogleDocLink(override)) {
    const link = override === null ? buildBook(recipient) : override;
    const googleDocument = DocumentApp.openByUrl(link);
    googleDocument.saveAndClose();
    const blobPDF = googleDocument.getAs("application/pdf");
    blobPDF.setName(`${recipient.firstName}-${recipient.lastName}-valuetales.pdf`);
    if (blobPDF.getBytes().length > MAX_BYTES_PDF) {
      const drivePDF = DriveApp.createFile(blobPDF);
      if (share) {
        Drive.Permissions.insert(
          {
            role: "reader",
            type: "user",
            value: recipient.email,
          },
          drivePDF.file.getId(),
          {
            sendNotificationEmails: "false",
          }
        );
      }
      return drivePDF.getUrl();
    }
    return blobPDF;
  }
  return null;
};

// Validates a Google Doc link.
const validateGoogleDocLink = (link) => {
  try {
    DocumentApp.openByUrl(link);
    return true;
  } catch (error) {
    console.error(`Invalid Google Doc link: ${link} - ${error}`);
    return false;
  }
};

// Gets the responder link for a specific weekly form.
const getFormResponderLink = (week) => {
  const weeklyFormDriveFile = getFormDriveFile(`${FORMS_TITLE_PREFIX} ${week}`);
  if (!weeklyFormDriveFile) {
    return null;
  }
  const weeklyFormDriveFileID = weeklyFormDriveFile.getId();
  const weeklyForm = FormApp.openById(weeklyFormDriveFileID);
  return weeklyForm.getPublishedUrl();
};
