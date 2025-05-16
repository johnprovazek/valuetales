# valuetales

## Description

This project was created as an open-source and cheaper alternative to the service [StoryWorth](https://welcome.storyworth.com/).
[StoryWorth](https://welcome.storyworth.com/) is a service that sends a loved-one weekly email question prompts throughout the course of the year.
At the end of the year [StoryWorth](https://welcome.storyworth.com/) bundles up the questions and responses into a book and sends you that book.

This project utilizes the products Gmail, Google Forms, Google Docs, Google Drive, and the [Google Apps Script platform](https://developers.google.com/apps-script) to orchestrate this service.
This project handles sending a loved-one weekly email question prompts to answer via a Google Form.
At the end of the year this project compiles the questions and responses into a Google Doc.
A third-party service such as [Barnes and Noble](https://press.barnesandnoble.com/print-on-demand) could then be used to print out this Google Doc as a physical book.

Built using Google Apps Script.

<div align="center">
  <picture>
    <img src="https://repository-images.githubusercontent.com/748037446/3c8c2418-113e-4bb0-934b-2328f3986f22" width="830px">
  </picture>
</div>

## Installation

### Adding Project Files

- Start by creating a new [Apps Script](https://script.google.com/) project named _valuetales_.
- Add the files [valuetales.gs](./valuetales.gs) and [template.html](./template.html) to the project. Make sure that these files are named correctly in the Apps Script project.
- You will need you to fill out a few variables at the top of the [valuetales.gs](./valuetales.gs) script that are unique to you and your loved-ones.

### Add Drive API Service

A [limitation](https://support.google.com/drive/thread/329524349/image-info-in-google-drive-on-windows-swaps-width-and-height?hl=en) with Apps Script and Google Drive is that image height and width values will occasionally be swapped if the image was rotated.
With Apps Script there is no means of getting the rotation value to set the appropriate height and width of an image.
This will occasionally lead to distorted images when adding them to the Google Doc at the end of year.
The work around for this problem is adding the Drive v2 API service to your project.
This API can be used to get the height, width, and rotation values of an image to then set the appropriate height and width of the image.

There is a second limitation in Apps Script that can be solved with adding the Drive v2 API.
At the end of the year a book is created as a Google Doc then compiled and sent as a PDF in an email attachment.
In the event the PDF attachment is too large the PDF is added to your Google Drive and shared with the recipient via a Google Drive share link.
With Apps Script you cannot disable sending a Google Drive share notification email.
This results in sending two emails to your recipient.
One directly from you on behalf of _valuetales_ and one share notification email from Google Drive.
Using the Drive API you can configure this share notification email not to be sent.

- To add this API to your project click on the plus sign or "Add a service" button next to Services.
- Find and select the "Drive API" service.
- Set the API version to "v2".
- Take special care to set the identifier as "Drive" without any spaces in the name.
  By default there may be a space before and after "Drive".
- Finally, add this service to your project.

### Google Forms Setup

This project uses Google Forms to handle your loved-one's responses to questions.
Due to limitations within Apps Script much of the setup for the Google Forms used in this project will need to be done manually.
Google Forms in Apps Script does not allow creating "File upload" questions, setting the "Collect email addresses" option to "Verified" or setting the "Responder view" to "Restricted".
The work around for this problem is to manually create a Google Forms template that will then be copied for each question with a function.

- Start by creating a new folder in your [Google Drive](https://drive.google.com/drive/u/0/my-drive) called "valuetales".
- Next, create a new Google Form inside this folder. This Google Form needs to be named exactly "Valuetales Google Forms Template".
  Note that the title of the Google Form and the name of the Google Form can be different values.
- In this Google Form add a new "File upload" question with the question title of "Add an image with your response".
- Enable "Allow only specific file types" then select the "Image" checkbox.
- "Maximum number of files" should be set to 1.
- "Maximum file size" should be set to 10MB.
- The new "File upload" question you just created should be the only question on this Google Form. If there are any other questions on the Google Form remove those questions.
- Next, open the settings for this Google Form.
- Set "Collect email addresses" to "Verified".
- Enable "Allow response editing".
- Enable "Limit to 1 response".
- Next, click on the "Publish" button.
- Under "Responders" click on "Manage" and change the "Responder view" from "Anyone with the link" to "Restricted".
- Click on "Done" then on "Publish". Your template should now be successfully setup.
- In the Apps Script project you can now run the function _createForms_.
  This will use the template you just created to create 50 Google Forms, one for each question.
- When running a function in this project for the first time it will first ask for authorization.
  Click "Review permissions" and sign into your Google account.
  This project will need access to Gmail, Google Drive, Google Docs, and Google Forms before it is able to run.
- Once all the Google Forms are created you will now have to manually open each one to restore the missing upload folder. Take special care not to miss any Google Forms during this process. This is a tedious process and an unfortunate limitation with the Apps Script platform.

### Weekly Trigger

- In your Apps Script project add a weekly trigger by selecting "Triggers" then selecting "Add Trigger".
- Under "Choose which function to run" set the function to _weeklyEmail_.
- "Choose which deployment should run" should be set to "Head"
- "Select event source" should be set to "Time-driven".
- "Select type of time based trigger" should be set to "Week timer".
- "Select day of week" can be set to any day you like. For example "Every Monday".
- "Select time of day" can be set to any time you like. For example "8am to 9am".
- Verify that the _startDate_ value in [valuetales.gs](./valuetales.gs) is inline with the first date you plan to run this script.
- "Failure notification settings" can be set to any value you like. For example "Notify me immediately".
- Hit "Save" and your script should now be setup to execute on a weekly basis.

## Usage

At the end of the year this project compiles the questions and responses into a Google Doc which will then be converted into a PDF.
This project won't handle buying and printing this book for you.
A third-party service such as [Barnes and Noble](https://press.barnesandnoble.com/print-on-demand) will need to be used to print out this PDF as a physical book.

## Credits

The email template in this project was adapted from this [template](https://github.com/leemunroe/responsive-html-email-template).

## Bugs & Improvements

- Enhance error handling logic.
- Add in some images to the readme to help with installation instructions.
