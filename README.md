# valuetales

## Description

This project was created as an open-source and cheaper alternative to the service [StoryWorth](https://welcome.storyworth.com/). [StoryWorth](https://welcome.storyworth.com/) is a service that sends a loved-one weekly email question prompts throughout the course of the year. At the end of the year [StoryWorth](https://welcome.storyworth.com/) bundles up the questions and responses into a book and sends you that book.

This project utilizes the products Gmail, Google Docs, and the [Google Apps Script platform](https://developers.google.com/apps-script) to orchestrate this alternative to the service [StoryWorth](https://welcome.storyworth.com/). This project handles sending loved-one weekly emails. This project also handles compiling the questions and responses into a Google Doc at the end of the year. You would then need to use a third-party service such as [Barnes and Noble](https://press.barnesandnoble.com/print-on-demand) to print out this Google Doc as a physical book. This project assumes some JavaScript coding knowledge and experience with regular expressions on part of the user.

## Installation

- Start by creating a new [Apps Script project](https://script.google.com/) named valuetales.
- Add the files [valuetales.gs](./valuetales.gs) and [template.html](./template.html) to the new project.
- The [valuetales.gs](./valuetales.gs) script will need you to fill out a few variables at the top of the script unique to you and your loved-ones. There is an array of questions filled out that you can either use or modify with questions of your own. Take special care to set the _startDate_ variable as the same date as the date that the script will first be ran with a weekly trigger.
- Next add a weekly trigger by selecting "triggers" then selecting "add trigger" in your Apps Script valuetales project.
- In the new trigger setup, set the trigger function to _askQuestions_ and set the time based trigger to _Week timer_. You can set the day of the week and the time of day to any time you like. For example Mondays at 8am.
- Next hit save and then you will be asked to approve this script's access to Gmail and Google Docs.
- You're script is now setup to execute on a weekly basis.

## Usage

At the end of the year valuetales will build a book as a Google Doc from the questions and responses throughout the year.

This portion of the script may need to be modified by you. The script needs to be able to parse the email responses of each recipient. Email responses often have unique signatures and also may include a chain of the previous email responses. These will need to be filtered out from the responses to the questions. These are also unique to each person and email provider so this will need to be modified by you. This filtering can be done by adding the appropriate regex for each case to the _regexList_ variable. It's suggested to run the _buildBookTest_ function a few months out from year-end to test out any modifications you may need to make to the script.

If modifying the script proves too challenging, the variable _googleDocLinkOverride_ is included for each recipient. This will let you override the PDF book sent in the final week's message with a PDF book generated from a Google Doc manually edited by you. Just add the Google Doc link to this variable.

The variable _coverPhotoWeekOverride_ is also included for each recipient. This allows you to select a specific week's photo instead of a random photo for the book's cover photo.

This script also won't handle buying and printing a book. That will need to be done through a third party service. [Barnes and Noble](https://press.barnesandnoble.com/print-on-demand) is a possible provider for this service.

## Credits

The email template in this project was adapted from this [template](https://github.com/leemunroe/responsive-html-email-template).

## Bugs & Improvements

- Emailing PDF attachment size limits may need to be handled later. A possible solution is to send a link to the PDF in Google Drive instead for large PDF attachments.
