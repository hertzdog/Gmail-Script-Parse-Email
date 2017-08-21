# Parse email messages for extracting information

## Scenario: parsing mails coming from Wordpress Plugin Contact Form 7

A mailbox has a lot of email messages sent though CF7 that contain data like the name, email and address of the senders. I would like to parse these email messages, extract the relevant bits and save them to a Google Spreadsheet/Excel.

Since the email is on google gmail, we will use [Google Apps Script](https://developers.google.com/gmail/api/quickstart/apps-script)


We added a cronjob (named "trigger" in Google terms) for running the script each day, between 00:00 and 01:00 AM
