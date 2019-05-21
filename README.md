# Phone Message App

Small application that allows the user to input information about a phone call, resolve and email address to the Outlook address book, outputs everything to an email in Outlook, and then allows the user to send out the email. In short, it's a glorified phone message note-taking application.

## Getting Started

These instructions will show you how to download the python script, package it into an executable and run it on a Windows 10 machine.

### Prerequisites

This app was built to be run on a system with Windows 10.

## Deployment

Packaging into .exe:

Using pyinstaller, run the below command, modified with your directories for distpath and workpath.

Run:
```
 pyinstaller --noconsole --onedir --distpath 'path where the bundled app should go (do not keep quotes)' --workpath 'where to put temp work files (do not keep quotes)' --icon=images\gator.ico phoneapp.py
```
After it's packaged into an .exe, I just copy the phoneapp folder in the dist folder to the root of the C:\ drive and then create a shortcut to phoneapp.exe on the user's desktop that they can run.

## Instructions

* Launch the app by double-clicking the .exe
* The "To:" button will resolve a name or email address to the Outlook address book.
* Fill in desired fields and press "Send". An Outlook message window will appear to verify the informatin looks correct and then you can send the email.
* Click "Exit" to close the app.

![alt text](https://github.com/JacobG04/phone_message/blob/master/images/phone_message_screencap.PNG)

## Built With

* [Python](https://www.python.org/downloads/) - The programming language used to build the app

## Authors

* **Jacob Godwin** - *Initial work* - [JacobG04](https://github.com/JacobG04)

## License

* This app is free to use and is not for commercial use.

## Acknowledgments

* Thank you to whatever old 90s app that didn't work on Widnows 10 for giving me the push to learn Python and put this application together.
