# Email_bulk
This has the project of a python file that  sends the bulk email
Go to https://myaccount.google.com → Security.

Turn on 2-Step Verification (choose Authenticator or Google Prompt). 
Google Help

Go to App passwords (use search on the page or https://myaccount.google.com/apppasswords). 
Google Accounts

Select Mail and Other / DesktopPython → Generate → copy the 16-character code. 
Saleshandy Docs

Put that code into your .env as APP_PASSWORD=thecode and run the Python script.
 
env file:
SENDER_EMAIL=youremail@gmail.com
APP_PASSWORD=your_16_char_app_password_here

