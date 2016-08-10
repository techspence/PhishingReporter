# PhishingReporter

Report phishing emails and have the notification sent to your security team. A button is created in Outlook using the Microsoft Junk Reporting Add-in and Powershell for deployment across an enterprise. This is not my idea, I have taken inspiration from http://www.nerdosaur.com/network-security/add-a-report-phishing-button-in-outlook and hacked it to fit what I need.

This does not work on 64 bit installs of office, I have confirmed this on Windows 10 Professional 64 bit running Office 2013 64 bit.

I am a novice scripter so take this code with a grain of salt, there very well may be better ways to this, this is my way. If anyone knows of a better way to do this, that's free and simple to setup, please let me know.

Basic Outline:

1) Download & install the Junk Reporting Add-in (You need to install the Junk Reporting Add-in on any machine you want this Outlook button on. You can download the Junk Reporting Add-in for Office 2007, 2010, 2013, and 2016 (32-bit) here: https://www.microsoft.com/en-us/download/details.aspx?id=18275)

2) Edit the HKLM:SOFTWARE\wow6432node\Microsoft\Junk E-mail Reporting\Addins BccEmailAddress Registry key to include any email address you would like

3) On a test machine, add the 'Report Phishing' button to any ribbons you would like

4) Copy olkmailread.officeUI and olkexplorer.officeUI from C:\users\%username%\AppData\Local\Microsoft\Office on the test machine to every machine you want the Outlook button on

4) Restart Outlook

5) By default when you 'Report' the Phishing Email it will send an email to abuse@messaging.microsoft.com (If you want to avoid this I suggest you either find where in the Add-in code the email is configured and change it, or do what I did, create a rule in your mail filtering application to block those emails on the way out)
