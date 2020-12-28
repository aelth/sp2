# SP^2 ![Logo](https://i.imgur.com/ASCtfgk.png)

Outlook plugin/Add-In for forwarding suspicious emails (most likely phishing, scam emails and sometimes spam) to arbitrary email address.  

The goal of the Add-In is to make reporting process easier both for the users and the incident response/security/helpdesk teams.  

# Why SP^2?

Although it is relatively easy for advanced users to forward suspicious emails, many people still lack the knowledge of *correct* message forwarding inside the Outlook mail client.  
**SP^2** makes it very easy to forward/report one or more suspicious emails received and stored in *Inbox* or any other Outlook folder.  

Forwarded emails are automatically parsed and, along with the original email attached, contain relevant details for quick triage by the security team.

# Features

- Forward one or more suspicious emails
- Configurable parameters (through *config* file)
- Parse message metadata (subject, sender, receiver(s))
- Automatically extracts links (and neutralizes them)
- Detect attachments
- Parse and extract relevant headers
- Parse (best effort) mail path (ordered list of traversed servers)

# Building

In order to compile **SP^2**, you need:
- Microsoft Visual Studio (current project uses Visual Studio 2017)
- Microsoft Office Developer Tools (for installation, start the Visual Studio setup program, and choose the *Modify* button. Select the *Microsoft Office Developer Tools* and choose *Update*)
- .NET Framework 4.7.2

Open the solution and build *VSTO* add-in.

You can generate custom *Code Signing Certificate* using Linux (*openssl*) or Windows (*signtool*).  
Linux commands are given below:
```
openssl req -new -x509 -newkey rsa:2048 -keyout sp2.key -out sp2.cer -days 3650
openssl pkcs12 -export -out sp2.pfx -inkey sp2.key -in sp2.cer
```
Password for the present PFX file is **sp2sp2**.
Feel free to modify above arguments as desired.

# Installation

**SP^2** has been tested on Office 2016 and 2019, but should generally work on older versions of Outlook (2010 and 2013).  

Currently, no installer is available (todo).  
After building the code, transfer all generated files (or use release.zip) to desired workstation.  
All configuration variables are present in `app.config` file. Feel free to modify the defaults.  

After configuration, double click *VSTO* file.
**SP^2** should now be integrated with Outlook.

![Ribbon](https://i.imgur.com/epjT2Fc.png)

# Usage

Select one or more suspicious emails and click on the *Report spam/phishing* button.

![Suspicious email](https://i.imgur.com/3XIZaCa.png)

![Click](https://i.imgur.com/dfvl4Kt.png)

This is the mail that will be sent to configured email address (usually security/IR team) - notice the attached original message:

![Forwarded email](https://i.imgur.com/Oxhs9cz.png)

# Uninstall

If you want to uninstall **SP^2**, you can use either Outlook Add-In dialog or the following command (modify if needed):
```
C:\Program Files\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe" /uninstall <path_to_vsto>
```

# Similar projects

**SP^2** was heavily influenced by [Outlook Spam Add-In](https://github.com/milcert/Outlook-Spam-Add-In) by milCERT.ch and [Phishing-Reporter](https://github.com/0dteam/Phishing-Reporter) by 0dteam.

# Credits

Icons made by [Freepik](http://www.freepik.com/") from [Flaticon](https://www.flaticon.com/)
