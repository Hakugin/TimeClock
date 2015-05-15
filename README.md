# TimeClock
Program used for sending time punches via Outlook email. I originally wrote this in my free time for contractors to use at my current employment. It does the job, but occassionally encounters errors here and there which is most likely due to the ugliness of the coding, but hey, I'm still learning the programming language. That's it for the brief summary.

Dependencies:
  - Python
  - wxPython
  - win32com
  - SQLalchemy
  - py2exe (For comiling into a Windows executable)

The program is pretty straight foreward, however within _dialogs.py there are 2 lines that should be updated before use. Second, you will need to supply your own myicons.ico file before attempting to compile.

To use:
  - Launch program
  - Configure, if needed
  - Click the appropriate button
  - Close program

Current "Features":
  - Sends an email by using Outlook's COM API
     - Destination is required, however a "Carbon Copy" address can be added as well
  - Destination email address can be viewed / changed
  - "Carbon Copy" email can be viewed / changed
  - Receive notification from Outlook when email is sent successfully
  - Buttons disable upon use, but can be re-enabled via the menu
  - --debug command ling option for more verbose log entries
  - Basic "Bug Report" capability
  - Settings, logs, and Time Punch entries saved to seperate tables within a single db file

Documentation is not my strong suit so what you see is what you get at the moment.
