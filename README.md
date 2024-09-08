Developers:
Gilad Abitbul
Ova-Maria Klassen
Mor Merkrebs

The chosen presentation topic:
Phishing

The subject of the project:
Spam Detector
The project was developed in the VSTO Outlook add in environment

The code listens to the inbox and when a new email is received:
1. The email content is extracted.
2. All the URL references that exist in it are extracted from the content.
3. Each URL is sent for testing on a site that verifies its legitimacy against 38 different sources.
4. A number is attached to each URL, which is actually the number of sources that identified it as a threat.
5. Before the email goes to the inbox - the code makes changes to it such as:
A. If at least a URL is located that a certain source has identified as a threat - the subject of the message will be prefixed with "(Threat Detected!):"
B. For each message, the scan conclusions will be appended to the end of the message content - how many URLs were checked and their check status.
third. After all this - the received email will be transferred to the inbox and will be viewable by the user.

Changes made according to comments:
1. There is no action the user must perform - the code runs automatically as soon as an email is received in the inbox.
2. No status is sent in a return email - the changes are made on the received email before it is shown to the user
3. Colors - emails in which a threat has been detected will be highlighted at the beginning in red: Threat Detected
