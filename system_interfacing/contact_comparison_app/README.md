This app provides users with a GUI-based program that compares a study's contacts listed in a web system versus those listed in a 
shared drive document.

The app removes the need for users to download the web system contact list and manually compare both documents to determine if an update
is required. As long as a user keeps their login credentials in the `username` and `password` fields, they can compare both contact lists
for a study as many times as desired. Comparison is powered via `difflib` function `SequenceMatcher`.
