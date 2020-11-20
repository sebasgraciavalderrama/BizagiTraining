# Replace fonts of Word documents - Bizagi Training

This Visual Basic code swaps the body text, Heading One and Heading Two text fonts to the latest style and font changes provided by Bizagi.
This current version of the code on the test phase. It is **not** a final release.



### Step-by-step
1. Open up a document in Word
2. Enable Macros to run in Word by enabling developer tab (https://support.microsoft.com/en-us/office/show-the-developer-tab-in-word-e356706f-1891-4bb8-8d72-f57a51146792).
3. Copy and paste the code provided in the file  `ReplaceFonts.cls`.
4. Make sure the fonts are being properly identified in the variable declaration.
5. After you have make sure that all the fonts and font sizes have been properly identified **Run** the code.
6. Check the document and verify that changes have been successfully applied to the document.
7. Save and upload.

###Troubleshooting
As it is now, we have not encountered any issues with the code, however things might go sideways. If that is the case, please let us know, so we can fix the issue.

###Considerations
This code has a complexity of O(n) which means that the code changes letter by letter and then moves on to the next paragraph. This is not an issue for our documents but for big documents you might notice a little lag. You can even see the letters getting swapped.





