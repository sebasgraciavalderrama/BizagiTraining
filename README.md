# Replace fonts of Word documents - Bizagi Training

This Visual Basic code swaps the body text, Heading One and Heading Two text fonts to the latest style and font changes provided by Bizagi.
This current version of the code on the test phase. It is **not** a final release.

Also, if you would like to know which fonts and font sizes are being used in your Word document you can use the macro in [GetFonts.cls](https://github.com/sebasgraciavalderrama/BizagiTraining/blob/main/GetFonts.cls)



### Step-by-step
1. Go to the folder structure of your assigned workshop/course.
2. Identify client type documents (Enunciados, Agendas, tareas...).
    - For Agendas do not use the file `Bizagi Word Template (Branded Front Cover).docx` use instead `Bizagi Word Template (Simple - No Front Cover)`
3. Open up the document in Word.
5. Enable Macros to run in Word by enabling developer tab [How to enable Developer tab in Word](https://support.microsoft.com/en-us/office/show-the-developer-tab-in-word-e356706f-1891-4bb8-8d72-f57a51146792).
6. Copy and paste the code provided in the file `ReplaceFonts.cls`. [See this](https://postimg.cc/RW3HRKGP)
7. Make sure the fonts and font sizes are being properly identified in the variable declaration.
    - You can check them manually, however can use the code in [GetFonts.cls](https://github.com/sebasgraciavalderrama/BizagiTraining/blob/main/GetFonts.cls) to get the full list of font and font size.
8. You ought to make sure all the fonts and font sizes have been properly identified in the variable declaration, then you can [Run](https://postimg.cc/F1Rs3mTp) the code.
9. Check the result of the execution of the macro, make changes manually if needed.
10. Now, open the file `Bizagi Word Template (Branded Front Cover).docx`. 
11. Copy and paste into the template.
12. Eliminate blank spaces, unused tags, etc.
13. Update the `Table of contents` at the beginning of the document and make sure it makes sense.
10. Save and upload.
11. Repeate for all the other documents for both English and Spanish.


### Troubleshooting
As it is now, we have not encountered any issues with the code, however things might go sideways. If that is the case, please let us know, so we can fix the issue.

### Considerations
This code has a complexity of **O(n)** which means that the code changes letter by letter and then moves on to the next paragraph. This is not an issue for our documents but for big documents you might notice a little lag. You can even see the letters getting swapped.

Make sure to run this set of Macros with only one instance (maximum 2) of Word.

When you want to save the new document, make sure to **NOT** save it as a Macro file, save it like a regular `.docx` file.





