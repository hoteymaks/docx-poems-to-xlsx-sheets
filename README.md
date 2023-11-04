# docx-poems-to-xlsx-sheets
Create Excel sheets out of Word provided poems


## Description
You may experience adding poems data to your app: title, author and the poem itself.<br>
In case poems are provided in Word file, you can easily automate processing the mentioned data using my utility.


## How to use
1. Clone the project to your IDE (e.g. for IntellJ IDEA: "File" - "New" - "Project from Version Control" - "Repository URL" - paste `https://github.com/hoteymaks/docx-poems-to-xlsx-sheets.git` in URL field)
2. Place your .docx file of poems in the project directory
3. Open `CreateExcel.java` from `src` folder
4. Change file name on line 17 from `sample-document.docx` to your own
5. Run the code
6. Wait until Terminal will print `Excel file 'poems.xlsx' created successfully.`
7. Locate project directory, there you will find your processed Excel file

After that, you might like to use <a href="https://github.com/hoteymaks/xlsx-poems-to-java-array">my utility to turn the output Excel file to an array</a>.


## How it works
- Utility takes first paragraph in bold as an author name
- If provided, second paragraph in bold as a poem title
- Paragraphs of text until the next paragraph in bold is found as the poem itself

Collected data is filled in rows with the mentioned content respectively. Output will be an Excel file `poems.xlsx`.


## Sample document
Repository contains a sample Word file `sample-document.docx`.
