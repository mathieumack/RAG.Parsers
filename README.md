# RAG.Parsers

This library allow you to parse Word or Excel based documents towards a Markdown format

Library created in .NET which read the Word/Excel based documents with openXML/closedXML in order to write the equivalent file in markdown format.

==========

# Onboarding Instructions 

## RAG.Parsers.Docx

### Installation

1. Add nuget package: 

> Install-Package RAG.Parsers.Docx

2. In your application, you must instanciate a new DocxParser object, and call the method 'DocToMarkdown' with the path of your file to transform it to markdown string: 

```c#

var docxParser = new DocxParser();
var filePath = Path.Combine(Environment.CurrentDirectory, "TestFiles/TestSample.docx");

var result = docxParser.DocToMarkdown(filePath);

```

result value :

```string

My sample document
Creation Date:
Last Revised:
Version:1.0
## Index
### Sub Index 
#### Sub Sub Index
**Something bold**
*Something italic*
***Something*** ***bold in italic***
Something either **bold** OR *italic*
In **the** middle, [An hyperlink to ChatGPT](https://openai.com/chatgpt), but *nothing*
|First Cell header|||
|---|---|---|
||||Middle Cell 1|
|Middle Cell 2||||
||||Last Cell|

|-|-|-|-|
|---|---|---|---|
|Test1||||
|||||||
|||||||
||||||Test final|

```

## RAG.Parsers.Xlsx

### Installation

1. Add nuget package: 

> Install-Package RAG.Parsers.Xlsx

2. In your application, you must instanciate a new XlsxParser object, and call the method 'ExcelToMarkdown' with the path of your file to transform it to markdown string: 

```c#

var xlsxParser = new XlsxParser();
var filePath = Path.Combine(Environment.CurrentDirectory, "TestFiles/TestSample.xlsx");

var result = xlsxParser.ExcelToMarkdown(filePath);

```

result value :

```string

# Worksheet "First tab"

||A|B|C|D|E|F|G|H|I|J|
|---|---|---|---|---|---|---|---|---|---|---|
|**1**|This is a test|some cell filled||||||||||
|**3**||||||||||also here||
|**7**||||an other one here||||||||
|**12**|last one here|||||||||||

# Worksheet "An other tab"

||A|B|C|D|E|F|G|H|
|---|---|---|---|---|---|---|---|---|
|**1**|First cell in second tab|||||||||
|**3**||With a tab||||||||
|**4**||||Header first|Colonne2|Colonne3|Header last|||
|**16**||||||||toto||

```


# Support / Contribute

If you have any questions, problems or suggestions, create an issue or fork the project and create a Pull Request.

You want more ? Feel free to create an issue or contribute by adding new functionnalities by forking the project and create a pull request.

And if you like this project, don't forget to star it !

You can also support me with a coffee :

[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/mathieumack)