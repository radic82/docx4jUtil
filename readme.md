# Docx4jUtil - Utility class create MS office word document based on [DOCX4J Library](https://www.docx4java.org/trac/docx4j)
[Here can you find Class Diagram](ClassDiagram.png)

This Lib (classes) helps you to create docx document. 
The factory classes contained on util package are the core of lib. This classes have most methods to compose docx. 
The classes are:

-  **StyleFactory** this class allows to create "word" object to insert into document 
   - extFormatting
   - textFormatting
   - overrideDefaultStyle
   - setCellBorders
   - setCellWidth
   - setCellNoWrap
   - setCellVMerge
   - setCellHMerge
   - setCellColor
   - setCellMargins
   - setVerticalAlignment
   - setFontSize
   - setFontFamily
   - setFontColor
   - setHorizontalAlignment
   - addBoldStyle
   - addItalicStyle
   - addUnderlineStyle
   - addStyle
   - addTableCell
   - addTableCell
   - addCellStyle
   - addImageCellStyle
   - newImage
   - getImageBytes
   - createTableWithContent
   - createHyperlink
   - customizeLayoutSection
   - createSectionBreak
- **SectionFactory** this class compose several object to create document:
	- fromDocTableToTable
	- fromFieldListToParagraphList
	- fromDocTextToParagraph
	- fromDocFieldToParagrah
	- fromDocTextToParagraph
	- fromDocLinkToParagraph
	- fromDocImgToParagraph
	- fromDocImgToParagraph
	- populateP
- **Docx4JXmlWrapper** this class is mainly used to manage [revision](http://radic.altervista.org/docx4j-java-read-review-from-docx/) / bookmark and retrive text form document
	- getBookmark
	- getRevisionIns
	- getRevisionDel
	- getRevision
	- createSimpleField
	- bookmarkRun
	- createRevision
	- handleRels
	- readRevision
	- acceptRevision
	- getText
	- setText
	- internalGetText
	- internalSetText
	- insertAnchorRevision
	- createAnchor
	- createAnchorNoParagraph
	- createTextFormatting

##Code Example
On test directory you can find some Junit Test that create some word document. The test classes are:
- CreateReport
  - createReportTextFormat - create document with many text formatting
  - createReportStyle - create document that use / overwrite default style of word document
  - createTable - create document with table
  - createReportDoubleColumn - create document with several break page to create double column template document
  - completeReport - create complete word document with several object 
- CreateSectionReport
  - createDefaultReport
	
	
	
