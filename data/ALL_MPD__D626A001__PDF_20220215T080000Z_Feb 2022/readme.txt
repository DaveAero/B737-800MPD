 
 Table of Contents:

 	1. PDF Technical Publications Overview
	2. Contents of the Data Package
	3. Opening the Data Package
	4. Document Directory Structure
	5. HTML Supplemental Link Structure
	6. Supplement Management

 ===========================================================================
                   PDF Technical Publication Overview
 ===========================================================================

 The Portable Document Format (PDF) Technical Publication document package
 contains a HTML Table of Contents and PDF files.

 The document package is designed to be run from a local device (ie. hard
 disk drive, portable media, etc) or over a network.

 A PDF reader is required to render the PDF files. Some web browsers, like
 Chrome and Firefox, come with default PDF readers embedded within them.
 Should an issue occur when viewing PDF files, it is recommended that you
 download and install the most current version of Adobe Reader
 (https://acrobat.adobe.com/us/en/acrobat/pdf-reader.html), if compatible
 with your device.


 ===========================================================================
                  Contents of the Data Package
 ===========================================================================

 The document package contains the following components:
  - HTML Table of Contents (index.html)
  - PDF content files (*.pdf)

 The table of contents file contains links to the PDF content files of the
 document and to a default supplement folder location. The supplement folder
 location is where supplemental material may be stored. Examples of
 supplements include Temporary Revisions, Airline Revisions, Revision
 Notices, and Process Document Departures.


 ===========================================================================
                  Opening the Data Package
 ===========================================================================

 Open the index.html file with a supported web browser.


 ===========================================================================
                  Document Directory Structure
 ===========================================================================

 The document directory structure is comprised of the following components,
 in underscore-delimited fields:

 The first field is a customer code designation. In the case where a
 document or supplement is not tailored for any particular customer, a value
 of "ALL" is used.

 The second field is a document type.

 The third field is a volume type. This is used when a document type
 contains multiple volumes. In most cases, this will be empty.

 The fourth field is the document identifier.

 The fifth field is the supplement identifier, if the product is a
 supplement.

 The sixth field is the PDF product label.

 The seventh field is the issue date of a document or a supplement.


 ===========================================================================
                  HTML Supplemental Link Structure
 ===========================================================================

 The "Supplements" link inside the HTML file is always-provided in a static
 form per product, but in a default structured manner:

 "file:///C:" tells a web browser to read a local C: hard disk drive.
 "PDF" is the product label to help distinguish from other directories.
 <OEM> is the Original Equipment Manufacturer that supports the product.
 <volume type>, when appropriate; otherwise, the <document type>.
 The last field is the <document identifier>. 

 Example: file:///C:/PDF/TBC/SDS/D633W101-ZAP

 If you choose, you may edit the HTML file to change the location to meet
 your business need.


 ===========================================================================
                  Supplement Management
 ===========================================================================

 When you receive supplement material, you may extract the content files
 from the packages and put them in the location that the HTML Supplemental
 link references.


[SAS:20170509T215348Z]