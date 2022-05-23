# MSOfficeResource
An extension to manage MSOffice files.

**This Extension is provided as-is and without warranty or support. It is not part of the PTC product suite and there is no PTC support.**

## Description
This extension adds a Resource object able to provide basic functionalities to manage MSOffice files.

## Services
- *readExcel*: reads an excel file with a simple tabular structure
  - input
    - fileRepository: the repository containing the excel file - THINGNAME (No default value)
    - path: the full path of the excel file - STRING (No default value)
    - sheetIndex: the index of the sheet - INTEGER (Default = 0)
    - hasHeader: true if the sheet has an header, false otherwise - BOOLEAN (Default: false)
    - dataShape: the output DataShape - DATASHAPENAME (No default value)
  - output: INFOTABLE (No DataShape)

## Dependencies
  - poi-3.17.jar
  - poi-ooxml-3.17.jar
  - poi-ooxml-schemas-3.17.jar
  - xmlbeans-2.6.0.jar

## Donate
If you would like to support the development of this and/or other extensions, consider making a [donation](https://www.paypal.com/donate/?business=HCDX9BAEYDF4C&no_recurring=0&currency_code=EUR).
