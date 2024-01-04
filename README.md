# Fsl

Fsl was previously thought as plugin for FileMaker based on ScriptMaster by 360works. Sadly to say that I was not able to create a custom plugin. So I was searching for an alternative to ScriptMaster. Finally I made acquaintance with the Sparkjava framework. This framework allows to create REST services in an easy way. It is in no way just a FileMaker plugin. But it can be used in a similar, though not equal way.

## Common

FileMaker allows REST requests with it's command **Insert From URL**. This opportunity uses **Fsl** for several functions. The **Fsl** Server runs as a service on Windows, MacOS and Linux. So it is possible to run **Fsl** local or on a server, where several users need access. Starting with version 1.0.0 there are functions available to create and fill in Excel files, generate swiss qrbills, and convert some file formats. See details below.

**Fsl** makes use of several open source libraries to fulfill its tasks:

spark-core: 

## Installation

Simply download the zip file to a convenient directory and extract it. You will find a new directory named **Fsl**. This directory contains a number of files:

- ch.eugster.filemaker.fsl.plist
- fsl_spark_linux.sh
- fsl_spark_mac.sh
- fsl_spark_win.bat
- fsl-spark-X.X.X-jar-with-dependencies.jar
- fsl-spark-X.X.X.jar
- nssm-2.24.zip

Depending on the os you will install **Fsl** on, you need a subset of these files. Let me show you tested examples for Linux, MacOS and Windows. Suppose we are logged on as root, or, on Windows as Administrator and residing within the new directory **Fsl**.

### Linux

1. Create a directory /opt/fsl_spark/
2. Copy the file fsl-spark-X.X.X-jar-with-dependencies.jar to this directory
3. Copy the file fsl_spark_linux.sh
4. Check the path to java in this file and replace it with the real path if it does not point to an existing java executable
5. Check if the installation at this point works by executing the following command in a terminal
```
$ /opt/fsl_spark/fsl_spark_linux.sh
```
6. If you see as result something like --INFO: Started @585ms--, then the installation at this point is ok
7. Stop the program
8. Copy the file fsl-server-linux.service to the directory /etc/systemd/system
9. Enable and run the service
```
$ systemctl enable fsl-server-linux.service
$ systemctl start fsl-server-linux.service  // or reboot
```
10. Check if the service is running by executing
```
$ systemctl status fsl-server-linux.service
```
and/or
```
$ curl http://localhost:4567/fsl/Fsl.version --data {}
```
11. If it returns a version number in the form of X.X.X then the service is ok

### MacOS

1. Create a directory /opt/fsl_spark/
2. Copy the file fsl-spark-X.X.X-jar-with-dependencies.jar to this directory
3. Copy the file fsl_spark_mac.sh
4. Check the path to java in this file and replace it with the real path if it does not point to an existing java executable
5. Check if the installation at this point works by executing the following command in a terminal
```
$ /opt/fsl_spark/fsl_spark_mac.sh
```
6. If you see as result something like --INFO: Started @585ms--, then the installation at this point is ok
7. Stop the program
8. Copy the file ch.eugster.filemaker.fsl.plist to the directory /Library/LaunchAgents
9. Enable and run the service
```
$ launchctl load /Library/LaunchAgents/ch.eugster.filemaker.fsl.plist
```
10. Check if the service is running by executing
```
$ curl http://localhost:4567/fsl/Fsl.version --data {}
```
11. If it returns a version number in the form of X.X.X then the service is ok

### Windows

1. Create a directory C:\Program Files\fsl_spark\
2. Copy the file fsl-spark-X.X.X-jar-with-dependencies.jar to this directory
3. Copy the file fsl_spark_win.bat
4. Check the path to java in this file and replace it with the real path if it does not point to an existing java executable
5. Check if the installation at this point works by executing the following command in a terminal
```
$ ./fsl_spark_win.bat
```
6. If you see as result something like --INFO: Started @585ms--, then the installation at this point is ok
7. Stop the program
8. Unzip the file nssp-2.24.zip and execute the expanded file nssm.exe
```
$ nssm.exe install c:\Program Files\fsl_spark\fsl_spark_win.bat
```
9. Start the service
10. Check if the service is running by executing
```
$ curl http://localhost:4567/fsl/Fsl.version --data {}
```
11. If it returns a version number in the form of X.X.X then the service is ok

Then everything is ok and you can access the server from FileMaker. How do you do this?

## Example

Example: You want to create a swiss qrbill in FileMaker:

- Create a json object using JSONSetElement:

```
set variable [ $request ; value: 
JSONSetElement ( $request ; 
    [ "amount" ; 287.3 ; JSONNumber ] ;
    [ "currency" ; "CHF" ; JSONString ] ;
    [ "iban" ; "CH4431999123000889012" ; JSONString ] ;
    [ "reference" ; "000000000000000000000000000" ; JSONString] ;
    [ "message" ; "Rechnungsnr. 10978 / Auftragsnr. 3987" ; JSONString ] ;
    [ "creditor.name" ; "Schreinerei Habegger & Söhne" ; JSONString ] ;
    [ "creditor.address_line_1" ; "Uetlibergstrasse 138" ; JSONString ] ;
    [ "creditor.address_line_2" ; "8045 Zürich" ; JSONString ] ;
    [ "creditor.country" ; "CH" ; JSONString ] ;
    [ "debtor.name" ; "Simon Glarner" ; JSONString ] ;
    [ "debtor.address_line_1" ; "Bächliwis 55" ; JSONString ] ;
    [ "debtor.address_line_2" ; "8184 Bachenbülach" ; JSONString ] ;
    [ "debtor.country" ; "CH" ; JSONString ] ;
    [ "format.graphics_format" ; "PDF" ; JSONString ] ;
    [ "format.output_size" ; "QR_BILL_EXTRA_SPACE" ; JSONString ] ;
    [ "format.language" ; "DE" ; JSONString ]
)
```

- After setting the variable (e.g. $request) run the command:

```
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/QRBill.generate" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```
In the example above while **http://localhost:4567/fsl/** is fix, the part **QRBill.generate** is interchangeable with other commands, that are listed and explained below. It consists of two parts:

The **QRBill** part in the example is the so named the module part, while **generate** is the command to execute. There exist the following modules: **Xls**, **QRBill**, **Pdf** and **Convert**. The parameters are in json format and passed in the cURL-Options section (see above). Depending on the command there are no, optional or mandatory parameters, explained in the respective command section.

- Read the status of the response. $response is a json object, that returns the element **status** ("OK" or "Fehler"), if status is "Fehler" then the array "errors" contains an array of errors. Some commands return other elements as described in the command respective below.

```
Variable setzen [ $status ; Wert: JSONGetElement ( $response ; "status" ) ] 
```

If **$status** is "OK" then the command ended successfully. You may fill a container field/variable as:

```
 Set Variable [ $QRBill ; Base64Decode ( JSONGetElement ( $response ; "result" ) ; "QRBill.pdf" ) ]
```

In the case above the result has to be decoded to get the pdf content

If **$status** is "Fehler" then you can get the element "errors", which is a list of error messages.

```
If [ $status = "Fehler" ]
    Set Variable [ $errors ; Value: JSONGetElement ( $response ; "errors" ) ]
    Set Variable [ $count ; Value: ElementsCount ( $errors ) ]
    Set Variable [ $index ; Value: 1 ]
    Loop
        Set Variable [ $error ; ElementsMiddle ( $errors ; $index ; 1 ) ]
        Set Variable [ $index ; $index + 1 ) ]
        Exit Loop If [ $index > $count ]         
    End Loop
End If  
```

## Xls 

### Commands

#### activateSheet

Activates the sheet with given name if present, else returns error

##### Example

```
JSONSetElement ( $request ; "sheet" ; "mySheet" ; JSONString )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.activateSheet" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```
or, with optional parameter "workbook"

```
JSONSetElement ( $request ; 
    [ "sheet" ; "mySheet" ; JSONString ] ;
    [ "workbook" ; "myWorkbook" ; JSONString ] 
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.activateSheet" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters 

| name | necessity | value(s) | description
|---|---|---|---
| [`workbook`](#workbook) | optional | name or path of workbook, if not provided the active workbook is used. if no active workbook is present an error is returned
| [`sheet`](#sheet) | mandatory | string | the name of the sheet

##### Response parameters

| name | description
|---|---
| [`index`](#index) | sheet index in the workbook
| [`sheet`](#sheet) | name of the sheet
| [`workbook`](#workbook) | name or path of workbook

#### activateWorkbook

Activates the workbook with given name, if present, else returns error

##### Example

```
JSONSetElement ( $request ; "workbook" ; "myWorkbook" ; JSONString )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.activateWorkbook" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```
##### Request parameters 

| name | necessity | value(s) | description
|---|---|---|---
| [`workbook`](#workbook) | mandatory | string | given name of the workbook. Workbook must exist.

##### Response parameters

| key | description
|---|---
| [`workbook`](#workbook) | name or path of workbook

#### activeSheetPresent

Checks if an active sheet is present. If an active sheet is present, it returns index and name of sheet, and the workbook, it belongs to, else error

##### Request parameters

none 

##### Response parameters

| name | value(s)
|---|---
| index | the index of the sheet in workbook
| sheet | the name of the active sheet
| [`workbook`](#workbook) | the name of the workbook

#### activeWorkbookPresent

Checks if an active workbook is present.  If an active workbook is present, it returns the name of the workbook, if not error

##### Request parameters

none

##### Response parameters

| name | value(s)
|---|---
| [`workbook`](#workbook) | the (path) name of the workbook


#### copy

Copies a cell or a range of cells to the given target cell or range. If source range differs target range, returns an error, if same sheet for source and target ranges is used, if ranges intersect, an error is returned. The content of the fields stays unchanged, except the formula fields, where the formula are adapted to the new environment.

##### Examples

Example 1: in the following examples several forms to define **source** and **target** parameters are used: source has a range address as used in Excel too, target.top_left defines cell indices to define the top row and right column of the top_left cell of the range. target.right takes an integer (3) that defines the target bottom_right cells column index and target.bottom takes an integer too to define the bottom_right cell's row index. Take notice, that the numbers to define the cell's row and column indices are zero based, i.e. bottom 3 and right 3 means "D4", not "C3" as one could think.

```
JSONSetElement ( $request ; 
    [ "source.range" ; "A1:B2" ; JSONString ] ;
    [ "target.top_left" ; "C3" ; JSONString ] ;
    [ "target.right" ; 3 ; JSONNumber ] ;
    [ "target.bottom : 3 ; JSONNumber ]
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.copy" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

Example 2:  This examples defines the same ranges as above. I contains a source sheet and a target sheet too, that means, the source range on sheet_1 will be copied to the range in sheet_2. 

```
JSONSetElement ( $request ;
    [ "source.top" ; 0 ; JSONNumber ] ;
    [ "source.left" ; 0 ; JSONNumber ] ;
    [ "source.bottom_right : "D4" ; JSONNumber ;
    [ "source.sheet" ; "sheet_1" ; JSONString ] ;
    [ "target.range" ; "C3:D4" ; JSONString ] ;
    [ "target.sheet" ; "sheet_2" ; JSONString ]
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.copy" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

Example 3: If the copy is done withing one sheet, you can omit the keyword sheet. If want to use an other sheet then the active one, you provide the sheet at root of the json object. 

```
JSONSetElement ( $request ;
    [ "sheet" ; "mySheet" ; JSONString | ;
    [ "source.range" ; "A1:B2 ; JSONNumber ] ;
    [ "target.range" ; "C3:D4" ; JSONString ] 
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.copy" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

Example 4: If you want to copy the content of one cell to a range of cells:

```
JSONSetElement ( $request ;
    [ "source.cell" ; "A1" ; JSONString | ;
    [ "target.range" ; "C3:D4" ; JSONString ] 
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.copy" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

Example 5: If you want to copy one cell to one cell 

```
JSONSetElement ( $request ;
    [ "source.cell" ; "A1" ; JSONString | ;
    [ "target.cell" ; "D4" ; JSONString ] 
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.copy" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters
  
| name | necessity | description
|---|---|---
| workbook | optional | given name of workbook to use 
| sheet | optional | if not the active sheet should be used 
| source.sheet | optional | if source sheet differs from target sheet
| source.range | source.range.top_left and source.range.bottom_right | source.range.top and source.range.left and source.range.bottom and source.range.right
| source.cell | source.rcell.row and source.cell.col | 
| target.sheet | optional | if source sheet differs from target sheet
| target.range | target.range.top_left and target.range.bottom_right | target.range.top and target.range.left and target.range.bottom and target.range.right
| target.cell | target.cell.row and target.cell.col |

##### Response parameters

none

#### createAndActivateSheet

creates and activates a sheet

##### Example

```
JSONSetElement ( $request ;
    [ "sheet" ; "mySheet" ; JSONString | ;
    [ "workbook" ; "myWorkbook" ; JSONString ] 
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.createAndActivateSheet" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| keyword | necessity | value
|---|---|---
| sheet | mandatory | name of sheet as string
| workbook | optional | if not the active workbook is used, the name of an already created workbook

##### Response parameters

none

#### createAndActivateWorkbook

creates and activates a workbook

##### Example

```
JSONSetElement ( $request ;
    [ "workbook" ; "myWorkbook" ; JSONString ] 
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.createAndActivateWorkbook" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| keyword | necessity | value
|---|---
| workbook | mandatory | the (path) name of the worbook to create

##### Response parameters

none

#### createSheet

creates a sheet but does not activate it

##### Example

```
JSONSetElement ( $request ;
    [ "sheet" ; "mySheet" ; JSONString | ;
    [ "workbook" ; "myWorkbook" ; JSONString ] 
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.createSheet" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| keyword | necessity | value
|---|---|---
| sheet | mandatory | name of sheet as string
| workbook | optional | if not the active workbook is used, the name of an already created workbook

##### Response parameters

none

#### createWorkbook

creates a workbook but does not activate it

##### Example

```
JSONSetElement ( $request ;
    [ "workbook" ; "myWorkbook" ; JSONString ] 
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.createWorkbook" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| keyword | necessity | value
|---|---
| workbook | mandatory | the (path) name of the worbook to create

##### Response parameters

none

#### getActiveSheet

returns the name of the active sheet, if any, else error

##### Example
```
JSONSetElement ( "{}" )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.getActiveSheet" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

none

##### Response parameters

| name | value(s)
|---|---
| index | the index of the sheet in workbook
| sheet | the name of the active sheet
| [`workbook`](#workbook) | the name of the workbook

#### getActiveWorkbook

returns the name of the active workbook, if any, else error

##### Example

```
JSONSetElement ( "{}" )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.getActiveWorkbook" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

none

##### Response paramters

| name | value(s)
|---|---
| [`workbook`](#workbook) | the name of the workbook


#### getSheets

returns a list of sheet names and a list of sheet indices

##### Examples

Example 1: without parameters the active workbook is selected, else error

```
JSONSetElement ( "{}" )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.getSheets" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

Example 2: with parameter "workbook", the workbook of which the sheets should be return 

```
JSONSetElement ( $request ; "workbook" ; "myWorkbook" ; JSONString )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.getSheets" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| keyword | necessity | value
|---|---|---
| workbook | optional | the (path) name of the worbook to create

##### Response parameters

| keyword | value
|---|---
| sheet | an possibly zero-length array of sheet names
| index | an possibly zero-length array of sheet indices


#### getWorkbookNames

returns a list of the existing workbook names

##### Example

Without parameters the active workbook is selected, else error

```
JSONSetElement ( "{}" )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.getWorkbookNames" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

none

##### Response parameters

| keyword | value
|---|---
| workbook | an possibly zero-length array of workbook names


#### releaseWorkbook

releases a workbook if it exists, else returns an error

##### Example

```
JSONSetElement ( $request ; "workbook" ; "myWorkbook" ; JSONString )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.releaseWorkbook" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| keyword | value
|---|---
| workbook | the name of the workbook to release

##### Response parameters

none 

#### releaseWorkbooks

releases all existing workbooks

##### Example

```
JSONSetElement ( "{}" )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.releaseWorkbooks" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

none

##### Response parameters

none

#### saveAndReleaseWorkbook

saves and releases a workbook if it exists, else returns an error

##### Examples

```
JSONSetElement ( "{ "workbook" ; "myWorkbook" }" )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.saveAndReleaseWorkbook" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameter

| keyword | value
|---|---
| workbook | the name of the workbook to release

##### Response parameters

none

#### saveWorkbook

saves a workbook if it exists, else returns an error. The workbook will not be release, that means, one could edit it further

##### Example

```
JSONSetElement ( $request ; "workbook" ; "myWorkbook" ; JSONString )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.saveWorkbook" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| keyword | value
|---|---
| workbook | the name of the workbook to release

##### Response parameters

none

#### setCells

set cell values or formulae. There are two possibilies how to fill cells.
1. Define one cells address and as much values you want to write. The values are distributed, beginning with the provided cell in one direction. If direction is not defined, the values are written to the right of the given cell, the first value into the given cell, the second one into the cell to the right of the beginning cell and so on. You can provide the direction by defining a parameter named "direction", that can have one of the following values: "right" (default), "up", "left", and "down".

##### Example

```
JSONSetElement ( $request ; 
    [ "cell" ; "A1" ; JSONString ] ;
    [ "values[0] ; 123.45 ; JSONNumber ] ;
    [ "values[1] ; 76.54 ; JSONNumber ]
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.setCells" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| keyword | value
|---|---
| cell | the cell address in the form "A1" or cell.row as integer and cell.col as integer
| values[index] | (index is an integer) contains the values to put into the cells 

##### Response parameters

none

2. Define two arrays of same length. One array ("cell") holds the cell address, the other ("values") the values or formulae for the cell in the same place of the array.

##### Example

```
JSONSetElement ( $request ; 
    [ "cell[0]" ; "A1" ; JSONString ] ;
    [ "cell[1]" ; "B2" ; JSONString ] ;
    [ "values[0] ; 123.45 ; JSONNumber ] ;
    [ "values[1] ; 76.54 ; JSONNumber ]
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.setCells" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| keyword | value
|---|---
| cell[index] | the cell addresses in the form "A1" or cell.row as integer and cell.col as integer
| values[index] | (index is an integer) contains the values to put into the cells 

##### Response parameters

none

#### setPrintSetup

define print setup options

##### Example

```
JSONSetElement ( $request ; 
    [ "orientation" ; "PORTRAIT" ; JSONString ] ;
    [ "copies" ; 2 ; JSONNUMBER ] 
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.setPrintSetup" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| name | necessity | values
|---|---|---
| workbook | optional | string
| sheet | optional | string
| orientation | mandatory | "PORTRAIT" or "LANDSCAPE"
| copies | optional | integer

##### Response parameters

none

#### setHeaders

Set the headers of a sheet. There are three headers, one on the left, one in the middle, and one on the right.

##### Example
```
JSONSetElement ( $request ; 
    [ "left" ; "This is the left header" ; JSONString ] ;
    [ "center" ; "This is the center header" ; JSONString ] 
    [ "right" ; "This is the right header" ; JSONString ] 
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.setHeaders" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| name | necessity | values
|---|---|---
| workbook | optional | string
| sheet | optional | string
| left | optional | string
| center | optional | string
| right | optional | string

##### Response parameters

none

#### setFooters

Set the footers of a sheet. There are three footers, one on the left, one in the middle, and one on the right.

##### Example

```
JSONSetElement ( $request ; 
    [ "left" ; "This is the left footer" ; JSONString ] ;
    [ "center" ; "This is the center footer" ; JSONString ] 
    [ "right" ; "This is the right footer" ; JSONString ] 
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.setFooters" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| name | necessity | values
|---|---|---
| workbook | optional | string
| sheet | optional | string
| left | optional | string
| center | optional | string
| right | optional | string

##### Response parameters

none

#### applyFontStyles

To style fonts

##### Example

```
JSONSetElement ( $request ; 
    [ "cell" ; "A1" ; JSONString ] ;
    [ "name" ; "Courier New" ; JSONString ] ;
    [ "size" ; 12 ; JSONNumber ] ;
    [ "bold" ; 1 ; JSONNumber ] ;
    [ "italic" ; 0 ; JSONNumber ] ;
    [ "underline" ; 1 ; JSONNumber ] ;
    [ "strike_out" ; 0 ; JSONNumber ] ;
    [ "type_offset" ; 0 ; JSONNumber ] ;
    [ "color" ; 0 ; JSONNumber ] ;
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.applyFontStyles" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| name | necessity | values
|---|---|---
| workbook | optional | string
| sheet | optional | string
| cell or range | mandatory | string
| name | optional | string
| size | optional | string
| bold | optional | string
| italic | optional | string
| underline | optional | string
| strike_out | optional | string
| type_offset | optional | string
| [`color`](#color) | optional | integer

##### Response parameters

none

#### applyCellStyles

To style cells

##### Example 

```
JSONSetElement ( $request ; 
    [ "alignment.horizontal" ; "Courier New" ; JSONString ] ;
    [ "alignment.vertical" ; "A1" ; JSONString ] ;
    [ "border.style.top" ; "THIN" ; JSONString ] ;
    [ "border.style.left" ; 0x1 ; JSONNumber ] ;
    [ "border.style.bottom" ; "THIN" ; JSONString ] ;
    [ "border.style.right" ; 0x1 ; JSONNumber ] ;
    [ "border.color.top" ; "RED1" ; JSONString ] ;
    [ "border.color.left" ; 2 ; JSONNumber ] ;
    [ "border.color.bottom" ; "RED1" ; JSONString ] ;
    [ "border.color.right" ; 2 ; JSONNumber ] ;
    [ "data_format" ; 0 ; JSONNumber ] ;
    [ "foreground.color" ; 0 ; JSONNumber ] ;
    [ "background.color" ; "WHITE" ; JSONString ] ;
    [ "fill_pattern" ; "NO_FILL" ; JSONString ] ;
    [ "shrink_to_fit" ; 0 ; JSONNumber ] ;
    [ "wrap_text" ; 1 ; JSONNumber ] ;
    [ "font" ; 0 ; JSONNumber ] ;
)
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.applyCellStyles" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| name | necessity | values
|---|---|---
| workbook | optional | string
| sheet | optional | string
| alignment.[`horizontal`](#horizontal) | optional | see [`list`](#horizontal) of possible values 
| alignment.[`vertical`](#vertical) | optional | see [`list`](#vertical) of possible values
| border.[`style`](#style).top | optional | style of the top cell border (see [`list`](#style) of possible values)
| border.[`style`](#style).left | optional | style of the left cell border (see [`list`](#style) of possible values)
| border.[`style`](#style).right | optional | style of the right cell border (see [`list`](#style) of possible values)
| border.[`style`](#style).bottom | optional | style of the bottom cell border (see [`list`](#style) of possible values)
| border.[`color`](#color).top | optional | color of the top cell border (see [`list`](#color) of possible values)
| border.[`color`](#color).left | optional | color of the left cell border (see [`list`](#color) of possible values)
| border.[`color`](#color).right | optional | color of the right cell border (see [`list`](#color) of possible values)
| border.[`color`](#color).bottom | optional | color of the bottom cell border (see [`list`](#color) of possible values)
| [`data_format`](#data_format) | optional | the formatting of the content of the cell (see list of available [`data_formats`](#data_format))
| foreground.[`color`](#color) | optional | the color of the foreground, i.e. of the font of the cell (see list of available [`colors`](#color))
| background.[`color`](#color) | optional | the color of the background of the cell (see list of available [`colors`](#color))
| [`fill_pattern`](#fill_pattern) | optional | the fill pattern of the background of the cell (see list of available [`fill_patterns`](#fill_pattern))
| [`shrink_to_fit`](#shrink_to_fit) | optional | shrink content to fit the text in a cell
| [`wrap_text`](#wrap_text) | optional | wrap the text in the cell, if it overflows the cell 

##### Response parameters

none

#### autoSizeColumns

Autosize columns in the given range.

##### Example

```
JSONSetElement ( $request ; "range" ; "A1:G1" ; JSONString )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.autoSizeColumns" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| name | necessity | values
|---|---|---
| workbook | optional | string
| sheet | optional | string
| range or cell | mandatory |

##### Response parameters

none

#### rotateCells

Rotate the content of the cells to given degrees.

##### Example

```
JSONSetElement ( $request ; "rotation" ; 90 ; JSONNumber )
Insert From URL [ selection ; with dialog: off ; Target: $response ; "http://localhost:4567/fsl/Xls.rotateCells" ; cURL-Options: " --data @$request --header \"Content-Type: application/json\"" ] 
```

##### Request parameters

| name | necessity | values
|---|---|---
| workbook | optional | string
| sheet | optional | string
| rotation | mandatory | integer

##### Response parameters

none

### Keywords

To assemble a json object as parameter, there exists a controlled set of keywords for each module. Some of them have a restricted list of values.

#### alignment

text alignment in cells, either horizontally or vertically, allowed values, see [`horizontal`](#horizontal) or [`vertical`](#vertical)

Used by [`applyCellStyles`](#applyCellStyles)

| Keys
|---
| alignment.[`horizontal`](#horizontal)
| alignment.[`vertical`](#vertical)

##### background

Used by [`applyCellStyles`](#applyCellStyles)

| Keys
|---
| background.[`color`](#color)

##### bold

sets the font to bold or not

Used by [`applyFontStyles`](#applyFontStyles)

| values | explanation
|---|---
| 0 | set bold off
| 1 | set bold on

##### border

Used by [`applyCellStyles`](#applyCellStyles)

| child keys | explanation
|---|---
| border.[`style`](#style) | set style
| border.[`color`](#color) | set color

##### bottom_right

Means the bottom right cell of a range

Used by [`range`](#range)
 
| values | examples
|---|---
| range.bottom_right | in the form like "A1"

or

| child keys | values | explanation
|---|---
| range.bottom_right.row | integer | the cell's row index in the sheet
| range.bottom_right.col | integer | the cell's column index in the sheet

###### style

Used by [`applyCellStyles`](#applyCellStyles)

| values
|---
| NONE(0x0),
| THIN(0x1),
| MEDIUM(0x2),
| DASHED(0x3),
| DOTTED(0x4),
| THICK(0x5),
| DOUBLE(0x6),
| HAIR(0x7),
| MEDIUM_DASHED(0x8),
| DASH_DOT(0x9),
| MEDIUM_DASH_DOT(0xA),
| DASH_DOT_DOT(0xB),
| MEDIUM_DASH_DOT_DOT(0xC),
| SLANTED_DASH_DOT(0xD);
 
##### cell

| values | examples | explanation
|---|---
| string | "A1" | cell address

| child keys | explanation
|---|---
| [`row`](#row) | integer value
| [`col`](#col) | integer value

##### center

| values | explanation
|---|---
| string | a text string that is used as header in a Excel sheet

##### col

| values | explanation
|---|---
| integer | an integer value specifying the column index (zero-based) of a cell

##### color

Colors can be provided as strings or as integers, the table below shows the possible color names (strings) and their respective number (integer)

| string | integer
|---|---
| BLACK1 | 0
| WHITE1 | 1
| RED1 | 2
| BRIGHT_GREEN1 | 3
| BLUE1 | 4
| YELLOW1 | 5
| PINK1 | 6
| TURQUOISE1 | 7 
| BLACK | 8
| WHITE | 9
| RED | 10
| BRIGHT_GREEN | 11
| BLUE | 12
| YELLOW | 13
| PINK | 14
| TURQUOISE | 15
| DARK_RED | 16
| GREEN | 17
| DARK_BLUE | 18
|  DARK_YELLOW | 19
| VIOLET | 20
| TEAL | 21
| GREY_25_PERCENT | 22
| GREY_50_PERCENT | 23
| CORNFLOWER_BLUE | 24
| MAROON | 25
| LEMON_CHIFFON | 26
| LIGHT_TURQUOISE1 | 27
| ORCHID | 28
| CORAL | 29
| ROYAL_BLUE | 30
| LIGHT_CORNFLOWER_BLUE | 31
| SKY_BLUE | 40
| LIGHT_TURQUOISE | 41
| LIGHT_GREEN | 42
| LIGHT_YELLOW | 43
| PALE_BLUE | 44
| ROSE | 45
| LAVENDER | 46
| TAN | 47
| LIGHT_BLUE | 48
| AQUA | 49
| LIME | 50
| GOLD | 51
| LIGHT_ORANGE | 52
| ORANGE | 53
| BLUE_GREY | 54
| GREY_40_PERCENT | 55
| DARK_TEAL | 56
| SEA_GREEN | 57
| DARK_GREEN | 58
| OLIVE_GREEN | 59
| BROWN | 60
| PLUM | 61
| INDIGO | 62
| GREY_80_PERCENT | 63
| AUTOMATIC | 64

##### copies

| values | explanation
|---|---
| integer | number of copies

##### data_format

There are a number of predefined data formats in Excel, but it is also possible to create own ones. The next table shows the predefined data formats and their numeric identifiers

| value | format as string
|---|---
| 0 | "General"
| 1 | "0"
| 2 | "0.00"
| 3 | "#,##0"
| 4 | "#,##0.00"
| 5 | "$#,##0_);($#,##0)"
| 6 | "$#,##0_);[Red]($#,##0)"
| 7 | "$#,##0.00);($#,##0.00)"
| 8 | "$#,##0.00_);[Red]($#,##0.00)"
| 9 | "0%"
| 0xa | "0.00%"
| 0xb | "0.00E+00"
| 0xc | "# ?/?"
| 0xd | "# ??/??"
| 0xe | "m/d/yy"
| 0xf | "d-mmm-yy"
| 0x10 | "d-mmm"
| 0x11 | "mmm-yy"
| 0x12 | "h:mm AM/PM"
| 0x13 | "h:mm:ss AM/PM"
| 0x14 | "h:mm"
| 0x15 | "h:mm:ss"
| 0x16 | "m/d/yy h:mm"
| // 0x17 - 0x24 reserved for international and undocumented
| 0x25 | "#,##0_);(#,##0)"
| 0x26 | "#,##0_);[Red](\#,##0)"
| 0x27 | "#,##0.00_);(#,##0.00)"
| 0x28 | "#,##0.00_);[Red](\#,##0.00)"
| 0x29 | "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"
| 0x2a | "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)"
| 0x2b | "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"
| 0x2c | "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"
| 0x2d | "mm:ss"
| 0x2e | "[h]:mm:ss"
| 0x2f | "mm:ss.0"
| 0x30 | "##0.0E+0"
| 0x31 | "@" - This is text format.
| 0x31  "text" - Alias for "@"

##### direction

| values | explanation
|---|---
| right | direction to the right
| top | direction to top
| left | direction to the left
| bottom | direction to bottom

##### fill_pattern

Like colors fill_patterns have names and a numeric value. You can choose which one you want to provide

| string | integer | explanation
|---|---|---
| NO_FILL | 0 | No background
| SOLID_FOREGROUND | 1 | Solidly filled
| FINE_DOTS | 2 | Small fine dots
| ALT_BARS | 3 | Wide dots
| SPARSE_DOTS | 4 | Sparse dots
| THICK_HORZ_BANDS | 5 | Thick horizontal bands
| THICK_VERT_BANDS | 6 | Thick vertical bands
| THICK_BACKWARD_DIAG | 7 | Thick backward facing diagonals
| THICK_FORWARD_DIAG | 8 | Thick forward facing diagonals
| BIG_SPOTS | 9 | Large spots
| BRICKS | 10 | Brick-like layout
| THIN_HORZ_BANDS | 11 | Thin horizontal bands
| THIN_VERT_BANDS | 12 | Thin vertical bands
| THIN_BACKWARD_DIAG | 13 | Thin backward diagonal
| THIN_FORWARD_DIAG | 14 | Thin forward diagonal
| SQUARES | 15 | Squares
| DIAMONDS | (16 | Diamonds
| LESS_DOTS | 17 | Less Dots
| LEAST_DOTS | 18 | Least Dots

##### font

| child keys | examples
|---|---
| name | the name of the font

##### foreground

| child keys | explanation
|---|---
| color | see [`colors`](#color)

##### format

| child keys | explanation
|---|---
| color | one of [`colors`](#color)

##### horizontal

used in [`alignment`](#alignment)

| string | integer | description
|---|---|---
| GENERAL | 0 | The horizontal alignment is general-aligned. Text data is left-aligned. Numbers, dates, and times are rightaligned. Boolean types are centered. Changing the alignment does not change the type of data.
| LEFT | 1 | The horizontal alignment is left-aligned, even in Rightto-Left mode. Aligns contents at the left edge of the cell. If an indent amount is specified, the contents of the cell is indented from the left by the specified number of character spaces. The character spaces are based on the default font and font size for the workbook.
| CENTER | 2 | The horizontal alignment is centered, meaning the text is centered across the cell.
| RIGHT | 3 | The horizontal alignment is right-aligned, meaning that cell contents are aligned at the right edge of the cell, even in Right-to-Left mode.
| FILL | 4 | Indicates that the value of the cell should be filled across the entire width of the cell. If blank cells to the right also have the fill alignment, they are also filled with the value, using a convention similar to centerContinuous.<br>Additional rules:<br>Only whole values can be appended, not partial values.<br>The column will not be widened to 'best fit' the filled value<br>If appending an additional occurrence of the value exceeds the boundary of the cell<br>left/right edge, don't append the additional occurrence of the value.<br>The display value of the cell is filled, not the underlying raw number
| JUSTIFY | 5 | The horizontal alignment is justified (flush left and right). For each line of text, aligns each line of the wrapped text in a cell to the right and left (except the last line). If no single line of text wraps in the cell, then the text is not justified.
| CENTER_SELECTION | 6 | The horizontal alignment is centered across multiple cells. The information about how many cells to span is expressed in the Sheet Part, in the row of the cell in question. For each cell that is spanned in the alignment, a cell element needs to be written out, with the same style Id which references the centerContinuous alignment.
| DISTRIBUTED | 7 | Indicates that each 'word' in each line of text inside the cell is evenly distributed across the width of the cell, with flush right and left margins.<br>When there is also an indent value to apply, both the left and right side of the cell are padded by the indent value.<br> A 'word' is a set of characters with no space character in them.<br> Two lines inside a cell are separated by a carriage return.

##### index

**index** is only used to return sheet's index in a workbook 

| values | examples
|---|---
| integer | an integer value assigning the row index (zero-based)

##### italic

| values | explanation
|---|---
| 0 | set italic off
| 1 | set italic on

##### left

| string | hexadecimal | explanation
|---|---|---
| NONE | 0x0 | No border (default)
| THIN | 0x1 | Thin border
| MEDIUM | 0x2 | Medium border
| DASHED | 0x3 | dash border
| DOTTED | 0x4 | dot border
| THICK | 0x5 | Thick border
| DOUBLE | 0x6 | double-line border
| HAIR | 0x7 | hair-line border
| MEDIUM_DASHED | 0x8 | Medium dashed border
| DASH_DOT | 0x9 | dash-dot border
| MEDIUM_DASH_DOT | 0xA | medium dash-dot border
| DASH_DOT_DOT | 0xB | dash-dot-dot border
| MEDIUM_DASH_DOT_DOT | 0xC | medium dash-dot-dot border
| SLANTED_DASH_DOT | 0xD | slanted dash-dot border

##### name

| values | examples
|---|---
| string | the name of the font, e.g. "Courier New" (must be available on the computer)

##### orientation

**orientation** is used for the print layout

| string | integer | explanation
|---|---
| DEFAULT | 1 | orientation not specified
| PORTRAIT | 2 | portrait orientation
| LANDSCAPE | 3 | landscape orientations

##### range

| values | explanation
|---|---
| string | a cell range expressed like "A1:B2"

or

| child keys | values | explanation
|---|---
| [`top`](#top) | integer | the topmost cell index in the range
| [`left`](#left) | integer | the leftmost cell index in the range
| [`right`](#right) | integer | the rightmost cell index in the range
| [`bottom`](#bottom) | integer | the lowermost cell index in the range

| child keys
|---|---
| [`top_left`](#top_left) | the topmost cell index in the range
| [`bottom_right`](#bottom_right) | the rightmost cell index in the range

##### right



##### rotation

an integer that provides the degree of rotation for the text in the cell. Note: HSSF uses values from -90 to 90 degrees, whereas XSSF uses values from 0 to 180 degrees. The implementations of this method will map between these two value-ranges value-range as used by the type of Excel file-format that this CellStyle is applied to. 

Example

```
JSONSetElement ( $request ; "rotation" ; 90 ; JSONNumber )
```

##### row

an integer value containing the row index (zero-based) of a cell

Example

```
JSONSetElement ( $request ; "row" ; 24 ; JSONNumber )
```

##### sheet

a string containing the name of a sheet

Example

```
JSONSetElement ( $request ; "sheet" ; "mySheet" ; JSONString )
```

##### shrink_to_fit

| values | explanation
|---|---
| 0 | do not shrink to fit the text of a cell
| 1 | shrink to fit the text of a cell

```
JSONSetElement ( $request ; "shrink_to_fit" ; 0 ; JSONNumber )
```

##### size

contains the size of the font in points

```
JSONSetElement ( $request ; "size" ; 12 ; JSONNumber )
```

##### source

The source cell or cell range that will be copied to a [`target`](#target) cell or cell range

| child keys | explanation
|---|---
| [`cell`](#cell) | a cell 
| [`range`](#range) | a cell range

##### strike_out

| values | explanation
|---|---
| 0 | do not use a strikeout horizontal line through the text
| 1 | use a strikeout horizontal line through the text

##### style

used by |`border`](#border)

| child keys | value | explanation
|---|---
| border.[`top`](#top) | 
| border.[`left`](#left)
| border.[`right`](#right)
| border[`bottom`](#bottom)

##### target

The target cell or cell range that will be copied to from [`source`](#source) cell or cell range

| child keys | explanation
|---|---
| [`cell`](#cell) | a cell 
| [`range`](#range) | a cell range

##### top


##### top_left



##### type_offset

used by Xls.applyFontStyles

| values | explanation
|---|---
| SS_NONE | 0 | set normal
[ SS_SUPER | 1 | set super
| SS_SUB | 2 | set subscript

##### underline

| values | explanation
|---|---
| U_NONE | 0 | no underline
[ U_SINGLE | 1 | single underline
| U_DOUBLE | 2 | double underline
| U_SINGLE_ACCOUNTING | 33 | single accounting
| U_DOUBLE_ACCOUNTING | 34 | double accounting

##### values

##### vertical

used by [`alignment`](#alignment)

| string | integer | description
|---|---|---
| TOP | 0 | The vertical alignment is aligned-to-top.
| CENTER | 1 | The vertical alignment is centered across the height of the cell.
| BOTTOM | 2 | The vertical alignment is aligned-to-bottom. (typically the default value)
| JUSTIFY | 3 | When text direction is horizontal: the vertical alignment of lines of text is distributed vertically, where each line of text inside the cell is evenly distributed across the height of the cell, with flush top and bottom margins.<br>When text direction is vertical: similar behavior as horizontal justification. The alignment is justified (flush top and bottom in this case). For each line of text, each line of the wrapped text in a cell is aligned to the top and bottom (except the last line). If no single line of text wraps in the cell, then the text is not justified.
| DISTRIBUTED | 4 | When text direction is horizontal: the vertical alignment of lines of text is distributed vertically, where each line of text inside the cell is evenly distributed across the height of the cell, with flush top<br>When text direction is vertical: behaves exactly as distributed horizontal alignment. The first words in a line of text (appearing at the top of the cell) are flush with the top edge of the cell, and the last words of a line of text are flush with the bottom edge of the cell, and the line of text is distributed evenly from top to bottom.

##### workbook

a string containing the name of a workbook

Example

```
JSONSetElement ( $request ; "workbook" ; "myWorkbook" ; JSONString )
```

or

```
JSONSetElement ( $request ; "workbook" ; Get ( DocumentsPath ) & "myWorkbook.xlsx" ; JSONString )
```

##### wrap_text

get whether the text should be wrapped

| values | explanation
|---|---
| 0 | set wrap text off
| 1 | set wrap text on

```
JSONSetElement ( $request ; "wrap_text" ; 1 ; JSONNumber )
```

