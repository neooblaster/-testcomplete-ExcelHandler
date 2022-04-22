# Postinstall Current Working Directory


## Installation

- Considering our following project folder : ``/d/test``.
- Considering command line is open and in it this folder.

````
npm install @testcomplete/excelhandler
````

- Postinstall command ``dir > c:\Temp\excelhdl.log`` :

````plaintext
 Le volume dans le lecteur D s'appelle JetBrain
 Le num‚ro de s‚rie du volume est 6639-0590

 R‚pertoire de D:\test\node_modules\@testcomplete\excelhandler

21/04/2022  12:03    <DIR>          .
21/04/2022  12:03    <DIR>          ..
21/04/2022  12:03    <DIR>          bin
21/04/2022  12:03            31ÿ386 ExcelHandler.js
21/04/2022  12:03               578 package.json
21/04/2022  12:03             5ÿ779 README.md
               3 fichier(s)           37ÿ743 octets
               3 R‚p(s)  161ÿ977ÿ769ÿ984 octets libres
````

- Dependacies are install at the same level as the requested package :

````plaintext
[ test ] /
├─[ node_modules ] / 
| ├─[ .bin ] / 
| | ├─ tablejs 
| | ├─ tablejs.cmd 
| | └─ tablejs.ps1 
| | 
| ├─[ @neooblaster ] / 
| | └─[ tablejs ] / 
| |   ├─[ bin ] / 
| |   | └─ TableJs.js 
| |   | 
| |   ├─[ releases ] / 
| |   | ├─ latest.js 
| |   | ├─ v0.1.0.js 
| |   | └─ (...)
| | 
| ├─[ @testcomplete ] /             
| | └─[ excelhandler ] /                <<< Requested Package & Current WD
| |   ├─[ bin ] / 
| |   | └─ postinstall.js               <<< 
| |   | 
| |   ├─ ExcelHandler.js 
| |   ├─ package.json 
| |   └─ README.md 
| | 
| └─ .package-lock.json 
| 
├─ package-lock.json 
└─ package.json 
````