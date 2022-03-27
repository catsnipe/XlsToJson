# XlsToJson or Scriptable Object

Unity: Convert excel sheet to Json / Scriptable object

![xls_to_json](https://user-images.githubusercontent.com/85425896/160281004-9c5fcdf7-df48-4c67-b0cf-94aea121fc7d.jpg)

## requirement
unity2017 or later  
npoi2.5.1 or later  

## usage
1. Download sample project and open.
2. Right click on Sample.xlsx and `XlsToJson Settings...`.
3. Click 'CREATE Importer'.  
   Source code is automatically generated.  
   
      * **Assets/Class_Character.cs**  
      * **Assets/Class_Stage.cs**  
         Defines the table type.  
      * **Assets/Character.cs**  
      * **Assets/Stage.cs**  
         Allows singleton access to tables.  
      * **Editor/XlsToJson/importer/*.cs**  
         These files will link Excel and Json data.  

4. Menu: Tools/XlsToJson/[Create] Json Data.  
   Json data will be generated in 'Resources/'.  
   
   ![image](https://user-images.githubusercontent.com/85425896/160277279-0873c5eb-272c-41e2-a668-97ba0cb4fb81.png)

5. After editing Json data, Menu: Tools/XlsToJson/[Xlsx Update] Json Data -> SampleData.xlsx  
   Export Json data to Excel.  

see more detail (japanese): https://www.create-forever.games/xls-to-json/

## license
This sample project includes the work that is distributed in the Apache License 2.0 / MIT / MIT X11.  

NPOI (Apache2.0): https://www.nuget.org/packages/NPOI/2.5.1/License  
SharpZLib (MIT): https://licenses.nuget.org/MIT  
Portable.BouncyCastle (MIT X11): https://www.bouncycastle.org/csharp/licence.html  
