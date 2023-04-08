# excel-blazor-web-addin
A Blazor WebAssembly Excel Web AddIn 

![image](https://github.com/AlfredBr/excel-blazor-web-addin/blob/main/ExcelBlazorWebAddIn.png)

This is my first attempt at building a Microsoft Excel Web AddIn in Blazor WebAssembly.
I had a need to build an Microsoft Excel extension and I did not want to use VSTO.  The default
Excel Web AddIn template uses .NET Framework 4.8 and I wanted to use something more modern.

This builds on the base Blazor WebAssembly "Weather Forecast" project that is created by the default template.
I added some ideas that I found from around the internet and it works pretty well!

You can run this project both as a standard webpage in a browser and you can load the *manifest.xml* file into Microsoft Excel and run it in a TaskPane.
(To do this, you'll have to setup the */manifest* folder as a share on your machine and then use the Excel Trust Center to trust that add in folder. 
You have to do this step - it won't side-load in Visual Studio like the default template does.)
