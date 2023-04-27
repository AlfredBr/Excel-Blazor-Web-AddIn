# excel-blazor-web-addin
A Blazor WebAssembly Excel Web AddIn 

![image](https://github.com/AlfredBr/excel-blazor-web-addin/blob/main/ExcelBlazorWebAddIn.png)

This is a Microsoft Excel Web AddIn built in Blazor WebAssembly.
I had a need to build an Microsoft Excel extension and I did not want to use VSTO.  The default
Excel Web AddIn template uses .NET Framework 4.8 and I wanted to use something more modern.

This builds on the base Blazor WebAssembly "Weather Forecast" project that is created by the default template.
I added some ideas that I found from around the internet and it works pretty well!

You can run this project:
- as a standard webpage in a browser 
- in an Excel TaskPane after you load the *manifest.xml* file into Microsoft Excel

Important: To load the manifest.xml file, you'll have to setup the */manifest* folder as a share on your machine and then use the ```Excel Trust Center``` to add that shared folder name to the ```Trusted Catalogs Table```. *You must do this step! Visual Studio won't side-load the addin like the default template does.*
