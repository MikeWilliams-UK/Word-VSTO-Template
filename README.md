# Word-VSTO-Template

Windows 10 (20H2)

Microsoft 365 Apps for enterprise [13127.21668 CtR]

Using Visual Studio 2019 [16.10.3] create a new Word VSTO Template project

![image](https://user-images.githubusercontent.com/13162784/125304918-e1694600-e325-11eb-94da-e75fa6fce261.png)

![image](https://user-images.githubusercontent.com/13162784/125306078-ccd97d80-e326-11eb-941d-d34fd5446553.png)

![image](https://user-images.githubusercontent.com/13162784/125306098-d19e3180-e326-11eb-84cc-1c61d52a5a07.png)

![image](https://user-images.githubusercontent.com/13162784/125308562-eb407880-e328-11eb-8191-53ef8fa46194.png)

Once project is created add simple class to log events, then run the project

![image](https://user-images.githubusercontent.com/13162784/125314259-f21dba00-e32d-11eb-99da-3f9c093e9753.png)

Save documnt to C:\Temp

Browse to C:\Temp
Double click to open document from C:\Temp

![image](https://user-images.githubusercontent.com/13162784/125315094-b1727080-e32e-11eb-96a0-7638830836f1.png)
![image](https://user-images.githubusercontent.com/13162784/125315253-d5ce4d00-e32e-11eb-9cb6-5e5c5217d698.png)

Copy files from bin\Debug into C:\Temp

Double click to open document from C:\Temp

![image](https://user-images.githubusercontent.com/13162784/125315600-3493c680-e32f-11eb-9067-88b9cb8e455f.png)

Running code in project PostBuild/FullyQualifyAssemblyLocation solves the issue of splattering the dlls around

If I save a document based on this template to my OneDrive folder I get this when double clicking on it.

![image](https://user-images.githubusercontent.com/13162784/125317456-ebdd0d00-e330-11eb-830c-3e52dc4c5658.png)
