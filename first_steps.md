The following steps are required to create a Visual Basic 6 application with TX Spell .NET 7.0 and TX Text Control ActiveX 25.0:

1. Download the source code.

2. Download and install the [TX Spell .NET for Windows Forms trial version](http://www.textcontrol.com/en_US/downloads/trials/index/default/spelldotnet/).

3. Compile the ActiveX Package project.
   1. Open the project (AxTXSpell.vbproj) in Visual Studio 2017 and compile it using the *Rebuild Solution* menu entry of the *Build* main menu.
   2. Close Visual Studio again.
4. Open a command prompt with administrative privileges.
   1. Execute the following command in the same folder where the DLL has been created to register the type library:
   ```C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm AxTXSpell.dll /tlb:AxTXSpell.tlb```
5. Open the VB6 example
   1. Navigate to the installation folder and find the VB6 samples.
   2. Open the DialogAndContextMenu sample: C:\Program Files\Text Control GmbH\TX Spell .NET ActiveX Package\Samples\VB6\DialogAndContextMenu.
   3. Open the sample in Visual Basic 6.0 by double-clicking Project1.vbp.
   4. Press F5 to start the project.

6. To learn how to create new VB6 applications and how to add the reference to AxTXSpell.dll, please refer to the documentation.
