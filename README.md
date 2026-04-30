# MS-word-print-path-automator
A VBA-powered global template and Inno Setup installer that automatically prints a clean, non-intrusive file path on Microsoft Word documents for physical document traceability

! [Project Thumbnail](add file address in bottom auto.png)

**📄 Word Print-Path Automator**

A professional tool to solve document traceability in busy office environments.

**🔍 The Problem**

In many offices, physical documents get printed and separated from their digital sources. Finding the exact folder where a file is saved on a standalone PC can be a nightmare. Default Word "Insert Path" options are often too permanent or difficult for casual users to manage.

**✨ The Solution**

This project provides a Global Word Template (.dotm) that overrides the standard Ctrl+P behavior.

•	Smart Prompt: Asks the user if they want to include the file path before printing.

•	Invisible Digital Footprint: The path is added to the printed page but deleted immediately afterward, keeping the digital file clean.

•	Precise Positioning: Places a clean, grey 8pt Arial italic path at the absolute bottom margin.

**🚀 Installation**

I have included an Inno Setup script that compiles this tool into a one-click .exe installer.
1.	Download the PrintWithPathInstaller.exe from the Releases section.
2.	Run the installer (it will automatically place the template in your Word Startup folder).
3.	Restart Microsoft Word.
   
**🛠 Technical Details**

The core logic uses VBA to calculate the physical page height and inject a temporary text box:
VBA
' Anchored to the physical bottom of the page
pageHeight = doc.PageSetup.pageHeight
boxTop = pageHeight - 20 
Set tempTextBox = doc.Shapes.AddTextbox(msoTextOrientationHorizontal, 20, boxTop, 500, 22)

**📖 Learn More**

For a full breakdown of the code, the Inno Setup script logic, and a guide on how to customize the margins for your specific printer, visit my detailed blog post:
👉 Read the Full Guide on [AdvancedUser.net](https://advanceduser.net/add-file-path-in-word/)

**Screenshot**
! [Screenshot of result](won.JPG)

**👨‍💻 Developed By**

Aashir Mehmood
Founder of [AdvancedUser.net](https://advanceduser.net)
