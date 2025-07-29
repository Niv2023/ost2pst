OST to PST converter tool

OVERVIEW:
This open-source application can be used for:
  1)	Exporting selected OST email folder tree into a PST file.
  2)	Compacting a PST file with the option to reset its password

USER INTERFACE:
  1)	Click in the <Select OST (or PST)> File button
      -	The open file dialog box defaults to “*.ost”… you can change the filetype to *.pst
      -	After the selected file is opened the file details are prompted
      -	And the input file tree is displayed
  2)	In case of an OST file:
      -	You may scroll the folder tree to select the (email) folder to be exported
      - Then click to the <Export “folder” to PST> to export it 
  3)	In case you selected a PST folder:
      -	You may select the “reset password" option before exporting to a new PST

NOTES:
  1)	The tool only supports UNICODE OST/PST
  2)	The conversion process can take a few minutes (depending on the email's nr/size)
  3)	This tool converts the OST data to PST, it does not provide emails previews/exports
      -	For that you may check https://github.com/Dijji/XstReader (I’ve learned a lot about OST/PST files from this tool)
  4)	OST files are usually located on C:\Users\<username>\AppData\Local\Microsoft\Outlook
      -	It might be advisable to copy the OST file to another location before running
  5)	You may use Microsoft’s SCANPST tool to check the new PST
      -	The tool should only report: “Only minor inconsistencies were found in he this file. Repairing the file is optional…”
      -	Please let me know if you get any error…. I will need the OST file to debug it.
      -	SCANPST is usually found within the Microsoft Office installation directory
6)	Please feel free to contact me in case you face any issues. I’ll try to respond as quickly as possible
7)	I would be glad to assist anyone wishing to improve the tool
8)	Please let me know if you liked it and wishes to “buy me a coffee” 😊


BASIC DESIGN PRINCIPLES
  1)	The first step is to read the OST NBT and BBT trees
  2)	After the user selects the top folder to export, the program marks al the NIDs (folders and messages) that will be exported. These are children from the selected folder NID
  3)	Then the process starts to convert these NIDs to the PST file format
      -	The NID’s ids are kept the same (LTP, PC, TCs NID’s reference remains unchanged)
      -	The OST BIDs will be rebuilt with the PST page/block sizes. This may have an impact on the message size
  4)	The program then reopens this temp PST file to recalculate and update the message sizes (on the related PC’s and TC’s)
  5)	In case the user selects a PST input file
      - the folder selection is disabled (exports all the PST)
      - the user may choose to reset the password
      - the program just exports all the NIDs/BIDs but in a new compact (no empty gaps) PST file
