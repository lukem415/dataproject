Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim CurrentDirectory
CurrentDirectory = fso.GetAbsolutePathName(".")
Dim Directory
Directory = fso.BuildPath(CurrentDirectory, "Kinetic.xlsm")    'declares path of Kinetic.xlsm

Dim filesys
Set filesys=CreateObject("Scripting.FileSystemObject")

Dim iCount
iCount = 1

Dim Directory1
Dim xl1
Dim wb1
Dim newFile1

For Each objFolder In filesys.GetFolder("C:\Users\Admin\Desktop\Janus").SubFolders

Wscript.Echo objFolder
Redim Preserve myArray(iCount)
myArray(iCount) = objFolder.Path
   
iCount=iCount + 1
	Set xl1 = CreateObject("Excel.Application")
	newFileName = "Kinetic" & iCount & ".xlsm"
	oldFile = Directory
	newFile = objFolder & "\" & newFileName
	filesys.CopyFile oldFile, newFile    
	
	Set wb1 = xl1.Workbooks.Open(newFile)
	xl1.Run newFilename & "!ImportCSVs"   'opens workbook at this path and runs ImportCSVs
	
	wb1.Save
	wb1.Close
	xl1.Quit
	
	newFile1 = "C:\Users\Admin\Desktop\Janus\hello\" & "Kinetic" & iCount & ".xlsm" 
	filesys.CopyFile newFile, newFile1   'copies path of newfile to path of newFile1
	
Next

Dim wb2
Dim xl2

Set xl2 = CreateObject("Excel.Application")
Set wb2 = xl2.Workbooks.Open(CurrentDirectory & "\hello\testing1.xlsm")
'xl2.Run "'testing1.xlsm'!Cycle"
xl2.Run "testing1.xlsm" & "!Cycle"

wb2.Save
wb2.Close
xl2.Quit