Sub Rename_Snips()
' inserts the date after the file stem for all files in the snips subfolder
    
    'get the current date and parse ISO yyyy-mm-dd format from it
    myDate = Left(CStr(Now()), 10)
    
    'get a file list
    myFolder = "C:\temp\ncsu\snips\"
    fileSpec = "*.jpg"
    
    'loop through files, renaming as you go
    fName = Dir(myFolder & fileSpec)
    Do While fName <> ""
        newName = Left(fName, Len(fName) - 4) & " " & myDate & Right(fName, 4)
        Name myFolder & fName As myFolder & newName
        fName = Dir()
    Loop
    
End Sub
