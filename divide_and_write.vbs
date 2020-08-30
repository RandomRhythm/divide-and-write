Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim dicSort: Set dicSort = CreateObject("Scripting.Dictionary")
Dim dicExclude: Set dicExclude = CreateObject("Scripting.Dictionary")
Dim dicInclude: Set dicInclude = CreateObject("Scripting.Dictionary")
dim dictOut: Set DictOut = CreateObject("Scripting.Dictionary")



strOutFile = "C:\path\to\file\containing\list.txt"
strOutFileExt = ".txt"
boolCaseSensitive = False
intDivide = 8


wscript.echo "Please open master list"
strfileOpen1 = SelectFile( )

wscript.echo "Please open exclusion list"
strfileOpen2 = SelectFile( )


if objFSO.fileexists(strfileOpen2) then
  populateDict dicExclude, strfileOpen2
end if

if objFSO.fileexists(strfileOpen1) then
  populateDict dicSort, strfileOpen1



  for each item in dicSort
    if dicExclude.exists(item) = False then
      dicInclude.add item, 1
    end if
  next  

  if dicInclude.count > intDivide then
    intPerFile = dicInclude.count / intDivide
    intOutCount = 1
    for each item in dicInclude
      dictOut.add item, 0
      if dictOut.count > intPerFile then
        writeFile strOutFile & cstr(intOutCount) & strOutFileExt
        intOutCount = intOutCount +1
      end if
    next
    writeFile strOutFile & cstr(intOutCount) & strOutFileExt
  end if
end if

sub populateDict(emptyDict, openFilePath)
    Set objFile = objFSO.OpenTextFile(openFilePath)
    Do While Not objFile.AtEndOfStream

        strData = objFile.ReadLine
        if boolCaseSensitive = false then strData = lcase(strData)
        if emptyDict.exists(strData) = false then

          emptyDict.Add strData, 1
        end if
    Loop

    objFile.Close
end sub

sub writeFile(strFilePath)

      Set objFile = objFSO.CreateTextFile(strFilePath)
          for each iOut in dictOut
      objFile.WriteLine iOut
      next
      dictOut.RemoveAll
      objFile.close
end sub



Function SelectFile( )
    ' File Browser via HTA
    ' Author:   Rudi Degrande, modifications by Denis St-Pierre and Rob van der Woude
    ' Features: Works in Windows Vista and up (Should also work in XP).
    '           Fairly fast.
    '           All native code/controls (No 3rd party DLL/ XP DLL).
    ' Caveats:  Cannot define default starting folder.
    '           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
    '           Dialog title says "Choose file to upload".
    ' Source:   http://social.technet.microsoft.com/Forums/scriptcenter/en-US/a3b358e8-15&ælig;-4ba3-bca5-ec349df65ef6

    Dim objExec, strMSHTA, wshShell

    SelectFile = ""

    ' For use in HTAs as well as "plain" VBScript:
    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    ' For use in "plain" VBScript only:
    ' strMSHTA = "mshta.exe ""about:<input type=file id=FILE>" _
    '          & "<script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
    '          & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>"""

    Set wshShell = CreateObject( "WScript.Shell" )
    Set objExec = wshShell.Exec( strMSHTA )

    SelectFile = objExec.StdOut.ReadLine( )

    Set objExec = Nothing
    Set wshShell = Nothing
End Function