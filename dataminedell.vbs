'========================
strInputFile = "D:\Dell Research\list.txt"
strOutputFile = "Results.csv"
 
arrHeadings = Array("Service Tag:", "Days Left</td>")
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Const intForReading = 1
Set objHTTP = CreateObject("Msxml2.XMLHTTP")
 
strDetails = """Service Tag"",""System Type"",""Ship Date"",""Dell IBU"",""Description"",""Provider"",""Warrenty Extension Notice *"",""Start Date"",""End Date"",""Days Left"""
Set objInputFile = objFSO.OpenTextFile(strInputFile, intForReading, False)
While Not objInputFile.AtEndOfStream
      strServiceTag = objInputFile.ReadLine
      strCurrentTag = ""
      strURL = "http://support.dell.com/support/topics/global.aspx/support/my_systems_info/details?c=us&cs=RC968571&l=en&s=bsdr&~ck=anavml&~wsf=tabs&servicetag=" & strServiceTag 
      objHTTP.open "GET", strURL, False
      objHTTP.send
      strPageText = objHTTP.responseText
      For Each strHeading In arrHeadings
            intSummaryPos = InStr(LCase(strPageText), LCase(strHeading))
            If intSummaryPos > 0 Then
                  intSummaryTableStart = InStrRev(LCase(strPageText), "<table", intSummaryPos)
                  intSummaryTableEnd = InStr(intSummaryPos, LCase(strPageText), "</table>") + 8
                  strInfoTable = Mid(strPageText, intSummaryTableStart, intSummaryTableEnd - intSummaryTableStart)
                  strInfoTable = Replace(Replace(Replace(strInfoTable, VbCrLf, ""), vbCr, ""), vbLf, "")
                  arrCells = Split(LCase(strInfoTable), "</td>")
                  For intCell = LBound(arrCells) To UBound(arrCells)
                        arrCells(intCell) = Trim(arrCells(intCell))
                        intOpenTag = InStr(arrCells(intCell), "<")
                        While intOpenTag > 0
                              intCloseTag = InStr(intOpenTag, arrCells(intCell), ">") + 1
                              strNewCell = ""
                              If intOpenTag > 1 Then strNewCell = strNewCell & Trim(Left(arrCells(intCell), intOpenTag - 1))
                              If intCloseTag < Len(arrCells(intCell)) Then strNewCell = strNewCell & Trim(Mid(arrCells(intCell), intCloseTag))
                              arrCells(intCell) = Replace(Trim(strNewCell), " &nbsp;&nbsp;&nbsp;&nbsp;change service tag","")
                              intOpenTag = InStr(arrCells(intCell), "<")
                              
                        Wend
                  Next
                  'WScript.Echo Join(arrCells, "|")
                  If LCase(arrCells(0)) = LCase("Service Tag:") Then
                        'strCurrentTag = """" & strServiceTag & """"
                        strCurrentTag = ""
                        For intField = 1 To UBound(arrCells) Step 2
                              If strCurrentTag = "" Then
                                    strCurrentTag = """" & arrCells(intField) & """"
                              Else
                                    strCurrentTag = strCurrentTag & ",""" & arrCells(intField) & """"
                              End If
                        Next
                  ElseIf LCase(arrCells(0)) = LCase("Description") Then
                        For intField = 6 To UBound(arrCells)
                              strCurrentTag = strCurrentTag & ",""" & arrCells(intField) & """"
                        Next
                  End If
            Else
                  strCurrentTag = """" & strServiceTag & """,""No warranty information found."""
            End If
      Next
      strDetails = strDetails & VbCrLf & strCurrentTag
Wend
objInputFile.Close
Set objInputFile = Nothing
 
Set objOutputFile = objFSO.CreateTextFile(strOutputFile, True)
objOutputFile.Write strDetails
objOutputFile.Close
Set objOutputFile = Nothing
Set objFSO = Nothing
 
MsgBox "Done. Please see " & strOutputFile
'========================
