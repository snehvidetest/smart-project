Private result As String
Private formID As Integer
Private formName As String
Private stopFormTest As Boolean
Private parameters As Scripting.Dictionary
Private parametersAndCols As Scripting.Dictionary
Private spmCells() As Variant
Private popCells() As Variant
Private rulCells() As Variant
Private groCells() As Variant


Sub RunTests()

    formName = "frm043"
    formID = 43
    
    Set parametersAndCols = Global_Test_Func.getParamtersAndTheirCols(formID)
    
    Dim nrTC As Integer, i As Integer
    nrTC = Application.WorksheetFunction.CountIf(testWS.Range("A:A"), formID)
    
    For i = 1 To nrTC
        Set parameters = New Scripting.Dictionary
        Testcase i
    Next i


End Sub


Private Function Testcase(tc As Integer)
    Dim review As Boolean, tcid As String
    
    'Reset spørgeskema workbook
    Global_Test_Func.resetSheets ThisWorkbook
    
    'Create the TCID
    tcid = Global_Test_Func.GetTCID(tc, formID)
    If logging Then
        Write #1, tcid
    End If
    
    Set parameters = New Scripting.Dictionary 'Resets the testcase parameters
    Set parameters = Global_Test_Func.getData(tcid, parametersAndCols)
    

    ThisWorkbook.Activate
    
    If parameters("run") = 0 Then
         Exit Function
    End If
        
    Select Case parameters("testSubject")
        Case "nextStep"
            SetFields
            frm002.forkertData.SetFocus
            frm043.CommandButton1_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            frm043.CommandButton2_Click
            result = Global_Test_Func.IsLoaded(formName)
            
        Case "noExtraPrints"
            SetFields
            If (parameters("testParameter") = "buttonOne") Then
                frm043.CommandButton1_Click
            ElseIf (parameters("testParameter") = "buttonTwo") Then
                frm043.CommandButton2_Click
            End If
            CheckNoExtraPrints
            Sheet1.recordChangingCells = False
        Case "checkCaption"
            If (parameters("testParameter") = "buttonOne") Then
                result = frm043.CommandButton1.Caption
            ElseIf (parameters("testParameter") = "buttonTwo") Then
                result = frm043.CommandButton2.Caption
            ElseIf (parameters("testParameter") = "beskrivelse") Then
                result = frm043.Label1.Caption
            End If
            
        Case Else
            MsgBox "Error in 'testsubject' input: tcid " & tcid 'Msgbox to stop the code because you made a mistake in the inputs..
    End Select
    
    'Comparison
    If result = parameters("expected") Then
        review = True
    Else:
        review = False
    End If

    KillForms

    'Print results
    Global_Test_Func.PrintTestResults tcid, result, review
    
    
End Function
Private Function SetFields()
   'The folowing code inserts the inputs into the actual form
    If (parameters("testParameter") = "frm005") Then
        frm002.forkertData.Value = True
        frm002.korrektData.Value = False
    ElseIf (parameters("testParameter") = "frm003") Then
        frm002.forkertData.Value = False
        frm002.korrektData.Value = True
    End If
    
End Function

Private Function CheckNoExtraPrints()
    popCells = Array()
    rulCells = Array()
    groCells = Array()
    spmCells = Array()

    'returns a string which shows either true or has the input of the cells that changed that shouldn't have been changed
    result = Global_Test_Func.CheckPrintsInAllSheets(spmCells, popCells, rulCells, groCells)
    
    'Cleans up all arrays and dictionaries
    Erase popCells, rulCells, groCells, spmCells
    Sheet9.spmChangedCells.RemoveAll
    Sheet5.groChangedCells.RemoveAll
    Sheet3.rulChangedCells.RemoveAll
    Sheet1.popChangedCells.RemoveAll
    
End Function
Private Function KillForms()
    'Closes forms
    ThisWorkbook.Activate
    If Global_Test_Func.IsLoaded("frm043") Then
        Unload frm043
    End If
    If Global_Test_Func.IsLoaded("frm005") Then
        Unload frm005
    End If
    If Global_Test_Func.IsLoaded("frm003") Then
        Unload frm003
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
End Function




