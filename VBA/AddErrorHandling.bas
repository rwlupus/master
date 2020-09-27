Attribute VB_Name = "AddErrorHandling"
Option Explicit

Sub ErrorHandling()
    
    Dim sModuleName As String
    
    sModuleName = "Class1"
    
    ScanCodeModule sModuleName
' to do - check module name in error handling

End Sub

Private Sub ScanCodeModule(p_sModuleName As String, _
                            Optional p_bAddHandling As Boolean = False, _
                            Optional p_bFixHandlingNames As Boolean = False)

    Dim oCodeModule As Object

    Dim oRegexProcDefine As Object, _
        oRegexLineContinue As Object, _
        oRegexOnError As Object, _
        oRegexProcEnd As Object, _
        oRegexCountLines As Object, _
        oRegexProcedureName As Object
    
    Dim i As Long
    Dim lLinesCount As Long
    
    Dim sProcName As String, sProcType As String
    Dim lProcStartLine As Long, lProcEndLine As Long, lErrHandlProcName As Long
    
    Dim sErrHandlProcName As String
    
    Dim lNewLines As Long
    
    Dim bOnErrorExist As Boolean
    Dim bProcDefEnd As Boolean
    Dim bProcedureNameExist As Boolean
    
    Dim sOnError As String
    Dim sErrHandlerTemplate As String, sErrHandler As String
    Dim sModuleType As String, sModuleName As String
    
    Set oRegexProcDefine = CreateObject("vbscript.regexp")
    Set oRegexLineContinue = CreateObject("vbscript.regexp")
    Set oRegexOnError = CreateObject("vbscript.regexp")
    Set oRegexProcEnd = CreateObject("vbscript.regexp")
    Set oRegexProcedureName = CreateObject("vbscript.regexp")
    
    Set oCodeModule = ActiveWorkbook.VBProject.VBComponents(p_sModuleName).CodeModule
      
    oRegexProcDefine.Pattern = "\w*\s?(Sub|Procedure|Function|Property)\s(.*)\(.*"
    oRegexLineContinue.Pattern = "_$"
    oRegexOnError.Pattern = "^\s*On Error \w+"
    oRegexProcedureName.Pattern = "^\s*Elc_ErrHanlder.PropagateException.*,\s*""(\w+)\s*"".*"
    
    
    If ActiveWorkbook.VBProject.VBComponents(p_sModuleName).Type = 2 Then
        sModuleName = "Me"
    Else
        sModuleName = """" & p_sModuleName & """"
    End If
    
    sOnError = vbNewLine & vbTab & "On Error Goto ErrHandler"
    sErrHandlerTemplate = vbNewLine & vbTab & "Exit ProcTypePlacholder" & vbNewLine & _
                  "ErrHandler:" & vbNewLine & _
                  vbTab & "Elc_ErrHanlder.GetLineNumber" & vbNewLine & _
                  vbTab & "Elc_ErrHanlder.PropagateException " & sModuleName & ", ""ProcNamePlaceholder""" & vbNewLine
    
    lNewLines = 9

    lLinesCount = oCodeModule.CountOfLines
    i = 1
    
    Do
        If oRegexProcDefine.test(oCodeModule.Lines(i, 1)) Then
            sProcType = oRegexProcDefine.Replace(oCodeModule.Lines(i, 1), "$1")
            sProcName = oRegexProcDefine.Replace(oCodeModule.Lines(i, 1), "$2")
            
            If oRegexLineContinue.test(oCodeModule.Lines(i, 1)) Then
                bProcDefEnd = False
            Else
                bProcDefEnd = True
                lProcStartLine = i
            End If
        End If
        
        If sProcName <> vbNullString And Not bProcDefEnd Then
            If oRegexLineContinue.test(oCodeModule.Lines(i, 1)) Then
                bProcDefEnd = False
            Else
                bProcDefEnd = True
                lProcStartLine = i
            End If
        End If
        
        If oRegexOnError.test(oCodeModule.Lines(i, 1)) Then bOnErrorExist = True
        If oRegexProcedureName.test(oCodeModule.Lines(i, 1)) Then
            sErrHandlProcName = oRegexProcedureName.Replace(oCodeModule.Lines(i, 1), "$1")
            lErrHandlProcName = i
            bProcedureNameExist = True
        End If
        
        oRegexProcEnd.Pattern = "\s*End\s" & sProcType
        
        If oRegexProcEnd.test(oCodeModule.Lines(i, 1)) Then lProcEndLine = i
        
        If lProcStartLine > 0 And lProcEndLine > 0 Then
            If bProcedureNameExist Then
                If sErrHandlProcName <> sProcName Then
                    Debug.Print "Wrong procedure name in error handling", sProcName, sErrHandlProcName, lErrHandlProcName
                    If p_bFixHandlingNames Then
                        oCodeModule.ReplaceLine lErrHandlProcName, Replace(oCodeModule.Lines(lErrHandlProcName, 1), sErrHandlProcName, sProcName)
                    End If
                End If
            End If
            
            
            If Not bOnErrorExist Then
                Debug.Print "Missing error handling", sProcName, sProcType, lProcStartLine, lProcEndLine
            
                sErrHandler = Replace(sErrHandlerTemplate, "ProcTypePlacholder", sProcType)
                sErrHandler = Replace(sErrHandler, "ProcNamePlaceholder", sProcName)
                
                If p_bAddHandling Then
                    oCodeModule.InsertLines lProcStartLine + 1, sOnError
                    oCodeModule.InsertLines lProcEndLine + 2, sErrHandler
                    i = i + lNewLines
                    lLinesCount = lLinesCount + lNewLines
                End If
                
            End If
        

       
            lProcStartLine = 0
            lProcEndLine = 0
            bOnErrorExist = False
        
        End If
        i = i + 1
        If i > lLinesCount Then Exit Do
        
    Loop

End Sub
