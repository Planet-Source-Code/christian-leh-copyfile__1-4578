Attribute VB_Name = "File"

'MODULE: File
'=============================
'Programmed by Lehnix 24/11/99
'(NC)NON-Copyright Lehnix 1999
'=============================
'
'COPYFILE:
'Use in Bytes: 1 KB = 1024 Byte
'Transfering over Netzwork: CPY_BUFFER = 2048  (2 * 1024)
'Normal Transfer:       CPY_BUFFER = 10240 (10 * 2048)
'
'The buffersize is various, but recommended size is not
'more than 10240 Bytes per block

Private Const CPY_BUFFER = 10240
Private Const DRY_FILE = ":\Temp.$$$"

Public Function CopyFile(ByVal sSource As String, ByVal sDestination As String) As Boolean
    Dim iSNr As Integer
    Dim iDNr As Integer
    Dim iPercent As Integer
    Dim dCount As Double
    Dim dRest As Double
    Dim sTemp As String
    
    CopyFile = True
    If Not DiskReady(sDestination) Then GoTo ERROR_HANDLER
    If Not FileExist(sSource) Then GoTo ERROR_HANDLER
    If FileExist(sDestination) Then Kill sDestination
    
    iSNr = FreeFile
    iDNr = FreeFile + 1
    Open sSource For Binary As iSNr
        If LOF(iSNr) = 0 Then
            Close iSNr
            GoTo ERROR_HANDLER
        End If
        
        Open sDestination For Binary As iDNr
            If LOF(iSNr) > CPY_BUFFER Then
                sTemp = Space(CPY_BUFFER)
                dCount = Int(LOF(iSNr) / CPY_BUFFER)
                dRest = LOF(iSNr)
                For I = 1 To dCount
                    DoEvents
                    Get iSNr, , sTemp
                    Put iDNr, , sTemp
                    dRest = dRest - CPY_BUFFER
                    iPercent = Int((LOF(iDNr) / LOF(iSNr) * 100) + 1)
                    'progressbar
                Next I
                sTemp = String(dRest, CPY_SPACE)
                Get iSNr, , sTemp
                Put iDNr, , sTemp
            ElseIf LOF(iSNr) <= CPY_BUFFER And LOF(iSNr) > 0 Then
                sTemp = Space(LOF(iSNr))
                Get iSNr, , sTemp
                Put iDNr, , sTemp
                iPercent = Int((LOF(iDNr) / LOF(iSNr) * 100) + 1)
                'progressbar
            End If
            sTemp = Empty
            
        Close iSNr
    Close iDNr
        Exit Function

ERROR_HANDLER:
    CopyFile = False
    Close iSNr
    Close iDNr
End Function

Private Function FileExist(ByVal sFile As String) As Boolean
    Dim sTemp As String
    
    FileExist = True
    sTemp = Dir(sFile)
    If sTemp = Empty Then FileExist = False
End Function

Private Function DiskReady(ByVal sDisk As String) As Boolean
    Dim iDNr As Integer
    Dim sFile As String
    
    DiskReady = True
    iDNr = FreeFile
    sFile = Left(sDisk, 1) & DRY_FILE
    
    On Local Error GoTo ERROR_HANDLER
    Open sFile For Random As iDNr Len = 1: Close iDNr
    Kill sFile
    Exit Function
    
ERROR_HANDLER:
    DiskReady = False
End Function
