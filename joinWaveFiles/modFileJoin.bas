Attribute VB_Name = "modFileJoin"
'***************************************
'WRITTEN BY GERRY MC DONNELL 2002
'WEB WWW.GERRYMCDONNELL.COM
'DUNDALK,IRELAND
'DATE SEP 15 2002

'NOTES:
'========================================================================
'IF YOU MODIFY/USE THIS FUNCTION YOU MUST EMAIL ME THE NEW VERSION OF THIS
'FUNTION. AS YOU MIGHT BE ABLE TO MAKE IT BETTER THAN MY EFFORT.

'EMAIL: GERRYMCD@ OCEANFREE.NET
'***************************************

'JOIN FILES
'USES A BUFFER OF SIZE LENGTH TO READ IN DATA IN A FILE AND WRITE IT
'READSPOT AND WRITE SPOT ARE THE POSITIONS OF THE FILE POINTER IN EACH FILE


Public Const MB As Long = 1048576   'I MEGB AS BYTES
Public CancelJoin As Boolean        'FALG WHETER TO STOP THE FILE JOIN OP
Public Const DEFAULT_BUFFER As Long = MB * 4 'DEFAULT BUFFER IS  4 MEGS


Public Sub Join_File(ByRef FileArray() As String, ByVal Outputfile As String, Optional BUFFER_SIZE, Optional PROGRESSBAR, Optional lblProgress, Optional OVERWRITEPROMPT)
    Dim i As Long, x As Long
    Dim FILE_BUFFER() As Byte
    Dim Readspot As Long, SavedSpot As Long

    Dim PROGRESS As Integer
    Dim OUTPUTFILESIZE As Long, WRITTENFILESIZE As Long
    Dim ext As String
    
    CancelJoin = False
    
    'FIEL EXTENSION OF OUTOUT FILE IS THE SAME AS THE FIRST FILE
    '**********************************************
    ext = Get_file_Ext(FileArray(1))            'GET EXT
    Outputfile = Set_File_Ext(Outputfile, ext)  'SET IT
    Debug.Print "OUTPUT FILE IS=>"; Outputfile
    '**********************************************
    
    'CHECK FOR THE SIZE OF THE BUFFER
    '**********************************************
    If IsMissing(BUFFER_SIZE) Then
        BUFFER_SIZE = DEFAULT_BUFFER   'SIZE OF BUFFER IN BYTES
    ElseIf BUFFER_SIZE = 0 Then
        BUFFER_SIZE = DEFAULT_BUFFER   'SIZE OF BUFFER IN BYTES
    End If
    '**********************************************


    'DOES THE FILE AREADY EXISTS
    '***********************************************
    If IsMissing(OVERWRITEPROMPT) Then
      state = OverWrite_File(Outputfile)
      If state = False And Dir(Outputfile) <> "" Then
            Exit Sub
      End If
    ElseIf OVERWRITEPROMPT = True Then
        state = OverWrite_File(Outputfile)
        If state = False Then
            Exit Sub
        End If
    End If
    '**********************************************
    
    
    'CALULATE TOTAL OUTPUTSIZE OF OUR JOINED FILE
    '**********************************************
    OUTPUTFILESIZE = Get_TotalFileSize(FileArray())
    WRITTENFILESIZE = 0
    frmMain.lblStatus = "Expected Output Size: " & SizeString(OUTPUTFILESIZE)
    '************************************************


    'RESIZE BYTE ARRAY TO SIZE OF BUFFER
    '************************************************
    ReDim FILE_BUFFER(BUFFER_SIZE)
    SavedSpot = 1
    '************************************************

    'FILE POINTERS FOR THE READ FILE AND OUTPUT FILE
    '************************************************
    readfile = 1
    WriteFile = 2
    '************************************************
    
    
    Open Outputfile For Binary Access Write As #WriteFile 'Opens the file to write to it in binary

    For i = 1 To UBound(FileArray)

        Debug.Print "READING FILE:  " & FileArray(i)
        Debug.Print "FILE LENGHT=> "; FileLen(FileArray(i)); " BYTES"

        Open FileArray(i) For Binary Access Read As #readfile 'Opens the file to read from it in binary
            
            'IF THE FILE IS < THAN OUR BUFFER WE RESIZE THE BYTE ARRAY TO THE FILE
            '**************************************************************
            If FileLen(FileArray(i)) > BUFFER_SIZE Then
                ReDim FILE_BUFFER(BUFFER_SIZE)
                Debug.Print "FILE IS BIGGER THAN BUFFER"
            Else
                TMP = FileLen(FileArray(i))
                ReDim FILE_BUFFER(TMP)
                Debug.Print "FILE IS LESS THAN BUFFER"
            End If
            '**************************************************************

            'RESET READFILE POINT TO THE FIRST BYTE
            '**************************************************************
            Readspot = 1
            '**************************************************************


            'FIX WAV FILE HEADER SIZE OTHERWIZE OT WONT PLAY CORRECTLY
            '**************************************************************
            If ext = "wav" Then
                Call CreateWavHeader(OUTPUTFILESIZE, Outputfile)
                Readspot = 45
                writespot = 45
            End If
            '**************************************************************


            'READIN BUFFER OF FILE
            '**************************************************************
            NUMBUFFERS = Int(FileLen(FileArray(i)) / BUFFER_SIZE)
            If NUMBUFFERS = 0 Then NUMBUFFERS = 1
            Debug.Print NUMBUFFERS; " BUFFER(S) OF SIZE "; UBound(FILE_BUFFER); " ARE NEEDED."
            '**************************************************************

            For x = 1 To NUMBUFFERS
            
                'READ FROM READSPOT TO THE SIZE OF THE FILEBUFFER INTO THE FILEBUFFER ARRAY
                '************************************************************
                Get #readfile, Readspot, FILE_BUFFER()
                Readspot = Readspot + UBound(FILE_BUFFER)
                Debug.Print "READ=>"; Readspot
                '************************************************************

                'WRITE BUFFER
                '************************************************************
                Debug.Print "WRITING TO ADDRESS =>"; SavedSpot
                Put #WriteFile, SavedSpot, FILE_BUFFER()
                SavedSpot = SavedSpot + UBound(FILE_BUFFER)
                Debug.Print "NEXT WRITE ADDRESS "; SavedSpot
                '************************************************************

                DoEvents
                
                
                'CHECK TO SEE IF WE SHOULD STOP COZ OF USER CANCEL
                '************************************************
                If CancelJoin = True Then
                    Close
                    Erase FILE_BUFFER
                    Exit Sub
                End If
                '************************************************
               
                
                'PERCENT PROGRESS STILL NOT QUITE RIGHT
                '************************************************************
                If IsMissing(PROGRESSBAR) = False Then
                    WRITTENFILESIZE = WRITTENFILESIZE + (NUMBUFFERS * UBound(FILE_BUFFER))
                    N = Int((WRITTENFILESIZE / OUTPUTFILESIZE) * 100)
                    Debug.Print "% PROGRESS= "; N
                    
                    If N <= PROGRESSBAR.Max Then
                        PROGRESSBAR.Value = N
                    Else
                        PROGRESSBAR.Value = PROGRESSBAR.Max
                    End If
                    
                    'PROGRESS LABLE
                    If IsMissing(lblProgress) = False Then
                        lblProgress.Caption = N & " %"
                    End If
                    
                End If
                '************************************************************
                
            Next
        
            WRITENBYTES = NUMBUFFERS * UBound(FILE_BUFFER)
            Debug.Print "FILE LENGHT: "; FileLen(FileArray(i))
            Debug.Print "NOT WRITTEN: "; FileLen(FileArray(i)) - WRITENBYTES
        
            'IF WE HAVE A REMAINING AMOUNT LEFTOVER
            remaining = FileLen(FileArray(i)) - WRITENBYTES
            If remaining <> 0 Then
                ReDim FILE_BUFFER(remaining)
                Get #readfile, Readspot, FILE_BUFFER()
                Put #WriteFile, SavedSpot, FILE_BUFFER()
                SavedSpot = SavedSpot + UBound(FILE_BUFFER)
            End If
                
        Close #readfile
    
    Next
    
    
    'PERCENT PROGRESS
    '************************************************************
    If IsMissing(PROGRESSBAR) = False Then
        PROGRESSBAR.Value = 100
    End If
    'PROGRESS LABLE
    If IsMissing(lblProgress) = False Then
        lblProgress.Caption = 100 & " %"
    End If
    '************************************************************

    
    Erase FILE_BUFFER
    Close
    
    'CHECK OUR JOINED FILE SIZE
    '***********************************************************
    Debug.Print "THE OUTPUTTED FILE SIZE: "; FileLen(Outputfile)
    If FileLen(Outputfile) <> OUTPUTFILESIZE Then
        MsgBox "ERROR: Estmated Filesize of " & OUTPUTFILESIZE & " doe not match our outputted file of " & FileLen(Outputfile), vbExclamation, "Error"
    End If
    '***********************************************************
End Sub


Function OverWrite_File(sfile As String) As Boolean
On Error GoTo err

  If Dir(sfile) <> "" Then
            ask = MsgBox("File Exists" & vbCrLf & vbCrLf & sfile & vbCrLf & "Do You want to Replace the existing File?", vbExclamation + vbYesNo)
            If ask = vbYes Then
                Kill sfile
                OverWrite_File = True
            Else
                OverWrite_File = False
                Exit Function
            End If
    End If

err:
If err.Number <> 0 Then
    MsgBox "Error: OverWrite_File()" & vbCrLf & vbCrLf & err.Description, vbExclamation
End If
End Function


Public Function SizeString(ByVal num_bytes As Double) As String
    Const SIZE_KB As Double = 1024
    Const SIZE_MB As Double = 1024 * SIZE_KB
    Const SIZE_GB As Double = 1024 * SIZE_MB
    Const SIZE_TB As Double = 1024 * SIZE_GB
    If num_bytes < SIZE_KB Then
        SizeString = Format$(num_bytes) & " bytes"
    ElseIf num_bytes < SIZE_MB Then
        SizeString = Format$(num_bytes / SIZE_KB, "0.00") & " KB"
    ElseIf num_bytes < SIZE_GB Then
        SizeString = Format$(num_bytes / SIZE_MB, "0.00") & " MB"
    Else
        SizeString = Format$(num_bytes / SIZE_GB, "0.00") & " GB"
    End If
End Function

'CALULATES TEH TOTAL FILESIZE OF AN ARRAY OF FILEPATHS
Public Function Get_TotalFileSize(myFileArray() As String) As Long
    Dim OUTPUTFILESIZE As Long

    For i = 1 To UBound(myFileArray)
        OUTPUTFILESIZE = OUTPUTFILESIZE + FileLen(myFileArray(i))
    Next
    Debug.Print "OUTPUTTED FILE WILL BE "; OUTPUTFILESIZE
    Get_TotalFileSize = OUTPUTFILESIZE
End Function
