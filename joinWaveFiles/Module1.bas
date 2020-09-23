Attribute VB_Name = "Module1"
Type waveFileInfo
    fileName As String
    fileSize As Long
    holder As String
    firstVal As Long
    compressionType As Integer
    channel As Integer
    sampFrequen As Long
    multipleSampFrequen As Long
    divisor As Integer
    numOfBytes As Integer
    numOfDataBytes As Long
    lenData As Long
    inData As Long
End Type
Public holder As String

Public wavefiles() As waveFileInfo
Public curFile As Integer
Public numOfFiles As Integer

Public newWaveFile As waveFileInfo


Public Function getWaveFileData(tmpWaveFile As waveFileInfo)

    'On Error GoTo Errhandler


    Open tmpWaveFile.fileName For Binary Access Read As #1
    wavefiles(curFile).fileSize = LOF(1)

    For n = 1 To 100
        X$ = Input(4, #1)
    If n = 2 Then tmpWaveFile.holder = X$  ' Hold This for Saving a New Wav
    If X$ = "fmt " Then Exit For 'Ignore everything else till this
    Next n
    'Get the Wave File Header Info
    Get #1, , tmpWaveFile.firstVal  ' 16
    Get #1, , tmpWaveFile.compressionType  'Compression Type (1=PCM)
    Get #1, , tmpWaveFile.channel  'is Channels, 1 if mono and 2 if stereo

    Get #1, , tmpWaveFile.sampFrequen  'is the Sampling frequency of the file

    Get #1, , tmpWaveFile.multipleSampFrequen  'is a multiple of the sample frequency

    Get #1, , tmpWaveFile.divisor  'is the divisor of the number of bytes of
                                'data which gives the number of Samples in the .wav
    Get #1, , tmpWaveFile.numOfBytes  'is the number of bits (8 or 16)
    
        'find data
    For n = 1 To 100
        Y$ = Input(1, #1)
        If Y$ = "d" Then Exit For ' Seek for start of Wav Data
    Next n
    
    Z$ = Input(3, #1)
    
    Get #1, , tmpWaveFile.numOfDataBytes  '= num of bytes of data, start reading data here.

    tmpWaveFile.lenData = tmpWaveFile.numOfDataBytes / tmpWaveFile.divisor
  
    LenTemp = tmpWaveFile.lenData / tmpWaveFile.sampFrequen
    Extemp = (Int(LenTemp * 1000)) / 1000 ' time
    If LenTemp - Extemp >= 0.0005 Then
        Extemp = Extemp + 0.001
    End If
    
    FimPlay = Int(LenTemp * 1000)
    tmpWaveFile.inData = Seek(1)  'Loc(1) + 1 is the number of the first sound data byte of the file
  
    Close #1
 
End Function
Public Function GraphWave(channel As Integer, lenData As Long, fileName As String, _
numOfBytes As Integer, inData As Long)
    Dim yByte As Byte
    Dim yzero As Double, xmax As Double, xmult As Double, ySelFat As Double
    Dim yint As Integer, yPos As Integer, yGraf As Integer
    Dim limsup As Integer
    Dim ySel As Long
    Dim nMult As Double, xPos As Integer
    
    frmMain.Picture2.Cls
    frmMain.Picture4.Cls
    frmMain.Picture5.Cls
    frmMain.Picture7.Cls
        
    Open fileName For Binary Access Read As #1
    
    If channel = 2 Then 'stereo
        frmMain.Picture2.Height = 1240
        frmMain.Picture5.Visible = True
      Else
        frmMain.Picture2.Height = 2480
        frmMain.Picture5.Visible = False
    End If
    ySelFat = lenData / frmMain.Picture2.ScaleWidth

    xzero = 0
    yzero = frmMain.Picture2.ScaleHeight / 2
    xmax = frmMain.Picture2.ScaleWidth
    ymax = 128
    ymaxgraf = frmMain.Picture2.ScaleHeight * 3 / 8
    ymult = ymaxgraf / ymax
    yPos = Int(yzero + 15 * 128)
    frmMain.Picture2.Line (0, yzero)-(xmax, yzero)
    frmMain.Picture2.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    frmMain.Picture2.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    frmMain.Picture4.Line (0, yzero)-(xmax, yzero)
    frmMain.Picture4.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    frmMain.Picture4.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    If channel = 2 Then GoTo Stereo8
    If numOfBytes = 16 Then GoTo Mono16
Mono8:
    Get #1, inData, yByte
    yGraf = yPos - 15 * yByte
    frmMain.Picture2.PSet (xzero, yGraf)
    frmMain.Picture4.PSet (xzero, yGraf)
    If lenData <= frmMain.Picture2.ScaleWidth Then
        nMult = frmMain.Picture2.ScaleWidth / lenData
        For n = 1 To lenData - 1
            xPos = Int(n * nMult)
            Get #1, , yByte
            yGraf = yPos - 15 * yByte
            frmMain.Picture2.Line -(xPos, yGraf)
            frmMain.Picture4.Line -(xPos, yGraf)
        Next n
      Else
        For n = 1 To 9999
            ySel = inData + Int(n * ySelFat)
            Get #1, ySel, yByte
            yGraf = (yPos - 15 * yByte)
            frmMain.Picture2.Line -(n, yGraf)
            frmMain.Picture4.Line -(n, yGraf)
        Next n
    End If
    GoTo Done
Mono16:
    Get #1, inData, yint
    yGraf = yzero - yint / 17
    frmMain.Picture2.PSet (xzero, yGraf)
    frmMain.Picture4.PSet (xzero, yGraf)
    If lenData <= frmMain.Picture2.ScaleWidth Then
        nMult = frmMain.Picture2.ScaleWidth / lenData
        For n = 1 To lenData - 1
            xPos = Int(n * nMult)
            Get #1, , yint
            yGraf = yzero - yint / 17
            frmMain.Picture2.Line -(xPos, yGraf)
            frmMain.Picture4.Line -(xPos, yGraf)
        Next n
      Else
        For n = 1 To 9999
            ySel = inData + 2 * Int(n * ySelFat)
            Get #1, ySel, yint
            yGraf = yzero - yint / 17
            frmMain.Picture2.Line -(n, yGraf)
            frmMain.Picture4.Line -(n, yGraf)
        Next n
    End If
    GoTo Done
Stereo8:
    frmMain.Picture2.CurrentX = 0
    frmMain.Picture2.CurrentY = 0
    frmMain.Picture2.Print "Left"
    frmMain.Picture4.CurrentX = 0
    frmMain.Picture4.CurrentY = 0
    frmMain.Picture4.Print "Left"
    frmMain.Picture5.Line (0, yzero)-(xmax, yzero)
    frmMain.Picture5.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    frmMain.Picture5.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    frmMain.Picture5.CurrentX = 0
    frmMain.Picture5.CurrentY = 0
    frmMain.Picture5.Print "Right"
    frmMain.Picture7.Line (0, yzero)-(xmax, yzero)
    frmMain.Picture7.Line (0, yzero - ymax * ymult)-(xmax, yzero - ymax * ymult), &H8000000F
    frmMain.Picture7.Line (0, yzero + ymax * ymult)-(xmax, yzero + ymax * ymult), &H8000000F
    frmMain.Picture7.CurrentX = 0
    frmMain.Picture7.CurrentY = 0
    frmMain.Picture7.Print "Right"
    If numOfBytes = 16 Then GoTo Stereo16
    yPos = Int(yzero + 7 * 128)
    Get #1, inData, yByte 'left Channel
    yGraf = yPos - 7 * yByte '15 * yByte
    frmMain.Picture2.PSet (xzero, yGraf)
    frmMain.Picture4.PSet (xzero, yGraf)
    Get #1, , yByte 'right Channel
    yGraf = yPos - 7 * yByte
    frmMain.Picture5.PSet (xzero, yGraf)
    frmMain.Picture7.PSet (xzero, yGraf)
    If lenData <= frmMain.Picture2.ScaleWidth Then
        nMult = frmMain.Picture2.ScaleWidth / lenData
        For n = 1 To lenData - 1
            xPos = Int(n * nMult)
            Get #1, , yByte 'left Channel
            yGraf = yPos - 7 * yByte '15 * yByte
            frmMain.Picture2.Line -(xPos, yGraf)
            frmMain.Picture4.Line -(xPos, yGraf)
            Get #1, , yByte 'right Channel
            yGraf = yPos - 7 * yByte
            frmMain.Picture5.Line -(xPos, yGraf)
            frmMain.Picture7.Line -(xPos, yGraf)
        Next n
      Else
        For n = 1 To 9999
            ySel = inData + 2 * Int(n * ySelFat)
            Get #1, ySel, yByte 'left Channel
            yGraf = yPos - 7 * yByte '15 * yByte
            frmMain.Picture2.Line -(n, yGraf)
            frmMain.Picture4.Line -(n, yGraf)
            Get #1, , yByte 'right Channel
            yGraf = yPos - 7 * yByte
            frmMain.Picture5.Line -(n, yGraf)
            frmMain.Picture7.Line -(n, yGraf)
        Next n
    End If
    GoTo Done

Stereo16:
    Get #1, inData, yint 'left Channel
    yGraf = yzero - yint / 35 '17
    frmMain.Picture2.PSet (xzero, yGraf)
    frmMain.Picture4.PSet (xzero, yGraf)
    Get #1, , yint 'right Channel
    yGraf = yzero - yint / 35
    frmMain.Picture5.PSet (xzero, yGraf)
    frmMain.Picture7.PSet (xzero, yGraf)
    If lenData <= frmMain.Picture2.ScaleWidth Then
        nMult = frmMain.Picture2.ScaleWidth / lenData
        For n = 1 To lenData - 1
            xPos = Int(n * nMult)
            Get #1, , yint 'left Channel
            yGraf = yzero - yint / 35 '17
            frmMain.Picture2.Line -(xPos, yGraf)
            frmMain.Picture4.Line -(xPos, yGraf)
            Get #1, , yint 'right Channel
            yGraf = yzero - yint / 35
            frmMain.Picture5.Line -(xPos, yGraf)
            frmMain.Picture7.Line -(xPos, yGraf)
        Next n
      Else
        For n = 1 To 9999
            ySel = inData + 4 * Int(n * ySelFat)
            Get #1, ySel, yint 'left Channel
            yGraf = yzero - yint / 35 '17
            frmMain.Picture2.Line -(n, yGraf)
            frmMain.Picture4.Line -(n, yGraf)
            Get #1, , yint 'right Channel
            yGraf = yzero - yint / 35
            frmMain.Picture5.Line -(n, yGraf)
            frmMain.Picture7.Line -(n, yGraf)
        Next n
    End If
    
Done:
    Close #1
   'draw tick marks
    'For n = 1 To 100
        'Picture1.Line (n * 100, Picture2.Top)-(n * 100, Picture2.Top - 100)
    'Next n

End Function

Private Sub WriteHeader(Chan As Integer, SampFreq As Long, Nbits As Integer, lenData As Long)
    
    Dim TmpR As Long
    Put #2, , "RIFF" ' RIFF Header Layer
    Put #2, 5, wavefiles(curFile).holder$
    Put #2, 9, "WAVE" ' WAVE Header Layer
    Put #2, 13, "fmt "
    Put #2, 17, 16 '16
    Put #2, 21, 1 ' Compression (None=1(PCM))
    
    Put #2, 23, Chan ' Channels 1 or 2
    Put #2, 25, SampFreq ' Sampling Rate
    TmpR = SampFreq * (Chan * (Nbits / 8))
    Put #2, 29, TmpR '  Calculation
    TmpR = (Nbits / 8) * Chan
    Put #2, 33, TmpR 'Calculation
    Put #2, 35, Nbits ' Sampling bits
             ' End of WAVE Header Layer
    Put #2, 37, "data" ' Sound Data Layer
    Put #2, , lenData * TmpR ' Number of Samples in Wav
    'Starts a Binary Copy from the Selected Area in the Wav File
            'to the Newly created Untitled Wav File.
End Sub

Public Sub SaveWave(FName As String)

    Dim yByte As Byte
    Dim yint As Integer
    Dim allData As Long
    
    Open wavefiles(curFile).fileName For Binary Access Read As #1

    Open FName For Binary Access Write As #2
    ' Create or Overwrite a File Named Untitled(FormInstance).wav
  
    For i = 0 To numOfFiles - 1
        allData = allData + wavefiles(i).lenData
    Next
    
    WriteHeader wavefiles(curFile).channel, wavefiles(curFile).sampFrequen, wavefiles(curFile).numOfBytes, allData ' Write Header Info
    
    If wavefiles(curFile).channel = 2 Then GoTo Stereo8
    If wavefiles(curFile).numOfBytes = 16 Then GoTo Mono16
Mono8:
    Get #1, wavefiles(curFile).inData, yByte ' Points to First Block of Selection in source wav
    Put #2, , yByte ' Writes to Next Block in New File
        For n = 1 To wavefiles(curFile).lenData - 1
            Get #1, , yByte ' Points to Next Block of Selection in source wav
            Put #2, , yByte ' Writes to Next Block in New File
        Next n
           ' my add
        For i = 1 To numOfFiles - 1
            Open wavefiles(i).fileName For Binary Access Read As #(i + 2)
            Get #(i + 2), wavefiles(i).inData, yByte ' Points to First Block of Selection in source wav
            Put #2, , yByte ' Writes to Next Block in New File
            
            For n = 1 To wavefiles(i).lenData - 1
            
                Get #(i + 2), , yByte ' Points to Next Block of Selection in source wav
                Put #2, , yByte ' Writes to Next Block in New File
        
            Next n
            Close #(i + 2)
        Next
    GoTo Done
    
Mono16:
    Get #1, wavefiles(curFile).inData, yint ' Points to First Block of Selection in source wav
    Put #2, , yint ' Writes to Next Block in New File
       
        For n = 1 To wavefiles(curFile).lenData - 1
            
            Get #1, , yint ' Points to Next Block of Selection in source wav
            Put #2, , yint ' Writes to Next Block in New File
        
        Next n
            
       ' my add
        For i = 1 To numOfFiles - 1
            Open wavefiles(i).fileName For Binary Access Read As #(i + 2)
            Get #(i + 2), wavefiles(i).inData, yint ' Points to First Block of Selection in source wav
            Put #2, , yint ' Writes to Next Block in New File
            
            For n = 1 To wavefiles(i).lenData - 1
            
                Get #(i + 2), , yint ' Points to Next Block of Selection in source wav
                Put #2, , yint ' Writes to Next Block in New File
        
            Next n
            Close #(i + 2)
        Next
    GoTo Done

Stereo8:
    
    If wavefiles(curFile).numOfBytes = 16 Then GoTo Stereo16

    Get #1, wavefiles(curFile).inData, yByte 'left Channel
    Put #2, , yByte ' Writes to Next Block in New File
   
    Get #1, , yByte 'right Channel
    Put #2, , yByte ' Writes to Next Block in New File
    
        For n = 1 To wavefiles(curFile).lenData - 1
            
            Get #1, , yByte 'left Channel
            Put #2, , yByte ' Writes to Next Block in New File
            
           
            Get #1, , yByte 'right Channel
            Put #2, , yByte ' Writes to Next Block in New File
        Next n
    GoTo Done

Stereo16:
    Get #1, wavefiles(curFile).inData, yint 'left Channel
    Put #2, , yint ' Writes to Next Block in New File
    
    Get #1, , yint 'right Channel
    Put #2, , yint ' Writes to Next Block in New File
    
        For n = 1 To wavefiles(curFile).lenData - 1
            
            Get #1, , yint 'left Channel
            Put #2, , yint ' Writes to Next Block in New File
         
            Get #1, , yint 'right Channel
            Put #2, , yint ' Writes to Next Block in New File
            
        Next n
         
Done:
    Close #1
    Close #2

End Sub

Public Function getWaveFileTitle()
    Dim fileName, channel, khz, bits As String
    fileName = wavefiles(curFile).fileName
    If wavefiles(curFile).channel = 1 Then
        channel = "Mono"
    Else
        channel = "Stereo"
    End If
    tmp = Int(wavefiles(curFile).sampFrequen / 1000)
    khz = tmp & "khz"
    bits = wavefiles(curFile).numOfBytes & "bits"
    getWaveFileTitle = khz & " " & bits & " " & channel & "     " & fileName

End Function

Public Function checkAllWaveFiles()
    Dim errorType, tmp As String
    Dim i As Integer
    ' check for same channel
    For i = 1 To UBound(wavefiles)
        If wavefiles(i - 1).channel <> wavefiles(i).channel Then
            errorType = "Channel"
            GoTo matchError
        End If
    Next
    ' check for same khz
    For i = 1 To UBound(wavefiles)
        If wavefiles(i - 1).sampFrequen <> wavefiles(i).sampFrequen Then
            errorType = "khz"
            GoTo matchError
        End If
    Next
        ' check for same bits
    For i = 1 To UBound(wavefiles)
        If wavefiles(i - 1).numOfBytes <> wavefiles(i).numOfBytes Then
            errorType = "Bits"
            GoTo matchError
        End If
    checkAllWaveFiles = True
    Next
    Exit Function

matchError:
 
            tmp = wavefiles(i - 1).fileName & " and " & wavefiles(i).fileName
            frmMain.lstCommands.AddItem "Error Message:"
            frmMain.lstCommands.AddItem tmp
            frmMain.lstCommands.AddItem errorType & " does not match"
            checkAllWaveFiles = False

End Function
