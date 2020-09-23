Attribute VB_Name = "reDimMud"
Public Function reDimWaveFilesArray()
    If numOfFiles = 1 Then Exit Function
    ReDim temp(numOfFiles - 2) As waveFileInfo
    For i = 0 To numOfFiles - 2
        temp(i) = wavefiles(i)
    Next
    ReDim wavefiles(numOfFiles - 1)
    For i = 0 To numOfFiles - 2
        wavefiles(i) = temp(i)
    Next
    
End Function
