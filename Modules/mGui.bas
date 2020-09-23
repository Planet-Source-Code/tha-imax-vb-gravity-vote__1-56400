Attribute VB_Name = "mGui"
' Gui Handler Module for Gravity

Public Pnts As New Collection
Public vGrav As Single
Public Wind As Single

Public Sub SetControls()
    Main.cmbBall.Clear
    For i = 1 To Pnts.Count
        Main.cmbBall.AddItem Pnts(i).sName
    Next
    Main.cmbBall.AddItem "ALL"
    Main.slGravity.Value = vGrav * 100
    Debug.Print Main.slGravity.Value
    Main.slWind.Value = Wind * 100
End Sub

Public Sub RefreshNFO()
On Error Resume Next
    With Main
    .nfoX = "X: " & Pnts(.cmbBall).X
    .nfoY = "Y: " & Pnts(.cmbBall).Y
    .nfofX = "fSx: " & Pnts(.cmbBall).fSx
    .nfofY = "fSy: " & Pnts(.cmbBall).fSy
    .nfoRX = "RealX: " & .pDraw.ScaleWidth - Pnts(.cmbBall).X
    .nfoRY = "RealY: " & .pDraw.ScaleHeight - Pnts(.cmbBall).Y
    
    If Pnts(.cmbBall).Y < 0 Then
     .nfoB.Visible = True
    Else
     .nfoB.Visible = False
    End If
    End With
End Sub



Public Sub SetPoint(sName As String, WhichP As Integer, X As Integer, Y As Integer, g As Single, Col As Long)
    Pnts(WhichP).sName = sName
    Pnts(WhichP).X = X
    Pnts(WhichP).Y = Y
    Pnts(WhichP).g = g
    Pnts(WhichP).Col = Col
End Sub
