VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    ConfigureWorksheet
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim r As Integer
    Dim c As Integer
    Dim cel As gameCell
    r = Target.Row
    c = Target.Column
    Debug.Print "(" & r & "," & c & ")"
    If Not Board Is Nothing Then
        If Board.IsInGame Then
            Cancel = True
            If Board.IsCoordinatesInGameBoard(r - OffsetRows, c - OffsetColumns) Then
                Board.CellUncovered r - OffsetRows, c - OffsetColumns
                Board.CheckEndGame
            Else
                MsgBox ("Veuillez cliquer dans la partie de jeu")
            End If
        End If
    End If
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim r As Integer
    Dim c As Integer
    r = Target.Row
    c = Target.Column
    Debug.Print "(" & r & "," & c & ")"
    If Not Board Is Nothing Then
        If Board.IsInGame Then
            Cancel = True
            If Board.IsCoordinatesInGameBoard(r - OffsetRows, c - OffsetColumns) Then
                Board.SetFlag r - OffsetRows, c - OffsetColumns
                Board.CheckEndGame
            Else
                MsgBox ("Veuillez cliquer dans la partie de jeu")
            End If
        End If
    End If
End Sub
