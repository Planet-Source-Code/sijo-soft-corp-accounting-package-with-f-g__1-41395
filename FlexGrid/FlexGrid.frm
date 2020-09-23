VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FlexGridControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FlexGrid"
   ClientHeight    =   5700
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8685
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5235
      Left            =   105
      TabIndex        =   2
      Top             =   390
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   9234
      _Version        =   393216
      Rows            =   100
      Cols            =   100
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1020
      TabIndex        =   1
      Top             =   30
      Width           =   3720
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   1.17491e-38
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   4650
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu FileNew 
         Caption         =   "New"
      End
      Begin VB.Menu FileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu FileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu FileSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu FileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "Edit"
      Begin VB.Menu EditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu EditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu EditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu EditClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu EditSelect 
         Caption         =   "Select All"
      End
   End
   Begin VB.Menu FormatMenu 
      Caption         =   "Format"
      Begin VB.Menu FormatFont 
         Caption         =   "Font"
      End
      Begin VB.Menu FormatCellColor 
         Caption         =   "CellColor"
      End
      Begin VB.Menu Separator2 
         Caption         =   "-"
      End
      Begin VB.Menu FormatInteger 
         Caption         =   "###"
      End
      Begin VB.Menu Format2decimal 
         Caption         =   "###.00"
      End
      Begin VB.Menu FormatComaInteger 
         Caption         =   "#,###.00"
      End
      Begin VB.Menu FormatDollar 
         Caption         =   "$#,###.00"
      End
   End
   Begin VB.Menu MergeMenu 
      Caption         =   "Merge"
      Begin VB.Menu MergeFree 
         Caption         =   "Free Merge"
      End
      Begin VB.Menu MergeRows 
         Caption         =   "Merge Rows"
      End
      Begin VB.Menu MergeCols 
         Caption         =   "Merge Columns"
      End
      Begin VB.Menu MergeBoth 
         Caption         =   "Merge Both"
      End
      Begin VB.Menu MergeNone 
         Caption         =   "Do not Merge"
      End
   End
   Begin VB.Menu SortMenu 
      Caption         =   "Sort"
      Begin VB.Menu SortAsc 
         Caption         =   "Ascending"
         Begin VB.Menu AscNumeric 
            Caption         =   "Numeric"
         End
         Begin VB.Menu AscString 
            Caption         =   "String"
            Begin VB.Menu AscStringSensitive 
               Caption         =   "Case Sensitive"
            End
            Begin VB.Menu AscStringNonsensitive 
               Caption         =   "case insensitive"
            End
         End
         Begin VB.Menu AscGeneric 
            Caption         =   "All"
         End
      End
      Begin VB.Menu SortDesc 
         Caption         =   "Descending"
         Begin VB.Menu DescNumeric 
            Caption         =   "Numeric"
         End
         Begin VB.Menu DescString 
            Caption         =   "String"
            Begin VB.Menu DescStringSensitive 
               Caption         =   "Case Sensitive"
            End
            Begin VB.Menu DescStringNonSensitive 
               Caption         =   "case insensitive"
            End
         End
         Begin VB.Menu DescGeneric 
            Caption         =   "All"
         End
      End
   End
   Begin VB.Menu AlignMenu 
      Caption         =   "Align"
      Begin VB.Menu AlignLeft 
         Caption         =   "Align Left"
      End
      Begin VB.Menu AlignCenter 
         Caption         =   "Center"
      End
      Begin VB.Menu AlignRight 
         Caption         =   "Align Right"
      End
   End
End
Attribute VB_Name = "FlexGridControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  *************************************************
'  *************************************************
'  ** Sijo K Jose                                 **
'  ** sijo@softhome.net                           **
'  ** www.sijosoft.com publish on 1/1/2003        **
'  ** http://groups.msn.com/sijosoft our community**
'  *************************************************
'  *************************************************

Option Explicit
Dim OpenFile As String

Sub NumberCells()
Dim i As Integer

    For i = 1 To Grid.Rows - 1
        Grid.TextMatrix(0, i) = Format$(i, "000")
    Next
    For i = 1 To Grid.Cols - 1
        Grid.TextMatrix(i, 0) = " " & Format$(i, "000")
    Next
    Grid.ColWidth(0) = TextWidth("99999")
End Sub

Sub FormatCells(formatString)
Dim irow, icol As Integer

    For irow = Grid.Row To Grid.RowSel
        For icol = Grid.Col To Grid.ColSel
            Grid.TextMatrix(irow, icol) = Format$(Grid.TextMatrix(irow, icol), formatString)
        Next
    Next
End Sub

Private Sub AlignCenter_Click()
    Grid.FillStyle = flexFillRepeat
    Grid.CellAlignment = 4
    Grid.FillStyle = flexFillSingle
End Sub

Private Sub AlignLeft_Click()
    Grid.FillStyle = flexFillRepeat
    Grid.CellAlignment = 1
    Grid.FillStyle = flexFillSingle
End Sub

Private Sub AlignRight_Click()
    Grid.FillStyle = flexFillRepeat
    Grid.CellAlignment = 7
    Grid.FillStyle = flexFillSingle
End Sub

Private Sub AscGeneric_Click()
    Grid.Sort = 1
End Sub

Private Sub AscNumeric_Click()
    Grid.Sort = 3
End Sub

Private Sub AscStringNonsensitive_Click()
    Grid.Sort = 5
End Sub

Private Sub AscStringSensitive_Click()
    Grid.Sort = 7
End Sub

Private Sub DescGeneric_Click()
    Grid.Sort = 2
End Sub

Private Sub DescNumeric_Click()
    Grid.Sort = 4
End Sub

Private Sub DescStringNonSensitive_Click()
    Grid.Sort = 6
End Sub

Private Sub DescStringSensitive_Click()
    Grid.Sort = 8
End Sub

Private Sub EditClear_Click()
Dim irow As Integer, icol As Integer

    For irow = Grid.Row To Grid.RowSel
        For icol = Grid.Col To Grid.ColSel
            Grid.TextMatrix(irow, icol) = ""
        Next
    Next
End Sub

Private Sub EditCopy_Click()
Dim tmpText As String

    tmpText = Grid.Clip
    Clipboard.Clear
    Clipboard.SetText tmpText
End Sub

Private Sub EditCut_Click()
Dim tmpText As String

    tmpText = Grid.Clip
    Clipboard.Clear
    Clipboard.SetText tmpText
    EditClear_Click
End Sub

Private Sub EditPaste_Click()
Dim tmpText As String

    tmpText = Clipboard.GetText
    Grid.Clip = tmpText
End Sub

Private Sub EditSelect_Click()
    Grid.Row = 1
    Grid.Col = 1
    Grid.RowSel = Grid.Rows - 1
    Grid.ColSel = Grid.Cols - 1
End Sub

Private Sub FileNew_Click()
    Grid.Clear
    Text1.Text = ""
End Sub

Private Sub FileOpen_Click()
Dim allCells As String
Dim fnum As Integer
Dim curRow, curCol As Integer

On Error GoTo NoFileSelected
    CommonDialog1.Filter = "FlexGrid Files|*.grd|All Files|*.*"
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowOpen
    OpenFile = CommonDialog1.FileName
    fnum = FreeFile
    Open OpenFile For Input As #fnum
    Input #fnum, allCells
    EditSelect_Click
    Grid.Clip = allCells
    Close #fnum

    Grid.Row = 1
    Grid.Col = 1
    Grid.RowSel = Grid.Row
    Grid.ColSel = Grid.Col

    NumberCells
    Exit Sub

NoFileSelected:
    Exit Sub
    
End Sub

Private Sub FileSave_Click()
Dim fnum As Integer
Dim allCells As String

    EditSelect_Click
    allCells = Grid.Clip
    fnum = FreeFile
    If OpenFile = "" Then
        CommonDialog1.DefaultExt = "FLEXGRID Files|GDT"
        CommonDialog1.Action = 1
        OpenFile = CommonDialog1.FileName
        If OpenFile = "" Then Exit Sub
    End If
    Open OpenFile For Input As #fnum
    Input #fnum, allCells
    EditSelect_Click
    Grid.Clip = allCells
    Close #fnum
End Sub

Private Sub FileSaveAs_Click()
Dim allCells As String
Dim fnum As Integer
Dim curRow, curCol As Integer

    curRow = Grid.Row
    curCol = Grid.Col
    
    CommonDialog1.DefaultExt = "GRD"
    CommonDialog1.Action = 2
    If CommonDialog1.FileName = "" Then Exit Sub
    EditSelect_Click
    allCells = Grid.Clip
    fnum = FreeFile
    Open CommonDialog1.FileName For Output As #fnum
    Write #fnum, allCells
    Close #fnum
    
    Grid.Row = curRow
    Grid.Col = curCol
    Grid.RowSel = Grid.Row
    Grid.ColSel = Grid.Col
End Sub

Private Sub Form_Load()
    NumberCells
End Sub

Private Sub Format2decimal_Click()
    FormatCells ("###.00")
End Sub

Private Sub FormatCellColor_Click()
    CommonDialog1.ShowColor
    Grid.FillStyle = flexFillRepeat
    Grid.CellBackColor = CommonDialog1.Color
    Grid.FillStyle = flexFillSingle
End Sub

Private Sub FormatComaInteger_Click()
    FormatCells ("#,###.00")
End Sub

Private Sub FormatDollar_Click()
    FormatCells ("$#,###.00")
End Sub

Private Sub FormatFont_Click()
On Error GoTo NoFontSelected

    CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
    CommonDialog1.Color = Grid.CellForeColor
    CommonDialog1.CancelError = True
    CommonDialog1.FontName = Grid.CellFontName
    CommonDialog1.FontBold = Grid.CellFontBold
    CommonDialog1.FontItalic = Grid.CellFontItalic
    CommonDialog1.FontSize = Grid.CellFontSize
    CommonDialog1.Color = Grid.CellForeColor
    CommonDialog1.ShowFont
    Grid.FillStyle = flexFillRepeat
    Grid.CellFontName = CommonDialog1.FontName
    Grid.CellFontBold = CommonDialog1.FontBold
    Grid.CellFontItalic = CommonDialog1.FontItalic
    Grid.CellFontSize = CommonDialog1.FontSize
    Grid.CellForeColor = CommonDialog1.Color
    Grid.FillStyle = flexFillSingle
    Exit Sub
    
NoFontSelected:
    Exit Sub

End Sub

Private Sub FormatInteger_Click()
    FormatCells ("###")
End Sub

Private Sub Grid_Click()
    Label1.Caption = Grid.TextMatrix(Grid.Col, 0) & " : " & Grid.TextMatrix(0, Grid.Row)
    Text1.Text = Grid.Text
    Text1.SetFocus
End Sub

Private Sub Grid_EnterCell()
On Error Resume Next

    Text1.Text = Grid.Text
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Grid_LeaveCell()
    Grid.Text = Text1.Text
End Sub

Private Sub MergeBoth_Click()
Dim irow, icol As Integer

    For irow = Grid.Row To Grid.RowSel
        Grid.MergeRow(irow) = True
    Next
    Grid.MergeCells = 4
    
    For icol = Grid.Col To Grid.ColSel
        Grid.MergeCol(icol) = True
    Next
    Grid.MergeCells = 4
    
End Sub

Private Sub MergeCols_Click()
Dim icol As Integer

    For icol = Grid.Col To Grid.ColSel
        Grid.MergeCol(icol) = True
    Next
    Grid.MergeCells = 3
End Sub

Private Sub MergeFree_Click()
Dim irow, icol As Integer

    For irow = Grid.Row To Grid.RowSel
        Grid.MergeRow(irow) = True
    Next
    For icol = Grid.Col To Grid.ColSel
        Grid.MergeCol(icol) = True
    Next
    Grid.MergeCells = 1
End Sub

Private Sub MergeNone_Click()
    Grid.MergeCells = 0
End Sub

Private Sub MergeRows_Click()
Dim irow As Integer
    
    For irow = Grid.Row To Grid.RowSel
        Grid.MergeRow(irow) = True
    Next
    Grid.MergeCells = 2
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim SRow, SCol As Integer

    If KeyAscii = 13 Then
        Grid.Text = Text1.Text
        SRow = Grid.Row + 1
        SCol = Grid.ColSel
        If SRow = Grid.Rows Then
            SRow = Grid.FixedCols
            If SCol < Grid.Cols - Grid.FixedCols Then SCol = SCol + 1
        End If
    
        Grid.Row = SRow
        Grid.Col = SCol
        Grid.RowSel = SRow
        Grid.ColSel = SCol
        Text1.Text = Grid.Text
        Text1.SetFocus
        KeyAscii = 0
    End If
End Sub
