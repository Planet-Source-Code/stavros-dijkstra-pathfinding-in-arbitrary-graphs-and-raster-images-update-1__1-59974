VERSION 5.00
Begin VB.Form frmPic2at 
   BorderStyle     =   0  'None
   Caption         =   "Pic2at"
   ClientHeight    =   825
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   930
   Icon            =   "frmPic2at.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   825
   ScaleWidth      =   930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   315
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      Top             =   0
      Width           =   315
   End
End
Attribute VB_Name = "frmPic2at"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Image -> .at CONVERTER (the simple way)

Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

Private Sub Form_Load()
    Picture1.BackColor = ImpassableColor
End Sub

Public Function mConvert(ByVal FileName As String) As String
Dim lButton As Long

    With frmMDI.CommonDialog1
        On Error GoTo Err
        Picture1.Picture = LoadPicture(FileName)
        
        lButton = MsgBox("Convert to grayscale (recommended for correct path finding)?", vbYesNoCancel)
        
        If lButton <> vbCancel Then
            .FileName = vbNullString
            .Filter = "Tile Map (*.at)|*.at"
            .ShowSave
            
            If .FileName <> vbNullString Then
                DoEvents
                MousePointer = vbHourglass
                If Convert(.FileName, lButton) Then
                    mConvert = .FileName
                Else
                    MsgBox "Error converting image"
                End If
                MousePointer = vbDefault
            End If
        End If
    End With
    
    Unload Me
    Exit Function
Err:
    MsgBox "Error converting image"
    Unload Me
End Function

Private Function Convert(ByVal FileName As String, lButton As Long) As Boolean
Dim TileMap()         As Long
Dim i                 As Long
Dim j                 As Long
Dim px                As Long
Dim lDim              As Long
Dim lDimX             As Long
Dim lDimY             As Long
Dim lhDC              As Long
Dim lImpassableColor  As Long

On Error GoTo Err
    With Picture1
        DoEvents
        lDimX = .ScaleWidth
        lDimY = .ScaleHeight
        
        lImpassableColor = modPublic.ImpassableColor '//Update 1.1
        
        If lDimX < lDimY Then lDim = lDimY Else lDim = lDimX
            
        ReDim TileMap(lDim, lDim) '//Currently, only square TileMaps are supported.
                                  '//We assign the square aray as impassable. The unused space will be deleted when imported.
        For i = 0 To lDim
            For j = 0 To lDim
                TileMap(j, i) = lImpassableColor
            Next j
        Next i
        
        lhDC = .hDC
        For i = 0 To lDimX
            For j = 0 To lDimY
                px = GetPixel(lhDC, i, j)
                
                '//Update 1.1
                If px <> lImpassableColor Then
                '//End Update
                
                    If lButton = vbYes Then
                                                                    
                        px = ((px And &HFF&) + _
                              (px And &HFF00&) \ &H100& + _
                              (px And &HFF0000) \ &H10000) \ 3      '/Greyscale conversion
                    End If
                    '//Update 1.1
                    If lButton = vbYes Then px = 256 - px           '/WhiteCost=1 --> BlackCost = 256
                    TileMap(i, j) = px
                    '//End update
                End If
            Next j
        Next i
    End With
    
    Dim AtPtTColor As AtPoint
    Dim AtPtTGscale As AtPoint
    
    If lButton = vbYes Then
        AtPtTGscale.End = -1                 '/Is GreyScale
        AtPtTGscale.Start = lImpassableColor '//Update 1.1
    Else
        AtPtTColor.End = -1                  '/Is TrueColor
        AtPtTColor.Start = lImpassableColor  '//Update 1.1
    End If
    
    lButton = vbYes
    If Dir$(FileName) <> vbNullString Then
        lButton = MsgBox("File exists. Overwrite?", vbYesNo)
        If lButton = vbYes Then
            If GetAttr(FileName) = vbNormal Or vbArchive Then
                Kill (FileName)
            Else
                lButton = vbNo
            End If
        End If
    End If
    If lButton = vbYes Then
        i = FreeFile
        Open FileName For Binary Access Write As i
            Put i, , AtPtTColor
            Put i, , AtPtTGscale
            Put i, , TileMap()
        Close i
        Convert = True
    End If
Exit Function
Err:
Close i
End Function
