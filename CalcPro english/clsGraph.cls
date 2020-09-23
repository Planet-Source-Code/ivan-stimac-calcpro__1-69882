VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'props
Private objPic As PictureBox
Private maxX As Single, maxY As Single, scaleX As Single, scaleY As Single
Private BC As OLE_COLOR, GC As OLE_COLOR, SC As OLE_COLOR
'pomocne varijable
Private centX As Long, centY As Long
Private xScUnitSz As Long, yScUnitSz As Long
Private unitsNum As Integer, i As Integer

'props
'max x
Public Property Get xMax() As Single

    xMax = maxX
    
End Property

Public Property Let xMax(ByVal nV As Single)

    maxX = nV
    
End Property

'max y
Public Property Get yMax() As Single

    yMax = maxY
    
End Property

Public Property Let yMax(ByVal nV As Single)

    maxY = nV
    
End Property

'scale x
Public Property Get xScale() As Single

    xScale = scaleX
    
End Property

Public Property Let xScale(ByVal nV As Single)

    scaleX = nV
    
End Property

'scale y
Public Property Get yScale() As Single

    yScale = scaleY
    
End Property

Public Property Let yScale(ByVal nV As Single)

    scaleY = nV
    
End Property

'colors
'back color
Public Property Get BackColor() As OLE_COLOR

    BackColor = BC
    
End Property

Public Property Let BackColor(ByVal nV As OLE_COLOR)

    BC = nV
    
End Property

'graph grid color
Public Property Get GraphColor() As OLE_COLOR

    GraphColor = GC
    
End Property

Public Property Let GraphColor(ByVal nV As OLE_COLOR)

    GC = nV
    
End Property

'graph scale color
Public Property Get ScaleColor() As OLE_COLOR

    ScaleColor = SC
    
End Property

Public Property Let ScaleColor(ByVal nV As OLE_COLOR)

    SC = nV
    
End Property

'
'set pictureBox object
Public Property Set PicBox(objPcBox As PictureBox)

    Set objPic = objPcBox
    
End Property

'crtanje grafa
Public Sub drawGraph()

    'podešavanje objekta
    objPic.Cls
    objPic.BackColor = BC
    objPic.ScaleMode = 3    'pixel
    objPic.ForeColor = SC
    '
    centX = objPic.ScaleWidth / 2
    centY = objPic.ScaleHeight / 2
    xScUnitSz = centX / maxX '- 1
    yScUnitSz = centY / maxY '- 1
    
    'crtanje koordinatnih osi
    objPic.Line (0, centY)-(objPic.ScaleWidth, centY), GC
    objPic.Line (centX, 0)-(centX, objPic.ScaleHeight), GC
    
    'crtanje x skale
    unitsNum = maxX / scaleX
    For i = 1 To unitsNum
    
        objPic.Line (centX - xScUnitSz * i * scaleX, centY - 5)-(centX - xScUnitSz * i * scaleX, centY + 5), GC
        objPic.Line (centX + xScUnitSz * i * scaleX, centY - 5)-(centX + xScUnitSz * i * scaleX, centY + 5), GC
        'ispis skale
        objPic.CurrentX = centX - xScUnitSz * i * scaleX - objPic.TextWidth(Str$("-" & i * scaleX))
        objPic.CurrentY = centY + 5
        objPic.Print "-" & i * scaleX
        
        objPic.CurrentX = centX + xScUnitSz * i * scaleX - objPic.TextWidth(Str$(i * scaleX))
        objPic.CurrentY = centY + 5
        objPic.Print i * scaleX
        
    Next i
    
    'crtanje y skale
    unitsNum = maxY / scaleY
    For i = 1 To unitsNum
    
        objPic.Line (centX - 5, centY - yScUnitSz * i * scaleY)-(centX + 5, centY - yScUnitSz * i * scaleY), GC
        objPic.Line (centX - 5, centY + yScUnitSz * i * scaleY)-(centX + 5, centY + yScUnitSz * i * scaleY), GC
        'ispis skale
        objPic.CurrentX = centX - 15 - objPic.TextWidth(Str$("-" & i * scaleY))
        objPic.CurrentY = centY + yScUnitSz * i * scaleY - objPic.TextHeight("1") / 2
        objPic.Print "-" & i * scaleY
        
        objPic.CurrentX = centX - 15 - objPic.TextWidth(Str$(i * scaleY))
        objPic.CurrentY = centY - yScUnitSz * i * scaleY - objPic.TextHeight("1") / 2
        objPic.Print i * scaleY
        
    Next i
End Sub

Public Sub DrawLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal Color1 As OLE_COLOR)
    
    On Error GoTo errH
    
    centX = objPic.ScaleWidth / 2
    centY = objPic.ScaleHeight / 2
    xScUnitSz = centX / maxX '- 1
    yScUnitSz = centY / maxY '- 1
    objPic.Line (centX + x1 * xScUnitSz, centY - y1 * yScUnitSz)-(centX + x2 * xScUnitSz, centY - y2 * yScUnitSz), Color1
errH:

End Sub

Private Sub Class_Initialize()

    maxX = 10
    maxY = 10
    scaleX = 1
    scaleY = 1
    BC = vbWhite
    GC = vbBlack
    SC = &H808080
    
End Sub
