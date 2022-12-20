VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarkerBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim A() As Boolean
Dim M() As EntityMarker
Dim Max As Long
Dim X As Long

Sub Reset()
Max = -1
Erase M
Erase A
End Sub

Function MarkerCount(Optional CountAliveAmount As Boolean) As Long
If CountAliveAmount Then
    MarkerCount = -1
    For X = 0 To Max
    MarkerCount = MarkerCount - A(X)
    Next
Else
MarkerCount = Max
End If
End Function

Friend Function AddMarker(PosX As Long, PosY As Long, EType As EntityType_E) As Long
For X = 0 To Max
    If Not A(X) Then
    M(X).X = PosX
    M(X).Y = PosY
    M(X).W = EPs(EType).W
    M(X).H = EPs(EType).H
    M(X).T = EType
    A(X) = True
    AddMarker = X
    Exit Function
    End If
Next
Max = Max + 1
ReDim Preserve M(Max)
ReDim Preserve A(Max)
M(Max).X = PosX
M(Max).Y = PosY
M(Max).T = EType
M(Max).W = EPs(EType).W
M(Max).H = EPs(EType).H
A(Max) = True
AddMarker = Max
End Function

Friend Function GetMarker(index As Long) As EntityMarker
On Error Resume Next
GetMarker.T = -1
If A(index) Then GetMarker = M(index)
End Function

Friend Sub SetMarker(index As Long, EM As EntityMarker)
On Error Resume Next
M(index) = EM
End Sub

Sub KillMarker(index As Long)
On Error Resume Next
A(index) = False
End Sub

Function FindMarkerInPos(PosX As Long, PosY As Long) As Long
For X = 0 To Max
    If A(X) Then
        With M(X)
            If .T = GreenKoopa Then
            .T = .T
            End If
            If .X <= PosX And .Y <= PosY And .X + .W > PosX And .Y + .H > PosY Then
            FindMarkerInPos = X
            Exit Function
            End If
        End With
    End If
Next
FindMarkerInPos = -1
End Function

Function FindMarkerInRect(PosX As Long, PosY As Long, Width As Long, Height As Long) As Long
For X = 0 To Max
    If A(X) Then
        With M(X)
            If .T = GreenKoopa Then
            .T = .T
            End If
            If PosX >= .X + .W And PosY >= .Y + .H And .X >= PosX + Width And .Y >= PosY + Height Then
            FindMarkerInRect = X
            Exit Function
            End If
        End With
    End If
Next
FindMarkerInRect = -1
End Function

Sub RenderMarkers(Buff As GDIBuffer, OffX As Long, OffY As Long)
For X = 0 To Max
    If A(X) Then
        With EPs(M(X).T)
        EntitiesTexture.RenderTo Buff.Hdc, M(X).X * PPB - OffX + .TextureOffX, M(X).Y * PPB - OffY + .TextureOffY, .TextureW, .TextureH, .TextureX, .TextureY
        End With
    End If
Next
End Sub

Private Sub Class_Initialize()
Max = -1
End Sub

