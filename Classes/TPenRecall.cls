VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TPenRecall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'  TPenRecall = Class(TRecall)
Implements TRecall
Private mBase As TRecall
Private Sub Class_Initialize()
  Set mBase = New TRecall
End Sub
'  public
'    constructor Create(APen: TPen);
Public Sub NewC(APen As TPen)
Dim NewPen As New TPen
  Call mBase.NewC(NewPen, APen)
End Sub
'so der Rest kommt vom Interface und wird an mBase weitergeleitet
Private Sub TRecall_NewC(AStorage As Object, AReference As Object)
  'hier muﬂ ma nix mehr machen
End Sub
Private Sub TRecall_Store()
  Call mBase.Store
End Sub
Private Property Get TRecall_Reference() As Object
  Set TRecall_Reference = mBase.Reference
End Property
Private Sub TRecall_Forget()
  Call mBase.Forget
End Sub

