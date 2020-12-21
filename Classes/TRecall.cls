VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRecall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'  TRecall = Class(TObject)
'  private
'    FStorage, FReference: TPersistent;
Private FStorage As Object 'TBrush, TFont, oder TPen
Private FReference As Object 'TBrush, TFont, oder TPen
'  public
Private Sub Class_Initialize()
  '
End Sub
'    constructor Create(AStorage, AReference: TPersistent);
Public Sub NewC(AStorage As Object, AReference As Object)
  Set FStorage = AStorage
  Set FReference = AReference
  Call Store
End Sub
'    destructor Destroy; override;
Private Sub Class_Terminate()
  If Not FReference Is Nothing Then
    FReference.Assign (FStorage)
  End If
  Call Forget
End Sub
'    procedure Store;
Public Sub Store()
  If Not FReference Is Nothing Then
    Call FStorage.Assign(FReference)
  End If
End Sub
'    procedure Forget;
Public Sub Forget()
  Set FReference = Nothing
  'Call FreeAndNil(FStorage)
  Set FStorage = Nothing
End Sub
'    property Reference: TPersistent read FReference;
Public Property Get Reference() As Object
  Set Reference = FReference
End Property
'  end;
'brauchma nich:
'procedure FreeAndNil(var Obj);
'Var
'  Temp: TObject;
'begin
'  Temp := TObject(Obj);
'  Pointer(Obj) := nil;
'  Temp.Free;
'end;

