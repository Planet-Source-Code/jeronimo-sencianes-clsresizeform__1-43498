VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResizeFormCtlCollec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sControlName As Object
Private nFunctionResize As Integer

Public Property Get ControlName() As Object
   Set ControlName = sControlName
End Property

Public Property Let ControlName(ByVal vNewValue As Object)
   Set sControlName = vNewValue
End Property

Public Property Get FunctionResize() As Integer
   FunctionResize = nFunctionResize
End Property

Public Property Let FunctionResize(ByVal vNewValue As Integer)
   nFunctionResize = vNewValue
End Property
