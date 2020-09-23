VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResizeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'variable local para contener colección
Private mCol As Collection
Private oForm As Form
Dim PrevResizeX As Long
Dim PrevResizeY As Long

Dim nMinResizeX As Long
Dim nMinResizeY As Long

Public Function Add(ByVal sControlName As Object, ByVal nMode As Double) As clsResizeFormCtlCollec
   Dim objNew As clsResizeFormCtlCollec


   On Error Resume Next
   Set objNew = mCol(sControlName.Name)
   If Err Then
      Err.Clear
      On Error GoTo 0
      Set objNew = New clsResizeFormCtlCollec
      mCol.Add objNew, sControlName.Name
      objNew.ControlName = sControlName
      objNew.FunctionResize = nMode
   End If
   
   Set Add = objNew
   Set objNew = Nothing

End Function

Public Property Get Count() As Long

    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Sub Clear()
   With mCol
      Do While .Count > 0
         .Remove 1
      Loop
   End With
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Let FormName(FormName As Form)
   Set oForm = FormName
End Property

Public Property Let MinResizeX(nValue As Long)
    nMinResizeX = nValue
End Property
Public Property Let MinResizeY(nValue As Long)
    nMinResizeY = nValue
End Property
Public Property Get MinResizeX() As Long
    MinResizeX = nMinResizeX
End Property
Public Property Get MinResizeY() As Long
    MinResizeY = nMinResizeY
End Property


Public Function ResizeAll()
    'Dim tmpControl As Control
    Dim tmpControl As New clsResizeFormCtlCollec
    On Error Resume Next
    
    'si el formulario esta minimizado salimos
    If oForm.WindowState = vbMinimized Then
       Exit Function
    End If
    
    'controlamos el tamaño minimo del formulario
    If oForm.Height < nMinResizeY Then
       oForm.Height = nMinResizeY
    End If
    
    If oForm.Width < nMinResizeX Then
       oForm.Width = nMinResizeX
    End If
    
        
    'si no ha cambiado el tamaño del formulario salimos, evitamos recursividad
    If oForm.Width = PrevResizeX And oForm.Height = PrevResizeY Then
       Exit Function
    End If
        
    If PrevResizeX = 0 Then PrevResizeX = oForm.ScaleWidth
    If PrevResizeY = 0 Then PrevResizeY = oForm.ScaleHeight
    
    'ignora los errorres en los controles que no tengan WIDTH, HEIGHT, ETC...


    For Each tmpControl In mCol
    
        If TypeOf tmpControl.ControlName Is Line Then
            'Chequea si el control es una linea  (esto es temporal!!!)
            'Line change X1, X2, Y1, Y2
            Select Case tmpControl.FunctionResize
                Case 1
                    tmpControl.ControlName.X1 = tmpControl.ControlName.X1 / PrevResizeX * oForm.ScaleWidth
                    tmpControl.ControlName.X2 = tmpControl.ControlName.X2 / PrevResizeX * oForm.ScaleWidth
                    tmpControl.ControlName.Y1 = tmpControl.ControlName.Y1 / PrevResizeY * oForm.ScaleHeight
                    tmpControl.ControlName.Y2 = tmpControl.ControlName.Y2 / PrevResizeY * oForm.ScaleHeigh
                Case 2
                   ' tmpControl.ControlName.X1 = tmpControl.ControlName.X1 - (PrevResizeX - oForm.ScaleWidth)
                    tmpControl.ControlName.X2 = tmpControl.ControlName.X2 - (PrevResizeX - oForm.ScaleWidth)
                    'tmpControl.ControlName.Y1 = tmpControl.ControlName.Y1 / PrevResizeY * oForm.ScaleHeight
                    'tmpControl.ControlName.Y2 = tmpControl.ControlName.Y2 / PrevResizeY * oForm.ScaleHeigh
            End Select

            
        Else
           'Para el resto de controles: left, top, Width y height

            Select Case tmpControl.FunctionResize
                Case 1 'pegado a los cuatro puntos
                   tmpControl.ControlName.Width = tmpControl.ControlName.Width - (PrevResizeX - oForm.ScaleWidth)
                   tmpControl.ControlName.Height = tmpControl.ControlName.Height - (PrevResizeY - oForm.ScaleHeight)
                   
                Case 2 'pegado a la derecha
                   tmpControl.ControlName.Left = tmpControl.ControlName.Left - (PrevResizeX - oForm.ScaleWidth)
                   
                Case 3 'pegado a la derecha y abajo
                   tmpControl.ControlName.Left = tmpControl.ControlName.Left - (PrevResizeX - oForm.ScaleWidth)
                   tmpControl.ControlName.Top = tmpControl.ControlName.Top - (PrevResizeY - oForm.ScaleHeight)
                   
                Case 4 'pegado izquierda y derecha
                   tmpControl.ControlName.Width = tmpControl.ControlName.Width - (PrevResizeX - oForm.ScaleWidth)
                   
                Case 5 'pegado izquierda, derecha y abajo
                   tmpControl.ControlName.Width = tmpControl.ControlName.Width - (PrevResizeX - oForm.ScaleWidth)
                   tmpControl.ControlName.Top = tmpControl.ControlName.Top - (PrevResizeY - oForm.ScaleHeight)
                   
                Case 6 'pegado izquierda y abajo
                   tmpControl.ControlName.Top = tmpControl.ControlName.Top - (PrevResizeY - oForm.ScaleHeight)
                      
                Case 9 'metodo resize proporcional
                   tmpControl.ControlName.Left = tmpControl.ControlName.Left / PrevResizeX * oForm.ScaleWidth
                   tmpControl.ControlName.Top = tmpControl.ControlName.Top / PrevResizeY * oForm.ScaleHeight
                   tmpControl.ControlName.Width = tmpControl.ControlName.Width / PrevResizeX * oForm.ScaleWidth
                   tmpControl.ControlName.Height = tmpControl.ControlName.Height / PrevResizeY * oForm.ScaleHeight
            End Select
            
        End If
        
    Next tmpControl
    
    PrevResizeX = oForm.ScaleWidth
    PrevResizeY = oForm.ScaleHeight
End Function

Private Sub Class_Initialize()
    'crea la colección cuando se crea la clase
    Set mCol = New Collection
    nMinResizeX = 0
    nMinResizeY = 0
End Sub

Private Sub Class_Terminate()
    'destruye la colección cuando se termina la clase
    Set mCol = Nothing
End Sub

