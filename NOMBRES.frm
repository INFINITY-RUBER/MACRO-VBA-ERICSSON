VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NOMBRES 
   Caption         =   "UserForm1"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4380
   OleObjectBlob   =   "NOMBRES.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "NOMBRES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ToggleButton1_Click()

Dim numero_ot, Nuevo_nombre As String
   
  Ruta = GUARDAR_EXCEL
 
  numero_ot = OT
  Nuevo_nombre = SITIO
  
 Name Dir(Ruta & "\*EX*.xlsm") As Dir(Ruta & "\MP_EX_& Nuevo_nombre & _ & numero_ot.xlsm")
 NOMBRES.Hide
 MACRO_MPA.Hide
 
End Sub
