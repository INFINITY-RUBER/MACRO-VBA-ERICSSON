VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MACRO_MPA 
   Caption         =   "UserForm1"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9210
   OleObjectBlob   =   "MACRO_MPA.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "MACRO_MPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'********************************Desarrollado por  RUBER HERNANDEZ*********************************************
Dim TanqCombustible, procesados, Motor, BatArranque, Generador, Transferencia, Power, BcoBaterias, UPS As Integer
Dim Inversor, TabElectrico, Medicion, Protecciones, AcomPpal, RedComercial, Fusible, IntBaja, IntMedia, Pararrayo As Integer
Dim Seccionador, Tierra, Trafo, BcoCondensad, CeldaTransfMT, AA, Chiller, UMA, UdadCondensad, MotBomba, SistSupervision As Integer
Dim Servidor, ConcentDisposi, RegSolar, PanelSol As Integer
Dim Codigo, Tipo_red, Ciudad, Departamento, Dirección, Tipo, coordenadas, Responsable, Numero, Correo, fecha As String
Dim REGIONAL, RMC, NOMBRE_SITIO, salida, origen, HOJA_SALIDA, hoja_origen, Celda1, Celda2, Ruta As String
Dim equipo(100), trabajo, ruta_salida, VERSION, PAIS, LIBRE, VALVULA, CAPACIDAD As String
Dim Dato_equipo(100), atributo(100) As String
Dim i, j, N_atributo, X, PEE, conta, condicion, N_columna, TOTAL, CANT, cuentahojas As Integer

Private Sub Combo_Change()
Text_OPERADOR = Combo
End Sub

Private Sub Combo2_Change()
Text_REGION = Combo2
End Sub

Private Sub Combo3_Change()
Text_DEPARTAMENTO = Combo3
End Sub

'************************** MACRO MPA V18_5  MODIFICACION DE ARCHIVAR SI ES VERSION VIEJA Y QUITA LOS /  DEL NOMBRE
' **************************7-06-2018  se agrega ademas de la zona se agrega fecha y rmc
Private Sub UserForm_Activate()
Combo.AddItem ("MOVIL")
Combo.AddItem ("FIJA")

Combo2.AddItem ("BOGOTA")
Combo2.AddItem ("CUNDINAMARCA")
Combo2.AddItem ("SUROCCIDENTE")
Combo2.AddItem ("NOROCCIDENTE")
Combo2.AddItem ("SURORIENTE")

Combo3.AddItem ("AMAZONAS")
Combo3.AddItem ("ANTIOQUIA")
Combo3.AddItem ("BOGOTA")
Combo3.AddItem ("BOYACA")
Combo3.AddItem ("CALDAS")
Combo3.AddItem ("CAQUETA")
Combo3.AddItem ("CAUCA")
Combo3.AddItem ("CHOCO")
Combo3.AddItem ("CUNDINAMARCA")
Combo3.AddItem ("GUAINIA")
Combo3.AddItem ("GUAVIARE")
Combo3.AddItem ("HUILA")
Combo3.AddItem ("META")
Combo3.AddItem ("NARIÑO")
Combo3.AddItem ("PUTUMAYO")
Combo3.AddItem ("QUINDIO")
Combo3.AddItem ("RISARALDA")
Combo3.AddItem ("SANTANDER")
Combo3.AddItem ("TOLIMA")
Combo3.AddItem ("VALLE DEL CAUCA")
Combo3.AddItem ("VAUPES")
Combo3.AddItem ("VICHADA")

End Sub
Private Sub UP_Click()
Dim id, nombre, region, depto, municipio, operador, coordenadas, direcion As String

On Error Resume Next

   Sheets(1).Select
    For Each ws In ActiveWorkbook.Worksheets 'MOSTRAS HOJAS OCULTAS
        ws.Visible = xlSheetVisible
    Next ws

Text_OPERADOR = Combo

id = Text_ID
nombre = Text_NOMBRE
region = Text_REGION
depto = Text_DEPARTAMENTO
municipio = Text_MUNICIPIO
operador = Text_OPERADOR
direcion = Text_DIRECCION
coordenadas = Text_COORDENADAS

If id <> "" And nombre <> "" And region <> "" And depto <> "" And municipio <> "" And operador <> "" And coordenadas <> "" Then
 ActiveWorkbook.Worksheets("SITIOS").Select
 Range("A2").Select
 ActiveCell.Value = id
 Range("B2").Select
 ActiveCell.Value = id & " " & nombre
 Range("C2").Select
 ActiveCell.Value = region
 Range("D2").Select
 ActiveCell.Value = depto
 Range("E2").Select
 ActiveCell.Value = municipio
 Range("F2").Select
 ActiveCell.Value = nombre
 Range("G2").Select
 ActiveCell.Value = id
 Range("H2").Select
 ActiveCell.Value = operador
 Range("I2").Select
 ActiveCell.Value = direcion
 Range("J2").Select
 ActiveCell.Value = coordenadas
 
 Range("AE2").Select
 ActiveCell.Value = id & " " & nombre
 MsgBox "DATOS CARGADOS  "
Else
   MsgBox "INGRESE TODOS LOS DATOS "
End If
MACRO_MPA.Hide
End Sub


Sub Macro_INVENTARIOS_2017()
  cuentahojas = 1
   
  TOTAL = 0
  Text1 = cuentahojas
 avance = 0
UpdateProgressBar avance
 Application.ScreenUpdating = False
  
 hoja_origen = "INVENTARIO"
  
 Ruta = GUARDAR_EXCEL
 ChDir Ruta
 ruta_salida = Ruta
 salida = Dir(ruta_salida & "\Plantilla_Electromecanico_Final_v2_Vacio.xls")
 ChDir ruta_salida
   Workbooks.Open Filename:=salida, UpdateLinks:=0
  
 ChDir Ruta
 trabajo = Dir(Ruta & "\INVENT*.xlsx")
  origen = trabajo

While trabajo <> ""
 ChDir Ruta
   Workbooks.Open Filename:=trabajo, UpdateLinks:=0
    'ChDir ruta_salida   ---------------------CAMBIO NUEVA MACRO--
     'Workbooks.Open Filename:=salida
   origen = trabajo
   COMPROBAR_SI '--------------------------------------COMPROBAR SI LA VERSION ACTUAL
   
If TOTAL = 0 Then  '--------------------------ingresa si es la version de formato de inventario es la actual
   
   Datos  '---------- CARGA LOS DATOS BASICOS DEL SITIO
   Datos_Sitio
   METODOS
   
   carpeta = "INVENTARIOS REVISADOS"
   X = Dir(carpeta, vbDirectory)
   If X = "" Then
   '**************************Comprueba que la carpeta no exista para crear el directorio.**************************************
   MkDir (carpeta)
  End If
  ruta1 = Ruta & "\" & carpeta & "\"
  ActiveWorkbook.SaveAs Filename:=ruta1 & trabajo, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
  Windows(trabajo).Activate
  
   Application.DisplayAlerts = False
   
        ActiveWorkbook.Close
        Kill (trabajo)
        trabajo = Dir(Ruta & "\INVENTA*.xlsx")
        Application.DisplayAlerts = True
        'trabajo = Dir
  Windows(salida).Activate
  ActiveWorkbook.Save
End If
   
If TOTAL <> 0 Then
      Windows(origen).Activate
      ActiveWorkbook.Worksheets("INVENTARIO").Select
      
     carpeta = "INVENTARIO VERSION VIEJA"
      X = Dir(carpeta, vbDirectory)
      If X = "" Then
      MkDir (carpeta)
      End If
      ruta1 = Ruta & "\" & carpeta & "\"
     ActiveWorkbook.SaveAs Filename:=ruta1 & trabajo, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        'On Error Resume Next
        'ActiveWorkbook.Save
        ActiveWorkbook.Close
          Kill (trabajo)
      trabajo = Dir(Ruta & "\INVENTA*.xlsx")
    
End If '---------------------------------------------FIN DE COMPROBAR
   cuentahojas = cuentahojas + 1
   
   CONTADOR_LABEL = cuentahojas
   Text1 = cuentahojas
 avance = avance + 0.01
UpdateProgressBar avance
  
Wend
Application.ScreenUpdating = True
    MsgBox " PROCESO TERMINADO Y SE ENCONTRO   " & cuentahojas & " INVENTARIOS EN VERSION NO ACTUALIZADA"
 MACRO_MPA.Hide
 

End Sub

Sub COMPROBAR_SI() ' ************************************COMPROBAR LA VERSION
    Windows(origen).Activate
    ActiveWorkbook.Worksheets("INVENTARIO").Select
    
    Range("I1").Select '--------VERSION
    
   VERSION = ActiveCell.Value
If VERSION = "Version 3.2.5" Then
                         
          TOTAL = 0
End If
           
If VERSION <> "Version 3.2.5" Then
                         
   TOTAL = 1
   CANT = CANT + 1
   
End If
         


End Sub


Sub METODOS()
'------------------------------------------------------------- inicio metodos de pegado hojas----------------
If PEE <> 0 Then
  PEE_
End If  '----
If TanqCombustible <> 0 Then
  TANQUE_
End If
  
If Motor <> 0 Then
  MOTOR_
End If

If BatArranque <> 0 Then
  BATARRANQUE_
End If

   '
If Generador <> 0 Then
  GENERADOR_
End If
   '
If Transferencia <> 0 Then
  TRANSFERENCIA_
End If
'POWER_
If Power <> 0 Then
  POWER_
End If
   'BCOBATERIAS_
If BcoBaterias <> 0 Then
  BCOBATERIAS_
End If
   'UP_S
If UPS <> 0 Then
  UP_S
End If
 'INVERSOR_
If Inversor <> 0 Then
  INVERSOR_
End If
   'MEDICION_
If Medicion <> 0 Then
  MEDICION_
End If
   'TABELECTRICO_
If TabElectrico <> 0 Then
  TABELECTRICO_
End If
   'PROTECCIONES_
If Protecciones <> 0 Then
  PROTECCIONES_
End If
'ACOMPPAL_
If AcomPpal <> 0 Then
  ACOMPPAL_
End If
   'REDCOMERCIAL_
If RedComercial <> 0 Then
  REDCOMERCIAL_
End If
   'FUSIBLE_
If Fusible <> 0 Then
  FUSIBLE_
End If
   'INTBAJA_
If IntBaja <> 0 Then
  INTBAJA_
End If
   'INTMEDIA_
If IntMedia <> 0 Then
  INTMEDIA_
End If
   'PARARRAYO_
If Pararrayo <> 0 Then
  PARARRAYO_
End If
   'SECCIONADOR_
If Seccionador <> 0 Then
  SECCIONADOR_
End If
   'TIERRA_
If Tierra <> 0 Then
  TIERRA_
End If
   'TRAFO_
If Trafo <> 0 Then
  TRAFO_
End If
   'BCO_CONDENSAD_
If BcoCondensad <> 0 Then
  BCO_CONDENSAD_
End If
   'CELDA_TRANSF_
If CeldaTransfMT <> 0 Then
  CELDA_TRANSF_
End If
   'AA_
If AA <> 0 Then
  AA_
End If
   'CHILLER_
If Chiller <> 0 Then
  CHILLER_
End If
   'UMA_
If UMA <> 0 Then
  UMA_
End If
   'UDAD_CODENSAD_
If UdadCondensad <> 0 Then
  UDAD_CODENSAD_
End If
   'MOT_BOMBA_
If MotBomba <> 0 Then
  MOT_BOMBA_
End If
   'SIST_SUPERVISION_
If SistSupervision <> 0 Then
  SIST_SUPERVISION_
End If
    'SERVIDOR_
If Servidor <> 0 Then
  SERVIDOR_
End If
   'CONCET_DISPOSI_
If ConcentDisposi <> 0 Then
  CONCET_DISPOSI_
End If
    'PANEL_SOLAR_
If PanelSol <> 0 Then
  PANEL_SOLAR_
End If
    'REG_SOLAR_
If RegSolar <> 0 Then
  REG_SOLAR_
End If
'-------------------------------------------------------------metodos de pegado hojas----------------
End Sub
Sub Datos_Sitio()
 HOJA_SALIDA = "Datos del Sitio"
 Windows(salida).Activate
 ActiveWorkbook.Worksheets(HOJA_SALIDA).Select
 'Range("A4").Select ------------------------------------------------------------CAMBIO DE NUEVO INVENTARIO
 
  With ActiveSheet
    lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
  End With
    
 Fila = lastRow + 1
 
 
 Cells(Fila, "A").Select
 '---------------------------------------------------------------------------------------------------------------
  
 ActiveCell.Value = Codigo
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = NOMBRE_SITIO
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = Tipo_red
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = Ciudad
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = Departamento
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = Dirección
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = Tipo
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = coordenadas
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = Responsable
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = Numero
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = Correo
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = fecha
 ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.Value = REGIONAL
 
End Sub
Sub REG_SOLAR_()
      condicion = RegSolar
      HOJA_SALIDA = "RegSolar_" '----------- CAMBIAR HOJA SALIDA
    
      Celda2 = "I858"
      N_atributo = 31 '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA
End Sub
Sub PANEL_SOLAR_()
      condicion = PanelSol
      HOJA_SALIDA = "PanelSol_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I829"
      N_atributo = 28 '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA

End Sub
Sub CONCET_DISPOSI_()
      condicion = ConcentDisposi
      HOJA_SALIDA = "ConcentDisposi_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I813"
      N_atributo = 15 '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA
End Sub
Sub SERVIDOR_()
      condicion = Servidor
      HOJA_SALIDA = "Servidor_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I780"
      N_atributo = 32 '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA
End Sub
Sub SIST_SUPERVISION_()
      condicion = SistSupervision
      HOJA_SALIDA = "SistSupervision_" '----------- CAMBIAR HOJA SALIDA
     
      Celda2 = "I765"
      N_atributo = 14 '----------- mirar el numero de celda E ATRIBUTOS
 
      PEGAR_SALIDA
End Sub
Sub MOT_BOMBA_()
      condicion = MotBomba
      HOJA_SALIDA = "MotBomba_" '----------- CAMBIAR HOJA SALIDA
  
      Celda2 = "I734"
      N_atributo = 30 '----------- mirar el numero de celda E ATRIBUTOS
     
      PEGAR_SALIDA
End Sub
Sub UDAD_CODENSAD_()
      condicion = UdadCondensad
      HOJA_SALIDA = "UdadCondensad_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I698"
      N_atributo = 35 '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA

End Sub
Sub UMA_()
      condicion = UMA
      HOJA_SALIDA = "UMA_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I651"
      N_atributo = 46  '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA
End Sub
Sub CHILLER_()
      condicion = Chiller
      HOJA_SALIDA = "Chiller_" '----------- CAMBIAR HOJA SALIDA
     
      Celda2 = "I596"
      N_atributo = 54  '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA
End Sub
Sub AA_()
      condicion = AA
      HOJA_SALIDA = "AA_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I558"
      N_atributo = 37  '----------- mirar el numero de celda E ATRIBUTOS
     
      PEGAR_SALIDA

End Sub
Sub CELDA_TRANSF_()
      condicion = CeldaTransfMT
      HOJA_SALIDA = "CeldaTransfMT_" '----------- CAMBIAR HOJA SALIDA
     
      Celda2 = "I541"
      N_atributo = 16 '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA
End Sub
Sub BCO_CONDENSAD_()

      condicion = BcoCondensad
      HOJA_SALIDA = "BcoCondensad_" '----------- CAMBIAR HOJA SALIDA
     
      Celda2 = "I523"
      N_atributo = 17  '----------- mirar el numero de celda E ATRIBUTOS
    
      PEGAR_SALIDA

End Sub
Sub TRAFO_()

      condicion = Trafo
      HOJA_SALIDA = "Trafo_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I498"
      
      N_atributo = 24 '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA
      
End Sub
Sub TIERRA_()

      condicion = Tierra
      HOJA_SALIDA = "Tierra_" '----------- CAMBIAR HOJA SALIDA
     
      Celda2 = "I481"
      N_atributo = 16 '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA

End Sub
Sub SECCIONADOR_()

      condicion = Seccionador
      HOJA_SALIDA = "Seccionador_" '----------- CAMBIAR HOJA SALIDA
    
      Celda2 = "I463"
      N_atributo = 17  '----------- mirar el numero de celda E ATRIBUTOS
     
      PEGAR_SALIDA
End Sub
Sub PARARRAYO_()

      condicion = Pararrayo
      HOJA_SALIDA = "Pararrayo_" '----------- CAMBIAR HOJA SALIDA
    
      Celda2 = "I447"
      N_atributo = 15  '----------- mirar el numero de celda E ATRIBUTOS
     
      PEGAR_SALIDA

End Sub
Sub INTMEDIA_()
      condicion = IntMedia
      HOJA_SALIDA = "IntMedia_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I429"
      N_atributo = 17  '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA

End Sub
Sub INTBAJA_()
      condicion = IntBaja
      HOJA_SALIDA = "IntBaja_" '----------- CAMBIAR HOJA SALIDA
     
      Celda2 = "I413"
      N_atributo = 15  '----------- mirar el numero de celda E ATRIBUTOS
    
      PEGAR_SALIDA

End Sub
Sub FUSIBLE_()
      condicion = Fusible
      HOJA_SALIDA = "Fusible_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I398"
      N_atributo = 14 '----------- mirar el numero de celda E ATRIBUTOS
     
      PEGAR_SALIDA

End Sub
Sub REDCOMERCIAL_()
      condicion = RedComercial
      HOJA_SALIDA = "RedComercial_" '----------- CAMBIAR HOJA SALIDA
     
      Celda2 = "I380"
      N_atributo = 17 '----------- mirar el numero de celda E ATRIBUTOS
    
      PEGAR_SALIDA

End Sub
Sub ACOMPPAL_()
      condicion = AcomPpal
      HOJA_SALIDA = "AcomPpal_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I359"
      N_atributo = 20  '----------- mirar el numero de celda E ATRIBUTOS
     
      PEGAR_SALIDA

End Sub
Sub PROTECCIONES_()
      condicion = Protecciones
      HOJA_SALIDA = "Protecciones_" '----------- CAMBIAR HOJA SALIDA
     
      Celda2 = "I340"
      N_atributo = 18  '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA

End Sub
Sub TABELECTRICO_()
      condicion = TabElectrico
      HOJA_SALIDA = "TabElectrico_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I316"
      N_atributo = 23  '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA


End Sub
Sub MEDICION_()
      condicion = Medicion
      HOJA_SALIDA = "Medicion_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I301"
      N_atributo = 14  '----------- mirar el numero de celda E ATRIBUTOS
     
      PEGAR_SALIDA

End Sub
Sub INVERSOR_()
      condicion = Inversor
      HOJA_SALIDA = "Inversor_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I276"
      N_atributo = 24 '----------- mirar el numero de celda E ATRIBUTOS
     
      PEGAR_SALIDA

End Sub
Sub UP_S()

      condicion = UPS
      HOJA_SALIDA = "UPS_" '----------- CAMBIAR HOJA SALIDA
     
      Celda2 = "I246"
      N_atributo = 29  '----------- mirar el numero de celda E ATRIBUTOS
     
      PEGAR_SALIDA

End Sub
Sub BCOBATERIAS_()
      condicion = BcoBaterias
      HOJA_SALIDA = "BcoBaterias_" '----------- CAMBIAR HOJA SALIDA
     
      Celda2 = "I215"
      N_atributo = 30  '----------- mirar el numero de celda E ATRIBUTOS
    
      PEGAR_SALIDA

End Sub
Sub POWER_()
      condicion = Power
      HOJA_SALIDA = "Power_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I170"
      N_atributo = 44  '----------- mirar el numero de celda E ATRIBUTOS
    
      PEGAR_SALIDA
      
      
End Sub
Sub PEGAR_SALIDA()
   
If condicion <> 0 Then ' ---------DATOS DE POGER_---------------------------------------------------------
      Windows(origen).Activate
      ActiveWorkbook.Worksheets(hoja_origen).Select
      '  Range(Celda1).Select  '----------- CAMBIAR
      '  TOMA_ATRIBUTO
           
      Range(Celda2).Select '----------- CAMBIAR
      
      
 For i = 1 To condicion '----------- mirar el numero de ELEMENTOS
      
      For j = 0 To N_atributo
      Dato_equipo(j) = ActiveCell.Value
      ActiveCell.Offset(1, 0).Range("A1").Select
      Next j
      ActiveCell.Offset(-j, 0).Range("A1").Select
      
      '*************************************************MODIFICACION 12-04-2018**************
      If (HOJA_SALIDA = "Trafo_") Then
         Range("I890").Select
         PAIS = ActiveCell.Value
         ActiveCell.Offset(1, 0).Range("A1").Select
         LIBRE = ActiveCell.Value
         ActiveCell.Offset(1, 0).Range("A1").Select
         VALVULA = ActiveCell.Value
      End If
      
       '*******************************************************
  Windows(salida).Activate
  ActiveWorkbook.Worksheets(HOJA_SALIDA).Select
      X = 0
      
         '----------- ----------------------------------------------CAMBIO INVENTARIO NUEVO---------------------
        With ActiveSheet
            lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        End With
    
              Fila = lastRow + 1
            If (Fila <= 4) Then
               Fila = 5
            End If
 
            Cells(Fila, "B").Select
            ' Range("C3").Select
             ' conta = Fila - 3
      '-----------------------------MODIFICACION DICIEMBRE NUEVO INVENTARIO 2017 ---------------------------------------------------------------------------------------------
      For j = 0 To N_atributo
      ActiveCell.Value = Dato_equipo(j)
      ActiveCell.Offset(0, 1).Range("A1").Select
      Next j
      ActiveCell.Value = REGIONAL
      ' ***********agregado  8-03-2018 inicio--------------------
      ActiveCell.Offset(0, 1).Range("A1").Select
      ActiveCell.Value = RMC
      ActiveCell.Offset(0, 1).Range("A1").Select
      ActiveCell.Value = fecha
      j = j + 2
      ' ***********fin------------------------------
      ActiveCell.Offset(0, -j).Range("A1").Select
    
     '---------------------------------------------------------------------------------------------
     '-While (ActiveCell.Column <= N_columna) '----------- mirar el numero de columnas
      
      '-      If ActiveCell.Value = atributo(x) Then
     '-          If "Nombre del Sitio en Maximo" <> atributo(x) Then
     '-             ActiveCell.Offset(conta, 0).Range("A1").Select
    '-              ActiveCell.Value = Dato_equipo(x)
     '-             ActiveCell.Offset(-conta, 0).Range("A1").Select
     '-          End If
               
    '-       x = x + 1
   '-      End If
    '-    ActiveCell.Offset(0, 1).Range("A1").Select
   '-  Wend
   '  conta = conta + 1
    
     Windows(origen).Activate
     ActiveWorkbook.Worksheets(hoja_origen).Select
     ActiveCell.Offset(0, 1).Range("A1").Select
 Next i           '---------fin  el numero de ELEMENTOS------------
End If '-------------------------------------------fin -

End Sub

Sub PEE_()

      condicion = PEE
      HOJA_SALIDA = "PEE_" '----------- CAMBIAR HOJA SALIDA
      'Celda1 = "E25"
      Celda2 = "I25"
      N_atributo = 38 '----------- mirar el numero de celda E ATRIBUTOS
      ' N_columna = 39
      PEGAR_SALIDA


End Sub

Sub TRANSFERENCIA_()

      condicion = Transferencia
      HOJA_SALIDA = "Transferencia_" '----------- CAMBIAR HOJA SALIDA
     
      Celda2 = "I144"
      N_atributo = 25   '----------- mirar el numero de celda E ATRIBUTOS
     
      PEGAR_SALIDA
                   '---------fin -TRANSFERENCIA_

End Sub

Sub GENERADOR_()

      condicion = Generador
      HOJA_SALIDA = "Generador_" '----------- CAMBIAR HOJA SALIDA
    
      Celda2 = "I110"
      N_atributo = 15   '----------- mirar el numero de celda E ATRIBUTOS
   
      PEGAR_SALIDA
                   '---------fin -TRANSFERENCIA_

End Sub
Sub BATARRANQUE_()

      condicion = BatArranque
      HOJA_SALIDA = "BatArranque_" '----------- CAMBIAR HOJA SALIDA
      
      Celda2 = "I126"
      N_atributo = 17  '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA

End Sub

Sub MOTOR_()

      condicion = Motor
      HOJA_SALIDA = "Motor_" '----------- CAMBIAR HOJA SALIDA
      'Celda1 = "E51"
      Celda2 = "I85"
      N_atributo = 24  '----------- mirar el numero de celda E ATRIBUTOS
      'N_columna = 25
      PEGAR_SALIDA

End Sub

Sub TANQUE_()

      condicion = TanqCombustible
      HOJA_SALIDA = "TanqCombustible_" '----------- CAMBIAR HOJA SALIDA
      'Celda1 = "E42"
      Celda2 = "I64"
      N_atributo = 20  '----------- mirar el numero de celda E ATRIBUTOS
      
      PEGAR_SALIDA

End Sub


Sub TOMA_ATRIBUTO()

     For j = 0 To N_atributo
       atributo(j) = ActiveCell.Value
       ActiveCell.Offset(1, 0).Range("A1").Select
       Next j
     ActiveCell.Offset(-j, 0).Range("A1").Select

End Sub

Sub Datos()
'TOMA LOS DATOS GENERALES
     Windows(origen).Activate
     ActiveWorkbook.Worksheets("INVENTARIO").Select
     
     
    '---------------------DATOS INICIALES -------------
    Range("D4").Select
     PEE = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     TanqCombustible = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Motor = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     BatArranque = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Generador = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Transferencia = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Power = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     BcoBaterias = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     UPS = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Inversor = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
    Medicion = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
    TabElectrico = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Protecciones = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     AcomPpal = ActiveCell.Value
     
    ActiveCell.Offset(1, 0).Range("A1").Select
     RedComercial = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Fusible = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     IntBaja = ActiveCell.Value
     
   Range("H4").Select '-----------------SEGUNDA COLUMNA DATOS ENTEROS
     IntMedia = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Pararrayo = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Seccionador = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Tierra = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Trafo = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     BcoCondensad = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     CeldaTransfMT = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     AA = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Chiller = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     UMA = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     UdadCondensad = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     MotBomba = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     SistSupervision = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     Servidor = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     ConcentDisposi = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     PanelSol = ActiveCell.Value
    ActiveCell.Offset(1, 0).Range("A1").Select
     RegSolar = ActiveCell.Value
     '--------------------------------------------------PRIMERA----INFORMACION DEL SITIO
     Range("K7").Select
     Codigo = ActiveCell.Value
      ActiveCell.Offset(1, 0).Range("A1").Select
     Tipo_red = ActiveCell.Value
      ActiveCell.Offset(1, 0).Range("A1").Select
     Ciudad = ActiveCell.Value
      ActiveCell.Offset(1, 0).Range("A1").Select
     Departamento = ActiveCell.Value
      ActiveCell.Offset(1, 0).Range("A1").Select
     Dirección = ActiveCell.Value
      ActiveCell.Offset(1, 0).Range("A1").Select
     REGIONAL = ActiveCell.Value
     '------------------------------------------------SEGUENDA -INFORMACION DEL SITIO
     Range("M7").Select
               
     NOMBRE_SITIO = ActiveCell.Value
      ActiveCell.Offset(1, 0).Range("A1").Select
     coordenadas = ActiveCell.Value
      ActiveCell.Offset(1, 0).Range("A1").Select
     Responsable = ActiveCell.Value
      ActiveCell.Offset(1, 0).Range("A1").Select
     Numero = ActiveCell.Value
      ActiveCell.Offset(1, 0).Range("A1").Select
     Correo = ActiveCell.Value
          
     Range("O2").Select
     RMC = ActiveCell.Value
     ActiveCell.Offset(1, 0).Range("A1").Select
     fecha = ActiveCell.Value
    
    
      
     
     
    
     
 

End Sub


Private Sub BOTON_1_Click() ' TOMA DATOS BOGOTA

Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
    Range("C50:L50").Select
    Selection.Copy
    Windows("CONTROL_MP_BOGOTA_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("A19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
   Range("R36:Y36").Select
    Selection.Copy
    Windows("CONTROL_MP_BOGOTA_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("D19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    Range("D21").Select
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
    Range("C8").Select
    Selection.Copy
    Windows("CONTROL_MP_BOGOTA_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("B19").Select
    ActiveSheet.Paste
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("PLANTAS").Select
    Range("E30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_BOGOTA_2016.xlsx").Activate
     Sheets("RESUMEN").Select
    Range("C19").Select
    ActiveSheet.Paste
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("TANQUE").Select
    Range("C14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_BOGOTA_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("E19").Select
    ActiveSheet.Paste
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("BATERIAS").Select
 
    Range("O12:P12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_BOGOTA_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("F19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("F19").Select

    MACRO_MPA.Hide
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Application.Dialogs(xlDialogSaveAs).Show
    ActiveWorkbook.SaveAs Filename:="RED-FT-0104020106 Preventivo EX"


End Sub



Private Sub BOTON_2_Click()

    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
    Range("C50:L50").Select
    Selection.Copy
    Windows("CONTROL_MP_NOROCCIDENTE-SURORIENTE_2017.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("A19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
   Range("R36:Y36").Select
    Selection.Copy
    Windows("CONTROL_MP_NOROCCIDENTE-SURORIENTE_2017.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("D19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    Range("D21").Select
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
    Range("C8").Select
    Selection.Copy
    Windows("CONTROL_MP_NOROCCIDENTE-SURORIENTE_2017.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("B19").Select
    ActiveSheet.Paste
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("PLANTAS").Select
    Range("E30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_NOROCCIDENTE-SURORIENTE_2017.xlsx").Activate
     Sheets("RESUMEN").Select
    Range("C19").Select
    ActiveSheet.Paste
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("TANQUE").Select
    Range("C14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_NOROCCIDENTE-SURORIENTE_2017.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("E19").Select
    ActiveSheet.Paste
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("BATERIAS").Select
 
    Range("O12:P12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_NOROCCIDENTE-SURORIENTE_2017.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("F19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("F19").Select
    
    MACRO_MPA.Hide
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Application.Dialogs(xlDialogSaveAs).Show
    ActiveWorkbook.SaveAs Filename:="RED-FT-0104020106 Preventivo EX"

End Sub



Private Sub CommandButton4_Click()
Dim dato As String

On Error Resume Next

   Sheets(1).Select
For Each ws In ActiveWorkbook.Worksheets 'MOSTRAS HOJAS OCULTAS
ws.Visible = xlSheetVisible
Next ws
On Error Resume Next

'Worksheets(2).Hoja2.Range("E10").Value
' Region1 = Trim(Cells(Fila_o, 2))
 Sheets("SUBESTACION").Select
 Text1 = Trim(Cells(8, 10)) 'FECHA
 Text5 = Trim(Cells(9, 20)) ' CONTADOR
 Text3 = Trim(Cells(112, 3))
 Sheets("TABLERO AC-DC").Select
 Text2 = Trim(Cells(8, 36))
 RMC2 = Trim(Cells(73, 2))
 Sheets("PLANTAS").Select
 Text4 = Trim(Cells(30, 7))
  Sheets("TANQUE").Select
 Text6 = Trim(Cells(18, 3))
 Sheets("BATERIAS").Select
 Text7 = Trim(Cells(12, 15))
 Text8 = Trim(Cells(135, 15))
 Text9 = Trim(Cells(259, 15))

End Sub

Private Sub BOTON_3_Click()

    'Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
    Range("J8:N8").Select
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2017.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("A19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
    Range("R36:Y36").Select
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("D19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    Range("D21").Select
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
    Range("C8").Select
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("B19").Select
    ActiveSheet.Paste
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("PLANTAS").Select
    Range("E30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2016.xlsx").Activate
     Sheets("RESUMEN").Select
    Range("C19").Select
    ActiveSheet.Paste
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("TANQUE").Select
    Range("C14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("E19").Select
    ActiveSheet.Paste
    On Error Resume Next
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("BATERIAS").Select
    Range("O12:P12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("F19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    Range("F19").Select
    MACRO_MPA.Hide
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Application.Dialogs(xlDialogSaveAs).Show
    ActiveWorkbook.SaveAs Filename:="RED-FT-0104020106 Preventivo EX"
    
End Sub
' --------------------UNIR HOJAS RF------------------

Private Sub CheckBox1_Click()
 Call quitaNombres
 On Error Resume Next
          Windows("PREVENTIVO TX-1.xlsx").Activate
          hojas = Sheets.Count
         Windows("FORMATO TX.xlsx").Activate
          foto = Sheets.Count
       For i = 1 To foto
       Workbooks("FORMATO TX.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO TX-1").Sheets(hojas)
       hojas = hojas + 1
       Next i
       Windows("PREVENTIVO TX-1.xlsx").Activate
        ActiveWorkbook.Save
        MsgBox " SE A MOVIDO  " & foto & " HOJAS TX"

End Sub
' GUARDAR EXCEL
Private Sub CommandButton5_Click()
Dim V As Double
Dim a As Double
Dim r As Double

Dim RECT As Double
Dim V_Igualacion As Double
Dim A_Igualacion As Double
  
 
  
If rectificador1.Value = True Then ' PRIMER RECTIFICADOR 1
 
    ActiveWorkbook.Worksheets("RECTIFICADOR").Select
    Range("AK9").Select
    V = ActiveCell.Value
    Range("AK10").Select
    a = ActiveCell.Value
    Range("AS9").Select
    r = ActiveCell.Value
    Flota = a / r
    V_Igualacion = V * 1.05
    A_Igualacion = Flota * 1.09
 
    If r > 10 Then
       r = 10
    
    End If
    
   If r = 1 Then
   
    Range("I15").Select
    ActiveCell.FormulaR1C1 = V
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = V_Igualacion
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = Flota
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = A_Igualacion
    ActiveWorkbook.Worksheets("RECTIFICADOR").Select
    
     
     ElseIf r > 1 Then ' MAYOR A UN MODULO
     
     
     Range("I15").Select
    
     For i = 1 To r  ' VOLTAGE
     ActiveCell.FormulaR1C1 = V
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
     
     Range("I15").Select  ' VOLTAGE FLOTACIN
     ActiveCell.Offset(1, 0).Range("A1").Select
     For i = 1 To r
     ActiveCell.FormulaR1C1 = V_Igualacion
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
     
     Range("I15").Select   ' CORRIENTE
     ActiveCell.Offset(2, 0).Range("A1").Select
     For i = 1 To r
     ActiveCell.FormulaR1C1 = Flota
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
     '---
     Range("I15").Select  ' A_Igualacion
     ActiveCell.Offset(3, 0).Range("A1").Select
     For i = 1 To r
     ActiveCell.FormulaR1C1 = A_Igualacion
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
    
    ActiveWorkbook.Worksheets("RECTIFICADOR").Select
     
     End If
    
End If   ' FIN RECTIFICADOR 1-------------------------------------------------------
     
If rectificador2.Value = True Then ' PRIMER RECTIFICADOR 2 ------------------------
 
    ActiveWorkbook.Worksheets("RECTIFICADOR").Select
    Range("AK85").Select
    V = ActiveCell.Value
    Range("AK86").Select
    a = ActiveCell.Value
    Range("AS85").Select
    r = ActiveCell.Value
    Flota = a / r
    V_Igualacion = V * 1.05
    A_Igualacion = Flota * 1.09
 
    If r > 10 Then
    
    r = 10
    
    End If
    
   If r = 1 Then
   
    Range("I91").Select
    ActiveCell.FormulaR1C1 = V
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = V_Igualacion
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = Flota
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = A_Igualacion
    ActiveWorkbook.Worksheets("RECTIFICADOR").Select
    
     
     ElseIf r > 1 Then ' MAYOR A UN MODULO
     
     
     Range("I91").Select
    
     For i = 1 To r  ' VOLTAGE
     ActiveCell.FormulaR1C1 = V
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
     
     Range("I91").Select ' VOLTAGE FLOTACIN
     ActiveCell.Offset(1, 0).Range("A1").Select
     For i = 1 To r
     ActiveCell.FormulaR1C1 = V_Igualacion
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
     
     Range("I91").Select   ' CORRIENTE
     ActiveCell.Offset(2, 0).Range("A1").Select
     For i = 1 To r
     ActiveCell.FormulaR1C1 = Flota
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
     '---
     Range("I91").Select ' A_Igualacion
     ActiveCell.Offset(3, 0).Range("A1").Select
     For i = 1 To r
     ActiveCell.FormulaR1C1 = A_Igualacion
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
    
    ActiveWorkbook.Worksheets("RECTIFICADOR").Select
     
    End If '
    
End If   ' FIN RECTIFICADOR 2-----------------------------------------------
     
If rectificador3.Value = True Then ' PRIMER RECTIFICADOR 3 ------------------------
 
    ActiveWorkbook.Worksheets("RECTIFICADOR").Select
    Range("AK161").Select
    V = ActiveCell.Value
    Range("AK162").Select
    a = ActiveCell.Value
    Range("AS161").Select
    r = ActiveCell.Value
    Flota = a / r
    V_Igualacion = V * 1.07
    A_Igualacion = Flota * 1.09
 
    If r > 10 Then
    
    r = 10
    
    End If
    
   If r = 1 Then
   
    Range("I167").Select
    ActiveCell.FormulaR1C1 = V
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = V_Igualacion
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = Flota
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = A_Igualacion
    ActiveWorkbook.Worksheets("RECTIFICADOR").Select
    
     
     ElseIf r > 1 Then ' MAYOR A UN MODULO
     
     
     Range("I167").Select
    
     For i = 1 To r  ' VOLTAGE
     ActiveCell.FormulaR1C1 = V
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
     
     Range("I167").Select ' VOLTAGE FLOTACIN
     ActiveCell.Offset(1, 0).Range("A1").Select
     For i = 1 To r
     ActiveCell.FormulaR1C1 = V_Igualacion
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
     
     Range("I167").Select  ' CORRIENTE
     ActiveCell.Offset(2, 0).Range("A1").Select
     For i = 1 To r
     ActiveCell.FormulaR1C1 = Flota
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
     '---
     Range("I167").Select ' A_Igualacion
     ActiveCell.Offset(3, 0).Range("A1").Select
     For i = 1 To r
     ActiveCell.FormulaR1C1 = A_Igualacion
     ActiveCell.Offset(0, 1).Range("A1").Select
     Next i
    
    ActiveWorkbook.Worksheets("RECTIFICADOR").Select
     
    End If '
    
End If   ' FIN RECTIFICADOR 3-----------------------------------------------
     

  MsgBox "V " & V & Chr(13) + "A " & a & Chr(13) + "FLOTACION " & Flota & Chr(13) + "RECTIFICADOR " & r
  MACRO_MPA.Hide
End Sub

Private Sub CommandButton6_Click()

On Error Resume Next

   Sheets(1).Select
    For Each ws In ActiveWorkbook.Worksheets 'MOSTRAS HOJAS OCULTAS
        ws.Visible = xlSheetVisible
    Next ws

'Application.FileDialog(msoFileDialogOpen).Show
    
   MACRO_MPA.Hide
End Sub

Private Sub CommandButton7_Click() '  BOTON  DE SEPARACION DE HOJA INVENTARIOS
       Dim Ruta As String
       On Error Resume Next
       Ruta = GUARDAR_EXCEL
        If GUARDAR_EXCEL = "" Then
       Ruta = "C:\Users\ERICSSON\Downloads\DESCARGAS_MPA"
         End If
        GUARDAR_EXCEL = Ruta
       Sheets("INVENTARIO").Select
       Dim nombreHoja As String
       nombreHoja = "INVENTARIO"
                 'VALIDACION DE ERROR  EN LOS DATOS
        Range("K7").Select
        validar = ActiveCell.Value
         
         For i = 1 To 5
           If IsError(validar) Then
              MsgBox "EL INVENTARIO ESTA SIN INFORMACION POR FAVOR CORREGIR"
              Exit Sub
           End If
           ActiveCell.Offset(1, 0).Range("A1").Select
           validar = ActiveCell.Value
           
         Next
         
         
       
     
      If (BuscarHoja1(nombreHoja)) Then
          Sheets("INVENTARIO").Select
          Sheets("SITIOS").Visible = True
          Sheets(Array("INVENTARIO", "SITIOS")).Select
          Sheets(Array("INVENTARIO", "SITIOS")).Move
          ActiveWindow.SmallScroll Down:=3
          Sheets("INVENTARIO").Select
           Contraseña = "Inventario2018"
          ActiveWorkbook.Unprotect Contraseña
          ActiveSheet.Unprotect Contraseña
          Sheets("INVENTARIO").Select
          Range("J2").Select
          nombre = ActiveCell.Value
          texto = nombre
          Posicion = InStr(1, texto, "\")  'Busca la posición del primer caracter blanco
          nombre = Left(texto, Posicion - 1)
        

         'abre archivo de trabajo (donde se encuentran los datos)
         Range("B2:AF446").Select
         Selection.Copy
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Sheets("INVENTARIO").Select
        ActiveWorkbook.SaveAs Filename:=Ruta & "\INVENTARIOS_" & nombre & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        On Error Resume Next
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        
      ' Windows("EX CON INVENTAR*.xlsx").Activate
        'Sheets("SUBESTACION").Select
       ' Range("AH:AH").Select
       ' Selection.ClearContents
        
        'Sheets("TABLERO AC-DC").Select
        'Range("AP:AP").Select
        'Selection.ClearContents
        'Sheets("TIERRAS").Select
        'Range("T:T").Select
        'Selection.ClearContents
        'Sheets("PLANTAS").Select
       ' Range("AB:AB").Select
       ' Selection.ClearContents
       ' Sheets("TRANSFER").Select
       ' Range("AA:AA").Select
       ' Selection.ClearContents
       ' Sheets("RECTIFICADOR").Select
      '  Range("AU:AU").Select
       ' Selection.ClearContents
     End If
  
     MACRO_MPA.Hide
 
End Sub

Sub UpdateProgressBar(ava)
'Por.DAM
    MACRO_MPA.FProgress.Caption = Format(ava, "0%")
    MACRO_MPA.LProgress.Width = ava * (MACRO_MPA.FProgress.Width - 10)
    'UserForm1.FProgress.Caption = Format(ava, "0%")
    'UserForm1.LProgress.Width = ava * (UserForm1.FProgress.Width - 10)
    DoEvents
End Sub

Private Sub CONVERTIR_Click()

Dim Ruta As String
Dim trabajo As String
Dim trabajo1 As String
Dim texto As String
Dim inventario, nombre As String

MACRO_MPA.LProgress.Width = 0

avance = 0.1
UpdateProgressBar avance
Application.ScreenUpdating = False
Ruta = GUARDAR_EXCEL
    If GUARDAR_EXCEL = "" Then
       Ruta = "C:\Users\ERICSSON\Downloads\DESCARGAS_MPA"
    End If
GUARDAR_EXCEL = Ruta
TOTAL = 0

 
  ' ubica ruta en directorio de trabajo
    ChDir Ruta
    'se posiciona en directorio de trabajo
    trabajo = Dir(Ruta & "\MP_*.xlsm")
    inventario = Dir(Ruta & "\MP_EX*.xlsm")
    '// se hizo un cambio de codigo
While trabajo <> ""
             
     'Application.DisplayAlerts = False
        trabajo1 = trabajo
        texto = trabajo
        Posicion = InStr(1, texto, ".xlsm")  'Busca la posición del primer caracter blanco
        PrimeraPalabra = Left(texto, Posicion - 1)
        'abre archivo de trabajo (donde se encuentran los datos)
        ChDir Ruta
        Workbooks.Open Filename:=trabajo
      '---------MODIFICACION--------------------------------------
        If trabajo = inventario Then
            On Error Resume Next
            Dim nombreHoja As String
            nombreHoja = "INVENTARIO"
     
             If (BuscarHoja1(nombreHoja)) Then
                 Sheets("INVENTARIO").Select
                 
                 If Range("I1").Value = "Version 3.2.5" Then
                     
                        'Sheets("SITIOS").Visible = True
                         'Sheets(Array("INVENTARIO", "SITIOS")).Select
                        ' Sheets(Array("INVENTARIO", "SITIOS")).Move
                        ' ActiveWindow.SmallScroll Down:=3
                         'Sheets("INVENTARIO").Select
                      Contraseña = "Inventario2018"
                         'ActiveWorkbook.Unprotect Contraseña
                      ActiveSheet.Unprotect Contraseña
                      Sheets("INVENTARIO").Select
                      ActiveSheet.Range("$A$24:$S$889").AutoFilter Field:=3, Criteria1:=Array( _
                      "Mandatorio", "Mandatorio (Formulado)", "Mandatorio (Relacionado)"), Operator:= _
                      xlFilterValues
                      Range("D24").Select
                      ActiveSheet.Range("$A$24:$S$889").AutoFilter Field:=4, Criteria1:="SI"
                      Range("I26:I889").Select
                      Selection.SpecialCells(xlCellTypeBlanks).Select
                      
                         With Selection.Interior
                             .PatternColorIndex = xlAutomatic
                             .Color = 255
                         End With
                             '       Range("J2").Select
                             '       Nombre = ActiveCell.Value
                             '       Range("B2:AF446").Select
                             '       Selection.Copy
                             '          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                             '        :=False, Transpose:=False
                             '        Sheets("INVENTARIO").Select
                             '        ActiveWorkbook.SaveAs Filename:=ruta & "\INVENTARIOS_" & Nombre & ".xlsx", _
                             '        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                             '
                             '        origen = "INVENTARIOS_" & Nombre
                             ' COMPROBAR_SI '--------------------------------------COMPROBAR
                     End If
            
                         If TOTAL <> 0 Then
                         MsgBox " NO SE TIENE  EN   " & TOTAL & " EQUIPOS LAS CANTIDADES EN LA HOJA INVENTARIOS"
                       
                         'Exit Sub
                         End If '---------------------------------------------FIN DE COMPROBAR
                      
                     'ActiveWorkbook.Close
                     Else
                       MsgBox nombreHoja & " NO HAY HOJA INVENTARIOS"
             End If
   
        End If
      
        ' si los archivos tiene calificación del 100% solo se trae la información resumen.
        'ActiveWorkbook.Worksheets("Resultados de Auditoria").Select
        Windows(trabajo).Activate
        ActiveWorkbook.SaveAs Filename:=PrimeraPalabra & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        Sheets(1).Select
          'ActiveWindow.SmallScroll Down:=-9
    
        Application.DisplayAlerts = False
        Kill (trabajo1)
        trabajo = Dir(Ruta & "\MP_*.xlsm")
        avance = avance + 0.1
        UpdateProgressBar avance
       
  Wend
  
     avance = 1
     UpdateProgressBar avance
     Application.ScreenUpdating = True
     MACRO_MPA.Hide
     
End Sub

Function BuscarHoja1(nombreHoja As String) As Boolean
 
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = nombreHoja Then
            BuscarHoja1 = True
            Exit Function
        End If
    Next
     
    BuscarHoja1 = False
 
End Function

Private Sub FOTOS_EX_Click()

      On Error Resume Next
          Windows("PREVENTIVO EX-1.xlsx").Activate
          hojas = Sheets.Count
          Windows("EX FOTOS P1.xlsx").Activate
          foto = Sheets.Count
          
          For i = 1 To foto
          Windows("EX FOTOS P1.xlsx").Activate
          Worksheets(i).Select
          PESO = ActiveSheet.Shapes.Count
           If (PESO >= 1) Then
           ' Sheets(i).Tab.Color = vbGreen
           End If
          
          Next i
          
          
       For i = 1 To foto
          
         Workbooks("EX FOTOS P1.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(hojas)
         hojas = hojas + 1
       Next i
       
    On Error Resume Next
        
       Windows("EX FOTOS P2.xlsx").Activate
       foto2 = Sheets.Count
       
       For i = 1 To foto2
          Windows("EX FOTOS P2.xlsx").Activate
          Worksheets(i).Select
          PESO = ActiveSheet.Shapes.Count
           If (PESO >= 1) Then
             'Sheets(i).Tab.Color = vbGreen
           End If
          
          Next i
       
       For i = 1 To foto2
                   
         Workbooks("EX FOTOS P2.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(hojas)
         hojas = hojas + 1
       Next i

      
      MsgBox " SE HA ORDENADO Y GUARDADO   " & hojas & " HOJAS"
     
     MACRO_MPA.Hide
End Sub

Private Sub INVENTARIOS_Click()

Macro_INVENTARIOS_2017

End Sub



Private Sub nombre_formatos_Click()
 Dim numero_ot, Nuevo_nombre As String
   
      Ruta = GUARDAR_EXCEL
 
  numero_ot = InputBox("", "INGRESE NUMERO DE OT", "")
  Nuevo_nombre = InputBox("", "INGRESE NOMBRE DEL SITIO", "")
  
 Name Dir(Ruta & "\EX*.xlsm") As Dir(Ruta & "\MP_EX_& Nuevo_nombre & _ & numero_ot.xlsm")
 NOMBRES.Hide
 MACRO_MPA.Hide

End Sub

Private Sub ORDENAR_ZONA_Click()

    Ruta = GUARDAR_EXCEL
    trabajo = Dir(Ruta & "\INVENTA*.xlsx")
    
   If ZONA_INVENTARIOS.Value = False Then ' **********************************INICIO CHECKBOX
    
While trabajo <> ""
    ChDir Ruta
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=trabajo, UpdateLinks:=0
     'Workbooks.Open ThisWorkbook.Path & trabajo, UpdateLinks:=0
    
    origen = trabajo
    Windows(origen).Activate
    ActiveWorkbook.Worksheets("INVENTARIO").Select
    
    Range("K12").Select '--------CELDA DE LA REGIONAL
    salida = ActiveCell.Value
    
    
 If salida = "BOGOTA" Then
         
      carpeta = "BOGOTA"
      X = Dir(carpeta, vbDirectory)
      If X = "" Then
           MkDir (carpeta)
       End If
         
      ruta1 = Ruta & "\" & carpeta & "\"
     ActiveWorkbook.SaveAs Filename:=ruta1 & trabajo, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        'On Error Resume Next
        'ActiveWorkbook.Save
        ActiveWorkbook.Close
          Kill (trabajo)
 End If

If salida = "CUNDINAMARCA" Then
         
      carpeta = "CUNDINAMARCA"
      
      X = Dir(carpeta, vbDirectory)
      If X = "" Then
           MkDir (carpeta)
       End If
         
      ruta1 = Ruta & "\" & carpeta & "\"
     ActiveWorkbook.SaveAs Filename:=ruta1 & trabajo, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        'On Error Resume Next
        'ActiveWorkbook.Save
        ActiveWorkbook.Close
          Kill (trabajo)
End If

If salida = "NOROCCIDENTE" Then
         
      carpeta = "NOROCCIDENTE"
      
      X = Dir(carpeta, vbDirectory)
      If X = "" Then
           MkDir (carpeta)
       End If
         
      ruta1 = Ruta & "\" & carpeta & "\"
     ActiveWorkbook.SaveAs Filename:=ruta1 & trabajo, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        'On Error Resume Next
        'ActiveWorkbook.Save
        ActiveWorkbook.Close
          Kill (trabajo)
End If

If salida = "SUROCCIDENTE" Then
         
      carpeta = "SUROCCIDENTE"
      
      X = Dir(carpeta, vbDirectory)
      If X = "" Then
           MkDir (carpeta)
       End If
         
      ruta1 = Ruta & "\" & carpeta & "\"
     ActiveWorkbook.SaveAs Filename:=ruta1 & trabajo, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        'On Error Resume Next
        'ActiveWorkbook.Save
        ActiveWorkbook.Close
          Kill (trabajo)
End If

If salida = "SURORIENTE" Then
         
      carpeta = "SURORIENTE"
      
      X = Dir(carpeta, vbDirectory)
      If X = "" Then
           MkDir (carpeta)
       End If
         
      ruta1 = Ruta & "\" & carpeta & "\"
     ActiveWorkbook.SaveAs Filename:=ruta1 & trabajo, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        'On Error Resume Next
        'ActiveWorkbook.Save
        ActiveWorkbook.Close
          Kill (trabajo)
    End If
         trabajo = Dir(Ruta & "\INVENTA*.xlsx")
Wend
  
End If  '**************************************FIN DE CHECKBOX
If ZONA_INVENTARIOS.Value = True Then

    SACAR_ZONAS

End If

End Sub
Private Sub SACAR_ZONAS()
Dim bogota As String
Dim cundinarmaca As String
Dim noroccidente As String
Dim cundinarmaca As String
Dim noroccidente As String
   ruta_salida = Ruta
   salida = Dir(ruta_salida & "\Plantilla_Electromecanico_Final_v2_Vacio.xls")
   ChDir ruta_salida
   Workbooks.Open Filename:=salida, UpdateLinks:=0
      
   

End Sub



'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< FUNCIO UNIR  RF <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Sub UNIR_RF_Click()

Dim HOJAFINAL As Integer
Dim Datos As Integer
Dim HOJA As Integer
Application.ScreenUpdating = False ' para que trabaje sin mostrar imagen en pantalla


If INFO_RF.Value = True Then  'HOJA ÍNFORMACION RF
      Dim ws As Worksheet
      Application.DisplayAlerts = False
      Windows("INFORMACION RF.xlsx").Activate
     If ActiveWorkbook.ProtectStructure = True Then 'CLAVE DESBLOQUEO HOJA
       Contraseña = "Preventivos2016"
       ActiveWorkbook.Unprotect Contraseña
     
     End If
      'Recorrer cada una de las hojas del libro activo
      For Each ws In ActiveWorkbook.Worksheets 'MOSTRAS HOJAS OCULTAS
               ws.Visible = xlSheetVisible
      Next ws

            
      Windows("INFORMACION RF.xlsx").Activate
      HOJAFINAL = Sheets.Count
    
      Sheets(1).Name = "INFORMACION MP RF"
    
     If HOJAFINAL > 1 Then
           For i = 1 To HOJAFINAL
              ' MsgBox "HOJAS  " & HOJAFINAL & "  CONTADOR i  " & i
              If i > 1 And i <= HOJAFINAL Then
                 Windows("INFORMACION RF.xlsx").Activate
                 Worksheets(i).Select
                 HOJA = ActiveSheet.Shapes.Count
                 If HOJA < 5 Then
                    Windows("INFORMACION RF.xlsx").Activate
                    Worksheets(i).Delete
                    i = i - 1
                  Windows("INFORMACION RF.xlsx").Activate
                  HOJAFINAL = Sheets.Count
                 End If
                  Windows("INFORMACION RF.xlsx").Activate
                  HOJAFINAL = Sheets.Count
              End If
               
       
           Next i
      End If
    
End If
Application.DisplayAlerts = True

If FOTOS_RF.Value = True Then
  '--------------------------GSM------------------------------
  '------PEGAR FOTOS RF--PRINCIPAL *************************************

    
       If GSM.Value = True Then 'FUNCION PEGAR HOJAS GSM
        Windows("FOTOS RF GSM.xlsx").Activate
               For Each ws In ActiveWorkbook.Worksheets 'MOSTRAS HOJAS OCULTAS
               ws.Visible = xlSheetVisible
               Next ws

        Call quitaNombres
         Hojas1 = Sheets.Count
         Datos = 0
       
    For i = 1 To Hojas1
          Windows("FOTOS RF GSM.xlsx").Activate
   
          Worksheets(i).Select
          HOJA = ActiveSheet.Shapes.Count
          
        If HOJA > 4 Then
         Workbooks("FOTOS RF GSM.xlsx").Worksheets(i).Copy After:=Workbooks("INFORMACION RF.xlsx").Sheets(HOJAFINAL)
         HOJAFINAL = HOJAFINAL + 1
         Datos = Datos + 1
        End If
       
    Next i
         MsgBox "SE PASO " & Datos & " HOJAS GSM"
           Windows("FOTOS RF GSM.xlsx").Activate
           On Error Resume Next
           ActiveWorkbook.Save
           ActiveWorkbook.Close
           
           
            
       End If ' IF GSM
   
'----------------------UMTS-----------------------------------------
       If UMTS.Value = True Then 'FUNCION PEGAR HOJAS UMTS----
          Windows("FOTOS RF UMTS.xlsx").Activate
            For Each ws In ActiveWorkbook.Worksheets 'MOSTRAS HOJAS OCULTAS
               ws.Visible = xlSheetVisible
               Next ws
          Call quitaNombres
          Hojas1 = Sheets.Count
          Datos = 0
            For i = 1 To Hojas1
             Windows("FOTOS RF UMTS.xlsx").Activate
             Worksheets(i).Select
             HOJA = ActiveSheet.Shapes.Count
             
              If HOJA > 4 Then
                 Workbooks("FOTOS RF UMTS.xlsx").Worksheets(i).Copy After:=Workbooks("INFORMACION RF.xlsx").Sheets(HOJAFINAL)
                 HOJAFINAL = HOJAFINAL + 1
                 Datos = Datos + 1
              End If
            Next i
            MsgBox "SE PASO " & Datos & " HOJAS UMTS"
           Windows("FOTOS RF UMTS.xlsx").Activate
            On Error Resume Next
           ActiveWorkbook.Save
           ActiveWorkbook.Close
            
        End If ' IF UMTS--------------
 '----------------------LTE-----------------------------------------
        
        If LTE.Value = True Then 'FUNCION PEGAR HOJAS LTE----
               Windows("FOTOS RF LTE.xlsx").Activate
                For Each ws In ActiveWorkbook.Worksheets 'MOSTRAS HOJAS OCULTAS
               ws.Visible = xlSheetVisible
               Next ws
               Call quitaNombres
               Hojas1 = Sheets.Count
               Datos = 0
            For i = 1 To Hojas1
             Windows("FOTOS RF LTE.xlsx").Activate
             Worksheets(i).Select
             HOJA = ActiveSheet.Shapes.Count
             
              If HOJA > 4 Then
                 Workbooks("FOTOS RF LTE.xlsx").Worksheets(i).Copy After:=Workbooks("INFORMACION RF.xlsx").Sheets(HOJAFINAL)
                 HOJAFINAL = HOJAFINAL + 1
                 Datos = Datos + 1
              End If
            Next i
            MsgBox "SE PASO " & Datos & " HOJAS LTE"
           Windows("FOTOS RF LTE.xlsx").Activate
            On Error Resume Next
           ActiveWorkbook.Save
           ActiveWorkbook.Close
           
            
        End If ' IF LTE----------------
   
End If ' IF FOTOS RF FINAL*******************************************************************
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>---PANTALLAZOS-->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

If PANTALLAZO_RF.Value = True Then
  '------------------------PANTALLAZOS--GSM------------------------------
  '------PEGAR PANTALLAZO RF--PRINCIPAL *************************************
         Windows("INFORMACION RF.xlsx").Activate
          For Each ws In ActiveWorkbook.Worksheets 'MOSTRAS HOJAS OCULTAS
               ws.Visible = xlSheetVisible
          Next ws
         HOJAFINAL = Sheets.Count
    
       If GSM.Value = True Then 'FUNCION PEGAR HOJAS GSM
         Windows("PANTALLAZOS GSM.xlsx").Activate
          For Each ws In ActiveWorkbook.Worksheets 'MOSTRAS HOJAS OCULTAS
               ws.Visible = xlSheetVisible
          Next ws
         Call quitaNombres
         Hojas1 = Sheets.Count
         Datos = 0
        For i = 1 To Hojas1
          Windows("PANTALLAZOS GSM.xlsx").Activate
          Worksheets(i).Select
          HOJA = ActiveSheet.Shapes.Count
          
           If HOJA > 4 Then
             Workbooks("PANTALLAZOS GSM.xlsx").Worksheets(i).Copy After:=Workbooks("INFORMACION RF.xlsx").Sheets(HOJAFINAL)
             HOJAFINAL = HOJAFINAL + 1
             Datos = Datos + 1
           End If
       
        Next i
           MsgBox "SE PASO " & Datos & " HOJAS PANTALLAZOS GSM"
           Windows("PANTALLAZOS GSM.xlsx").Activate
           On Error Resume Next
           ActiveWorkbook.Save
           ActiveWorkbook.Close
         
           
            
       End If ' IF GSM
   
'----------------------UMTS-----------------------------------------
       If UMTS.Value = True Then 'FUNCION PEGAR HOJAS UMTS----
             Windows("PANTALLAZOS UMTS.xlsx").Activate
              For Each ws In ActiveWorkbook.Worksheets 'MOSTRAS HOJAS OCULTAS
               ws.Visible = xlSheetVisible
               Next ws
             Call quitaNombres
             Hojas1 = Sheets.Count
             Datos = 0
            For i = 1 To Hojas1
             Windows("PANTALLAZOS UMTS.xlsx").Activate
             Worksheets(i).Select
             HOJA = ActiveSheet.Shapes.Count
             
              If HOJA > 4 Then
                 Workbooks("PANTALLAZOS UMTS.xlsx").Worksheets(i).Copy After:=Workbooks("INFORMACION RF.xlsx").Sheets(HOJAFINAL)
                 HOJAFINAL = HOJAFINAL + 1
                 Datos = Datos + 1
              End If
            Next i
             MsgBox "SE PASO " & Datos & " HOJAS PANTALLAZOS UMTS"
           Windows("PANTALLAZOS UMTS.xlsx").Activate
             On Error Resume Next
           ActiveWorkbook.Save
           ActiveWorkbook.Close
          
            
        End If ' IF UMTS--------------
 '----------------------LTE-----------------------------------------
        
        If LTE.Value = True Then 'FUNCION PEGAR HOJAS LTE----
               Windows("PANTALLAZOS LTE.xlsx").Activate
                For Each ws In ActiveWorkbook.Worksheets 'MOSTRAS HOJAS OCULTAS
               ws.Visible = xlSheetVisible
               Next ws
               Call quitaNombres
               Hojas1 = Sheets.Count
               Datos = 0
            For i = 1 To Hojas1
             Windows("PANTALLAZOS LTE.xlsx").Activate
             Worksheets(i).Select
             HOJA = ActiveSheet.Shapes.Count
             
              If HOJA > 4 Then
                 Workbooks("PANTALLAZOS LTE.xlsx").Worksheets(i).Copy After:=Workbooks("INFORMACION RF.xlsx").Sheets(HOJAFINAL)
                 HOJAFINAL = HOJAFINAL + 1
                 Datos = Datos + 1
              End If
            Next i
            MsgBox "SE PASO " & Datos & " HOJAS PANTALLAZOS LTE"
           Windows("PANTALLAZOS LTE.xlsx").Activate
            On Error Resume Next
           ActiveWorkbook.Save
           ActiveWorkbook.Close
           
          
           
            
        End If ' IF LTE----------------
           
End If ' IF FOTOS RF FINAL*******************************************************************
         Application.ScreenUpdating = True ' MOSTRAR PANTALLAS PROCESO
         Windows("INFORMACION RF.xlsx").Activate
         MsgBox "EL ARCHIVO SE GUARDARA " & " CON EL NOMBRE FORMATO-MTTO_PREVENTIVO_ESTACION_RF_v4 "
          On Error Resume Next
         Application.Dialogs(xlDialogSaveAs).Show
         ActiveWorkbook.SaveAs Filename:="FORMATO-MTTO_PREVENTIVO_ESTACION_RF_00-00-2016_v4 1"
           

   MACRO_MPA.Hide
End Sub

Private Sub CommandButton1_Click() ' PEGAR 3 ACHIVOS CON BATERIAS

  
   
    Windows("PREVENTIVO EX-2.xlsx").Activate
    hojas3 = Sheets.Count
    
    For i = 1 To hojas3
    Workbooks("PREVENTIVO EX-2.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(1)
    Next i
        
    Windows("PREVENTIVO EX-3.xlsx").Activate
    hojas3 = Sheets.Count
    
    For i = 1 To hojas3
    Workbooks("PREVENTIVO EX-3.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(1)
    Next i

    ' ORDENA HOJAS
    Windows("PREVENTIVO EX-1.xlsx").Activate

    Sheets("SUBESTACION").Select
    Sheets("SUBESTACION").Move Before:=Sheets(1)
    Sheets("TRANSFER").Move Before:=Sheets(2)
    Sheets("PLANTAS").Move Before:=Sheets(3)
    Sheets("TANQUE").Move Before:=Sheets(4)
    Sheets("TABLERO AC-DC").Move Before:=Sheets(5)
    Sheets("RECTIFICADOR").Move Before:=Sheets(6)
    Sheets("TIERRAS").Move Before:=Sheets(7)
    Sheets("A.A.V").Move Before:=Sheets(8)
    Sheets("A.A.C").Move Before:=Sheets(9)
    Sheets("UPS").Move Before:=Sheets(10)
    Sheets("SOLARES").Move Before:=Sheets(11)
    Sheets("INSUMOS").Move Before:=Sheets(12)

 '--------------------------------------------------------
    VALOR = 0
    Windows("BATERIAS").Activate ' PEGADO DE BATERIAS
    hojas2 = Sheets.Count
    VALOR = hojas2 ' NUMERO DE HOJAS DE BATERIAS
    
    For i = 1 To hojas2
    Windows("BATERIAS").Activate
   
    
       If (i = 1) Then
       Workbooks("BATERIAS.xlsx").Worksheets(1).Move Before:=Workbooks("PREVENTIVO EX-1").Sheets(7)
       
          End If
      
       If (i = 2) Then
     Workbooks("BATERIAS.xlsx").Worksheets(1).Move Before:=Workbooks("PREVENTIVO EX-1").Sheets(8)
     Windows("PREVENTIVO EX-1.xlsx").Activate
     Sheets(8).Name = "BATERIAS2"
          End If
      
      If (i = 3) Then
     Workbooks("BATERIAS.xlsx").Worksheets(1).Move Before:=Workbooks("PREVENTIVO EX-1").Sheets(9)
     Windows("PREVENTIVO EX-1.xlsx").Activate
     Sheets(9).Name = "BATERIAS3"
          End If
          
      If (i = 4) Then
     Workbooks("BATERIAS.xlsx").Worksheets(1).Move Before:=Workbooks("PREVENTIVO EX-1").Sheets(10)
     Windows("PREVENTIVO EX-1.xlsx").Activate
     Sheets(10).Name = "BATERIAS4"
          End If
           
     Next i
     
If FOTOS_EX.Value = True Then
      
           On Error Resume Next
          Windows("PREVENTIVO EX-1.xlsx").Activate
          hojas = Sheets.Count
          Windows("EX FOTOS P1.xlsx").Activate
          foto = Sheets.Count
          
        If (foto > 1) Then
          For i = 1 To foto
          Windows("EX FOTOS P1.xlsx").Activate
          Worksheets(i).Select
          PESO = ActiveSheet.Shapes.Count
           If (PESO >= 1) Then
            Windows("EX FOTOS P1.xlsx").Activate
            ' Sheets(i).Tab.Color = vbGreen
           End If
          
          Next i
       End If
          
       For i = 1 To foto
          
         Workbooks("EX FOTOS P1.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(hojas)
         hojas = hojas + 1
       Next i
       
    On Error Resume Next
        
       Windows("EX FOTOS P2.xlsx").Activate
       foto2 = Sheets.Count
   If (foto2 > 1) Then
       For i = 1 To foto2
          Windows("EX FOTOS P2.xlsx").Activate
          Worksheets(i).Select
          PESO = ActiveSheet.Shapes.Count
           If (PESO >= 1) Then
             Windows("EX FOTOS P2.xlsx").Activate
            ' Sheets(i).Tab.Color = vbGreen
           End If
          
          Next i
     End If
       
       For i = 1 To foto2
                   
         Workbooks("EX FOTOS P2.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(hojas)
         hojas = hojas + 1
       Next i
              
End If
      Windows("PREVENTIVO EX-1.xlsx").Activate
      ActiveWorkbook.Save
      MsgBox " SE HA ORDENADO Y GUARDADO   " & hojas & " HOJAS"
      
   MACRO_MPA.Hide

End Sub
Private Sub CommandButton2_Click() ' PEGAR 3 ACHIVOS SIN BATERIAS

    VALOR = 0

    Windows("PREVENTIVO EX-2.xlsx").Activate
    hojas3 = Sheets.Count
    
    For i = 1 To hojas3
    Workbooks("PREVENTIVO EX-2.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(1)
    Next i
        
    Windows("PREVENTIVO EX-3.xlsx").Activate
    hojas3 = Sheets.Count
    
    For i = 1 To hojas3
    Workbooks("PREVENTIVO EX-3.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(1)
    Next i

    ' ORDENA HOJAS
    Windows("PREVENTIVO EX-1.xlsx").Activate

    Sheets("SUBESTACION").Select
    Sheets("SUBESTACION").Move Before:=Sheets(1)
    Sheets("TRANSFER").Move Before:=Sheets(2)
    Sheets("PLANTAS").Move Before:=Sheets(3)
    Sheets("TANQUE").Move Before:=Sheets(4)
    Sheets("TABLERO AC-DC").Move Before:=Sheets(5)
    Sheets("RECTIFICADOR").Move Before:=Sheets(6)
    Sheets("TIERRAS").Move Before:=Sheets(7)
    Sheets("A.A.V").Move Before:=Sheets(8)
    Sheets("A.A.C").Move Before:=Sheets(9)
    Sheets("UPS").Move Before:=Sheets(10)
    Sheets("SOLARES").Move Before:=Sheets(11)
    Sheets("INSUMOS").Move Before:=Sheets(12)
    
If FOTOS_EX.Value = True Then
      
           On Error Resume Next
          Windows("PREVENTIVO EX-1.xlsx").Activate
          hojas = Sheets.Count
          Windows("EX FOTOS P1.xlsx").Activate
          foto = Sheets.Count
          If (foto > 1) Then
          For i = 1 To foto
          Windows("EX FOTOS P1.xlsx").Activate
          Worksheets(i).Select
          PESO = ActiveSheet.Shapes.Count
           If (PESO >= 1) Then
           Windows("EX FOTOS P1.xlsx").Activate
            ' Sheets(i).Tab.Color = vbGreen
           End If
          
          Next i
          
         End If
       For i = 1 To foto
          
         Workbooks("EX FOTOS P1.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(hojas)
         hojas = hojas + 1
       Next i
       
    On Error Resume Next
        
       Windows("EX FOTOS P2.xlsx").Activate
       foto2 = Sheets.Count
       
      If (foto2 > 1) Then
         For i = 1 To foto2
           Windows("EX FOTOS P2.xlsx").Activate
           Worksheets(i).Select
          PESO = ActiveSheet.Shapes.Count
           If (PESO > 1) Then
           Windows("EX FOTOS P2.xlsx").Activate
             'Sheets(i).Tab.Color = vbGreen
           End If
          
          Next i
     End If
       
       For i = 1 To foto2
                   
         Workbooks("EX FOTOS P2.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(hojas)
         hojas = hojas + 1
       Next i

      
              
End If
      Windows("PREVENTIVO EX-1.xlsx").Activate
      ActiveWorkbook.Save
      MsgBox " SE HA ORDENADO Y GUARDADO   " & hojas & " HOJAS"
   
  MACRO_MPA.Hide


End Sub


'ORDENAR

Private Sub CommandButton3_Click()
On Error Resume Next

    Windows("PREVENTIVO EX-1.xlsx").Activate
  
     ' ----------ORDENAR HOJAS---------
 If (VALOR = 0) Then
    Sheets("SUBESTACION").Select
    Sheets("Indice").Move Before:=Sheets(1)
    Sheets("SUBESTACION").Move Before:=Sheets(2)
    Sheets("FOTOS SUBESTACION").Move Before:=Sheets(3)
    Sheets("TABLERO AC-DC").Move Before:=Sheets(4)
    Sheets("FOTOS TABLERO").Move Before:=Sheets(5)
    Sheets("TIERRAS").Move Before:=Sheets(6)
    Sheets("FOTOS TIERRAS").Move Before:=Sheets(7)
    Sheets("PLANTAS").Move Before:=Sheets(8)
    Sheets("FOTOS PLANTA").Move Before:=Sheets(9)
    Sheets("TRANSFER").Move Before:=Sheets(10)
    Sheets("FOTOS TRANSFERENCIA").Move Before:=Sheets(11)
    Sheets("TANQUE").Move Before:=Sheets(12)
    Sheets("FOTOS TANQUE").Move Before:=Sheets(13)
    Sheets("A.A.C").Move Before:=Sheets(14)
    Sheets("FOTOS A.A.C").Move Before:=Sheets(15)
    Sheets("A.A.V").Move Before:=Sheets(16)
    Sheets("FOTOS A.A.V").Move Before:=Sheets(17)
    Sheets("BATERIAS").Move Before:=Sheets(18)
    Sheets("FOTOS BATERIAS").Move Before:=Sheets(19)
    Sheets("RECTIFICADOR").Move Before:=Sheets(20)
    Sheets("FOTOS RECTIFICADOR").Move Before:=Sheets(21)
    Sheets("SOLARES").Move Before:=Sheets(22)
    Sheets("FOTOS PANELES").Move Before:=Sheets(23)
    Sheets("INSUMOS").Move Before:=Sheets(24)
    Sheets("FOTOS INSUMOS").Move Before:=Sheets(25)
    On Error Resume Next
    Sheets("BATERIAS1").Move Before:=Sheets(19)
    On Error Resume Next
    Sheets("BATERIAS2").Move Before:=Sheets(20)
   End If
   
    
    
    Windows("PREVENTIVO EX-1.xlsx").Activate
    hojas3 = Sheets.Count
      
      MsgBox " SE HA ORDENADO Y GUARDADO   " & hojas3 & " HOJA Y FOTOS"
End Sub



Private Sub PEGAR_HOJAS_Click() ' PEGAR 2 ARCHIVOS CON BATERIAS
  

    Windows("PREVENTIVO EX-2.xlsx").Activate
    hojas3 = Sheets.Count
    
    For i = 1 To hojas3
    Workbooks("PREVENTIVO EX-2.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(1)
    Next i

    ' ORDENA HOJAS
    Windows("PREVENTIVO EX-1.xlsx").Activate

    
    Sheets("SUBESTACION").Select
    Sheets("SUBESTACION").Move Before:=Sheets(1)
    Sheets("TRANSFER").Move Before:=Sheets(2)
    Sheets("PLANTAS").Move Before:=Sheets(3)
    Sheets("TANQUE").Move Before:=Sheets(4)
    Sheets("TABLERO AC-DC").Move Before:=Sheets(5)
    Sheets("RECTIFICADOR").Move Before:=Sheets(6)
    Sheets("TIERRAS").Move Before:=Sheets(7)
    Sheets("A.A.V").Move Before:=Sheets(8)
    Sheets("A.A.C").Move Before:=Sheets(9)
    Sheets("UPS").Move Before:=Sheets(10)
    Sheets("SOLARES").Move Before:=Sheets(11)
    Sheets("INSUMOS").Move Before:=Sheets(12)
    '--------------------------------------------------------
    Windows("BATERIAS").Activate ' PEGADO DE BATERIAS
    hojas2 = Sheets.Count
    For i = 1 To hojas2
    Windows("BATERIAS").Activate
       If (i = 1) Then
       Workbooks("BATERIAS.xlsx").Worksheets(1).Move Before:=Workbooks("PREVENTIVO EX-1").Sheets(7)
          End If
      
       If (i = 2) Then
     Workbooks("BATERIAS.xlsx").Worksheets(1).Move Before:=Workbooks("PREVENTIVO EX-1").Sheets(8)
          End If
      
      If (i = 3) Then
     Workbooks("BATERIAS.xlsx").Worksheets(1).Move Before:=Workbooks("PREVENTIVO EX-1").Sheets(9)
          End If
          
      If (i = 4) Then
     Workbooks("BATERIAS.xlsx").Worksheets(1).Move Before:=Workbooks("PREVENTIVO EX-1").Sheets(10)
          End If
           
    Next i
           
   
      hojas = Sheets.Count
      ActiveWorkbook.Save
  
  
    MsgBox "AHORA TIENE QUE BUSCAR Y REMPLAZAR MANUALMENTE LA PALABRA.. =SI(  EN TODO EL LIBRO. "
     
    MACRO_MPA.Hide
 


End Sub


Private Sub PEGAR_SIN_BATERIAS_Click() ' PEGAR 2 ARCHIVOS SIN BATERIAS


 
    Windows("PREVENTIVO EX-2.xlsx").Activate
    hojas3 = Sheets.Count
    
    For i = 1 To hojas3
    Workbooks("PREVENTIVO EX-2.xlsx").Worksheets(1).Move After:=Workbooks("PREVENTIVO EX-1").Sheets(1)
    Next i

    ' ORDENA HOJAS
    Windows("PREVENTIVO EX-1.xlsx").Activate

    Sheets("SUBESTACION").Select
    Sheets("SUBESTACION").Move Before:=Sheets(1)
    Sheets("TRANSFER").Move Before:=Sheets(2)
    Sheets("PLANTAS").Move Before:=Sheets(3)
    Sheets("TANQUE").Move Before:=Sheets(4)
    Sheets("TABLERO AC-DC").Move Before:=Sheets(5)
    Sheets("RECTIFICADOR").Move Before:=Sheets(6)
    Sheets("TIERRAS").Move Before:=Sheets(7)
    Sheets("A.A.V").Move Before:=Sheets(8)
    Sheets("A.A.C").Move Before:=Sheets(9)
    Sheets("UPS").Move Before:=Sheets(10)
    Sheets("SOLARES").Move Before:=Sheets(11)
    Sheets("INSUMOS").Move Before:=Sheets(12)
     
     MsgBox "AHORA TIENE QUE BUSCAR Y REMPLAZAR MANUALMENTE LA PALABRA...   SI(  EN TODO EL LIBRO. "
     
MACRO_MPA.Hide
End Sub

Private Sub CAMBIAR_FORMATO_Click()

   Windows("PREVENTIVO EX-1.xlsx").Activate
   hojas = Sheets.Count
     If (hojas > 12) Then
         hojas = VALOR + 16
    Else
     hojas = hojas
    End If
   
   
    For i = 1 To hojas
      
    Sheets(i).Select
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

     Next i

     Sheets(1).Select

     ActiveWorkbook.Save
     MsgBox " SE HA GUARDADO LOS CAMBIOS A " & i - 1 & " HOJAS"


End Sub

Private Sub TOMAR_DATOS_Click() ' TOMAR DATOS


' PEGA RMC
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
    Range("C48:L48").Select
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("A19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
    Range("R34:Y34").Select
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("D19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    Range("D21").Select
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("SUBESTACION").Select
    Range("C8").Select
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("B19").Select
    ActiveSheet.Paste
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("PLANTAS").Select
    Range("E30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2016.xlsx").Activate
     Sheets("RESUMEN").Select
    Range("C19").Select
    ActiveSheet.Paste
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets("TANQUE").Select
    Range("C14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("E19").Select
    ActiveSheet.Paste
    Windows("PREVENTIVO EX-1.xlsx").Activate
    Sheets(7).Select
    Sheets(7).Name = "BATERIAS"
    Range("O12:P12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("CONTROL_MP_SUROCCIDENTE_CENTRO_2016.xlsx").Activate
    Sheets("RESUMEN").Select
    Range("F19").Select
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("F19").Select


End Sub

Sub quitaNombres()

' Macro desarrollada el 21/10/2006 por Elsamatilde
'
Dim nroNbres, i As Integer
On Error Resume Next
nroNbres = ActiveWorkbook.Names.Count
'MsgBox nroNbres 'opcional
For i = 1 To nroNbres
ActiveWorkbook.Names(1).Delete
Next i
End Sub

Sub CrearCarpeta()
  Ruta = GUARDAR_EXCEL
  Ruta2 = Ruta
 CreaCarpeta Ruta, "Informes Agregados"
End Sub
 
Sub CreaCarpeta(Ruta2 As String, NomCarpeta As String)
 'Verificar si la carpeta existe.
  If Dir(Ruta, vbDirectory + vbHidden) = "" Then
   'Comprueba que la carpeta no exista para crear el directorio.
    If Dir(Ruta & "" & NomCarpeta, vbDirectory + vbHidden) = "" Then _
       MkDir Ruta & "" & NomCarpeta
  End If
End Sub


Private Sub ZONA_INVENTARIOS_Click()
MACRO_MPA.Hide
CORRECCION_INVENTARIO.Show
'UserForm1.Show

End Sub
