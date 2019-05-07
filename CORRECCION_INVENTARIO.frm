VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CORRECCION_INVENTARIO 
   Caption         =   "Correccion Del Inventario "
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6690
   OleObjectBlob   =   "CORRECCION_INVENTARIO.frx":0000
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "CORRECCION_INVENTARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim lastRow, VALOR1, VALOR2 As Double
Dim BUSCARLETRAS, L1, L2, L3, L4, L5, TEXTO_CELDA, MODIFICA As String



Sub BUCLE() '-------------------------------BUCLE1----MAYOR MENTE PARA PINTAR CELDAS CON ERROR-------------
  
 For i = 5 To lastRow
 
 If IsNumeric(ActiveCell.Value) Then
   If (ActiveCell.Value) < VALOR1 Then ActiveCell.Interior.Color = QBColor(12)
   If (ActiveCell.Value) > VALOR2 Then ActiveCell.Interior.Color = QBColor(12)
  End If
 If BUSCARLETRAS = "SI" Then
  
   If ActiveCell.Value = L1 Then ActiveCell.Interior.Color = QBColor(12)
   If ActiveCell.Value = L2 Then ActiveCell.Interior.Color = QBColor(12)
   If ActiveCell.Value = L3 Then ActiveCell.Interior.Color = QBColor(12)
   If ActiveCell.Value = L4 Then ActiveCell.Interior.Color = QBColor(12)
   
  End If
   If MODIFICA = "SI" Then
      If IsError(ActiveCell) Then ActiveCell = ""
     If ActiveCell.Value = L1 Then ActiveCell = TEXTO_CELDA
     If ActiveCell.Value = L2 Then ActiveCell = TEXTO_CELDA
     If ActiveCell.Value = L3 Then ActiveCell = TEXTO_CELDA
     If ActiveCell.Value = L4 Then ActiveCell = TEXTO_CELDA
   
  End If
    
   ActiveCell.Offset(1, 0).Select
 Next i
End Sub
Sub BUCLE1() '-------------------------------BUCLE1----MAYOR MENTE PARA CORREGIR VALORES-------------
  
   
  For i = 5 To lastRow
   
 If IsNumeric(ActiveCell.Value) Then
   If (ActiveCell.Value) < VALOR1 Then ActiveCell = VALOR1
   If (ActiveCell.Value) >= VALOR2 Then ActiveCell = VALOR2
  End If
 
  If MODIFICA = "SI" Then
     If IsError(ActiveCell) Then ActiveCell = ""
     If ActiveCell.Value = L1 Then ActiveCell = TEXTO_CELDA
     If ActiveCell.Value = L2 Then ActiveCell = TEXTO_CELDA
     If ActiveCell.Value = L3 Then ActiveCell = TEXTO_CELDA
     If ActiveCell.Value = L4 Then ActiveCell = TEXTO_CELDA
   
  End If
  
      ActiveCell.Offset(1, 0).Select
 Next i

End Sub
Sub UpdateProgressBar(ava)
'Por.DAM
    CORRECCION_INVENTARIO.FProgress.Caption = Format(ava, "0%")
    CORRECCION_INVENTARIO.LProgress.Width = ava * (CORRECCION_INVENTARIO.FProgress.Width - 10)
    'UserForm1.FProgress.Caption = Format(ava, "0%")
    'UserForm1.LProgress.Width = ava * (UserForm1.FProgress.Width - 10)
    DoEvents
End Sub



Private Sub Boton_Click() '**************************************************************************

 CORRECCION_INVENTARIO.LProgress.Width = 0
 
 avance = 0

UpdateProgressBar avance
Application.ScreenUpdating = False
'*******************************************************PRUEBAS---------------------
 
'////////////////////////////////////////////////////////////////////////////////////////

'SELECCIONA HOJA EN LA QUE SE VA A TRABAJAR Y FONDO DE CELDA SIN COLOR
'//////////////////////////////////////////////////////////////////////  HOJA  PEE_    *******************************************************

avance = 0.1
UpdateProgressBar avance


Sheets("PEE_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With

ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0

'Atributo "ALTURA DE INSTALACIÓN" con celdas vacías fecha 12/04/2018
 Range("E5").Select
 
  VALOR1 = 3
  VALOR2 = 3500
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "SI"
    BUCLE
        
 
'No es valido el valor de "0" y "NA" en "CARGA ACTUAL KVA". Independientemente de si está apagado o prendido se requiere conocer cuánto puede soportar FECHA=12/04/2018
 Range("F5").Select
   VALOR1 = 0
  VALOR2 = 500
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "SI"
    BUCLE
  '-----------------"CORRIENTE DERRATEADA Y CORRIENTE NOMINAL".
  Range("G5").Select
  VALOR1 = 5
  VALOR2 = 350
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
    BUCLE
   Range("H5").Select '-----------------"CORRIENTE DERRATEADA Y
  VALOR1 = 5
  VALOR2 = 350
  L1 = ""
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
    BUCLE
 
'Atributo "FACTOR POTENCIA NOMINAL" valores en "0" no son permitidos
Range("J5").Select
  VALOR1 = 0.6
  VALOR2 = 0.9
 
  BUSCARLETRAS = "NO"
    BUCLE
'Atributo "FACTOR DE EFICIENCIA" valores en "0" no son permitidos
Range("K5").Select
  VALOR1 = 0.7
  VALOR2 = 1
 
  BUSCARLETRAS = "NO"
    BUCLE

'Atributo "HORAS DE TRABAJO" Valores en "0" y "NA" no permitidos
Range("N5").Select
For i = 5 To lastRow
 
If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "N/A" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "NA" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = 0 Then Cells(ActiveCell.Row, 15) = "NO"

If ActiveCell.Value = "NO VISIBLE" Then ActiveCell.Interior.ColorIndex = 0
ActiveCell.Offset(1, 0).Select
Next i

'SI LA MARCA  , MARCA PANEL CONTROL, MODELO  ESTAN EN N/A O N/V CAMBIAR POR NO VISIBLE

Range("P5").Select

For i = 5 To lastRow
    Select Case ActiveCell.Value
        Case "N/A"
            ActiveCell = "NO VISIBLE"
    End Select
ActiveCell.Offset(1, 0).Select
Next i




'SI LA MARCA  , MARCA PANEL CONTROL, MODELO  ESTAN EN N/A O N/V CAMBIAR POR NO VISIBLE
Range("Q5").Select
For i = 5 To lastRow
    Select Case ActiveCell.Value
        Case "N/A"
            ActiveCell = "NO VISIBLE"
        Case "N/V"
            ActiveCell = "NO VISIBLE"
        Case "NO VERIFICABLE"
            ActiveCell = "NO VISIBLE"
        Case "NO TIENE"
            ActiveCell = "NO VISIBLE"
    End Select
ActiveCell.Offset(1, 0).Select
Next i

'SI LA MARCA  , MARCA PANEL CONTROL, MODELO  ESTAN EN N/A O N/V CAMBIAR POR NO  VISIBLE 12/04/2018
Range("R5").Select
For i = 5 To lastRow
Select Case ActiveCell.Value
        Case "0"
            ActiveCell = "NO VISIBLE"
        Case "N/A"
            ActiveCell = "NO VISIBLE"
         Case "NA"
            ActiveCell = "NO VISIBLE"
         Case ""
            ActiveCell = "NO VISIBLE"
    End Select
ActiveCell.Offset(1, 0).Select
Next i
'-------------------------------------------RUIDO---------------
Range("U5").Select
 VALOR1 = 30
  VALOR2 = 100
 
  BUSCARLETRAS = "NO"
    BUCLE



'Atributo "NÚMERO DE FASES" valores en  "0", "NA" Y "REF" fecha 12/04/2018
Range("V5").Select
For i = 5 To lastRow
If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "N/A" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "NA" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "REF" Then ActiveCell.Interior.Color = QBColor(12)
ActiveCell.Offset(1, 0).Select
Next i

'si no hay POTENCIA DERRATEADA KVA esta sera igual a POTENCIA NOMINAL KVARange("Y5").Select
For i = 5 To lastRow
If ActiveCell.Value = 0 Then ActiveCell = Cells(ActiveCell.Row, 28)
ActiveCell.Offset(1, 0).Select
Next i


'Atributo "POTENCIA DERRATEADA" Validar valores menores a 10 KVA  fecha  --------------
Range("Y5").Select

 VALOR1 = 0
 VALOR2 = 500
 BUSCARLETRAS = "NO"
 BUCLE


Range("Z5").Select
For i = 5 To lastRow
If Int(ActiveCell.Value) < 10 Then ActiveCell.Interior.Color = QBColor(12)
ActiveCell.Offset(1, 0).Select
Next i


'SI CARGA ACTUAL KVA MAYOR QUE POTENCIA DERRATEADA KVA EN ROJO CARGA ACTUAL
Range("F5").Select
For i = 5 To lastRow
 If IsNumeric(ActiveCell.Value) And IsNumeric(Cells(ActiveCell.Row, 25)) Then
   If Int(ActiveCell.Value) > Int(Cells(ActiveCell.Row, 25)) Then ActiveCell.Interior.Color = QBColor(12)
 End If

ActiveCell.Offset(1, 0).Select
Next i


'Atributo "POTENCIA NOMINAL" Valores en "0" y "NA" no permitidos   ------completar error valor
Range("AA5").Select


For i = 5 To lastRow
If IsError(ActiveCell) Then ActiveCell = "NO VISIBLE"
If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "N/A" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "NA" Then ActiveCell.Interior.Color = QBColor(12)
ActiveCell.Offset(1, 0).Select
Next i

Range("AB5").Select
  VALOR1 = 0
  VALOR2 = 500
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "SI"
    BUCLE



' Atributo "TENSIÓN NOMINAL" Valores en "0" no permitidos FECHA
  VALOR1 = 100
  VALOR2 = 240
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "SI"
    BUCLE


'VACANCIA 0-100 DE LO CONTRARIO ROJO
Range("AN5").Select
For i = 5 To lastRow
ActiveCell.NumberFormat = "#,##0.000"
If IsNumeric(ActiveCell.Value) Then
  If Int(ActiveCell.Value) < 0 Then ActiveCell.Interior.Color = QBColor(12)
  If Int(ActiveCell.Value) > 100 Then ActiveCell.Interior.Color = QBColor(12)
End If
ActiveCell.NumberFormat = "00%"
ActiveCell.Offset(1, 0).Select
Next i


avance = 0.2
UpdateProgressBar avance

'***************************************************************************HOJA TANQUE_COMBUSTIBLE-----------------------------------------------------------------
Sheets("TanqCombustible_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With

ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0

'"Atributo "ALTURA TANQUE cm" Valores "0" Y "NA" no permitidos. Valores de 1.8 cm?

Range("E5").Select
  VALOR1 = 40
  VALOR2 = 300
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

'"   Atributo "ANCHO TANQUE cm" Valores "0" Y "NA" no permitidos. Valores de 2.9 cm?
Range("F5").Select

VALOR1 = 40
  VALOR2 = 200
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

' "   Atributo "AUTONOMIA ESTIMADA" Valores mayores a 200 h
Range("G5").Select
For i = 5 To lastRow

 If IsNumeric(ActiveCell.Value) Then
  If Int(ActiveCell.Value) >= 300 Then ActiveCell.Interior.Color = QBColor(12)
 End If

ActiveCell.Offset(1, 0).Select
Next i

'"   Atributo "CAPACIDAD DEL TANQUE" valores en "0" no permitidos
Range("H5").Select

  VALOR1 = 0
  VALOR2 = 1000
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'"   Atributo "CIECUNFERENCIA DEL TANQUE cm" Valores elevados ejemplo: 19100.
Range("I5").Select

  VALOR1 = 0
  VALOR2 = 150
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'"   Atributo "LARGO DEL TANQUE cm" Valores elevados ejemplo: 19100.
Range("N5").Select

  VALOR1 = 0
  VALOR2 = 300
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

avance = 0.3
UpdateProgressBar avance
'********************************************************************************HOJA *MOTOR******************--------------------------------
Sheets("Motor_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0

'"   Atributo "VOLUMEN ACTUAL COMBUSTUIBLE" Valores "NA" no permitidos
Range("E5").Select
  VALOR1 = 0
  VALOR2 = 500
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("F5").Select

  VALOR1 = 0
  VALOR2 = 50
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

'"   Atributo "CONSUMO PROMEDIO DE CARGA" VLORES EN "0", se encuentra un valor de "401 Gl/h" es válido?  no pueden ser superiores a 10 gl hora

Range("G5").Select

  VALOR1 = 0
  VALOR2 = 50
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'"   Atributo "NIVEL DE RUIDO" Valores en "0" y "1"
Range("O5").Select

  VALOR1 = 30
  VALOR2 = 100
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

'"   Atributo "POTENCIA DERRATEADA" Valores en "0" y "1"
Range("P5").Select

  VALOR1 = 0
  VALOR2 = 800
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

'"   Atributo "POTENCIA NOMINAL" Valores en "NA"

Range("Q5").Select

  VALOR1 = 0
  VALOR2 = 1000
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

'"   Atributo "TIPO DE COMBUSTIBLE" Valores en "NA"

Range("W5").Select

For i = 5 To lastRow

If ActiveCell.Value = "" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "N/A" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "NO VERIFICABLE" Then ActiveCell.Interior.Color = QBColor(12)
ActiveCell.Offset(1, 0).Select
Next i

'"   Atributo "TIPO MOTOR" valores que no coinciden (T:39201912, B, L)
Range("X5").Select
For i = 5 To lastRow
If ActiveCell.Value = "T" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "B" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "L" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "N/A" Then ActiveCell.Interior.Color = QBColor(12)
ActiveCell.Offset(1, 0).Select
Next i

avance = 0.4
UpdateProgressBar avance
'**********************************************************************HOJA GENERADROR******************************************************
Sheets("Generador_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0

'"   Atributo "MARCA" Valores en "0", "NA" y Vacías. Valores no permitidos
 Range("F5").Select
For i = 5 To lastRow

If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "N/A" Then ActiveCell = "NO VERIFICABLE"
ActiveCell.Offset(1, 0).Select
Next i
'"   Atributo "POTENCIA ACTIVA NOMINAL" Valores "0" y "NA" no permitidos
 Range("J5").Select
For i = 5 To lastRow

If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "N/A" Then ActiveCell = "ND"
ActiveCell.Offset(1, 0).Select
Next i

'"   Atributo "POTENCA DERRATEADA NOMINAL" Valores en "0". ¿Valor de 1361 KVA Válido?  nota : la derrateada es menor a  la nominal
 Range("K5").Select
For i = 5 To lastRow
   If IsNumeric(ActiveCell.Value) And IsNumeric(Cells(ActiveCell.Row, 12)) Then
      If Int(ActiveCell.Value) >= Int(Cells(ActiveCell.Row, 12)) Then ActiveCell.Interior.Color = QBColor(12)
   End If
 If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
 If ActiveCell.Value = "" Then ActiveCell.Interior.Color = QBColor(12)
 If ActiveCell.Value = "N/A" Then ActiveCell = "ND"
 ActiveCell.Offset(1, 0).Select
Next i

'"   Atributo "POTENCIA NOMINAL KVA" Valores en "0" no permitidos
 Range("L5").Select
For i = 5 To lastRow
If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "N/A" Then ActiveCell = "ND"
ActiveCell.Offset(1, 0).Select
Next i

'"   Atributo "TENSIÓN SALIDA ENTRE FASES" Valores en "0" y "NA" No permitidos

 Range("P5").Select
For i = 5 To lastRow
If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "N/A" Then ActiveCell = "ND"
ActiveCell.Offset(1, 0).Select
Next i



 
avance = 0.5
UpdateProgressBar avance
'**********************************************************************HOJA Bateria_Arranque_******************************************************

Sheets("BatArranque_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0

'"   Atributo "MARCA BATERIA DE ARRANQIE" Valores "NA" No permitidos

 Range("G5").Select
For i = 5 To lastRow
If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "N/D" Then ActiveCell.Interior.Color = QBColor(12)
If ActiveCell.Value = "N/A" Then ActiveCell = "NO VISIBLE"
If ActiveCell.Value = "NA" Then ActiveCell = "NO VISIBLE"
ActiveCell.Offset(1, 0).Select

Next i

'"   Atributo "TENSIÓN BATERIA DE ARRANQUE" Valores en "0" y menores a "9.5" V ESTARIA MAL

 Range("P5").Select
 
  VALOR1 = 11
  VALOR2 = 25
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("Q5").Select
  VALOR1 = 11
  VALOR2 = 25
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

'"   Atributo "TENSIÓN CAGADOR MOTOR" Valores en "0" y "NA" No permitidos
 Range("R5").Select
  VALOR1 = 11
  VALOR2 = 31
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
   
   avance = 0.6
UpdateProgressBar avance
  '********************************************************************************HOJA TRANFERENCIA***************************************************
Sheets("Transferencia_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0
'"  Atributo "CARGA ACTUAL" Valores en "0" y "NA" no validos
  Range("E5").Select
  VALOR1 = 0
  VALOR2 = 500
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'"  Atributo "POTENCIA NOMINAL" Valores en "0" y "NA" no validos

Range("N5").Select
  VALOR1 = 0
  VALOR2 = 300
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
 
'"   Atributo "TENSIÓN NOMINAL FASE-FASE" y "TENSIÓN NOMINAL FASE- NEUTRO" Valores en "0" y "NA" No disponible
  Range("V5").Select
     VALOR1 = 100
     VALOR2 = 240
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

 Range("W5").Select
     VALOR1 = 115
     VALOR2 = 126
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
 avance = 0.7
UpdateProgressBar avance
'********************************************************************************HOJA POWER***************************************************
Sheets("Power_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0
'"   Atributo "CANTIDAD DE SLOTS LIBRES" Valores negativos, "NA"  y errores de formula.
 Range("E5").Select
For i = 5 To lastRow
   If IsNumeric(ActiveCell.Value) Then
      If Int(ActiveCell.Value) < 0 Then ActiveCell.Interior.Color = QBColor(12)
   End If
 If IsError(ActiveCell) Then ActiveCell = "NO VISIBLE"
 If ActiveCell.Value = "" Then ActiveCell.Interior.Color = QBColor(12)
 If ActiveCell.Value = "N/A" Then ActiveCell.Interior.Color = QBColor(12)
 ActiveCell.Offset(1, 0).Select
Next i
'"  Atributo "CARGA ACTUAL", "CARGA DC", "CORRIENTE CARGA", "CORRIENTE ENTRADA R", "CORRIENTE ENTRADA S", "CORRIENTE ENTRADA T" Valores en "0" , texto y errores de formula.
  Range("N5").Select

For i = 5 To lastRow
   j = 0
   For j = 0 To 5
      If IsError(ActiveCell) Then ActiveCell = "ND"
      If ActiveCell.Value = 0 Then ActiveCell = "0"
      If Not IsNumeric(ActiveCell.Value) Then
        
        If ActiveCell.Value = "" Then ActiveCell = "ND"
        If ActiveCell.Value = "N/A" Then ActiveCell = "ND"
      End If
     
      ActiveCell.Offset(0, 1).Select
   Next j
  ActiveCell.Offset(0, -j).Select
  ActiveCell.Offset(1, 0).Select
Next i

 Range("N5").Select
  VALOR1 = 0
  VALOR2 = 30
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
 
 Range("O5").Select
  VALOR1 = 0
  VALOR2 = 300
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
 Range("P5").Select
  VALOR1 = 0
  VALOR2 = 300
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

'"   Atributo "CORRIENTE NOMINAL MODULO" Valores en "0", "valores por encima de "1000" validos?
Range("T5").Select
VALOR1 = 0
  VALOR2 = 30
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

'"   Atributo "FACTOR DE EFICIENCIA DEL MODULO" Valores menores a 0.8
Range("V5").Select

VALOR1 = 0
  VALOR2 = 1
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

'"   Atributo " NÚMEROS DE MÓDULOS INSTALADOS" Valores en texto y error de formula.
Range("AC5").Select
If IsError(ActiveCell) Then ActiveCell = ""
  VALOR1 = 0
  VALOR2 = 20
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE


'"   Atributo "POTENCIA NÓMINAL MÓDULO" Valores menores a 1500 y mayores a 3000. Validar
Range("AF5").Select
  VALOR1 = 0
  VALOR2 = 5
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'" Atributo OCUPACION
Range("AE5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'" Atributo "TENSIÓN DE CARGA",
Range("AN5").Select
  VALOR1 = 46
  VALOR2 = 54
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'" Atributo "TENSION ENTRADA ENTRE LINEAS
Range("AO5").Select
  VALOR1 = 100
  VALOR2 = 240
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
''" Atributo  "TENSION ENTRADA ENTRE LINEAS", "TENSION FLOTACION",
Range("AP5").Select
  VALOR1 = 47
  VALOR2 = 60
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("AQ5").Select
  VALOR1 = 47
  VALOR2 = 60
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'  VACANCIA

Range("AT5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE


'" Atributo "TENSIÓN DE CARGA", "TENSION ENTRADA ENTRE LINEAS", "TENSION FLOTACION",
'" TENSION NOMINAL" y "TENSION IGUALACION"  Valores negativos, valores "0", texto y error de formula
Range("AN5").Select
For i = 5 To lastRow
   j = 0
   For j = 0 To 4
      If IsNumeric(ActiveCell.Value) Then
        If Int(ActiveCell.Value) < 1 Then ActiveCell.Interior.Color = QBColor(12)
      End If
      If IsError(ActiveCell) Then ActiveCell.Interior.Color = QBColor(12)
      If Not IsNumeric(ActiveCell.Value) Then ActiveCell.Interior.Color = QBColor(12)
      If ActiveCell.Value = "" Then ActiveCell.Interior.Color = QBColor(12)
   
      ActiveCell.Offset(0, 1).Select
   Next j
  ActiveCell.Offset(0, -j).Select
  ActiveCell.Offset(1, 0).Select
Next i

avance = 0.8
UpdateProgressBar avance
'********************************************************************************HOJA BANCO DE BATERIAS***************************************************
Sheets("BcoBaterias_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0
'"   Atributo "ALTO CELDA",  "LARGO CELDA" Y " NUMERO CELDAS BANCO" Valores en "0" no permitidos
Range("E5").Select

  VALOR1 = 20
  VALOR2 = 80
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'"ANCHO CELDA",

Range("F5").Select
  VALOR1 = 10
  VALOR2 = 40
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'AUTONOMIA ESTIMADA
Range("G5").Select
  VALOR1 = 0
  VALOR2 = 500
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

'CAPACIDAD BANCO

Range("H5").Select
  VALOR1 = 0
  VALOR2 = 1000
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'LARGO CELDA
Range("K5").Select
  VALOR1 = 20
  VALOR2 = 80
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
' NUMERO CELDAS BANCO
Range("N5").Select
  VALOR1 = 4
  VALOR2 = 24
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'   PESO
Range("Q5").Select
  VALOR1 = 0
  VALOR2 = 500
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
' TENSION_FLOTACION_BATERIA
Range("Y5").Select
  VALOR1 = 47
  VALOR2 = 55
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'TENSION_IGUALACION_BATERIA
Range("Z5").Select
  VALOR1 = 47
  VALOR2 = 60
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'TENSION_NOMINAL_BANCO
Range("AA5").Select
  VALOR1 = 47
  VALOR2 = 55
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'TENSION_NOMINAL_BATERIA

Range("AB5").Select
  VALOR1 = 11
  VALOR2 = 14
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE


'------------------
Range("E5").Select
For i = 5 To lastRow
   j = 0
   For j = 0 To 1
      If IsNumeric(ActiveCell.Value) Then
        If Int(ActiveCell.Value) < 0 Then ActiveCell.Interior.Color = QBColor(12)
      End If

      If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
      If ActiveCell.Value = "N/A" Then ActiveCell.Interior.Color = QBColor(12)
      
      ActiveCell.Offset(0, 1).Select
    Next j
      L = 2
      ' "LARGO CELDA"
      ActiveCell.Offset(0, 4).Select
        If IsNumeric(ActiveCell.Value) Then
        If Int(ActiveCell.Value) < 0 Then ActiveCell.Interior.Color = QBColor(12)
      End If
      If IsError(ActiveCell) Then ActiveCell = ""
      If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
      If ActiveCell.Value = "N/A" Then ActiveCell.Interior.Color = QBColor(12)
      L = L + 4
      ' NUMERO CELDAS BANCO"
      ActiveCell.Offset(0, 3).Select
      If IsNumeric(ActiveCell.Value) Then
        If Int(ActiveCell.Value) < 0 Then ActiveCell.Interior.Color = QBColor(12)
      End If
      If IsError(ActiveCell) Then ActiveCell = ""
      If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
      If ActiveCell.Value = "N/A" Then ActiveCell.Interior.Color = QBColor(12)
      L = L + 3
      
      
  ActiveCell.Offset(0, -L).Select
  ActiveCell.Offset(1, 0).Select
Next i

'"   Atributo "TENSION_FLOTACION_BATERIA", "TENSION_FLOTACION_BATERIA", "TENSION_NOMINAL_BANCO",
 '"TENSION_NOMINAL_BATERIA" Valores "0" y "NA" No permitidos
Range("Y5").Select
For i = 5 To lastRow
   j = 0
   For j = 0 To 3
      If IsError(ActiveCell) Then ActiveCell = ""
      If ActiveCell.Value = 0 Then ActiveCell.Interior.Color = QBColor(12)
      If ActiveCell.Value = "" Then ActiveCell.Interior.Color = QBColor(12)
      If ActiveCell.Value = "N/A" Then ActiveCell.Interior.Color = QBColor(12)
      ActiveCell.Offset(0, 1).Select
   Next j
  ActiveCell.Offset(0, -j).Select
  ActiveCell.Offset(1, 0).Select
Next i

'********************************************************************************HOJA TABLERO ELECTRICO ***************************************************
Sheets("TabElectrico_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0
'"   Atributo "CANTIDAD DE CIRCUITOS INSTALADOS" Valores "0" y "NA" no permitidos
'"   Atributo "CANTIDAD POSICIONES LIBRES" Valores "NA" no permitidos
'"   Atributo "CARGA TOTAL" Valores en "0" y "NA" No permitidos
'"   Atributo  "CORRIENTE NOMINAL" Valores "NA" Y error de formula
'"   Atributo "TENSIÓN NOMINLA" Valores en "0" no permitidos

'"   Atributo "CANTIDAD DE CIRCUITOS
Range("E5").Select
  VALOR1 = 1
  VALOR2 = 20
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

'"   Atributo CANTIDAD POSICIONES LIBRES

Range("F5").Select
  VALOR1 = 1
  VALOR2 = 10
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
' CARGA TOTAL
Range("G5").Select
  VALOR1 = 0
  VALOR2 = 200
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'CORRIENTE NOMINAL
Range("I5").Select
  VALOR1 = 0
  VALOR2 = 500
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
' TENSION NOMINAL
Range("W5").Select
  VALOR1 = 100
  VALOR2 = 240
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("E5").Select
For i = 5 To lastRow
    j = 0
   For j = 0 To 2  'CANTIDAD DE CIRCUITOS INSTALADOS", "CANTIDAD POSICIONES LIBRES", CARGA TOTAL"
       If IsError(ActiveCell) Then ActiveCell = ""
     Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      ActiveCell.Offset(0, 1).Select
   Next j
   L = 4
   ActiveCell.Offset(0, 1).Select '"CORRIENTE NOMINAL"
    If IsError(ActiveCell) Then ActiveCell = ""
     Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      L = L + 1
   ActiveCell.Offset(0, 14).Select '"TENSIÓN NOMINLA"
         If IsError(ActiveCell) Then ActiveCell = ""
     Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      L = L + 13
   
  ActiveCell.Offset(0, -L).Select
  ActiveCell.Offset(1, 0).Select
   
Next i

avance = 0.8
UpdateProgressBar avance
'********************************************************************************HOJA PROTECCIONES ***************************************************
Sheets("Protecciones_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0
'"   Atributos "CORRIENTE IMPULSO", "CORRIENTE MAXIMA", "CORRIENTE NOMINAL" Valores en 0 no permitidos. Valores de 1000 A correctos?
'"   Atributo "REFERENCIA" Celdas vacías.
'"   Atributo "TENSIÓN NOMINAL" Valores "0" No permitidos
'"   Atributo "TIPO PROTECCIÓN"  Valores (B, 150VAC, i IP20 CE)
Range("E5").Select
For i = 5 To lastRow
j = 0
   For j = 0 To 2  '"CORRIENTE IMPULSO", "CORRIENTE MAXIMA", "CORRIENTE NOMINAL"
       If IsError(ActiveCell) Then ActiveCell = ""
          If IsNumeric(ActiveCell.Value) Then
          If Int(ActiveCell.Value) > 1000 Then ActiveCell.Interior.Color = QBColor(12)
          End If
         Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
         End Select
      ActiveCell.Offset(0, 1).Select
   Next j
    L = j
    ActiveCell.Offset(0, 7).Select '"   Atributo "REFERENCIA" Celdas vacías.
      Select Case ActiveCell.Value
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
         End Select
     L = L + 7
     ActiveCell.Offset(0, 2).Select '"   Atributo "TENSIÓN NOMINAL" Valores "0" No permitidos
       If IsError(ActiveCell) Then ActiveCell = ""
     
       Select Case ActiveCell.Value
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
         End Select
      L = L + 2
      ActiveCell.Offset(0, 3).Select '"   Atributo "TIPO PROTECCIÓN"  Valores (B, 150VAC, i IP20 CE)
      If IsError(ActiveCell) Then ActiveCell = ""
       Select Case ActiveCell.Value
           Case "B"
              ActiveCell.Interior.Color = QBColor(12)
           Case "150VAC"
              ActiveCell.Interior.Color = QBColor(12)
           Case "IP20"
              ActiveCell.Interior.Color = QBColor(12)
           Case "CE"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
          Case "0"
              ActiveCell.Interior.Color = QBColor(12)
         End Select
      L = L + 3
      ActiveCell.Offset(0, -L).Select
      ActiveCell.Offset(1, 0).Select
Next i
'********************************************************************************HOJA ACOMETIDA PRINCIPAL ***************************************************
Sheets("AcomPpal_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0
'"   Atributo "CALIBRE CONDUTOR" Valores como fechas, indicar valores con números. Validar información
'"   Atributo "CORRIENTE ENTREGADA R", "CORRIENTE ENTREGADA S", "CORRIENTE ENTREGADA T" valores en "0" y "NA" No permitidos
'"   Atributo "FACTOR DE POTENCIA" Valores fuera de rango (0.8-0.95)
'"   Atributo "NUMERO DE CONDUCTORES POR FASE" Validar valores mayores a 3
'"   Atributo "NUMERO DE FASES" Valores "0" no permitidos
'CORRIENTE ENTREGADA R

Range("G5").Select
  VALOR1 = 0
  VALOR2 = 300
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'CORRIENTE ENTREGADA S
Range("H5").Select
BUCLE
'CORRIENTE ENTREGADA T
Range("I5").Select
BUCLE
'FACTOR POTENCIA
Range("L5").Select
  VALOR1 = 0.7
  VALOR2 = 1
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'NUM CONDUCTORES POR FASE
Range("P5").Select
  VALOR1 = 0
  VALOR2 = 3
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'TENSION NOMINAL
Range("T5").Select
  VALOR1 = 100
  VALOR2 = 240
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE


Range("E5").Select

For i = 5 To lastRow
     ' "CALIBRE CONDUTOR" Valores como fechas, indicar valores con números. Validar información
       
       If IsError(ActiveCell) Then ActiveCell = ""
       If IsDate(ActiveCell.Value) Then ActiveCell = "ND"
       If ActiveCell.Value = "" Then ActiveCell.Interior.Color = QBColor(12)
       ActiveCell.Offset(0, 2).Select 'CORRIENTE ENTREGADA R", "CORRIENTE ENTREGADA S", "CORRIENTE ENTREGADA T" valores en "0" y "NA" No permitidos
       j = 0
   For j = 0 To 2
       
       Select Case ActiveCell.Value
        Case "0"
            ActiveCell.Interior.Color = QBColor(12)
        Case "N/A"
            ActiveCell.Interior.Color = QBColor(12)
        Case ""
            ActiveCell.Interior.Color = QBColor(12)
        End Select
         ActiveCell.Offset(0, 1).Select
    Next j
    L = j + 2
    
   ActiveCell.Offset(0, 2).Select  '"   Atributo "FACTOR DE POTENCIA" Valores fuera de rango (0.8-0.95)
    Select Case ActiveCell.Value
        Case "0.8"
            ActiveCell.Interior.Color = QBColor(12)
        Case "0.95"
            ActiveCell.Interior.Color = QBColor(12)
        Case "0"
            ActiveCell.Interior.Color = QBColor(12)
        Case "N/A"
            ActiveCell.Interior.Color = QBColor(12)
        Case ""
            ActiveCell.Interior.Color = QBColor(12)
        End Select
        L = L + 2
   ActiveCell.Offset(0, 4).Select         '"   Atributo "NUMERO DE CONDUCTORES POR FASE" Validar valores mayores a 3
     If Not IsNumeric(ActiveCell) Then ActiveCell.Interior.Color = QBColor(12)
     If IsNumeric(ActiveCell.Value) Then
          If Int(ActiveCell.Value) > 3 Then ActiveCell.Interior.Color = QBColor(12)
     End If
     L = L + 4
  ActiveCell.Offset(0, 1).Select    '"   Atributo "NUMERO DE FASES" Valores "0" no permitidos
     If Not IsNumeric(ActiveCell) Then ActiveCell.Interior.Color = QBColor(12)
     If IsNumeric(ActiveCell.Value) Then
          If Int(ActiveCell.Value) = 0 Then ActiveCell.Interior.Color = QBColor(12)
     End If
  
   L = L + 1
     
      ActiveCell.Offset(0, -L).Select
      ActiveCell.Offset(1, 0).Select

 Next i
 
 

'********************************************************************************HOJA RED COMERCIAL ***************************************************
Sheets("RedComercial_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0
'"   La cantidad de registros no es igual a la cantidad de sitios reportados en la hoja "DATOS DEL SITIO"
'"   Atributo "TIPO RED ELECTRICA" y "VOLTAJE NOMINAL" Valores "0" y "NA" No permitidos
'VOLTAJE NOMINAL
Range("S5").Select
  VALOR1 = 11.4
  VALOR2 = 34.5
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE



Range("R5").Select
For i = 5 To lastRow
      Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
   ActiveCell.Offset(0, 1).Select
      Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
 ActiveCell.Offset(0, -1).Select
ActiveCell.Offset(1, 0).Select
Next i
'********************************************************************************HOJA INTERRUCTOR BAJANTE ***************************************************
Sheets("IntBaja_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0
'"   Atributo "CORRIENTE CORTO CIRCUITO" Valores en "0" no permitidos, Valor de "800 KA" Valido?
'"   Atributo "CORRIENTE NOMINAL" Valores "NA" no permitidos
Range("E5").Select 'CORRIENTE CORTO CIRCUITO

  VALOR1 = 8
  VALOR2 = 85
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
Range("F5").Select 'CORRIENTE NOMINAL
  VALOR1 = 6
  VALOR2 = 630
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
Range("P5").Select 'TENSION NOMINAL

  VALOR1 = 100
  VALOR2 = 240
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("E5").Select
For i = 5 To lastRow
 If IsError(ActiveCell) Then ActiveCell = ""
    If IsNumeric(ActiveCell.Value) Then
      If Int(ActiveCell.Value) >= 800 Then ActiveCell.Interior.Color = QBColor(12)
    End If

      Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
   ActiveCell.Offset(0, 1).Select
   If IsError(ActiveCell) Then ActiveCell = ""
      Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
 ActiveCell.Offset(0, -1).Select
ActiveCell.Offset(1, 0).Select
Next i
'********************************************************************************HOJA PARARRAYO***************************************************
Sheets("Pararrayo_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0
'------------------------------------------'Pararrayo--------------
'"   Atributo "CANTIDAD DE PARARRAYOS" Valores "0" y "NA" no permitidos. Validar sitio con "6" pararrayos.
'"   Atributo "CORRIENTE NOMINAL DESCARGA" Valores " 0" y "NA" No permitidos. Validar valor "110KA"
'"   Atributo "TENSIÓN NOMINAL" Valores "0" y celdas vacías no permitidos. Validar valores fuera de rango (11.4 a 13.2 KV)
'"   Atributo "TENSIÓN RESIDUAL" Valores "0" y celdas vacías no permitidas
Range("E5").Select '"CANTIDAD PARARRAYOS
  VALOR1 = 1
  VALOR2 = 3
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
Range("F5").Select 'CORRIENTE NOMINAL DESCARGA
  VALOR1 = 1
  VALOR2 = 25
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("N5").Select 'TENSION MAXIMA DE OPERACION
  VALOR1 = 11.4
  VALOR2 = 34.5
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
Range("O5").Select '
BUCLE
'--------------------
Range("E5").Select '"   Atributo "CANTIDAD DE PARARRAYOS" Valores "0" y "NA" no permitidos. Validar sitio con "6" pararrayos.
For i = 5 To lastRow
 If IsError(ActiveCell) Then ActiveCell = ""
    If IsNumeric(ActiveCell.Value) Then
      If Int(ActiveCell.Value) >= 6 Then ActiveCell.Interior.Color = QBColor(12)
    End If

      Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      
   ActiveCell.Offset(0, 1).Select '"   Atributo "CORRIENTE NOMINAL DESCARGA" Valores " 0" y "NA" No permitidos. Validar valor "110KA"
    If IsError(ActiveCell) Then ActiveCell = ""
    If IsNumeric(ActiveCell.Value) Then
      If Int(ActiveCell.Value) >= 110 Then ActiveCell.Interior.Color = QBColor(12)
    End If

      Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      L = 2
     ActiveCell.Offset(0, 9).Select '" '"   Atributo "TENSIÓN NOMINAL" Valores "0" y celdas vacías no permitidos. Validar valores fuera de rango (11.4 a 13.2 KV)
      If IsError(ActiveCell) Then ActiveCell = ""
    If IsNumeric(ActiveCell.Value) Then
         If CInt(ActiveCell.Value) < 11 Then ActiveCell.Interior.Color = QBColor(12)
    End If

      Select Case ActiveCell.Value
          
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      L = L + 9
     ActiveCell.Offset(0, 1).Select  'Atributo "TENSIÓN RESIDUAL" Valores "0" y celdas vacías no permitidas
            Select Case ActiveCell.Value
          
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
    
      
 ActiveCell.Offset(0, -L).Select
ActiveCell.Offset(1, 0).Select
Next i
'********************************************************************************HOJA TIERRA***************************************************
Sheets("Tierra_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0
'------------------------------------------Tierra
'"   Atributo "DIAMETRO CONDUCTOR" y "DIAMETRO VARILLA" Valores en "0" Y celdas vacias no permitidos.
'"   Atributo "MEDIDA PUESTA A TIERRA" valores fuera de rango (1-5 OHM)
Range("G5").Select 'DIAMETRO CONDUCTOR
  VALOR1 = 25
  VALOR2 = 150
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("O5").Select 'MEDIDA PUESTA A TIERRA
  VALOR1 = 0
  VALOR2 = 15
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("G5").Select '"  Atributo "DIAMETRO CONDUCTOR" y "DIAMETRO VARILLA" Valores en "0" Y celdas vacias no permitidos.
For i = 5 To lastRow
    If IsError(ActiveCell) Then ActiveCell = ""
        Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
               ActiveCell.Interior.Color = QBColor(12)
           Case ""
               ActiveCell.Interior.Color = QBColor(12)
      End Select
      
   ActiveCell.Offset(0, 1).Select '"
     If IsError(ActiveCell) Then ActiveCell = ""
        Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
               ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      L = 1
      
      
     ActiveCell.Offset(0, 7).Select 'Atributo "MEDIDA PUESTA A TIERRA" valores fuera de rango (1-5 OHM)
     
      If IsError(ActiveCell) Then ActiveCell = ""
    If IsNumeric(ActiveCell.Value) Then
         If CInt(ActiveCell.Value) > 1.5 Then ActiveCell.Interior.Color = QBColor(12)
    End If

      Select Case ActiveCell.Value
          
           Case "0"
               ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      L = L + 7
             
 ActiveCell.Offset(0, -L).Select
ActiveCell.Offset(1, 0).Select
Next i
'*********************************************************************************HOJA TRAFOR**********************************************
Sheets("Trafo_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0
'------------------------------------------Transferencia
'"   Atributo "CARGA ACTUAL" Valores "0" y celdas vacías no permitidos -
'"   Atributo "CLASE TRANSFORMADOR" (Seco o aceite) valores "NA" no permitidos-
'"   Atributo "NÚMERO DE FASES" Celdas vacías no permitidas
'"   Atributo "POTENCIA NOMINAL" Valores "0" no permitidos
Range("E5").Select 'CARGA ACTUAL
  VALOR1 = 0.01
  VALOR2 = 300
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("N5").Select 'OCUPACION
  VALOR1 = 0
  VALOR2 = 1
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
Range("O5").Select 'POTENCIA NOMINAL
  VALOR1 = 0
  VALOR2 = 300
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("U5").Select 'TENSION PRIMARIO
  VALOR1 = 11.4
  VALOR2 = 34.5
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
Range("V5").Select 'TENSION SECUNDARIO
  VALOR1 = 100
  VALOR2 = 240
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE


Range("E5").Select '"  Atributo "CARGA ACTUAL" Valores "0" y celdas vacías no permitidos
For i = 5 To lastRow
    If IsError(ActiveCell) Then ActiveCell = ""
        Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      
   ActiveCell.Offset(0, 1).Select '" "CLASE TRANSFORMADOR" (Seco o aceite) valores "NA" no permitidos
     If IsError(ActiveCell) Then ActiveCell = ""
        Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      L = 1
           
     ActiveCell.Offset(0, 6).Select '"NÚMERO DE FASES" Celdas vacías no permitidas
     
      If IsError(ActiveCell) Then ActiveCell = ""
  
      Select Case ActiveCell.Value
          
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      
      L = L + 6
      ActiveCell.Offset(0, 3).Select    '"   Atributo "POTENCIA NOMINAL" Valores "0" no permitidos
      
      If IsError(ActiveCell) Then ActiveCell = ""
        Select Case ActiveCell.Value
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      L = L + 3
 ActiveCell.Offset(0, -L).Select
ActiveCell.Offset(1, 0).Select
Next i
'*********************************************************************************HOJA AA**********************************************
Sheets("AA_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0
'-----------------------------------------'Aire Acondicionado
'"   Filas sin nombre del sitio ni código de Máximo--
'"   Atributo "CAPACIDAD TERMICA" Valores inconsistentes, validar.--
'"   Atributo "CARGA ACTUAL ESTIMADA" Valores fuera de rango (1-50 TR)
'"   Atributo "FACTOR DE EFICIENCIA" Valores fuera de rango (0.8 - 1)
'"   Atributo "MARCA" valores en "0" no permitidos
'"   Atributo "MARCA SISTEMA DE GESTIÓN" Valores "NA" y "SI" No permitidos
'"   Atributo "NÚMERO DE COMPRESORES" y "NÚMEROD E SALONES" Valores "0", "NA" y celdas vacías no permitidas
'"   Atributo "NUMERO DE UMAS" Valores "0" y celdas vacías no permitidos
'"   Atributo "NUMERO CONDENSADORAS" Valores fuera de rango (1-3)
'"   Atributo "POTENCIA CONSUMIDA" y "POTENCIA NOMINAL" Valores fuera de rango (1-50 KVA)
'"    Atributo "TEM SALON REFRIGERADO" Valores "NA" no permitidos. Valor 38°C correcto?

Range("E5").Select ' CAPACIDAD TERMICA
  VALOR1 = 0.1
  VALOR2 = 40
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("F5").Select 'CARGA ACTUAL ESTIMADA
BUCLE
Range("I5").Select 'FACTOR EFICIENCIA
  VALOR1 = 0.1
  VALOR2 = 3
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
Range("P5").Select 'NUMERO DE COMPRESORES
  VALOR1 = 1
  VALOR2 = 4
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("Q5").Select 'NUMERO DE SALONES
BUCLE

Range("R5").Select 'NUMERO DE UMAS
BUCLE

Range("S5").Select 'NUMERO_CONDENSADORAS
BUCLE

Range("U5").Select ' OCUPACION
  VALOR1 = 0
  VALOR2 = 1
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("V5").Select 'POTENCIA CONSUMIDA
  VALOR1 = 0.001
  VALOR2 = 10
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

Range("W5").Select '
BUCLE


Range("D5").Select 'Filas sin nombre del sitio ni código de Máximo
For i = 5 To lastRow
    If IsError(ActiveCell) Then ActiveCell = ""
        Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
          Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      
   ActiveCell.Offset(0, 1).Select ' Atributo "CAPACIDAD TERMICA" Valores inconsistentes, validar.
     If IsError(ActiveCell) Then ActiveCell = ""
     
        Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      
           
     ActiveCell.Offset(0, 1).Select '"CARGA ACTUAL ESTIMADA" Valores fuera de rango (1-50 TR)
     
      If IsError(ActiveCell) Then ActiveCell = ""
  
      Select Case ActiveCell.Value
          
           Case "NA"
              ActiveCell.Interior.Color = QBColor(12)
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
              Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      
      L = 3
      ActiveCell.Offset(0, 3).Select    '"   Atributo "FACTOR DE EFICIENCIA" Valores fuera de rango (0.8 - 1)
      
      If IsError(ActiveCell) Then ActiveCell = ""
     ' If Not IsNumeric(ActiveCell.Value) Then ActiveCell.Interior.Color = QBColor(12)
      If IsNumeric(ActiveCell.Value) Then
         If (ActiveCell.Value) > 3 Then ActiveCell.Interior.Color = QBColor(12)
         If (ActiveCell.Value) < 0.1 Then ActiveCell.Interior.Color = QBColor(12)
      End If
       Select Case ActiveCell.Value
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
            L = L + 3
            
 ActiveCell.Offset(0, 2).Select    '"   Atributo "MARCA" valores en "0" no permitidos
         If IsError(ActiveCell) Then ActiveCell = ""
      Select Case ActiveCell.Value
           Case "0"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
      L = L + 2
      
  ActiveCell.Offset(0, 1).Select '"   Atributo "MARCA SISTEMA DE GESTIÓN" Valores "NA" y "SI" No permitidos
         If IsError(ActiveCell) Then ActiveCell = ""
      Select Case ActiveCell.Value
           Case "SI"
              ActiveCell.Interior.Color = QBColor(12)
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
          Case "NA"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
              
      End Select
      L = L + 1
   ActiveCell.Offset(0, 4).Select '"   Atributo "NÚMERO DE COMPRESORES" y "NÚMEROD E SALONES" Valores "0", "NA" y celdas vacías no permitidas
       j = 0                              '"   Atributo "NUMERO DE UMAS" Valores "0" y celdas vacías no permitidos
   For j = 0 To 2
       If IsError(ActiveCell) Then ActiveCell = "" '"
      Select Case ActiveCell.Value
           Case "0"
               ActiveCell.Interior.Color = QBColor(12)
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
          Case "NA"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
        End Select
        ActiveCell.Offset(0, 1).Select
    Next j
     L = L + 7
                                '""   Atributo "NUMERO CONDENSADORAS" Valores fuera de rango (1-3)
      
       
      If IsError(ActiveCell) Then ActiveCell = "ND"
      'If Not IsNumeric(ActiveCell.Value) Then ActiveCell.Interior.Color = QBColor(12)
      If IsNumeric(ActiveCell.Value) Then
         If CInt(ActiveCell.Value) > 4 Then ActiveCell.Interior.Color = QBColor(12)
         If CInt(ActiveCell.Value) < 1 Then ActiveCell.Interior.Color = QBColor(12)
      End If
      Select Case ActiveCell.Value
           Case "0"
               ActiveCell.Interior.Color = QBColor(12)
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
          Case "NA"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
        End Select
          L = L + 1
    ActiveCell.Offset(0, 3).Select '"   Atributo "POTENCIA CONSUMIDA" y "POTENCIA NOMINAL" Valores fuera de rango (1-50 KVA)
       If IsError(ActiveCell) Then ActiveCell = "ND"
   j = 0
   For j = 0 To 1
      If IsNumeric(ActiveCell.Value) Then
         If Int(ActiveCell.Value) > 50 Then ActiveCell.Interior.Color = QBColor(12)
      End If
      Select Case ActiveCell.Value
           Case "0"
               ActiveCell.Interior.Color = QBColor(12)
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
          Case "NA"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
        End Select
        ActiveCell.Offset(0, 1).Select
    Next j
      L = L + 5
ActiveCell.Offset(0, 12).Select '""    Atributo "TEM SALON REFRIGERADO" Valores "NA" no permitidos. Valor 38°C correcto?
     If IsError(ActiveCell) Then ActiveCell = "ND"
     If IsNumeric(ActiveCell.Value) Then
         If Int(ActiveCell.Value) > 38 Then ActiveCell.Interior.Color = QBColor(12)
      End If
    Select Case ActiveCell.Value
           Case "0"
               ActiveCell.Interior.Color = QBColor(12)
           Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
          Case "NA"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell.Interior.Color = QBColor(12)
        End Select
       L = L + 10

 ActiveCell.Offset(0, -L).Select
ActiveCell.Offset(1, 0).Select
Next i

avance = 0.9
UpdateProgressBar avance

'*********************************************************************************HOJA UMA  **********************************************

Sheets("UMA_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0

'"   Atributo "CAPACIDAD NOMINAL UMA", "CONSUMO ACTUAL", "CORRIENTE CARGA R", "CORRIENTE CARGA S" , "CORRIENTE CARGA T" Valores en "0" no permitidos
'"   Validar sitio "PTO LEGUIZAMO" mayor cantidad de atributos en "NA"

Range("E5").Select 'Atributo "CAPACIDAD NOMINAL UMA", "CONSUMO ACTUAL", "CORRIENTE CARGA R", "CORRIENTE CARGA S" , "CORRIENTE CARGA T" Valores en "0" no permitidos
 For i = 5 To lastRow
     j = 0
      For j = 0 To 4

             If IsError(ActiveCell) Then ActiveCell = ""
               Select Case ActiveCell.Value
               Case "N/A"
              ActiveCell.Interior.Color = QBColor(12)
               Case "0"
              ActiveCell.Interior.Color = QBColor(12)
               Case ""
              ActiveCell.Interior.Color = QBColor(12)
      End Select
       ActiveCell.Offset(0, 1).Select
      Next j
  L = j
  
  ActiveCell.Offset(0, -L).Select
  ActiveCell.Offset(1, 0).Select
Next i

'*********************************************************************************HOJA Unidad Condensadora **********************************************
Sheets("UdadCondensad_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "D").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0 'Unidad Condensadora
'Atributo "CAPACIDAD NOMINAL" Valor "0" y "220" no permitidos
'Validar ocupaciones, se presentan valores por encima del 100%

Range("E5").Select 'Atributo "CAPACIDAD NOMINAL" Valor "0" y "220" no permitidos

 For i = 5 To lastRow
    
      
             If IsError(ActiveCell) Then ActiveCell = ""
               Select Case ActiveCell.Value
               Case "220"
              ActiveCell.Interior.Color = QBColor(12)
               Case "0"
               ActiveCell.Interior.Color = QBColor(12)
               Case ""
              ActiveCell.Interior.Color = QBColor(12)
           End Select
      
     
       
     ActiveCell.Offset(0, 17).Select
     If IsError(ActiveCell) Then ActiveCell = ""
      'ActiveCell.NumberFormat = "#,##0.000"
      If IsNumeric(ActiveCell.Value) Then
          If Int(ActiveCell.Value) > 1 Then ActiveCell = ActiveCell / 100
         
          End If
        Select Case ActiveCell.Value
        Case ""
             ActiveCell.Interior.Color = QBColor(12)
         End Select
       
       'ActiveCell.NumberFormat = "00%"
    L = 17
  ActiveCell.Offset(0, -L).Select
  ActiveCell.Offset(1, 0).Select
  
Next i


'*********************************************************************************HOJA UPS **********************************************
Sheets("UPS_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "D").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0 'Unidad Condensadora
'AUTONOMIA_POT

Range("E5").Select
  VALOR1 = 0
  VALOR2 = 60
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'CORRIENTE ENTREGADA R
Range("F5").Select
  VALOR1 = 0
  VALOR2 = 50
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'CORRIENTE ENTREGADA S

Range("G5").Select
BUCLE
'CORRIENTE ENTREGADA T
Range("H5").Select
BUCLE
'FACTOR EFICIENCIA
Range("J5").Select
  VALOR1 = 0.7
  VALOR2 = 1
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'FACTOR POTENCIA
Range("K5").Select
BUCLE
' OCUPACION
Range("Q5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'POTENCIA CARGA
Range("R5").Select
  VALOR1 = 0
  VALOR2 = 80
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'POTENCIA NOMINAL KVA

Range("S5").Select
  VALOR1 = 0
  VALOR2 = 100
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'TENSION BANCO BATERIAS
Range("Y5").Select
  VALOR1 = 12
  VALOR2 = 24
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'TENSION CARGA
Range("Z5").Select
  VALOR1 = 100
  VALOR2 = 240
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE
'TENSION ENTRADA AC
Range("AA5").Select
BUCLE
'VACANCIA
Range("Z5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = "0"
  L2 = "N/A"
  L3 = "NA"
  L4 = ""
  BUSCARLETRAS = "NO"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
BUCLE

avance = 0.9
UpdateProgressBar avance

'-----------------------------FIN----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
MsgBox "Corregido con exito: SE MARCO EN ROJO LOS ERRORES"
CORRECCION_INVENTARIO.Hide
'CIERRA MACRO AL FINALIZAR


End Sub


Private Sub CommandButton2_Click()
Unload Me
Load REVISION_CELDA
REVISION_CELDA.Show

End Sub

Private Sub CommandButton3_Click() '**++++++++++++++++++++++++++++++++CORREGIR INFORME*++++++++++++++++++++++++++++++++++++++++++++++

CORRECCION_INVENTARIO.LProgress.Width = 0
 
 avance = 0

UpdateProgressBar avance
Application.ScreenUpdating = False

'//////////////////////////////////////////////////////////////////////  HOJA  PEE_    *******************************************************

avance = 0.1
UpdateProgressBar avance


Sheets("PEE_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With

ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0

'Atributo "ALTURA DE INSTALACIÓN" con celdas vacías fecha 12/04/2018
 Range("E5").Select
 
  VALOR1 = 3
  VALOR2 = 3500
   L1 = ""
  L2 = "NA"
  L3 = "NO VISIBLE"
  L4 = "NO VERIFICABLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "3"
    BUCLE1
        
 
'No es valido el valor de "0" y "NA" en "CARGA ACTUAL KVA". Independientemente de si está apagado o prendido se requiere conocer cuánto puede soportar FECHA=12/04/2018
 Range("F5").Select
  VALOR1 = 0
  VALOR2 = 500
   L1 = ""
  L2 = "NA"
  L3 = "NO VISIBLE"
  L4 = "NO VERIFICABLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
  '-----------------"CORRIENTE DERRATEADA Y CORRIENTE NOMINAL".
  Range("G5").Select
  VALOR1 = 5
  VALOR2 = 350
  L1 = ""
  L2 = "NA"
  L3 = "NO VISIBLE"
  L4 = "NO VERIFICABLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
      BUCLE1
      
   Range("H5").Select '-----------------"CORRIENTE DERRATEADA Y
    VALOR1 = 5
  VALOR2 = 350
   L1 = ""
  L2 = "NA"
  L3 = "NO VISIBLE"
  L4 = "NO VERIFICABLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
      BUCLE1
 
'Atributo "FACTOR POTENCIA NOMINAL" valores en "0" no son permitidos
Range("J5").Select
  VALOR1 = 0.6
  VALOR2 = 0.9
    L1 = ""
  L2 = "NA"
  L3 = "NO VISIBLE"
  L4 = "NO VERIFICABLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
    
'Atributo "FACTOR DE EFICIENCIA" valores en "0" no son permitidos
  Range("K5").Select
  VALOR1 = 0.7
  VALOR2 = 0.9
  L1 = ""
  L2 = "NA"
  L3 = "NO VISIBLE"
  L4 = "NO VERIFICABLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0.7"
    BUCLE1

'Atributo "HORAS DE TRABAJO" Valores en "0" y "NA" no permitidos
Range("N5").Select
  VALOR1 = 0
  VALOR2 = 1000000
  L1 = ""
  L2 = "NA"
  L3 = "NO VISIBLE"
  L4 = "NO VERIFICABLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'SI LA MARCA  , MARCA PANEL CONTROL, MODELO  ESTAN EN N/A O N/V CAMBIAR POR NO VISIBLE

Range("Q5").Select
  VALOR1 = 0
  VALOR2 = 0
  L1 = ""
  L2 = "NA"
  L3 = "NO VISIBLE"
  L4 = "NO VERIFICABLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "GENERICO"
    BUCLE1

'SI LA MARCA  , MARCA PANEL CONTROL, MODELO  ESTAN EN N/A O N/V CAMBIAR POR NO  VISIBLE 12/04/2018
Range("R5").Select

  L1 = ""
  L2 = "0"
  L3 = "N/A"
  L4 = "NA"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1


'-------------------------------------------RUIDO---------------
                  'Atributo "NIVEL RUIDO

Range("U5").Select
  VALOR1 = 30
  VALOR2 = 100
  L1 = "0"
  L2 = "NA"
  L3 = ""
  L4 = "NO VERIFICABLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
                   'Atributo "OCUPACION
Range("W5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = " "
  L2 = "NA"
  L3 = ""
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'si no hay POTENCIA DERRATEADA KVA esta sera igual a POTENCIA NOMINAL KVA

Range("Y5").Select

 VALOR1 = 0
 VALOR2 = 500
   L1 = "NO VERIFICABLE"
  L2 = "NA"
  L3 = ""
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1


'Atributo "POTENCIA NOMINAL" Valores en "0" y "NA" no permitidos   ------completar error valor

Range("AB5").Select
 VALOR1 = 0
 VALOR2 = 500
  L1 = "NO VERIFICABLE"
  L2 = "NA"
  L3 = ""
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

' Atributo "TENSIÓN NOMINAL" Valores en "0" no permitidos FECHA
Range("AK5").Select
  VALOR1 = 100
  VALOR2 = 240
  L1 = "NO VERIFICABLE"
  L2 = "NA"
  L3 = ""
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "100"
    BUCLE1
 
'VACANCIA 0-100 DE LO CONTRARIO ROJO
Range("AN5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = " "
  L2 = "NA"
  L3 = ""
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

avance = 0.2
UpdateProgressBar avance

'***************************************************************************HOJA TANQUE_COMBUSTIBLE-----------------------------------------------------------------
Sheets("TanqCombustible_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With

ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0

'"Atributo "ALTURA TANQUE cm" Valores "0" Y "NA" no permitidos. Valores de 1.8 cm?

Range("E5").Select
  VALOR1 = 40
  VALOR2 = 300
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'"   Atributo "ANCHO TANQUE cm" Valores "0" Y "NA" no permitidos. Valores de 2.9 cm?
Range("F5").Select
VALOR1 = 40
VALOR2 = 200
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

' "   Atributo "AUTONOMIA ESTIMADA" Valores mayores a 200 h
Range("G5").Select
VALOR1 = 0
VALOR2 = 300
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1


'"   Atributo "CAPACIDAD DEL TANQUE" valores en "0" no permitidos
Range("H5").Select
  VALOR1 = 0
  VALOR2 = 1000
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
 
'"   Atributo "CIECUNFERENCIA DEL TANQUE cm" Valores elevados ejemplo: 19100.
Range("I5").Select

  VALOR1 = 0
  VALOR2 = 150
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
    
'"   Atributo "LARGO DEL TANQUE cm" Valores elevados ejemplo: 19100.
Range("N5").Select

  VALOR1 = 0
  VALOR2 = 300
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
    
'"   Atributo


avance = 0.3
UpdateProgressBar avance
'********************************************************************************HOJA *MOTOR******************--------------------------------
Sheets("Motor_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0

'"   Atributo "VOLUMEN ACTUAL COMBUSTUIBLE" Valores "NA" no permitidos
Range("E5").Select
  VALOR1 = 0
  VALOR2 = 500
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
'"   Atributo CONSUMO NOMINAL

Range("F5").Select
  VALOR1 = 0
  VALOR2 = 50
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1

   '"Atributo "CONSUMO PROMEDIO DE CARGA" VLORES EN "0", se encuentra un valor de "401 Gl/h" es válido?  no pueden ser superiores a 10 gl hora

Range("G5").Select

  VALOR1 = 0
  VALOR2 = 50
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
'"   Atributo "NIVEL DE RUIDO" Valores en "0" y "1"
Range("O5").Select

  VALOR1 = 30
  VALOR2 = 100
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "0"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1

'"   Atributo "POTENCIA DERRATEADA" Valores en "0" y "800"
Range("P5").Select

  VALOR1 = 0
  VALOR2 = 800
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "0"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1

'"   Atributo "POTENCIA NOMINAL" Valores en "NA"

Range("Q5").Select

  VALOR1 = 0
  VALOR2 = 1000
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "0"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
'"   Atributo "TIPO DE COMBUSTIBLE" Valores en "NA"



'"   Atributo "TIPO MOTOR" valores que no coinciden (T:39201912, B, L)
Range("X5").Select

  VALOR1 = 0
  VALOR2 = 0
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "0"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
avance = 0.4
UpdateProgressBar avance
'**********************************************************************HOJA GENERADROR******************************************************
Sheets("Generador_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(655, 22)).Interior.ColorIndex = 0

'"   Atributo "MARCA" Valores en "0", "NA" y Vacías. Valores no permitidos
 Range("F5").Select
 
For i = 5 To lastRow
If IsError(ActiveCell) Then ActiveCell = ""
If ActiveCell.Value = 0 Then ActiveCell = "ND"
If ActiveCell.Value = "" Then ActiveCell = "ND"
If ActiveCell.Value = "N/A" Then ActiveCell = "ND"
ActiveCell.Offset(1, 0).Select
Next i
'"   Atributo "POTENCIA ACTIVA NOMINAL" Valores "0" y "NA" no permitidos

 Range("G5").Select
  VALOR1 = 0
  VALOR2 = 0
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "0"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
   BUCLE1
'-
Range("H5").Select
BUCLE1

Range("I5").Select
BUCLE1


 Range("J5").Select
For i = 5 To lastRow
If IsError(ActiveCell) Then ActiveCell = ""
If ActiveCell.Value = 0 Then ActiveCell = "ND"
If ActiveCell.Value = "" Then ActiveCell = "ND"
If ActiveCell.Value = "N/A" Then ActiveCell = "ND"
ActiveCell.Offset(1, 0).Select
Next i

'"   Atributo "POTENCA DERRATEADA NOMINAL" Valores en "0". ¿Valor de 1361 KVA Válido?  nota : la derrateada es menor a  la nominal


'"   Atributo "POTENCIA NOMINAL KVA" Valores en "0" no permitidos
 Range("L5").Select
For i = 5 To lastRow
If IsError(ActiveCell) Then ActiveCell = ""
If ActiveCell.Value = 0 Then ActiveCell = "ND"
If ActiveCell.Value = "" Then ActiveCell = "ND"
If ActiveCell.Value = "N/A" Then ActiveCell = "ND"
ActiveCell.Offset(1, 0).Select
Next i

'"   Atributo "TENSIÓN SALIDA ENTRE FASES" Valores en "0" y "NA" No permitidos

 Range("P5").Select
For i = 5 To lastRow
If IsError(ActiveCell) Then ActiveCell = ""
If ActiveCell.Value = 0 Then ActiveCell = "ND"
If ActiveCell.Value = "" Then ActiveCell = "ND"
If ActiveCell.Value = "N/A" Then ActiveCell = "ND"
ActiveCell.Offset(1, 0).Select
Next i





 
avance = 0.5
UpdateProgressBar avance
'**********************************************************************HOJA Bateria_Arranque_******************************************************

Sheets("BatArranque_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 22)).Interior.ColorIndex = 0

'"   Atributo "MARCA BATERIA DE ARRANQIE" Valores "NA" No permitidos

 Range("G5").Select
For i = 5 To lastRow
If IsError(ActiveCell) Then ActiveCell = ""
If ActiveCell.Value = 0 Then ActiveCell = "ND"
If ActiveCell.Value = "" Then ActiveCell = "ND"
If ActiveCell.Value = "N/D" Then ActiveCell = "ND"
If ActiveCell.Value = "N/A" Then ActiveCell = "ND"
If ActiveCell.Value = "NA" Then ActiveCell = "ND"
ActiveCell.Offset(1, 0).Select

Next i

'"   Atributo "TENSIÓN BATERIA DE ARRANQUE" Valores en "0" y menores a "9.5" V ESTARIA MAL

 Range("P5").Select
 
  VALOR1 = 11
  VALOR2 = 25
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "0"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1

Range("Q5").Select
  VALOR1 = 11
  VALOR2 = 25
   L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "0"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1

'"   Atributo "TENSIÓN CAGADOR MOTOR" Valores en "0" y "NA" No permitidos
 Range("R5").Select
  VALOR1 = 11
  VALOR2 = 31
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "0"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
   
   avance = 0.6
UpdateProgressBar avance
  '********************************************************************************HOJA TRANFERENCIA***************************************************
Sheets("Transferencia_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 27)).Interior.ColorIndex = 0

'"  Atributo "CARGA ACTUAL" Valores en "0" y "NA" no validos
  Range("E5").Select
  VALOR1 = 0
  VALOR2 = 500
   L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
    
  Range("H5").Select
  VALOR1 = 0
  VALOR2 = 0
   L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
    
  Range("I5").Select
    BUCLE1
 Range("J5").Select
    BUCLE1
 Range("K5").Select
    BUCLE1
 Range("L5").Select
    BUCLE1
'"  Atributo "POTENCIA NOMINAL" Valores en "0" y "NA" no validos

Range("N5").Select
  VALOR1 = 0
  VALOR2 = 300
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'"   Atributo "TENSIÓN NOMINAL FASE-FASE" y "TENSIÓN NOMINAL FASE- NEUTRO" Valores en "0" y "NA" No disponible
  Range("V5").Select
     VALOR1 = 100
     VALOR2 = 240
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

 Range("W5").Select
     VALOR1 = 115
     VALOR2 = 125
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
    
 avance = 0.7
UpdateProgressBar avance
'********************************************************************************HOJA POWER***************************************************
Sheets("Power_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 46)).Interior.ColorIndex = 0
'"   Atributo "CANTIDAD DE SLOTS LIBRES" Valores negativos, "NA"  y errores de formula.
 Range("E5").Select
     VALOR1 = 0
     VALOR2 = 30
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'"  Atributo "CARGA ACTUAL", "CARGA DC", "CORRIENTE CARGA", "CORRIENTE ENTRADA R", "CORRIENTE ENTRADA S", "CORRIENTE ENTRADA T" Valores en "0" , texto y errores de formula.

 Range("N5").Select
  VALOR1 = 0
  VALOR2 = 30
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
                '"CARGA DC"
 Range("O5").Select
  VALOR1 = 0
  VALOR2 = 300
  L1 = ""
  L2 = "N/A"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
    
Range("P5").Select  'CORRIENTE CARGA ADC (Deben ser los mismos valores de la columna anterior ya que corresponden al mismo atributo)
 

For i = 5 To lastRow

ActiveCell = Cells(ActiveCell.Row, 15)
ActiveCell.Offset(1, 0).Select
Next i

Range("Q5").Select   '"   Atributo "CORRIENTE ENTRADA R
  VALOR1 = 0
  VALOR2 = 200
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
Range("R5").Select   '"   Atributo "CORRIENTE ENTRADA S
    BUCLE1
Range("S5").Select   '"   Atributo "CORRIENTE ENTRADA T
    BUCLE1
    
'"   Atributo "CORRIENTE NOMINAL MODULO" Valores en "0", "valores por encima de "1000" validos?
Range("T5").Select
VALOR1 = 0
VALOR2 = 104
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'"   Atributo "FACTOR DE EFICIENCIA DEL MODULO" Valores menores a 0.8
Range("V5").Select

VALOR1 = 0
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'"   Atributo " NÚMEROS DE MÓDULOS INSTALADOS" Valores en texto y error de formula.
Range("AC5").Select
If IsError(ActiveCell) Then ActiveCell = ""
  VALOR1 = 0
  VALOR2 = 20
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
            '" Atributo OCUPACION
Range("AE5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'"   Atributo "POTENCIA NÓMINAL MÓDULO" Valores menores a 1500 y mayores a 3000. Validar
Range("AF5").Select
  VALOR1 = 0
  VALOR2 = 5
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'" Atributo "TENSIÓN DE CARGA",
Range("AN5").Select
  VALOR1 = 24
  VALOR2 = 54
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'" Atributo "TENSION ENTRADA ENTRE LINEAS
Range("AO5").Select
  VALOR1 = 100
  VALOR2 = 240
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
  
''" Atributo  "TENSION ENTRADA ENTRE LINEAS", "TENSION FLOTACION",
Range("AP5").Select
  VALOR1 = 47
  VALOR2 = 60
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
       ''" Atributo TENSION NOMINAL
Range("AQ5").Select
  VALOR1 = 47
  VALOR2 = 60
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'  VACANCIA

Range("AT5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1



avance = 0.8
UpdateProgressBar avance
'********************************************************************************HOJA BANCO DE BATERIAS***************************************************
Sheets("BcoBaterias_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 57)).Interior.ColorIndex = 0
'"   Atributo "ALTO CELDA",  "LARGO CELDA" Y " NUMERO CELDAS BANCO" Valores en "0" no permitidos
Range("E5").Select
  VALOR1 = 20
  VALOR2 = 80
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
'"ANCHO CELDA",

Range("F5").Select
  VALOR1 = 10
  VALOR2 = 40
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
'AUTONOMIA ESTIMADA
Range("G5").Select
  VALOR1 = 0
  VALOR2 = 500
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'CAPACIDAD BANCO

Range("H5").Select
  VALOR1 = 0
  VALOR2 = 1000
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'LARGO CELDA
Range("K5").Select
  VALOR1 = 20
  VALOR2 = 80
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
' NUMERO CELDAS BANCO
Range("N5").Select
  VALOR1 = 4
  VALOR2 = 24
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
'   PESO
Range("Q5").Select
  VALOR1 = 0
  VALOR2 = 500
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
' TENSION_FLOTACION_BATERIA
Range("Y5").Select
  VALOR1 = 47
  VALOR2 = 55
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'TENSION_IGUALACION_BATERIA
Range("Z5").Select
  VALOR1 = 48
  VALOR2 = 60
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
'TENSION_NOMINAL_BANCO
Range("AA5").Select
  VALOR1 = 47
  VALOR2 = 55
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'TENSION_NOMINAL_BATERIA

Range("AB5").Select
  VALOR1 = 11
  VALOR2 = 14
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'*********************************************************************************HOJA UPS **********************************************
Sheets("UPS_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "D").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0 'Unidad Condensadora
'AUTONOMIA_POT

Range("E5").Select
  VALOR1 = 0
  VALOR2 = 60
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'CORRIENTE ENTREGADA R
Range("F5").Select
  VALOR1 = 0
  VALOR2 = 50
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'CORRIENTE ENTREGADA S
Range("G5").Select
 BUCLE1
'CORRIENTE ENTREGADA T
Range("H5").Select
 BUCLE1
'FACTOR EFICIENCIA
Range("J5").Select
  VALOR1 = 0.7
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
'FACTOR POTENCIA
Range("K5").Select
 BUCLE1
' OCUPACION
Range("Q5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'POTENCIA CARGA
Range("R5").Select
  VALOR1 = 0
  VALOR2 = 80
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'POTENCIA NOMINAL KVA

Range("S5").Select
  VALOR1 = 0
  VALOR2 = 100
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'TENSION BANCO BATERIAS
Range("Y5").Select
  VALOR1 = 12
  VALOR2 = 24
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

'TENSION CARGA
Range("Z5").Select
  VALOR1 = 100
  VALOR2 = 240
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'TENSION ENTRADA AC
Range("AA5").Select
 BUCLE1
 
'VACANCIA
Range("AE5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

avance = 0.82
UpdateProgressBar avance
'********************************************************************************HOJA Inversor_*****************************************

Sheets("Inversor_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 67)).Interior.ColorIndex = 0

Range("E5").Select 'CORRIENTE ENTREGADA R
  VALOR1 = 0
  VALOR2 = 50
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
Range("F5").Select 'CORRIENTE ENTREGADA S
    BUCLE1
Range("G5").Select 'CORRIENTE ENTREGADA T
    BUCLE1
'FACTOR EFICIENCIA
Range("I5").Select
  VALOR1 = 0.7
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
'OCUPACION
Range("P5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'POTENCIA CARGA
Range("Q5").Select
  VALOR1 = 0
  VALOR2 = 100
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'POTENCIA NOMINAL
Range("R5").Select
    BUCLE1
'TENSION ENTRADA DC
Range("X5").Select
  VALOR1 = 48
  VALOR2 = 60
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'TENSION SALIDA AC
Range("Y5").Select
  VALOR1 = 100
  VALOR2 = 240
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'VACANCIA
Range("Z5").Select
  VALOR1 = 0
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

avance = 0.84
UpdateProgressBar avance

'********************************************************************************HOJA TABLERO ELECTRICO ***************************************************
Sheets("TabElectrico_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 157)).Interior.ColorIndex = 0
'"   Atributo "CANTIDAD DE CIRCUITOS INSTALADOS" Valores "0" y "NA" no permitidos
'"   Atributo "CANTIDAD POSICIONES LIBRES" Valores "NA" no permitidos
'"   Atributo "CARGA TOTAL" Valores en "0" y "NA" No permitidos
'"   Atributo  "CORRIENTE NOMINAL" Valores "NA" Y error de formula
'"   Atributo "TENSIÓN NOMINLA" Valores en "0" no permitidos

'"   Atributo "CANTIDAD DE CIRCUITOS
Range("E5").Select
  VALOR1 = 1
  VALOR2 = 20
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1


'"   Atributo CANTIDAD POSICIONES LIBRES

Range("F5").Select
  VALOR1 = 1
  VALOR2 = 10
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
' CARGA TOTAL
Range("G5").Select
  VALOR1 = 0
  VALOR2 = 200
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'CORRIENTE NOMINAL
Range("I5").Select
  VALOR1 = 0
  VALOR2 = 500
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
' TENSION NOMINAL
Range("W5").Select
  VALOR1 = 100
  VALOR2 = 240
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1


avance = 0.86
UpdateProgressBar avance

'********************************************************************************HOJA ACOMETIDA PRINCIPAL ***************************************************
Sheets("AcomPpal_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 127)).Interior.ColorIndex = 0
'"   Atributo "CALIBRE CONDUTOR" Valores como fechas, indicar valores con números. Validar información
'"   Atributo "CORRIENTE ENTREGADA R", "CORRIENTE ENTREGADA S", "CORRIENTE ENTREGADA T" valores en "0" y "NA" No permitidos
'"   Atributo "FACTOR DE POTENCIA" Valores fuera de rango (0.8-0.95)
'"   Atributo "NUMERO DE CONDUCTORES POR FASE" Validar valores mayores a 3
'"   Atributo "NUMERO DE FASES" Valores "0" no permitidos
'CORRIENTE ENTREGADA R

Range("G5").Select
  VALOR1 = 0
  VALOR2 = 300
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'CORRIENTE ENTREGADA S
Range("H5").Select
 BUCLE1
'CORRIENTE ENTREGADA T
Range("I5").Select
 BUCLE1
'FACTOR POTENCIA
Range("L5").Select
  VALOR1 = 0.7
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
'NUM CONDUCTORES POR FASE
Range("P5").Select
  VALOR1 = 0
  VALOR2 = 3
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'TENSION NOMINAL
Range("T5").Select
  VALOR1 = 100
  VALOR2 = 240
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

avance = 0.87
UpdateProgressBar avance


'********************************************************************************HOJA RED COMERCIAL ***************************************************
Sheets("RedComercial_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(536, 127)).Interior.ColorIndex = 0
'"   La cantidad de registros no es igual a la cantidad de sitios reportados en la hoja "DATOS DEL SITIO"
'"   Atributo "TIPO RED ELECTRICA" y "VOLTAJE NOMINAL" Valores "0" y "NA" No permitidos
'VOLTAJE NOMINAL
Range("S5").Select
  VALOR1 = 11.4
  VALOR2 = 34.5
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1


avance = 0.88
UpdateProgressBar avance
 
 
'********************************************************************************HOJA INTERRUCTOR BAJANTE ***************************************************
Sheets("IntBaja_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 127)).Interior.ColorIndex = 0
'"   Atributo "CORRIENTE CORTO CIRCUITO" Valores en "0" no permitidos, Valor de "800 KA" Valido?
'"   Atributo "CORRIENTE NOMINAL" Valores "NA" no permitidos
Range("E5").Select 'CORRIENTE CORTO CIRCUITO

  VALOR1 = 8
  VALOR2 = 85
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
Range("F5").Select 'CORRIENTE NOMINAL
  VALOR1 = 6
  VALOR2 = 630
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
Range("P5").Select 'TENSION NOMINAL

  VALOR1 = 100
  VALOR2 = 240
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1


avance = 0.89
UpdateProgressBar avance
 '********************************************************************************HOJA IntMedia ***********************************
 Sheets("IntMedia_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 127)).Interior.ColorIndex = 0
 'CORRIENTE CORTO CIRCUITO
Range("E5").Select

  VALOR1 = 1
  VALOR2 = 25
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'CORRIENTE NOMINAL
Range("F5").Select
  VALOR1 = 6
  VALOR2 = 630
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
'TENSION DE AISLAMIENTO
Range("P5").Select
  VALOR1 = 1
  VALOR2 = 17.5
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
'TENSION NOMINAL
Range("Q5").Select
  VALOR1 = 11.4
  VALOR2 = 34.5
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
avance = 0.9
UpdateProgressBar avance
 
  
'********************************************************************************HOJA PARARRAYO***************************************************
Sheets("Pararrayo_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 99)).Interior.ColorIndex = 0
'------------------------------------------'Pararrayo--------------
'"   Atributo "CANTIDAD DE PARARRAYOS" Valores "0" y "NA" no permitidos. Validar sitio con "6" pararrayos.
'"   Atributo "CORRIENTE NOMINAL DESCARGA" Valores " 0" y "NA" No permitidos. Validar valor "110KA"
'"   Atributo "TENSIÓN NOMINAL" Valores "0" y celdas vacías no permitidos. Validar valores fuera de rango (11.4 a 13.2 KV)
'"   Atributo "TENSIÓN RESIDUAL" Valores "0" y celdas vacías no permitidas
Range("E5").Select '"CANTIDAD PARARRAYOS
  VALOR1 = 1
  VALOR2 = 5
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
Range("F5").Select 'CORRIENTE NOMINAL DESCARGA
  VALOR1 = 1
  VALOR2 = 25
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

Range("N5").Select 'TENSION MAXIMA DE OPERACION
  VALOR1 = 11.4
  VALOR2 = 34.5
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
' TENSION NOMINAL
Range("O5").Select
    BUCLE1
avance = 0.91
UpdateProgressBar avance


'********************************************************************************HOJA TIERRA***************************************************
Sheets("Tierra_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 99)).Interior.ColorIndex = 0

'*****************************************************-Tierra *****************************************
'"   Atributo "DIAMETRO CONDUCTOR" y "DIAMETRO VARILLA" Valores en "0" Y celdas vacias no permitidos.
'"   Atributo "MEDIDA PUESTA A TIERRA" valores fuera de rango (1-5 OHM)
Range("G5").Select 'DIAMETRO CONDUCTOR
  VALOR1 = 25
  VALOR2 = 150
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
    Range("H5").Select 'DIAMETRO CONDUCTOR
  VALOR1 = 0
  VALOR2 = 150
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1

Range("O5").Select 'MEDIDA PUESTA A TIERRA
  VALOR1 = 0.1
  VALOR2 = 15
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1


avance = 0.92
UpdateProgressBar avance
 
'*********************************************************************************HOJA TRAFOR**********************************************
Sheets("Trafo_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 99)).Interior.ColorIndex = 0
'------------------------------------------Transferencia
'"   Atributo "CARGA ACTUAL" Valores "0" y celdas vacías no permitidos -
'"   Atributo "CLASE TRANSFORMADOR" (Seco o aceite) valores "NA" no permitidos-
'"   Atributo "NÚMERO DE FASES" Celdas vacías no permitidas
'"   Atributo "POTENCIA NOMINAL" Valores "0" no permitidos
Range("E5").Select 'CARGA ACTUAL
  VALOR1 = 0.01
  VALOR2 = 300
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
    
Range("N5").Select 'OCUPACION
  VALOR1 = 0
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
    
Range("O5").Select 'POTENCIA NOMINAL
  VALOR1 = 0.01
  VALOR2 = 300
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

Range("U5").Select 'TENSION PRIMARIO
  VALOR1 = 11.4
  VALOR2 = 34.5
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
Range("V5").Select 'TENSION SECUNDARIO
  VALOR1 = 100
  VALOR2 = 240
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1
Range("Z5").Select 'TENSION SECUNDARIO
  VALOR1 = 0
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

avance = 0.93
UpdateProgressBar avance


'*********************************************************************************HOJA AA**********************************************
Sheets("AA_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 99)).Interior.ColorIndex = 0
'-----------------------------------------'Aire Acondicionado
'"   Filas sin nombre del sitio ni código de Máximo--
'"   Atributo "CAPACIDAD TERMICA" Valores inconsistentes, validar.--
'"   Atributo "CARGA ACTUAL ESTIMADA" Valores fuera de rango (1-50 TR)
'"   Atributo "FACTOR DE EFICIENCIA" Valores fuera de rango (0.8 - 1)
'"   Atributo "MARCA" valores en "0" no permitidos
'"   Atributo "MARCA SISTEMA DE GESTIÓN" Valores "NA" y "SI" No permitidos
'"   Atributo "NÚMERO DE COMPRESORES" y "NÚMEROD E SALONES" Valores "0", "NA" y celdas vacías no permitidas
'"   Atributo "NUMERO DE UMAS" Valores "0" y celdas vacías no permitidos
'"   Atributo "NUMERO CONDENSADORAS" Valores fuera de rango (1-3)
'"   Atributo "POTENCIA CONSUMIDA" y "POTENCIA NOMINAL" Valores fuera de rango (1-50 KVA)
'"    Atributo "TEM SALON REFRIGERADO" Valores "NA" no permitidos. Valor 38°C correcto?

Range("E5").Select ' CAPACIDAD TERMICA
  VALOR1 = 0.1
  VALOR2 = 40
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

Range("F5").Select 'CARGA ACTUAL ESTIMADA
 BUCLE1
Range("I5").Select 'FACTOR EFICIENCIA
  VALOR1 = 0.1
  VALOR2 = 3
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "ND"
    BUCLE1
    
Range("P5").Select 'NUMERO DE COMPRESORES
  VALOR1 = 1
  VALOR2 = 4
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

Range("Q5").Select 'NUMERO DE SALONES
 BUCLE1

Range("R5").Select 'NUMERO DE UMAS
 BUCLE1

Range("S5").Select 'NUMERO_CONDENSADORAS
 BUCLE1

Range("U5").Select ' OCUPACION
  VALOR1 = 0
  VALOR2 = 1
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

Range("V5").Select 'POTENCIA CONSUMIDA
  VALOR1 = 0.001
  VALOR2 = 10
  L1 = ""
  L2 = "NA"
  L3 = "NO VERIFICABLE"
  L4 = "NO VISIBLE"
  MODIFICA = "SI"
  TEXTO_CELDA = "0"
    BUCLE1

Range("W5").Select '
 BUCLE1




avance = 0.99
UpdateProgressBar avance



'********************************************************************************HOJA PROTECCIONES ***************************************************
Sheets("Protecciones_").Select
 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 27)).Interior.ColorIndex = 0
'"   Atributos "CORRIENTE IMPULSO", "CORRIENTE MAXIMA", "CORRIENTE NOMINAL" Valores en 0 no permitidos. Valores de 1000 A correctos?
'"   Atributo "REFERENCIA" Celdas vacías.
'"   Atributo "TENSIÓN NOMINAL" Valores "0" No permitidos
'"   Atributo "TIPO PROTECCIÓN"  Valores (B, 150VAC, i IP20 CE)
Range("E5").Select
For i = 5 To lastRow
j = 0
   For j = 0 To 2  '"CORRIENTE IMPULSO", "CORRIENTE MAXIMA", "CORRIENTE NOMINAL"
       If IsError(ActiveCell) Then ActiveCell = ""
          If IsNumeric(ActiveCell.Value) Then
          If Int(ActiveCell.Value) > 1000 Then ActiveCell.Interior.Color = QBColor(12)
          End If
         Select Case ActiveCell.Value
           Case "N/A"
              ActiveCell = "ND"
           Case "0"
              ActiveCell = "ND"
           Case ""
              ActiveCell = "ND"
           Case "NO VERIFICABLE"
              ActiveCell = "ND"
              
         End Select
      ActiveCell.Offset(0, 1).Select
   Next j
    L = j
    ActiveCell.Offset(0, 7).Select '"   Atributo "REFERENCIA" Celdas vacías.
      Select Case ActiveCell.Value
           Case "0"
              ActiveCell = "ND"
           Case ""
              ActiveCell = "ND"
          Case "NO VERIFICABLE"
              ActiveCell = "ND"
         End Select
     L = L + 7
     ActiveCell.Offset(0, 2).Select '"   Atributo "TENSIÓN NOMINAL" Valores "0" No permitidos
       If IsError(ActiveCell) Then ActiveCell = ""
     
       Select Case ActiveCell.Value
           Case "0"
              ActiveCell = "ND"
           Case ""
              ActiveCell = "ND"
          Case "NO VERIFICABLE"
              ActiveCell = "ND"
         End Select
      L = L + 2
      ActiveCell.Offset(0, 3).Select '"   Atributo "TIPO PROTECCIÓN"  Valores (B, 150VAC, i IP20 CE)
      If IsError(ActiveCell) Then ActiveCell = ""
       Select Case ActiveCell.Value
           Case "B"
              ActiveCell.Interior.Color = QBColor(12)
           Case "150VAC"
              ActiveCell.Interior.Color = QBColor(12)
           Case "IP20"
              ActiveCell.Interior.Color = QBColor(12)
           Case "CE"
              ActiveCell.Interior.Color = QBColor(12)
           Case ""
              ActiveCell = "ND"
          Case "NO VERIFICABLE"
              ActiveCell = "ND"
         End Select
      L = L + 3
      ActiveCell.Offset(0, -L).Select
      ActiveCell.Offset(1, 0).Select
Next i

'*********************************************************************************HOJA UMA  **********************************************

Sheets("UMA_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0

'"   Atributo "CAPACIDAD NOMINAL UMA", "CONSUMO ACTUAL", "CORRIENTE CARGA R", "CORRIENTE CARGA S" , "CORRIENTE CARGA T" Valores en "0" no permitidos
'"   Validar sitio "PTO LEGUIZAMO" mayor cantidad de atributos en "NA"

Range("E5").Select 'Atributo "CAPACIDAD NOMINAL UMA", "CONSUMO ACTUAL", "CORRIENTE CARGA R", "CORRIENTE CARGA S" , "CORRIENTE CARGA T" Valores en "0" no permitidos
 For i = 5 To lastRow
     j = 0
      For j = 0 To 4

             If IsError(ActiveCell) Then ActiveCell = ""
               Select Case ActiveCell.Value
               Case "N/A"
              ActiveCell = "ND"
               Case "0"
              ActiveCell = "ND"
               Case "NO VERIFICABLE"
              ActiveCell = "ND"
      End Select
       ActiveCell.Offset(0, 1).Select
      Next j
  L = j
  
  ActiveCell.Offset(0, -L).Select
  ActiveCell.Offset(1, 0).Select
Next i

'*********************************************************************************HOJA Unidad Condensadora **********************************************
Sheets("UdadCondensad_").Select

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "D").End(xlUp).Row
  End With
ActiveSheet.Range(Cells(5, 1), Cells(65536, 40)).Interior.ColorIndex = 0 'Unidad Condensadora
'Atributo "CAPACIDAD NOMINAL" Valor "0" y "220" no permitidos
'Validar ocupaciones, se presentan valores por encima del 100%

Range("E5").Select 'Atributo "CAPACIDAD NOMINAL" Valor "0" y "220" no permitidos

 For i = 5 To lastRow
    
      
             If IsError(ActiveCell) Then ActiveCell = ""
               Select Case ActiveCell.Value
               Case "220"
              ActiveCell = "ND"
               Case "0"
               ActiveCell = "ND"
               Case ""
              ActiveCell = "ND"
              Case "NO VERIFICABLE"
              ActiveCell = "ND"
           End Select
      
     
       
     ActiveCell.Offset(0, 17).Select
     If IsError(ActiveCell) Then ActiveCell = ""
      'ActiveCell.NumberFormat = "#,##0.000"
      If IsNumeric(ActiveCell.Value) Then
          If Int(ActiveCell.Value) > 1 Then ActiveCell = ActiveCell / 100
         
          End If
        Select Case ActiveCell.Value
        Case ""
             ActiveCell = "ND"
        Case ""
             ActiveCell = "NO VERIFICABLE"
         End Select
       
       'ActiveCell.NumberFormat = "00%"
    L = 17
  ActiveCell.Offset(0, -L).Select
  ActiveCell.Offset(1, 0).Select
  
Next i



avance = 1
UpdateProgressBar avance
'-----------------------------FIN----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
MsgBox "FIN  DE PROCESO : SE CAMBIARON LOS ERRORES  MAS COMUNES "
CORRECCION_INVENTARIO.Hide
'CIERRA MACRO AL FINALIZAR



End Sub



