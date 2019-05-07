Attribute VB_Name = "Módulo3"
Sub BAJAR_PESO()
Attribute BAJAR_PESO.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BAJAR_PESO Macro
'

'
    Range("Q7").Select
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\ERICSSON\Downloads\DESCARGAS_MPA\MP_RF_ACOPI_II_OT-4356090_S12972.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Range("P10").Select
    ActiveWorkbook.Save
End Sub
