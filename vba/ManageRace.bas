' ================================================================
'  ManageRace  -  Race management actions from Excel
'  Compatible with: Formato_Carrera_de_drones.xlsm
'
'  WHAT THIS MODULE DOES:
'    Opens frmManageHeat to edit a specific heat's pilot slots. This is
'    the only entry point left in this module -- EditHeatPilots (InputBox
'    chain to pick class/heat by id), RemixClass, DeleteClass, and
'    DeleteAllHeats were removed: the per-heat/per-group buttons on the
'    "Carrera" sheet (E/R per heat, R/MX/X per group) already cover the
'    same actions without having to type an id into an InputBox, so the
'    column-A buttons that called those Subs were removed as redundant.
'
'  HOW TO INSTALL:
'    1. Alt+F11 to open VBA Editor
'    2. File > Import File... > select this .bas file
' ================================================================

Option Explicit

' -- Editar los pilotos de un heat especifico (llamado desde el boton "E"
' de ese heat en la hoja Carrera). Abre frmManageHeat -- requiere haberlo
' creado, ver frmManageHeat_instructions.txt --------------------------
Public Sub EditHeatPilotsById(heatId As Long)
    Dim frm As frmManageHeat
    Set frm = New frmManageHeat
    frm.LoadHeat heatId
    frm.Show
End Sub
