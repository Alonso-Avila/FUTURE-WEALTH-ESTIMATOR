VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Future Wealth Estimator"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10575
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Instrucciones al iniciarse el formulario
Private Sub UserForm_Initialize()
Hoja1.Range("A2:H5200").Clear 'Limpiar calculo anterior si aún sigue

'Llenar combobox de periodos
Dim a As Integer
For a = 1 To 5200
    ComboBox1.AddItem a
Next

'Deshabilitar espacios para agregar cambios antes de ingresar datos iniciales
PeriodoMonto.Locked = True
NvoMonto.Locked = True
PeriodoTasa.Locked = True
NvaTasa.Locked = True
PeriodoRetiro.Locked = True
Monto_retiro.Locked = True

'Habilitar botón para guardar datos iniciales
Guardar.Locked = False

'Agregar tips a textboxes diciendo que primero se deben guardar datos iniciales
Me.ComboBox1.ControlTipText = "Puede teclearlo o seleccionarlo"
Me.PeriodoTasa.ControlTipText = "Para editar debe primero GUARDAR los datos iniciales"
Me.NvaTasa.ControlTipText = "Para editar debe primero GUARDAR los datos iniciales"
Me.PeriodoMonto.ControlTipText = "Para editar debe primero GUARDAR los datos iniciales"
Me.NvoMonto.ControlTipText = "Para editar debe primero GUARDAR los datos iniciales"
Me.PeriodoRetiro.ControlTipText = "Para editar debe primero GUARDAR los datos iniciales"
Me.Monto_retiro.ControlTipText = "Para editar debe primero GUARDAR los datos iniciales"

'Deshabilitar espacios donde aparecen los resultados
capital_final.Locked = True
ahorro.Locked = True

End Sub

'Macro para guardar datos iniciales
Private Sub Guardar_Click()
Workbooks("Calculadora Capital Futuro aportaciones y tasa variable.xlsm").Activate 'Prevenir editar otro libro abierto

With Hoja1
'Declaración de variables
Dim i As Integer
Dim P As Integer
Dim Rate As Double
Dim Payment As Long
Dim iI As Integer
Dim CO As Integer
Dim iCO As Integer
Dim Pay0 As Long

'Asegurarse de que se llenen todos los campos iniciales evitando así un error
If ComboBox1 = "" Or Tasa = "" Or Aportacion = "" Or Capital_inicial = "" Then
MsgBox "Faltan casillas por llenar", vbCritical, "FUTURE WEALTH ESTIMATOR"
Exit Sub
End If

'Establecer limite de periodos para eficientización de cálculos
If ComboBox1 > 5200 Then
MsgBox "El numero maximo de periodos es 5200", vbCritical, "FUTURE WEALTH ESTIMATOR"
Exit Sub
End If

'Agregar datos iniciales a listbox para consulta
Me.ListBox1.AddItem ComboBox1 & " " & "periodos"
Me.ListBox1.AddItem Tasa & "%" & " " & "de tasa de rendimiento"
Me.ListBox1.AddItem Aportacion & "$" & " " & "de aportaciones iniciales"
Me.ListBox1.AddItem Capital_inicial & "$" & " " & "de capital inicial invertido"

'mapeo de captura con contador
i = 3 'porque los registros inician el la fila 3
iCO = 4 'Contador para aportaciones después del capital inicial
P = ComboBox1.Value
iI = ComboBox1.Value  'Contador inverso para periodos restantes

Rate = Tasa.Value
Payment = Aportacion.Value
Pay0 = Capital_inicial.Value

'captura de variables excepto aportaciones
Do While i <= P + 2 ' 'P+2 porque los registros inician el la fila 3
    .Cells(3, 4).Value = Pay0
    .Cells(3, 4).NumberFormat = "$#,##0.00" 'Formato moneda
    .Cells(i, 1).Value = i - 2
    .Cells(i, 2).Value = iI
    .Cells(i, 3).Value = Rate
    
    i = i + 1
    iI = iI - 1
Loop

'captura de aportaciones
Do While iCO <= P + 2
 .Cells(iCO, 4).Value = Payment
 .Cells(iCO, 4).NumberFormat = "$#,##0.00"
 iCO = iCO + 1

Loop
End With

'Una vez que los datos iniciales son guardados se pueden agregar cambios
PeriodoMonto.Locked = False
NvoMonto.Locked = False
PeriodoTasa.Locked = False
NvaTasa.Locked = False
PeriodoRetiro.Locked = False
Monto_retiro.Locked = False

'Remover recomendación flotante
Me.PeriodoTasa.ControlTipText = ""
Me.NvaTasa.ControlTipText = ""
Me.PeriodoMonto.ControlTipText = ""
Me.NvoMonto.ControlTipText = ""
Me.PeriodoRetiro.ControlTipText = ""
Me.Monto_retiro.ControlTipText = ""

Guardar.Locked = True 'Deshabilitar el boton de guardar para evitar errores de cálculo

End Sub

'Macro para agregar cambios de aportaciones
Private Sub Agregar_Click()

With Hoja1

Dim NewPayment As Integer
Dim paymentperiod As Integer
Dim P As Integer
Dim NoconflictMontoperiod As Integer

'Asegurarse de que se llenen ambos campos para evitar un error
If NvoMonto.Value = "" Or PeriodoMonto.Value = "" Then
MsgBox "Espacio(s) por llenar, si no puede editar primero debe GUARDAR los datos iniciales", vbCritical, "FUTURE WEALTH ESTIMATOR"
Exit Sub
End If

'Agregar un '0' antes del periodo de cambio por motivos de interpretación de VBA al compararlo con el total de periodos.
NoconflictMontoperiod = "0" & PeriodoMonto

'Evitar que se agregue un cambio en un periodo inexistente o mayor al total.
If NoconflictMontoperiod > ComboBox1 Then
MsgBox "Sólo hay" & " " & ComboBox1 & " " & "periodos en este cálculo", vbCritical, "FUTURE WEALTH ESTIMATOR"
Exit Sub
End If

'Agregar cambio a listbox para consulta
Me.ListBox1.AddItem NvoMonto & "$" & " " & "de aportación a partir del periodo" & " " & (Me.PeriodoMonto)

P = ComboBox1.Value
NewPayment = NvoMonto.Value
paymentperiod = PeriodoMonto.Value

'Agregar cambio a la tabla de cálculo
 For Each Cell In .Range("A3:A5203")
 If Cell.Value >= paymentperiod And Cell.Value <= P Then
    Cell.Offset(0, 3).Value = NewPayment
 End If
 Next
End With

'Vaciar espacios para facilitar el ingreso de nuevos cambios
NvoMonto = ""
PeriodoMonto = ""
End Sub
'Macro para agregar cambios de tasa
Private Sub add_rate_change_Click()
With Hoja1

Dim Newrate As Double
Dim Rateperiod As Integer
Dim P As Integer
Dim noconflictrateperiod As Integer

'Asegurarse de que se llenen ambos campos para evitar un error
If NvaTasa.Value = "" Or PeriodoTasa.Value = "" Then
MsgBox "Espacio(s) por llenar, si no puede editar primero debe GUARDAR los datos iniciales", vbCritical, "FUTURE WEALTH ESTIMATOR"
Exit Sub
End If

'Agregar un '0' antes del periodo de cambio por motivos de interpretación de VBA al compararlo con el total de periodos.
noconflictrateperiod = "0" & PeriodoTasa

'Evitar que se agregue un cambio en un periodo inexistente o mayor al total.
If noconflictrateperiod > ComboBox1 Then
    MsgBox "Sólo hay" & " " & ComboBox1 & " " & "periodos en este cálculo", vbCritical, "FUTURE WEALTH ESTIMATOR"
    Exit Sub
End If

'Agregar cambio a listbox para consulta
Me.ListBox1.AddItem NvaTasa & "%" & " " & "de rendimiento a partir del periodo" & " " & (Me.PeriodoTasa)

P = ComboBox1.Value
Newrate = NvaTasa.Value
Rateperiod = PeriodoTasa.Value

'Agregar cambio a la tabla de cálculo
 For Each Cell In .Range("A3:A5203")
 If Cell.Value >= Rateperiod And Cell.Value <= P Then
    Cell.Offset(0, 2).Value = Newrate
 End If
 Next
End With

'Vaciar espacios para facilitar el ingreso de nuevos cambios
NvaTasa = ""
PeriodoTasa = ""
End Sub

'Macro para agregar retiro
Private Sub add_withdraw_Click()
With Hoja1

Dim withdraw As Double
Dim Period_withdraw As Integer
Dim no_conflict_withdraw_period As Integer

'Asegurarse de que se llenen ambos campos para evitar un error
If Monto_retiro.Value = "" Or PeriodoRetiro.Value = "" Then
MsgBox "Espacio(s) por llenar, si no puede editar primero debe GUARDAR los datos iniciales", vbCritical, "FUTURE WEALTH ESTIMATOR"
Exit Sub
End If

'Agregar un '0' antes del periodo de retiro por motivos de interpretación de VBA al compararlo con el total de periodos.
no_conflict_withdraw_period = "0" & PeriodoRetiro

'Evitar que se agregue un retiro en un periodo inexistente o mayor al total.
If no_conflict_withdraw_period > ComboBox1 Then
    MsgBox "Sólo hay" & " " & ComboBox1 & " " & "periodos en este cálculo", vbCritical, "FUTURE WEALTH ESTIMATOR"
    Exit Sub
End If

'Agregar retiro a listbox para consulta.
Me.ListBox1.AddItem "Retiro de" & " " & Monto_retiro & "$" & " " & "en el periodo" & " " & (Me.PeriodoRetiro)

withdraw = Monto_retiro.Value
Period_withdraw = Me.PeriodoRetiro.Value

'Agregar retiro a la tabla de cálculo, aquí solo se captura en el periodo indicado y no en todos a partir del mismo.
 For Each Cell In .Range("A3:A5203")
 If Cell.Value = Period_withdraw Then
    Cell.Offset(0, 5).Value = withdraw
                        
 End If
 Next
End With

'Vaciar espacios para facilitar el ingreso de nuevos cambios.
PeriodoRetiro = ""
Monto_retiro = ""
End Sub

'Macro para hacer el cálculo.
Private Sub Calcular_Click()
With Hoja1

'Declaración de variables.
Dim P As Integer
Dim c As Integer
Dim i As Double
Dim t As Integer
Dim pp As Integer
Dim LastRow As Integer
Dim Suma As Double
Dim sumrow As Integer
Dim LastsavesRow As Integer
Dim savessum As Double
Dim answer As Long
Dim Last_accRow As Integer

'Evitar error no permitiendo ejecutar la macro sin datos guardados.
If ComboBox1 = "" Or Tasa = "" Or Aportacion = "" Or Capital_inicial = "" Then
    MsgBox "No ha guardado los parametros iniciales ni ha agregado cambios", vbCritical, "FUTURE WEALTH ESTIMATOR"
    Exit Sub
End If

'Si no se han agregado cambios en los datos (detectados en listbox de consulta) preguntar si aún así se desea hacer el "cálculo".
If ListBox1.ListCount < 5 Then
answer = MsgBox("No ha agregado ningun cambio de tasa, de aportaciones ni retiros, ¿Desea continuar?", vbYesNo, "FUTURE WEALTH ESTIMATOR") 'MsgBox "No ha agregado ningun cambio de tasa ni de aportaciones, ¿Desea continuar?", vbYesNo, "FUTURE WEALTH ESTIMATOR"
    If answer = vbNo Then
    Exit Sub
    End If
End If

'Calcular flujos individuales y sin retiros, de momento sin influencia en resultados ni en los gráficos.
P = ComboBox1.Value
pp = 3
Do While pp <= P + 2
    .Cells(pp, 5).Value = .Cells(pp, 4).Value * (1 + (.Cells(pp, 3).Value / 100)) ^ .Cells(pp, 2).Value
     pp = pp + 1
Loop


pp = 3
Do While pp <= P + 2
    'Calcular el acumulado solo ahorrado.
    .Cells(pp, 7).Value = ((.Cells(pp, 4).Value + .Cells(pp - 1, 7).Value) * (1 + (.Cells(pp, 3).Value / 100))) - .Cells(pp, 6).Value
    .Cells(pp, 7).NumberFormat = "$#,##0.00"
    'Calcular capital acumulado en cada periodo con cambios y retiros.
    .Cells(pp, 8).Value = (.Cells(pp, 4).Value + .Cells(pp - 1, 8).Value) - .Cells(pp, 6).Value
     pp = pp + 1
Loop

 'Variables de mapeo y totales para mostrar resultados.
 LastRow = .Cells(Rows.Count, 5).End(xlUp).row
 Last_accRow = .Cells(Rows.Count, 7).End(xlUp).row
 sumrow = LastRow + 1
 Suma = WorksheetFunction.Sum(.Range("E1:E" & LastRow))
 .Cells(sumrow, 5).Value = Suma
 LastsavesRow = .Cells(Rows.Count, 4).End(xlUp).row
 
'Mostrar capital formado y solo ahorros en el formulario.
 capital_final.Value = Format(.Range("G" & Last_accRow).Value, "$#,##0.00")
 ahorro.Value = Format((WorksheetFunction.Sum(.Range("D1:D" & LastsavesRow)) - WorksheetFunction.Sum(.Range("F1:F5203"))), "$#,##0.00")
 
End With

'Bloquear todos los botones para modificar calculos generandose nuevos resultados incorrectos
Calcular.Locked = True
add_rate_change.Locked = True
Agregar.Locked = True
Guardar.Locked = True
End Sub

'macro para hacer un nuevo calculo, limpia los campos y llama al reinicio del formulario
Private Sub Nuevo_calc_Click()
ComboBox1 = ""
Tasa = ""
Aportacion = ""
Capital_inicial = ""
PeriodoTasa = ""
NvaTasa = ""
PeriodoMonto = ""
NvoMonto = ""
capital_final = ""
ahorro = ""
ListBox1.Clear

Calcular.Locked = False
add_rate_change.Locked = False
Agregar.Locked = False
Guardar.Locked = False


Call UserForm_Initialize
End Sub

'Botón para mostrar las opciones de los gráficos en otro formulario
Private Sub Graficos_command_Click()
UserForm2.Show
End Sub

'Botón para cerrar el formulario principal
Private Sub Cerrar_Click()
    Unload Me
End Sub

'Macro para darle formato de moneda al campo de aportaciones y para aclarar el ingreso de solo carácteres númericos
Private Sub Aportacion_AfterUpdate()

If Not IsNumeric(Aportacion.Value) Then
    Aportacion = ""
    MsgBox "Ingrese solo valores númericos"
Else
    Aportacion = Format(Aportacion, "$#,##0.00")
End If
End Sub

''Macro para darle formato de moneda al campo de capital inicial y para aclarar el ingreso de solo carácteres númericos
Private Sub Capital_inicial_AfterUpdate()
If Not IsNumeric(Capital_inicial.Value) Then
    Capital_inicial = ""
    MsgBox "Ingrese solo valores númericos"
Else
    Capital_inicial = Format(Capital_inicial, "$#,##0.00")
End If
End Sub

