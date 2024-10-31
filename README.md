presentado por: Hamilton Atehortua

Introducción
Este código VBA implementa un sistema de registro y verificación de apuestas en un juego de lotería estilo Baloto en Excel. Se compone de dos macros principales:

Sub jugar(): Captura y registra la apuesta ingresada por el usuario.
Sub VerificarGanadores(): Revisa todas las apuestas registradas y verifica cuáles cumplen las condiciones para ser ganadoras.
A continuación, se detalla cada una de estas macros, con explicaciones de las variables y procesos clave.

Macro 1: Sub jugar()
Propósito
La macro jugar() permite a los usuarios registrar sus apuestas en una hoja de Excel. La información capturada incluye seis números y un número adicional o "balota".

Sección 1: Declaración de Variables
vba
Copiar código
Dim ws As Worksheet
Dim apuesta(1 To 6) As Integer
Dim balota As Integer
Dim i As Integer, j As Integer
Dim repetido As Boolean
ws: Referencia a la hoja de trabajo donde se guardarán las apuestas.
apuesta(1 To 6): Array para almacenar los seis números de la apuesta.
balota: Número de la balota adicional.
i, j: Variables de control para los bucles.
repetido: Booleano para identificar números repetidos en la apuesta.
Sección 2: Captura de Números desde el Formulario
vba
Copiar código
apuesta(1) = f_mijuegoBalotto.ComboBoxNum1.Value
'...captura los valores de ComboBox del formulario f_mijuegoBalotto
balota = f_mijuegoBalotto.ComboBoxBalota.Value
Formulario f_mijuegoBalotto: Se utiliza un formulario de usuario con ComboBoxes donde el usuario ingresa su apuesta.
Los valores seleccionados en cada ComboBox se almacenan en el array apuesta y en la variable balota.
Sección 3: Verificación de Números Repetidos
vba
Copiar código
For i = 1 To 6
    For j = i + 1 To 6
        If apuesta(i) = apuesta(j) Then
            MsgBox "Hay números repetidos. Intenta nuevamente.", vbExclamation
            Exit Sub
        End If
    Next j
Next i
Propósito: Garantiza que no haya duplicados en la apuesta.
Funcionalidad: Los bucles anidados comparan cada número de la apuesta con los demás. Si encuentra números repetidos, muestra un mensaje de error y detiene la ejecución.
Sección 4: Identificación de Fila y Asignación de ID
vba
Copiar código
Dim lastRow As Long
Dim nuevoID As Long

lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
If lastRow = 6 Then ' La fila 5 sería la primera fila de apuestas
    nuevoID = 1
Else
    nuevoID = ws.Cells(lastRow - 1, 1).Value + 1
End If
lastRow: Encuentra la última fila ocupada en la columna 1 de la hoja de apuestas.
nuevoID: Calcula el ID para la nueva apuesta (incremental).
Sección 5: Registro de Apuesta en la Hoja
vba
Copiar código
ws.Cells(lastRow, 1).Value = nuevoID
For i = 1 To 6
    ws.Cells(lastRow, i + 1).Value = apuesta(i)
Next i
ws.Cells(lastRow, 8).Value = balota
MsgBox "Apuesta registrada con éxito.", vbInformation
Registro de datos: Guarda el ID en la columna 1, los números de la apuesta en las columnas 2 a 7, y la balota en la columna 8.
Confirmación: Muestra un mensaje al usuario cuando la apuesta se registra exitosamente.
Macro 2: Sub VerificarGanadores()
Propósito
La macro VerificarGanadores() revisa las apuestas registradas en la hoja de cálculo y verifica cuáles cumplen las condiciones de ganadoras comparándolas con los números ganadores ingresados.

Sección 1: Inicialización y Captura de Números Ganadores
vba
Copiar código
Dim numerosGanadores(1 To 6) As Integer
Dim balotaGanadora As Integer

numerosGanadores(1) = CInt(f_balotta.TextBoxNum1.Value)
'...captura los valores de los TextBoxes en el formulario f_balotta
balotaGanadora = CInt(f_balotta.TextBoxBalota.Value)
numerosGanadores y balotaGanadora: Almacenan los números ganadores y la balota, capturados desde el formulario f_balotta.
Sección 2: Verificación de Apuestas Registradas
vba
Copiar código
Dim idApuesta As Variant
Dim numAciertos As Integer
Dim aciertoBalota As String

For i = 5 To lastRow ' Revisa cada apuesta desde la fila 5
    idApuesta = ws.Cells(i, 1).Value
    numAciertos = 0
    ' Comparar cada número de la apuesta con los ganadores
    For j = 2 To 7
        For k = 1 To 6
            If ws.Cells(i, j).Value = numerosGanadores(k) Then
                numAciertos = numAciertos + 1
                Exit For
            End If
        Next k
    Next j

    If ws.Cells(i, 8).Value = balotaGanadora Then
        aciertoBalota = "Si"
    Else
        aciertoBalota = "No"
    End If
Propósito: Compara los números de cada apuesta con los números ganadores.
numAciertos: Cuenta la cantidad de coincidencias entre los números apostados y los ganadores.
aciertoBalota: Indica si la balota de la apuesta coincide con la balota ganadora.
Sección 3: Verificación de Apuestas Ganadoras
vba
Copiar código
If (numAciertos = 6 And aciertoBalota = "Si") Or _
   (numAciertos = 6 And aciertoBalota = "No") Or _
   (numAciertos = 5 And aciertoBalota = "Si") Then
   result = result & idApuesta & vbTab & vbTab & i & vbTab & vbTab & numAciertos & vbTab & vbTab & aciertoBalota & vbCrLf
End If
Condiciones de Ganador: Determina que una apuesta es ganadora si cumple uno de estos requisitos:
6 aciertos con o sin balota ganadora.
5 aciertos con balota ganadora.
Sección 4: Resultados Finales
vba
Copiar código
If result = "ID_Apuesta" & vbTab & "Fila" & vbTab & "Numero aciertos" & vbTab & "Acierto Balota" & vbCrLf Then
    MsgBox "No hubo ganadores correspondientes."
Else
    MsgBox result, vbOKOnly, "Resultados de Apuestas Ganadoras"
End If
Visualización: Muestra los resultados de apuestas ganadoras en un mensaje. Si no hay ganadores, informa "No hubo ganadores correspondientes."
