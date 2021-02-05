VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Introducir Pedidos"
   ClientHeight    =   5295
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7485
   OleObjectBlob   =   "UserForm_orders.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) = True Then ListBox2.AddItem ListBox1.List(i)
Next i
End Sub

Private Sub CommandButton2_Click()
Dim counter As Integer
counter = 0

For i = 0 To ListBox2.ListCount - 1
    If ListBox2.Selected(i - counter) Then
        ListBox2.RemoveItem (i - counter)
        counter = counter + 1
    End If
Next i

End Sub

Private Sub CommandButton3_Click()

Dim a As Long
Dim p As Long
Dim arr_1 As Variant
Dim num_pedido As String
Dim temporada As String
Dim codigosPedidos As String
Dim long_query As Long
Dim ArrayLen As Integer

long_query = 0
a = 0
p = 0
Sheets(1).Cells.Clear


If ListBox2.ListCount = 0 Then
MsgBox "No hay pedidos seleccionados"
Else

Dim nIndex As Integer
Dim vArray() As Variant
ReDim vArray(UserForm1.ListBox2.ListCount - 1)

For nIndex = 0 To UserForm1.ListBox2.ListCount - 1
    vArray(nIndex) = UserForm1.ListBox2.List(nIndex)
Next


codigosPedidos = "'" + Join(vArray, "','") + "'" 'this array contains the selected orders from listbox2 ready to add them to the sql query

ArrayLen = UBound(vArray) - LBound(vArray) + 1

temporada = UserForm1.TextBox1.Text 'parsing the season to add it to the query

Application.ScreenUpdating = False


Set remoteCon = CreateObject("ADODB.Connection")

remoteCon.Open "DRIVER={MySQL ODBC 8.0 Unicode Driver}" _
          & ";SERVER=" & "" _
          & ";DATABASE=" & "" _
          & ";UID=" & "" _
          & ";PWD=" & "" _
          & ";PORT = xxxx"

remoteCon.Execute ("USE ;")

Set rs = CreateObject("ADODB.Recordset")


For j = 1 To 40

Sql = "" 'sql query to know how large the array is going to be

Set rs = remoteCon.Execute(Sql)
long_query = long_query + rs!cuenta

Next j


ReDim arr_1(0 To long_query, 1 To 26) 'this part is completely optional but I use it to make the excel sheet more understandable
'basically it adds headers to the sheet
 arr_1(0, 1) = ""
     arr_1(0, 2) = ""
     arr_1(0, 3) = ""
     arr_1(0, 4) = ""
     arr_1(0, 5) = ""
     arr_1(0, 6) = ""
     arr_1(0, 7) = ""
     arr_1(0, 8) = ""
     arr_1(0, 9) = ""
     arr_1(0, 10) = ""
     arr_1(0, 11) = ""
     arr_1(0, 12) = ""
     arr_1(0, 13) = ""
     arr_1(0, 14) = ""
     arr_1(0, 15) = ""
     arr_1(0, 16) = ""
     arr_1(0, 17) = ""
     arr_1(0, 18) = ""
     arr_1(0, 19) = ""
     arr_1(0, 20) = ""
     arr_1(0, 21) = ""
     arr_1(0, 22) = ""
     arr_1(0, 23) = ""
     arr_1(0, 24) = ""
     arr_1(0, 25) = ""
     arr_1(0, 26) = ""

For i = 1 To 40


Sql = "" 'query we are going to show

Set rs = remoteCon.Execute(Sql)
     
If Not rs.BOF And Not rs.EOF Then
    myArray = rs.GetRows

    kolumner = UBound(myArray, 1) 'columnas
    rader = UBound(myArray, 2) 'filas
        
    For k = 1 To rader 'this loop is used to change some of the variables and format them so it's made for my particular case

     arr_1(k + a, 1) = Replace(myArray(1, k), "A", "")
     arr_1(k + a, 2) = ""
     arr_1(k + a, 3) = ""
     arr_1(k + a, 4) = ""
     arr_1(k + a, 5) = 0
     arr_1(k + a, 6) = ""
     arr_1(k + a, 7) = ""
     arr_1(k + a, 8) = myArray(3, k)
     arr_1(k + a, 9) = myArray(4, k)
     arr_1(k + a, 10) = myArray(11, k)
     arr_1(k + a, 11) = myArray(12, k)
     arr_1(k + a, 12) = ""
     
     
     If InStr(1, "PZ", myArray(19, k)) <> 0 Or InStr(1, "PA", myArray(19, k)) <> 0 Or InStr(1, "CF", myArray(19, k)) <> 0 Or IsNull(myArray(19, k)) = True Then
      arr_1(k + a, 13) = "1SIZ"
     Else
     arr_1(k + a, 13) = Replace(myArray(20, k), "½", ".5") 'Talla normal'
     End If
     
     arr_1(k + a, 14) = myArray(20, k)
     arr_1(k + a, 15) = myArray(15, k)
     
     If myArray(17, k) <> 0 Then
     arr_1(k + a, 16) = myArray(17, k)
     Else
     arr_1(k + a, 16) = myArray(18, k)
     End If
     
     arr_1(k + a, 17) = 0
     arr_1(k + a, 18) = ""
     arr_1(k + a, 19) = myArray(0, k)
     arr_1(k + a, 20) = ""
     arr_1(k + a, 21) = ""
     arr_1(k + a, 22) = temporada
     arr_1(k + a, 23) = 0
     arr_1(k + a, 24) = myArray(5, k)
     arr_1(k + a, 25) = myArray(1, k)
     arr_1(k + a, 26) = ""
    

    Next k
    
End If
a = a + rader

Next i


PrintArray arr_1, ActiveWorkbook.Worksheets(1).[A1]
     
Dim lastrow As Long
lastrow = Cells(Rows.Count, 2).End(xlUp).Row
Range("A1:Y" & lastrow).Sort _
key1:=Range("A1:A" & lastrow), _
order1:=xlAscending, Header:=xlYes, _
key2:=Range("J1:J" & lastrow), order2:=xlAscending, Header:=xlYes

rs.Close
Set rs = Nothing

remoteCon.Close
Set remoteCon = Nothing

Dim LastRow2 As Integer
Dim LastCol As Integer
Dim numero As Integer

LastRow2 = ActiveSheet.UsedRange.Rows.Count

codigoPedido = Range("Y2")
numPedidoAuto = 1
rangoColumna1 = "A2:A" + CStr(LastRow2)

For k = 2 To LastRow2

columnPedido = "Y" + CStr(k)
columnNumAuto = "A" + CStr(k)
currentItem = Range(columnPedido).Value
columnaRellenar = Range(columnNumAuto).Value

    If (currentItem <> codigoPedido) Then
        numPedidoAuto = numPedidoAuto + 1
        codigoPedido = currentItem
    End If
    
    Range(columnNumAuto).Value = CStr(numPedidoAuto)

Next k

Application.ScreenUpdating = True

UserForm1.Hide

MsgBox "Operacion terminada"

End If
End Sub

Sub PrintArray(Data As Variant, Cl As Range)
    Cl.Resize(UBound(Data, 1), UBound(Data, 2)) = Data

End Sub

Private Sub CommandButton4_Click()
Dim arrayPedidos As Variant
Dim temporada As String

temporada = UserForm1.TextBox1.Text


Application.ScreenUpdating = False

Set remoteCon = CreateObject("ADODB.Connection")

remoteCon.Open "DRIVER={MySQL ODBC 8.0 Unicode Driver}" _
          & ";SERVER=" & "" _
          & ";DATABASE=" & "" _
          & ";UID=" & "" _
          & ";PWD=" & "" _
          & ";PORT = xxxx"

remoteCon.Execute ("USE ;")

Set rs = CreateObject("ADODB.Recordset")

Sql = "" 'query to select the season

Set rs = remoteCon.Execute(Sql)

If Not rs.BOF And Not rs.EOF Then
    myArray = rs.GetRows

    columnas = UBound(myArray, 1) 'columnas
    filas = UBound(myArray, 2) 'filas
        
    For i = 0 To filas
        UserForm1.ListBox1.AddItem myArray(0, i)
    Next i
    
End If
End Sub
