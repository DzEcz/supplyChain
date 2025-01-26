Attribute VB_Name = "Module1"
Sub AnaProsedur()
    ' Optimizasyonlarý kapat
    OptimizeOperations False

    ' Hata yakalama
    On Error GoTo HataYakalama

    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet

    ' UserForm'u göster
    UserForm1.Show vbModeless
    UserForm1.Caption = "Ýlerleme Durumu"
    DoEvents

    ' Tüm düðmeleri pasif yap
    UserForm1.CommandButton1.Enabled = False
    UserForm1.CommandButton2.Enabled = False
    UserForm1.CommandButton3.Enabled = False

    ' Hesap sayfasýnýn kilidini aç
    Sheets("Hesap").Unprotect Password:="8142"
    Sheets("Pusula").Unprotect Password:="8142"

    ' Ýþlemleri gerçekleþtir
    Call Adshow
    GuncelleIlerleme 1
    Call PusulaSayfasiniGuncelle
    GuncelleIlerleme 2
    Call VeriKopyala
    GuncelleIlerleme 3
    Call KopyalaVeEkleHizli
    GuncelleIlerleme 4
    Call KutuMiktarKopyala
    GuncelleIlerleme 5
    Call EsdegerToplam
    GuncelleIlerleme 6
    Call DinamikSirala
    GuncelleIlerleme 7
    Call KopyalaHastaneleri
    GuncelleIlerleme 8
    Call UpdateDepoDurumu
    GuncelleIlerleme 9
    Call PivotTabloyuYenile
    GuncelleIlerleme 10

    ' Ýþlemler tamamlandýðýnda bildirim ekle
    UserForm1.ListBox.AddItem "Tüm iþlemler baþarýyla gerçekleþti."

    ' Hesap sayfasýný tekrar kilitle
    Sheets("Hesap").Protect Password:="8142"
    Sheets("Pusula").Protect Password:="8142"

    ' Baþlatýlan sayfaya geri dön
    currentSheet.Activate

    ' Kapatma butonunu aktif yap
    UserForm1.CommandButton1.Enabled = True
    UserForm1.CommandButton3.Enabled = True

    ' Optimizasyonlarý aç
    OptimizeOperations True
    Exit Sub

HataYakalama:
    ' Hata durumunda UserForm'u gizle ve hata mesajýný göster
    MsgBox "Bir hata oluþtu: " & Err.Description & vbCrLf & _
           "Prosedür: " & Err.Source & vbCrLf & _
           "Satýr: " & Erl, vbCritical

    ' Optimizasyonlarý aç
    OptimizeOperations True
End Sub

Sub OptimizeOperations(state As Boolean)
    Application.ScreenUpdating = state
    Application.Calculation = IIf(state, xlCalculationAutomatic, xlCalculationManual)
    Application.EnableEvents = state
End Sub

Sub GuncelleIlerleme(adim As Integer)
    With UserForm1.ProgressBar
        .Width = adim * (UserForm1.Frame1.Width / 10) ' Her adým Frame1'in geniþliðinin 1/10'u kadar
    End With
    DoEvents ' Güncellemelerin anlýk olarak görülmesini saðlar
End Sub

Sub Adshow()
    Application.StatusBar = "Ecz. Harun Topal"
End Sub

Sub PusulaSayfasiniGuncelle()
    UserForm1.ListBox.AddItem "Pusula sayfasý güncelleme iþlemi baþladý."
    Dim kaynakKitap As Workbook
    Dim hedefKitap As Workbook
    Dim kaynakSayfa As Worksheet
    Dim hedefSayfa As Worksheet
    Dim kaynakDosyaYolu As String

    kaynakDosyaYolu = ThisWorkbook.Path & "\Pusula.xlsx"
    Set kaynakKitap = Workbooks.Open(kaynakDosyaYolu)
    Set kaynakSayfa = kaynakKitap.Sheets("Sheet")
    Set hedefKitap = ThisWorkbook
    Set hedefSayfa = hedefKitap.Sheets("Pusula")

    hedefSayfa.Cells.Clear
    kaynakSayfa.UsedRange.Copy
    hedefSayfa.Range("A1").PasteSpecial Paste:=xlPasteValues
    kaynakKitap.Close False

    UserForm1.ListBox.AddItem "Pusula sayfasý güncelleme iþlemi tamamlandý."
End Sub

Sub VeriKopyala()
    UserForm1.ListBox.AddItem "Pusula sayfasýndan veri kopyalama iþlemi baþladý."
    Dim wsPusula As Worksheet
    Dim wsHesap As Worksheet
    Dim lastRow As Long
    Dim kodCol As Long
    Dim adCol As Long
    Dim miktarCol As Long
    Dim kodData As Variant
    Dim i As Long

    Set wsPusula = ThisWorkbook.Sheets("Pusula")
    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    lastRow = wsPusula.Cells(wsPusula.Rows.count, "A").End(xlUp).row

    If lastRow < 2 Then
        MsgBox "Lütfen Pusuladan çektiðiniz stok durum raporunu ayný klasöre kopyalayýnýz!", vbExclamation
        wsPusula.Activate
        OptimizeOperations True
        Exit Sub
    End If

    If wsHesap.Cells(2, 1).value <> "" Then
        wsHesap.Rows("2:" & wsHesap.Rows.count).ClearContents
    End If

    kodCol = wsPusula.Rows(1).Find("C. EMR Eþdeðer Ürün Grup Kodu").Column
    adCol = wsPusula.Rows(1).Find("Adý").Column
    miktarCol = wsPusula.Rows(1).Find("Miktar").Column

    kodData = wsPusula.Range(wsPusula.Cells(2, kodCol), wsPusula.Cells(lastRow, kodCol)).value

    For i = 1 To UBound(kodData, 1)
        If IsNumeric(kodData(i, 1)) Then
            kodData(i, 1) = Round(CDbl(kodData(i, 1)), 0)
        End If
    Next i

    wsHesap.Cells(1, 1).value = "EþdeðerKod"
    wsHesap.Cells(1, 2).value = "Müstahzar"
    wsHesap.Cells(1, 3).value = "Stok Miktar"

    wsHesap.Range("A2:A" & lastRow).value = kodData
    wsHesap.Range("B2:B" & lastRow).value = wsPusula.Range(wsPusula.Cells(2, adCol), wsPusula.Cells(lastRow, adCol)).value
    wsHesap.Range("C2:C" & lastRow).value = wsPusula.Range(wsPusula.Cells(2, miktarCol), wsPusula.Cells(lastRow, miktarCol)).value

    UserForm1.ListBox.AddItem "Pusula sayfasýndan veri kopyalama iþlemi tamamlandý."
End Sub

Sub KopyalaVeEkleHizli()
    UserForm1.ListBox.AddItem "Müstahzar sayýsýnýn üçlemesi iþlemi baþladý."
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim kod As String
    Dim kodCount As Object
    Dim data As Variant
    Dim result() As Variant
    Dim resultIndex As Long
    Dim esdegerKodCol As Long, mustahzarCol As Long, stokMiktarCol As Long

    Set kodCount = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets("Hesap")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    esdegerKodCol = Application.WorksheetFunction.Match("EþdeðerKod", ws.Rows(1), 0)
    mustahzarCol = Application.WorksheetFunction.Match("Müstahzar", ws.Rows(1), 0)
    stokMiktarCol = Application.WorksheetFunction.Match("Stok Miktar", ws.Rows(1), 0)

    data = ws.Range(ws.Cells(2, esdegerKodCol), ws.Cells(lastRow, stokMiktarCol)).value

    ReDim result(1 To (lastRow - 1) * 2, 1 To UBound(data, 2))
    resultIndex = 1

    For i = 1 To UBound(data, 1)
        kod = data(i, 1)
        If kodCount.Exists(kod) Then
            kodCount(kod) = kodCount(kod) + 1
        Else
            kodCount.Add kod, 1
        End If
    Next i

    For i = 1 To UBound(data, 1)
        kod = data(i, 1)
        If kodCount(kod) = 2 Then
            result(resultIndex, 1) = data(i, 1)
            result(resultIndex, 2) = data(i, 2) & "_kopya" & kodCount(kod)
            result(resultIndex, 3) = data(i, 3)
            resultIndex = resultIndex + 1
            kodCount(kod) = kodCount(kod) + 1
        End If
    Next i

    For i = 1 To UBound(data, 1)
        kod = data(i, 1)
        If kodCount(kod) = 1 Then
            For j = 1 To 2
                If kodCount(kod) < 3 Then
                    result(resultIndex, 1) = data(i, 1)
                    result(resultIndex, 2) = data(i, 2) & "_kopya" & kodCount(kod)
                    result(resultIndex, 3) = data(i, 3)
                    resultIndex = resultIndex + 1
                    kodCount(kod) = kodCount(kod) + 1
                End If
            Next j
        End If
    Next i

    ws.Range(ws.Cells(lastRow + 1, esdegerKodCol), ws.Cells(lastRow + resultIndex - 1, stokMiktarCol)).value = result
    ws.Range(ws.Cells(2, esdegerKodCol), ws.Cells(lastRow + resultIndex - 1, stokMiktarCol)).Sort Key1:=ws.Cells(2, esdegerKodCol), Order1:=xlAscending, header:=xlNo

    UserForm1.ListBox.AddItem "Müstahzar sayýsýnýn üçlemesi iþlemi tamamlandý."
End Sub

Sub KutuMiktarKopyala()
    UserForm1.ListBox.AddItem "Kutu içi miktarlarýnýn kopyalanmasý iþlemi baþladý."
    Dim wsHesap As Worksheet
    Dim wsKutuiçi As Worksheet
    Dim hesapData As Variant
    Dim kutuiciData As Variant
    Dim i As Long
    Dim dict As Object

    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    Set wsKutuiçi = ThisWorkbook.Sheets("Kutuiçi")

    Dim colEsdegerKodHesap As Long
    Dim colKutuMiktarHesap As Long
    Dim colEsdegerKodKutuiçi As Long
    Dim colKutuIciKutuiçi As Long

    colEsdegerKodHesap = Application.Match("EþdeðerKod", wsHesap.Rows(1), 0)
    colKutuMiktarHesap = Application.Match("Kutu Miktar", wsHesap.Rows(1), 0)
    colEsdegerKodKutuiçi = Application.Match("Eþdeðer", wsKutuiçi.Rows(1), 0)
    colKutuIciKutuiçi = Application.Match("Kutu Ýçi", wsKutuiçi.Rows(1), 0)

    hesapData = wsHesap.Range(wsHesap.Cells(2, colEsdegerKodHesap), wsHesap.Cells(wsHesap.Rows.count, colEsdegerKodHesap).End(xlUp)).Resize(, colKutuMiktarHesap - colEsdegerKodHesap + 1).value
    kutuiciData = wsKutuiçi.Range(wsKutuiçi.Cells(2, colEsdegerKodKutuiçi), wsKutuiçi.Cells(wsKutuiçi.Rows.count, colEsdegerKodKutuiçi).End(xlUp)).Resize(, colKutuIciKutuiçi - colEsdegerKodKutuiçi + 1).value

    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(kutuiciData, 1)
        dict(kutuiciData(i, 1)) = kutuiciData(i, 2)
    Next i

    For i = 1 To UBound(hesapData, 1)
        If dict.Exists(hesapData(i, 1)) Then
            hesapData(i, colKutuMiktarHesap - colEsdegerKodHesap + 1) = dict(hesapData(i, 1))
        Else
            hesapData(i, colKutuMiktarHesap - colEsdegerKodHesap + 1) = 1
        End If
    Next i

    wsHesap.Range(wsHesap.Cells(2, colEsdegerKodHesap), wsHesap.Cells(UBound(hesapData, 1) + 1, colKutuMiktarHesap)).value = hesapData

    UserForm1.ListBox.AddItem "Kutu içi miktarlarýnýn kopyalanmasý iþlemi tamamlandý."
End Sub

Sub EsdegerToplam()
    UserForm1.ListBox.AddItem "Stok hesaplama iþlemleri baþladý."
    Dim wsHesap As Worksheet
    Dim wsPusula As Worksheet
    Dim hesesdegerkodverisi As Range
    Dim heskutumiktarverisi As Range
    Dim hesesdmiktoplam As Range
    Dim heskrimiktoplam As Range
    Dim hesmaxmiktartoplam As Range
    Dim hesgopithmik As Range
    Dim pusesdegerkodverisi As Range
    Dim pusmikverisi As Range
    Dim puskrimikverisi As Range
    Dim pusmaxmikverisi As Range
    Dim cell As Range
    Dim pCell As Range
    Dim toplam As Double
    Dim kod As String
    Dim miktar As Double
    Dim krimiktoplam As Double
    Dim maxmiktartoplam As Double

    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    Set wsPusula = ThisWorkbook.Sheets("Pusula")

    Set hesesdegerkodverisi = wsHesap.Rows(1).Find("EþdeðerKod")
    Set heskutumiktarverisi = wsHesap.Rows(1).Find("Kutu Miktar")
    Set hesesdmiktoplam = wsHesap.Rows(1).Find("Eþd.Mik. TOPLAM")
    Set heskrimiktoplam = wsHesap.Rows(1).Find("Kri.Mik. TOPLAM")
    Set hesmaxmiktartoplam = wsHesap.Rows(1).Find("Max.Mik TOPLAM")
    Set hesgopithmik = wsHesap.Rows(1).Find("Ýht. Mik.")

    Set pusesdegerkodverisi = wsPusula.Rows(1).Find("C. EMR Eþdeðer Ürün Grup Kodu")
    Set pusmikverisi = wsPusula.Rows(1).Find("Miktar")
    Set puskrimikverisi = wsPusula.Rows(1).Find("Kritik Miktar")
    Set pusmaxmikverisi = wsPusula.Rows(1).Find("Max Miktar")

    For Each cell In wsHesap.Range(hesesdegerkodverisi.Offset(1, 0), wsHesap.Cells(wsHesap.Rows.count, hesesdegerkodverisi.Column).End(xlUp))
        kod = Trim(UCase(cell.value))
        toplam = 0
        krimiktoplam = 0
        maxmiktartoplam = 0

        For Each pCell In wsPusula.Range(pusesdegerkodverisi.Offset(1, 0), wsPusula.Cells(wsPusula.Rows.count, pusesdegerkodverisi.Column).End(xlUp))
            If Trim(UCase(pCell.value)) = kod Then
                toplam = toplam + CDbl(pCell.Offset(0, pusmikverisi.Column - pusesdegerkodverisi.Column).value)
                krimiktoplam = krimiktoplam + CDbl(pCell.Offset(0, puskrimikverisi.Column - pusesdegerkodverisi.Column).value)
                maxmiktartoplam = maxmiktartoplam + CDbl(pCell.Offset(0, pusmaxmikverisi.Column - pusesdegerkodverisi.Column).value)
            End If
        Next pCell

        miktar = CDbl(cell.Offset(0, heskutumiktarverisi.Column - hesesdegerkodverisi.Column).value)
        If miktar <> 0 Then
            cell.Offset(0, hesesdmiktoplam.Column - hesesdegerkodverisi.Column).value = Round(toplam / miktar, 0)
            cell.Offset(0, heskrimiktoplam.Column - hesesdegerkodverisi.Column).value = Round(krimiktoplam / miktar, 0)
            cell.Offset(0, hesmaxmiktartoplam.Column - hesesdegerkodverisi.Column).value = Round(maxmiktartoplam / miktar, 0)
        Else
            cell.Offset(0, hesesdmiktoplam.Column - hesesdegerkodverisi.Column).value = 0
            cell.Offset(0, heskrimiktoplam.Column - hesesdegerkodverisi.Column).value = 0
            cell.Offset(0, hesmaxmiktartoplam.Column - hesesdegerkodverisi.Column).value = 0
        End If

        If cell.Offset(0, hesesdmiktoplam.Column - hesesdegerkodverisi.Column).value < cell.Offset(0, heskrimiktoplam.Column - hesesdegerkodverisi.Column).value Then
            cell.Offset(0, hesgopithmik.Column - hesesdegerkodverisi.Column).value = Round(cell.Offset(0, hesmaxmiktartoplam.Column - hesesdegerkodverisi.Column).value - cell.Offset(0, hesesdmiktoplam.Column - hesesdegerkodverisi.Column).value, 0)
        Else
            cell.Offset(0, hesgopithmik.Column - hesesdegerkodverisi.Column).value = "Pass"
        End If
    Next cell

    UserForm1.ListBox.AddItem "Stok hesaplama iþlemleri tamamlandý."
End Sub
'Data sayfasý ihtyiaç miktarlarý sýralamasý, istediðim gibi deðil ama sanýrým iþ görür
Sub DinamikSirala()
    UserForm1.ListBox.AddItem "Ýhtiyaç fazlasý sýralama iþlemleri baþladý."
    Dim ws As Worksheet
    Dim esdegerCol As Long
    Dim ihtiyacCol As Long
    Dim lastRow As Long
    Dim headerRow As Long
    Dim cell As Range

    Set ws = ThisWorkbook.Sheets("Data")
    headerRow = 1
    For Each cell In ws.Rows(headerRow).Cells
        If cell.value = "Eþdeðer" Then
            esdegerCol = cell.Column
        ElseIf cell.value = "Ýhtiyaç" Then
            ihtiyacCol = cell.Column
        End If
    Next cell

    lastRow = ws.Cells(ws.Rows.count, esdegerCol).End(xlUp).row
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Cells(headerRow + 1, ihtiyacCol), Order:=xlAscending
    ws.Sort.SortFields.Add key:=ws.Cells(headerRow + 1, esdegerCol), Order:=xlAscending
    With ws.Sort
        .SetRange ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column))
        .header = xlYes
        .Apply
    End With

    UserForm1.ListBox.AddItem "Ýhtiyaç fazlasý sýralama iþlemleri tamamlandý."
End Sub

Sub KopyalaHastaneleri()
    UserForm1.ListBox.AddItem "Ýhtiyaç fazlasý bulunan hastane tespiti iþlemleri baþladý."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wsData As Worksheet
    Dim wsHesap As Worksheet
    Dim i As Long, j As Long, k As Long
    Dim ihtiyacRow As Long
    Dim esdegerKodCol As Long, gopIhtMikCol As Long, ihtFazHastAdCol As Long, ihtFazMiktarCol As Long
    Dim hastaneAdiCol As Long, esdegerCol As Long, ihtiyacCol As Long
    Dim lastRow As Long
    Dim esdegerKod As String
    Dim ihtiyacList As Collection
    Dim ihtiyacDict As Object
    Dim ihtiyacArray() As Variant

    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    Set ihtiyacDict = CreateObject("Scripting.Dictionary")

    esdegerKodCol = Application.WorksheetFunction.Match("EþdeðerKod", wsHesap.Rows(1), 0)
    gopIhtMikCol = Application.WorksheetFunction.Match("Ýht. Mik.", wsHesap.Rows(1), 0)
    ihtFazHastAdCol = Application.WorksheetFunction.Match("Ýht. Faz. Hast AD", wsHesap.Rows(1), 0)
    ihtFazMiktarCol = Application.WorksheetFunction.Match("Ýht. Faz. Miktar", wsHesap.Rows(1), 0)
    hastaneAdiCol = Application.WorksheetFunction.Match("Hastane", wsData.Rows(1), 0)
    esdegerCol = Application.WorksheetFunction.Match("Eþdeðer", wsData.Rows(1), 0)
    ihtiyacCol = Application.WorksheetFunction.Match("Ýhtiyaç", wsData.Rows(1), 0)
    lastRow = wsData.Cells(wsData.Rows.count, esdegerCol).End(xlUp).row

    For ihtiyacRow = 2 To lastRow
        esdegerKod = wsData.Cells(ihtiyacRow, esdegerCol).value
        If Not ihtiyacDict.Exists(esdegerKod) Then
            Set ihtiyacDict(esdegerKod) = New Collection
        End If
        ihtiyacDict(esdegerKod).Add Array(wsData.Cells(ihtiyacRow, ihtiyacCol).value, wsData.Cells(ihtiyacRow, hastaneAdiCol).value)
    Next ihtiyacRow

    For i = 2 To wsHesap.Cells(wsHesap.Rows.count, esdegerKodCol).End(xlUp).row
        If wsHesap.Cells(i, gopIhtMikCol).value <> "Pass" Then
            esdegerKod = wsHesap.Cells(i, esdegerKodCol).value
            If ihtiyacDict.Exists(esdegerKod) Then
                Set ihtiyacList = ihtiyacDict(esdegerKod)
                ihtiyacArray = CollectionToArray(ihtiyacList)
                Call QuickSort(ihtiyacArray, LBound(ihtiyacArray, 2), UBound(ihtiyacArray, 2))

                For j = 1 To Application.Min(3, UBound(ihtiyacArray, 2))
                    wsHesap.Cells(i, ihtFazMiktarCol).value = Round(ihtiyacArray(1, j), 0)
                    wsHesap.Cells(i, ihtFazHastAdCol).value = ihtiyacArray(2, j)
                    i = i + 1
                Next j
                For k = j To 3
                    wsHesap.Cells(i, ihtFazMiktarCol).value = ""
                    wsHesap.Cells(i, ihtFazHastAdCol).value = ""
                    i = i + 1
                Next k
                Do While wsHesap.Cells(i, esdegerKodCol).value = esdegerKod
                    wsHesap.Cells(i, ihtFazMiktarCol).value = ""
                    wsHesap.Cells(i, ihtFazHastAdCol).value = ""
                    i = i + 1
                Loop
                i = i - 1
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    UserForm1.ListBox.AddItem "Ýhtiyaç fazlasý bulunan hastane tespiti iþlemleri tamamlandý."
End Sub

Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    ReDim arr(1 To 2, 1 To col.count)
    For i = 1 To col.count
        arr(1, i) = col(i)(0)
        arr(2, i) = col(i)(1)
    Next i
    CollectionToArray = arr
End Function

Sub QuickSort(arr As Variant, first As Long, last As Long)
    Dim low As Long, high As Long, mid As Variant, temp As Variant
    low = first
    high = last
    mid = arr(1, (first + last) \ 2)
    Do While low <= high
        Do While arr(1, low) < mid
            low = low + 1
        Loop
        Do While arr(1, high) > mid
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(1, low)
            arr(1, low) = arr(1, high)
            arr(1, high) = temp
            temp = arr(2, low)
            arr(2, low) = arr(2, high)
            arr(2, high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop
    If first < high Then Call QuickSort(arr, first, high)
    If low < last Then Call QuickSort(arr, low, last)
End Sub

Function IsInCollection(col As Collection, value As Variant) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = col(value)
    IsInCollection = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub UpdateDepoDurumu()
    UserForm1.ListBox.AddItem "Tedarikçi ecza deposu tespiti iþlemleri baþladý."
    Dim wsHesap As Worksheet
    Dim wsAnlMuad As Worksheet
    Dim lastRowHesap As Long
    Dim lastRowAnlMuad As Long
    Dim ihtMikCol As Long
    Dim esdegerKodCol As Long
    Dim depoDurumuCol As Long
    Dim esdegerCol As Long
    Dim tedarikciCol As Long
    Dim aciklamaCol As Long
    Dim i As Long
    Dim j As Long
    Dim esdegerKod As String
    Dim ihtMik As String
    Dim tedarikci As String
    Dim aciklama As String
    Dim esdegerCount As Object

    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    Set wsAnlMuad = ThisWorkbook.Sheets("AnlMuad")

    lastRowHesap = wsHesap.Cells(wsHesap.Rows.count, 1).End(xlUp).row
    lastRowAnlMuad = wsAnlMuad.Cells(wsAnlMuad.Rows.count, 1).End(xlUp).row

    ihtMikCol = wsHesap.Rows(1).Find("Ýht. Mik.").Column
    esdegerKodCol = wsHesap.Rows(1).Find("EþdeðerKod").Column
    depoDurumuCol = wsHesap.Rows(1).Find("Depo Adý & Durumu").Column
    esdegerCol = wsAnlMuad.Rows(1).Find("Eþdeðer").Column
    tedarikciCol = wsAnlMuad.Rows(1).Find("Tedarikçi").Column
    aciklamaCol = wsAnlMuad.Rows(1).Find("Açýklama").Column

    Set esdegerCount = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRowHesap
        ihtMik = wsHesap.Cells(i, ihtMikCol).value
        If ihtMik <> "Pass" Then
            esdegerKod = wsHesap.Cells(i, esdegerKodCol).value
            If Not esdegerCount.Exists(esdegerKod) Then
                esdegerCount(esdegerKod) = 1
            Else
                esdegerCount(esdegerKod) = esdegerCount(esdegerKod) + 1
            End If

            Dim foundCount As Long
            foundCount = 0
            For j = 2 To lastRowAnlMuad
                If wsAnlMuad.Cells(j, esdegerCol).value = esdegerKod Then
                    foundCount = foundCount + 1
                    If foundCount = esdegerCount(esdegerKod) Then
                        tedarikci = wsAnlMuad.Cells(j, tedarikciCol).value
                        aciklama = wsAnlMuad.Cells(j, aciklamaCol).value
                        wsHesap.Cells(i, depoDurumuCol).value = aciklama & " - " & tedarikci
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i

    UserForm1.ListBox.AddItem "Tedarikçi ecza deposu tespiti iþlemleri tamamlandý."
End Sub

Sub PivotTabloyuYenile()
    Dim wsPVT As Worksheet
    Dim wsDepo As Worksheet
    Dim ptHastane As PivotTable
    Dim ptDepo As PivotTable

    Set wsPVT = ThisWorkbook.Sheets("PVT")
    Set ptHastane = wsPVT.PivotTables("hastanepvt")
    Set wsDepo = ThisWorkbook.Sheets("depo")
    Set ptDepo = wsDepo.PivotTables("depopvt")

    UserForm1.ListBox.AddItem "Pivot tablo güncellemeleri baþladý."
    ptHastane.RefreshTable
    ptDepo.RefreshTable
    UserForm1.ListBox.AddItem "Pivot tablo güncellemeleri tamamlandý."
End Sub

Sub SendEmail()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim wsOrg As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim emailBody As String
    Dim hospitalName As String
    Dim emailAddress As String
    Dim pharmacistName As String
    Dim findRow As Range
    Dim senderEmail As String
    Dim senderHospitalName As String
    Dim searchRange As Range
    Dim foundCell As Range
    Dim firstAddress As String

    Set OutlookApp = CreateObject("Outlook.Application")
    Set ws = ThisWorkbook.Sheets("PVT")
    Set wsOrg = ThisWorkbook.Sheets("Org")
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).row
    Set rng = ws.Range("C2:I" & lastRow)

    ws.Columns("C").Hidden = False
    Set searchRange = ws.Range("C3:C" & ws.Cells(ws.Rows.count, "C").End(xlUp).row)
    hospitalName = ws.Range("C3").value

    Dim cell As Range
    For Each cell In searchRange
        If cell.value <> "" And cell.value <> hospitalName Then
            MsgBox "Ýhtiyaç fazlasý ilaçlarý içeren hastaneler sütununda farklý hastane adlarý tespit edildi." & vbCrLf & "Lütfen her iþlemde yalnýzca bir hastane seçiniz.", vbExclamation
            ws.Columns("C").Hidden = True
            Exit Sub
        End If
    Next cell

    ws.Columns("C").Hidden = True
    Set findRow = wsOrg.Columns("B").Find(What:=hospitalName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not findRow Is Nothing Then
        pharmacistName = findRow.Offset(0, 1).value
        emailAddress = findRow.Offset(0, 2).value

        If emailAddress <> "" Then
            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                .To = emailAddress
                .Cc = "umit.yazir@mlpcare.com;ceyda.simsek@mlpcare.com"
                .Subject = "Ýlaç Ýhtiyaç Fazlasý Talebi Hk."
                .Display

                senderEmail = .Session.Accounts.Item(1).SmtpAddress
                Set findRow = wsOrg.Columns("D").Find(What:=senderEmail, LookIn:=xlValues, LookAt:=xlWhole)
                If Not findRow Is Nothing Then
                    senderHospitalName = findRow.Offset(0, -2).value
                Else
                    senderHospitalName = "Bilinmiyor"
                End If

                Dim dataContent As String
                dataContent = "<table border='1' style='border-collapse:collapse;'>"
                dataContent = dataContent & "<tr><td colspan='4' style='font-weight:bold; background-color:lightblue; text-align:center;'>" & hospitalName & "</td><td colspan='3' style='font-weight:bold; background-color:lightblue; text-align:center;'>" & senderHospitalName & "</td></tr>"
                For Each cell In rng.Rows
                    dataContent = dataContent & "<tr>"
                    For Each dataCell In cell.Cells
                        If cell.row = 1 Or cell.row = 2 Then
                            If dataCell.Column = 7 Or dataCell.Column = 8 Or dataCell.Column = 9 Then
                                dataContent = dataContent & "<td style='font-weight:bold; background-color:lightgreen; word-wrap:break-word; text-align:center;'>" & dataCell.value & "</td>"
                            Else
                                dataContent = dataContent & "<td style='font-weight:bold; background-color:lightblue; word-wrap:break-word; text-align:center;'>" & dataCell.value & "</td>"
                            End If
                        ElseIf dataCell.Column = 6 Then
                            dataContent = dataContent & "<td style='word-wrap:break-word; width:1.8cm; text-align:right; background-color:" & IIf(cell.row Mod 2 = 0, "lightblue;", "white;") & "'>" & dataCell.value & "</td>"
                        ElseIf dataCell.Column = 7 Or dataCell.Column = 8 Then
                            dataContent = dataContent & "<td style='word-wrap:break-word; width:1.8cm; text-align:right; background-color:" & IIf(cell.row Mod 2 = 0, "lightgrey;", "white;") & "'>" & dataCell.value & "</td>"
                        ElseIf dataCell.Column = 9 Then
                            dataContent = dataContent & "<td style='word-wrap:break-word; width:1.8cm; text-align:right; background-color:" & IIf(cell.row Mod 2 = 0, "lightgreen;", "white;") & "'>" & dataCell.value & "</td>"
                        Else
                            dataContent = dataContent & "<td style='word-wrap:break-word; background-color:" & IIf(cell.row Mod 2 = 0, "lightgrey;", "white;") & "'>" & dataCell.value & "</td>"
                        End If
                    Next dataCell
                    If cell.row = 2 Then
                        dataContent = dataContent & "<td style='word-wrap:break-word; width:1.6cm; font-weight:bold; text-align:center;background-color:" & IIf(cell.row Mod 2 = 0, "lightblue;", "white;") & "'>Karþ. Miktar (Kt)</td>"
                        dataContent = dataContent & "<td style='word-wrap:break-word; width:5cm; font-weight:bold; text-align:center;background-color:" & IIf(cell.row Mod 2 = 0, "lightblue;", "white;") & "'>Açýklamalar</td>"
                    Else
                        dataContent = dataContent & "<td style='word-wrap:break-word; width:1.6cm; background-color:" & IIf(cell.row Mod 2 = 0, "lightgrey;", "white;") & "'></td>"
                        dataContent = dataContent & "<td style='word-wrap:break-word; width:5cm; background-color:" & IIf(cell.row Mod 2 = 0, "lightgrey;", "white;") & "'></td>"
                    End If
                    dataContent = dataContent & "</tr>"
                Next cell
                dataContent = dataContent & "</table>"

                emailBody = "<span style='font-size:12pt; font-family:Times New Roman;'>" & _
                            "Merhaba " & pharmacistName & "," & "<br><br>" & _
                            "Aþaðýdaki tabloda sizin ihtiyaç fazlanýz bizimse ihtiyaç duyduðumuz ilaçlarýn listesi ve ihtiyaç miktarlarýmýz görünmektedir." & "<br>" & _
                            "Mümkünse ihtiyacýmýz kadar deðilse sizin uygun gördüðünüz miktarlarda yardýmcý olmanýzý rica ediyoruz." & "<br><br>" & _
                            "Teþekkürler, iyi çalýþmalar." & "<br><br>" & _
                            dataContent & "<br><br>" & _
                            "* Bu mail Satýn Alma Çalýþmasý Beta 5.1 tarafýndan otomatik olarak oluþturulmuþtur. Yanlýþlýk olduðunu düþünüyorsanýz lütfen Ecz. Harun Topal ile iletiþime geçiniz." & _
                            "</span>"

                .HTMLBody = emailBody & "<br><br>" & .HTMLBody
            End With
        End If
    End If

    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub


