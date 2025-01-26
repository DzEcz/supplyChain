Attribute VB_Name = "Module1"
Sub AnaProsedur()
        ' Optimizasyonlar� kapat
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo HataYakalama ' Hata yakalama
    
    Dim currentSheet As Worksheet
    
    ' Mevcut aktif sayfay� belirle
    Set currentSheet = ActiveSheet
    
    ' UserForm'u g�ster
    UserForm1.Show vbModeless
    UserForm1.Caption = "�lerleme Durumu"
    DoEvents ' UserForm'un g�ncellenmesini sa�lar
    
    ' T�m d��meleri pasif yap
    UserForm1.CommandButton1.Enabled = False
    UserForm1.CommandButton2.Enabled = False
    UserForm1.CommandButton3.Enabled = False
    
    ' Hesap sayfas�n�n kilidini a�
    Sheets("Hesap").Unprotect Password:="8142" ' �ifreyi kendi belirledi�iniz �ifre ile de�i�tirin
    Sheets("Pusula").Unprotect Password:="8142" ' �ifreyi kendi belirledi�iniz �ifre ile de�i�tirin
    
    ' ��lemleri ger�ekle�tir
    Call Adshow
    DoEvents
    Call PusulaSayfasiniGuncelle
    DoEvents
    Call VeriKopyala
    DoEvents
    Call KopyalaVeEkleHizli
    DoEvents
    Call KutuMiktarKopyala
    DoEvents
    Call EsdegerToplam
    DoEvents
    Call DinamikSirala
    DoEvents
    Call KopyalaHastaneleri
    DoEvents
    Call UpdateDepoDurumu
    DoEvents
    Call PivotTabloyuYenile
    DoEvents
    
    ' ��lemler tamamland���nda bildirim ekle
    UserForm1.ListBox.AddItem "T�m i�lemler ba�ar�yla ger�ekle�ti."
    
    ' Hesap sayfas�n� tekrar kilitle
    Sheets("Hesap").Protect Password:="8142" ' �ifreyi kendi belirledi�iniz �ifre ile de�i�tirin
    Sheets("Pusula").Protect Password:="8142" ' �ifreyi kendi belirledi�iniz �ifre ile de�i�tirin
    
    ' Ba�lat�lan sayfaya geri d�n
    currentSheet.Activate
    
    ' Kapatma butonunu aktif yap
    UserForm1.CommandButton1.Enabled = True
    UserForm1.CommandButton3.Enabled = True

    Exit Sub

HataYakalama:
    ' Hata durumunda UserForm'u gizle ve hata mesaj�n� g�ster
    MsgBox "Bir hata olu�tu: " & Err.Description & vbCrLf & _
           "Prosed�r: " & Err.Source & vbCrLf & _
           "Sat�r: " & Erl, vbCritical
           
    ' Optimizasyonlar� a�
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


Sub Adshow()

    Application.StatusBar = "Ecz. Harun Topal"

End Sub

Sub PusulaSayfasiniGuncelle()

UserForm1.ListBox.AddItem "Pusula sayfas� g�ncelleme i�lemi ba�lad�."
    Dim kaynakKitap As Workbook
    Dim hedefKitap As Workbook
    Dim kaynakSayfa As Worksheet
    Dim hedefSayfa As Worksheet
    Dim kaynakDosyaYolu As String
    
    ' Kaynak dosya yolunu belirleyin
    kaynakDosyaYolu = ThisWorkbook.Path & "\Pusula.xlsx"
    
    ' Kaynak �al��ma kitab�n� a��n
    Set kaynakKitap = Workbooks.Open(kaynakDosyaYolu)
    Set kaynakSayfa = kaynakKitap.Sheets("Sheet")
    
    ' Hedef �al��ma kitab�n� ve sayfas�n� belirleyin
    Set hedefKitap = ThisWorkbook
    Set hedefSayfa = hedefKitap.Sheets("Pusula")
    
    ' Hedef sayfadaki mevcut verileri temizleyin
    hedefSayfa.Cells.Clear
    
    ' Kaynak sayfadaki verileri kopyalay�n
    kaynakSayfa.UsedRange.Copy
    
    ' Verileri hedef sayfaya yap��t�r�n
    hedefSayfa.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    ' Kaynak �al��ma kitab�n� kapat�n
    kaynakKitap.Close False
    
    ' Kullan�c�ya bildirimde bulunun
UserForm1.ListBox.AddItem "Pusula sayfas�g�ncelleme i�lemi tamamland�."
End Sub

Sub VeriKopyala()

UserForm1.ListBox.AddItem "Pusula sayfas�ndan veri kopyalama i�lemi ba�lad�."
   
    Dim wsPusula As Worksheet
    Dim wsHesap As Worksheet
    Dim lastRow As Long
    Dim kodCol As Long
    Dim adCol As Long
    Dim miktarCol As Long
    Dim kodData As Variant
    Dim i As Long
    
    ' �al��ma sayfalar�n� tan�mla
    Set wsPusula = ThisWorkbook.Sheets("Pusula")
    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    
    ' Pusula sayfas�ndaki son sat�r� bul
    lastRow = wsPusula.Cells(wsPusula.Rows.count, "A").End(xlUp).row
    
    ' Pusula sayfas�nda veri olup olmad���n� kontrol et
    If lastRow < 2 Then
        MsgBox "L�tfen Pusuladan �ekti�iniz stok durum raporunu ayn� klas�re kopyalay�n�z!", vbExclamation
        wsPusula.Activate
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Hesap sayfas�ndaki verileri kontrol et ve gerekirse sil
    If wsHesap.Cells(2, 1).value <> "" Then
        wsHesap.Rows("2:" & wsHesap.Rows.count).ClearContents
    End If
    
    ' S�tun numaralar�n� bul
    kodCol = wsPusula.Rows(1).Find("C. EMR E�de�er �r�n Grup Kodu").Column
    adCol = wsPusula.Rows(1).Find("Ad�").Column
    miktarCol = wsPusula.Rows(1).Find("Miktar").Column
    
    ' Pusula sayfas�ndaki kod verilerini diziye al
    kodData = wsPusula.Range(wsPusula.Cells(2, kodCol), wsPusula.Cells(lastRow, kodCol)).value
    
    ' Kod verilerini say�ya d�n��t�r ve ondal�k olmamas�n� sa�la
    For i = 1 To UBound(kodData, 1)
        If IsNumeric(kodData(i, 1)) Then
            kodData(i, 1) = Round(CDbl(kodData(i, 1)), 0)
        End If
    Next i
    
    ' Hesap sayfas�ndaki ba�l�klar� yaz
    wsHesap.Cells(1, 1).value = "E�de�erKod"
    wsHesap.Cells(1, 2).value = "M�stahzar"
    wsHesap.Cells(1, 3).value = "Stok Miktar"
    
    ' Pusula sayfas�ndaki verileri Hesap sayfas�na kopyala
    wsHesap.Range("A2:A" & lastRow).value = kodData
    wsHesap.Range("B2:B" & lastRow).value = wsPusula.Range(wsPusula.Cells(2, adCol), wsPusula.Cells(lastRow, adCol)).value
    wsHesap.Range("C2:C" & lastRow).value = wsPusula.Range(wsPusula.Cells(2, miktarCol), wsPusula.Cells(lastRow, miktarCol)).value
    
  
UserForm1.ListBox.AddItem "Pusula sayfas�ndan veri kopyalama i�lemi tamamland�."
End Sub

'e�de�erkodlar� ��e tamamla;

Sub KopyalaVeEkleHizli()

UserForm1.ListBox.AddItem "M�stahzar say�s�n�n ��lemesi i�lemi ba�lad�."
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
    
    Application.ScreenUpdating = False ' Ekran g�ncellemelerini kapat
    Application.Calculation = xlCalculationManual ' Otomatik hesaplamay� kapat
    
    Set ws = ThisWorkbook.Sheets("Hesap") ' �al��ma sayfas�n� tan�mla
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row ' Son sat�r� bul
    
    ' S�tun ba�l�klar�n� bul
    esdegerKodCol = Application.WorksheetFunction.Match("E�de�erKod", ws.Rows(1), 0)
    mustahzarCol = Application.WorksheetFunction.Match("M�stahzar", ws.Rows(1), 0)
    stokMiktarCol = Application.WorksheetFunction.Match("Stok Miktar", ws.Rows(1), 0)
    
    ' Verileri diziye al
    data = ws.Range(ws.Cells(2, esdegerKodCol), ws.Cells(lastRow, stokMiktarCol)).value
    
    ' Sonu� dizisini ba�lat
    ReDim result(1 To (lastRow - 1) * 2, 1 To UBound(data, 2))
    resultIndex = 1
    
    ' E�de�er Kodlar� say
    For i = 1 To UBound(data, 1)
        kod = data(i, 1) ' E�de�er Kod s�tunu
        If kodCount.exists(kod) Then
            kodCount(kod) = kodCount(kod) + 1
        Else
            kodCount.Add kod, 1
        End If
    Next i
    
    ' �ki adet olan E�de�er Kodlar� kopyala
    For i = 1 To UBound(data, 1)
        kod = data(i, 1)
        If kodCount(kod) = 2 Then
            result(resultIndex, 1) = data(i, 1) ' E�de�erKod
            result(resultIndex, 2) = data(i, 2) & "_kopya" & kodCount(kod) ' M�stahzar
            result(resultIndex, 3) = data(i, 3) ' Stok Miktar
            resultIndex = resultIndex + 1
            kodCount(kod) = kodCount(kod) + 1
        End If
    Next i
    
    ' Bir adet olan E�de�er Kodlar� kopyala
    For i = 1 To UBound(data, 1)
        kod = data(i, 1)
        If kodCount(kod) = 1 Then
            ' �ki kopya ekle
            For j = 1 To 2
                If kodCount(kod) < 3 Then
                    result(resultIndex, 1) = data(i, 1) ' E�de�erKod
                    result(resultIndex, 2) = data(i, 2) & "_kopya" & kodCount(kod) ' M�stahzar
                    result(resultIndex, 3) = data(i, 3) ' Stok Miktar
                    resultIndex = resultIndex + 1
                    kodCount(kod) = kodCount(kod) + 1
                End If
            Next j
        End If
    Next i
    
    ' Sonu�lar� �al��ma sayfas�na yaz
    ws.Range(ws.Cells(lastRow + 1, esdegerKodCol), ws.Cells(lastRow + resultIndex - 1, stokMiktarCol)).value = result
    
    ' E�de�erKod verilerini alfabetik olarak s�ralama
    ws.Range(ws.Cells(2, esdegerKodCol), ws.Cells(lastRow + resultIndex - 1, stokMiktarCol)).Sort Key1:=ws.Cells(2, esdegerKodCol), Order1:=xlAscending, header:=xlNo
    
    Application.ScreenUpdating = True ' Ekran g�ncellemelerini a�
    Application.Calculation = xlCalculationAutomatic ' Otomatik hesaplamay� a�

UserForm1.ListBox.AddItem "M�stahzar say�s�n�n ��lemesi i�lemi tamamland�."
End Sub

Sub KutuMiktarKopyala()

UserForm1.ListBox.AddItem "Kutu i�i miktarlar�n�n kpyalanmas� i�lemi ba�lad�."
    Dim wsHesap As Worksheet
    Dim wsKutui�i As Worksheet
    Dim rngHesap As Range
    Dim rngKutui�i As Range
    Dim cell As Range
    Dim matchRow As Variant
    Dim colEsdegerKodHesap As Long
    Dim colKutuMiktarHesap As Long
    Dim colEsdegerKodKutui�i As Long
    Dim colKutuIciKutui�i As Long
    Dim hesapData As Variant
    Dim kutuiciData As Variant
    Dim i As Long
    Dim dict As Object
    
    ' Sayfalar� tan�mla
    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    Set wsKutui�i = ThisWorkbook.Sheets("Kutui�i")
    
    ' S�tun ba�l�klar�n�n yerini bul
    colEsdegerKodHesap = Application.Match("E�de�erKod", wsHesap.Rows(1), 0)
    colKutuMiktarHesap = Application.Match("Kutu Miktar", wsHesap.Rows(1), 0)
    colEsdegerKodKutui�i = Application.Match("E�de�er", wsKutui�i.Rows(1), 0)
    colKutuIciKutui�i = Application.Match("Kutu ��i", wsKutui�i.Rows(1), 0)
    
    ' Verileri diziye al
    hesapData = wsHesap.Range(wsHesap.Cells(2, colEsdegerKodHesap), wsHesap.Cells(wsHesap.Rows.count, colEsdegerKodHesap).End(xlUp)).Resize(, colKutuMiktarHesap - colEsdegerKodHesap + 1).value
    kutuiciData = wsKutui�i.Range(wsKutui�i.Cells(2, colEsdegerKodKutui�i), wsKutui�i.Cells(wsKutui�i.Rows.count, colEsdegerKodKutui�i).End(xlUp)).Resize(, colKutuIciKutui�i - colEsdegerKodKutui�i + 1).value
    
    ' E�de�er kodlar� ve kutu i�i miktarlar�n� bir s�zl�kte sakla
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(kutuiciData, 1)
        dict(kutuiciData(i, 1)) = kutuiciData(i, 2)
    Next i
    
    ' Ekran g�ncellemelerini ve hesaplamalar� kapat
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Hesap sayfas�ndaki her bir E�de�erKod i�in
    For i = 1 To UBound(hesapData, 1)
        If dict.exists(hesapData(i, 1)) Then
            hesapData(i, colKutuMiktarHesap - colEsdegerKodHesap + 1) = dict(hesapData(i, 1))
        Else
            hesapData(i, colKutuMiktarHesap - colEsdegerKodHesap + 1) = 1
        End If
    Next i
    
    ' Sonu�lar� �al��ma sayfas�na yaz
    wsHesap.Range(wsHesap.Cells(2, colEsdegerKodHesap), wsHesap.Cells(UBound(hesapData, 1) + 1, colKutuMiktarHesap)).value = hesapData
    
    ' Ekran g�ncellemelerini ve hesaplamalar� a�
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

UserForm1.ListBox.AddItem "Kutu i�i miktarlar�n�n kopyalanmas� i�lemi tamamland�."
End Sub
Sub EsdegerToplam()

UserForm1.ListBox.AddItem "Stok hesaplama i�lemleri ba�lad�."
    
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
    
    ' Sayfalar� tan�mla
    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    Set wsPusula = ThisWorkbook.Sheets("Pusula")
    
    ' Hesap sayfas�ndaki s�tunlar� bul
    Set hesesdegerkodverisi = wsHesap.Rows(1).Find("E�de�erKod")
    Set heskutumiktarverisi = wsHesap.Rows(1).Find("Kutu Miktar")
    Set hesesdmiktoplam = wsHesap.Rows(1).Find("E�d.Mik. TOPLAM")
    Set heskrimiktoplam = wsHesap.Rows(1).Find("Kri.Mik. TOPLAM")
    Set hesmaxmiktartoplam = wsHesap.Rows(1).Find("Max.Mik TOPLAM")
    Set hesgopithmik = wsHesap.Rows(1).Find("�ht. Mik.")
    
    ' Pusula sayfas�ndaki s�tunlar� bul
    Set pusesdegerkodverisi = wsPusula.Rows(1).Find("C. EMR E�de�er �r�n Grup Kodu")
    Set pusmikverisi = wsPusula.Rows(1).Find("Miktar")
    Set puskrimikverisi = wsPusula.Rows(1).Find("Kritik Miktar")
    Set pusmaxmikverisi = wsPusula.Rows(1).Find("Max Miktar")
    
    ' Hesap sayfas�ndaki her bir E�de�erKod icin i�lemleri yap
    For Each cell In wsHesap.Range(hesesdegerkodverisi.Offset(1, 0), wsHesap.Cells(wsHesap.Rows.count, hesesdegerkodverisi.Column).End(xlUp))
        kod = Trim(UCase(cell.value))
        toplam = 0
        krimiktoplam = 0
        maxmiktartoplam = 0
        
        ' Pusula sayfas�nda e�le�en kodlar� bul ve miktarlar� topla
        For Each pCell In wsPusula.Range(pusesdegerkodverisi.Offset(1, 0), wsPusula.Cells(wsPusula.Rows.count, pusesdegerkodverisi.Column).End(xlUp))
            If Trim(UCase(pCell.value)) = kod Then
                toplam = toplam + CDbl(pCell.Offset(0, pusmikverisi.Column - pusesdegerkodverisi.Column).value)
                krimiktoplam = krimiktoplam + CDbl(pCell.Offset(0, puskrimikverisi.Column - pusesdegerkodverisi.Column).value)
                maxmiktartoplam = maxmiktartoplam + CDbl(pCell.Offset(0, pusmaxmikverisi.Column - pusesdegerkodverisi.Column).value)
            End If
        Next pCell
        
        ' Toplam� Kutu Miktar'a b�l ve sonucu ilgili s�tunlara yaz
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
        
        ' �ht. Mik. s�tununu hesapla
        If cell.Offset(0, hesesdmiktoplam.Column - hesesdegerkodverisi.Column).value < cell.Offset(0, heskrimiktoplam.Column - hesesdegerkodverisi.Column).value Then
            cell.Offset(0, hesgopithmik.Column - hesesdegerkodverisi.Column).value = Round(cell.Offset(0, hesmaxmiktartoplam.Column - hesesdegerkodverisi.Column).value - cell.Offset(0, hesesdmiktoplam.Column - hesesdegerkodverisi.Column).value, 0)
        Else
            cell.Offset(0, hesgopithmik.Column - hesesdegerkodverisi.Column).value = "Pass"
        End If
    Next cell
    
UserForm1.ListBox.AddItem "Stok hesaplama i�lemleri tamamland�."
End Sub
'Data sayfas� ihtyia� miktarlar� s�ralamas�, istedi�im gibi de�il ama san�r�m i� g�r�r
Sub DinamikSirala()

UserForm1.ListBox.AddItem "�htiya� fazlas� s�ralama i�lemleri ba�lad�."
    Dim ws As Worksheet
    Dim esdegerCol As Long
    Dim ihtiyacCol As Long
    Dim lastRow As Long
    Dim headerRow As Long
    Dim cell As Range

    ' �al��ma sayfas�n� belirle
    Set ws = ThisWorkbook.Sheets("Data") ' Sayfa ad�n� ihtiyac�n�za g�re de�i�tirin

    ' Ba�l�k sat�r�n� belirle
    headerRow = 1 ' Ba�l�k sat�r�n�n numaras�n� ihtiyac�n�za g�re de�i�tirin

    ' "E�de�er" ve "�htiya�" s�tunlar�n� bul
    For Each cell In ws.Rows(headerRow).Cells
        If cell.value = "E�de�er" Then
            esdegerCol = cell.Column
        ElseIf cell.value = "�htiya�" Then
            ihtiyacCol = cell.Column
        End If
    Next cell

    ' Son sat�r� bul
    lastRow = ws.Cells(ws.Rows.count, esdegerCol).End(xlUp).row

    ' S�ralama i�lemi
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Cells(headerRow + 1, ihtiyacCol), Order:=xlAscending
    ws.Sort.SortFields.Add key:=ws.Cells(headerRow + 1, esdegerCol), Order:=xlAscending
    With ws.Sort
        .SetRange ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column))
        .header = xlYes
        .Apply
    End With

UserForm1.ListBox.AddItem "�htiya� fazlas� s�ralama i�lemleri tamamland�."
End Sub

'ihtyia� fazlas� hastaneleri kopyalama
Sub KopyalaHastaneleri()

UserForm1.ListBox.AddItem "�htiya� fazlas� bulunan hastane tespiti i�lemleri ba�lad�."
    ' Optimizasyonlar� kapat
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
    
    ' Hesap sayfas�ndaki s�tun ba�l�klar�n� bul
    esdegerKodCol = Application.WorksheetFunction.Match("E�de�erKod", wsHesap.Rows(1), 0)
    gopIhtMikCol = Application.WorksheetFunction.Match("�ht. Mik.", wsHesap.Rows(1), 0)
    ihtFazHastAdCol = Application.WorksheetFunction.Match("�ht. Faz. Hast AD", wsHesap.Rows(1), 0)
    ihtFazMiktarCol = Application.WorksheetFunction.Match("�ht. Faz. Miktar", wsHesap.Rows(1), 0)
    
    ' Data sayfas�ndaki s�tun ba�l�klar�n� bul
    hastaneAdiCol = Application.WorksheetFunction.Match("Hastane", wsData.Rows(1), 0)
    esdegerCol = Application.WorksheetFunction.Match("E�de�er", wsData.Rows(1), 0)
    ihtiyacCol = Application.WorksheetFunction.Match("�htiya�", wsData.Rows(1), 0)
    
    lastRow = wsData.Cells(wsData.Rows.count, esdegerCol).End(xlUp).row
    
    ' Data sayfas�ndaki her bir E�de�erKod icin �htiya� ve Hastane Ad� bilgilerini topla
    For ihtiyacRow = 2 To lastRow
        esdegerKod = wsData.Cells(ihtiyacRow, esdegerCol).value
        If Not ihtiyacDict.exists(esdegerKod) Then
            Set ihtiyacDict(esdegerKod) = New Collection
        End If
        ihtiyacDict(esdegerKod).Add Array(wsData.Cells(ihtiyacRow, ihtiyacCol).value, wsData.Cells(ihtiyacRow, hastaneAdiCol).value)
    Next ihtiyacRow
    
    ' Hesap sayfas�ndaki her bir E�de�erKod icin i�lemleri yap
    For i = 2 To wsHesap.Cells(wsHesap.Rows.count, esdegerKodCol).End(xlUp).row
        If wsHesap.Cells(i, gopIhtMikCol).value <> "Pass" Then
            esdegerKod = wsHesap.Cells(i, esdegerKodCol).value
            If ihtiyacDict.exists(esdegerKod) Then
                Set ihtiyacList = ihtiyacDict(esdegerKod)
                ' �htiya� miktarlar�na g�re k���kten b�y��e s�rala
                ihtiyacArray = CollectionToArray(ihtiyacList)
                Call QuickSort(ihtiyacArray, LBound(ihtiyacArray, 2), UBound(ihtiyacArray, 2))
                
                ' �lk �� hastane ve ihtiya� miktar�n� alt alta kopyala
                For j = 1 To Application.Min(3, UBound(ihtiyacArray, 2))
                    wsHesap.Cells(i, ihtFazMiktarCol).value = Round(ihtiyacArray(1, j), 0)
                    wsHesap.Cells(i, ihtFazHastAdCol).value = ihtiyacArray(2, j)
                    i = i + 1
                Next j
                ' Di�er sat�rlar� bo� b�rak
                For k = j To 3
                    wsHesap.Cells(i, ihtFazMiktarCol).value = ""
                    wsHesap.Cells(i, ihtFazHastAdCol).value = ""
                    i = i + 1
                Next k
                ' Ayn� E�de�erKod icin kopyalamay� durdur
                Do While wsHesap.Cells(i, esdegerKodCol).value = esdegerKod
                    wsHesap.Cells(i, ihtFazMiktarCol).value = ""
                    wsHesap.Cells(i, ihtFazHastAdCol).value = ""
                    i = i + 1
                Loop
                i = i - 1
            End If
        End If
    Next i
    
    ' Optimizasyonlar� a�
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    
UserForm1.ListBox.AddItem "�htiya� fazlas� bulunan hastane tespiti i�lemleri tamamland�."
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

UserForm1.ListBox.AddItem "Tedarik�i ecza deposu tespiti i�lemleri ba�lad�."
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
    
    ' �al��ma sayfalar�n� tan�mla
    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    Set wsAnlMuad = ThisWorkbook.Sheets("AnlMuad")
    
    ' Son sat�rlar� bul
    lastRowHesap = wsHesap.Cells(wsHesap.Rows.count, 1).End(xlUp).row
    lastRowAnlMuad = wsAnlMuad.Cells(wsAnlMuad.Rows.count, 1).End(xlUp).row
    
    ' S�tun ba�l�klar�n�n yerlerini bul
    ihtMikCol = wsHesap.Rows(1).Find("�ht. Mik.").Column
    esdegerKodCol = wsHesap.Rows(1).Find("E�de�erKod").Column
    depoDurumuCol = wsHesap.Rows(1).Find("Depo Ad� & Durumu").Column
    esdegerCol = wsAnlMuad.Rows(1).Find("E�de�er").Column
    tedarikciCol = wsAnlMuad.Rows(1).Find("Tedarik�i").Column
    aciklamaCol = wsAnlMuad.Rows(1).Find("A��klama").Column
    
    ' E�de�er kodlar�n�n say�s�n� takip etmek i�in Scripting.Dictionary kullan
    Set esdegerCount = CreateObject("Scripting.Dictionary")
    
    ' Hesap sayfas�nda d�ng�
    For i = 2 To lastRowHesap
        ihtMik = wsHesap.Cells(i, ihtMikCol).value
        If ihtMik <> "Pass" Then
            esdegerKod = wsHesap.Cells(i, esdegerKodCol).value
            ' E�de�er kodunun say�s�n� art�r
            If Not esdegerCount.exists(esdegerKod) Then
                esdegerCount(esdegerKod) = 1
            Else
                esdegerCount(esdegerKod) = esdegerCount(esdegerKod) + 1
            End If
            
            ' AnlMuad sayfas�nda e�de�er kodu ara
            Dim foundCount As Long
            foundCount = 0
            For j = 2 To lastRowAnlMuad
                If wsAnlMuad.Cells(j, esdegerCol).value = esdegerKod Then
                    foundCount = foundCount + 1
                    If foundCount = esdegerCount(esdegerKod) Then
                        tedarikci = wsAnlMuad.Cells(j, tedarikciCol).value
                        aciklama = wsAnlMuad.Cells(j, aciklamaCol).value
                        ' Tedarik�i ve A��klama bilgilerini birle�tir ve yaz
                        wsHesap.Cells(i, depoDurumuCol).value = aciklama & " - " & tedarikci
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
    
UserForm1.ListBox.AddItem "Tedarik�i ecza deposu tespiti i�lemleri tamamland�."
End Sub

Sub PivotTabloyuYenile()
    Dim wsPVT As Worksheet
    Dim wsDepo As Worksheet
    Dim ptHastane As PivotTable
    Dim ptDepo As PivotTable
    
    ' PVT sayfas�ndaki pivot tabloyu tan�mlay�n
    Set wsPVT = ThisWorkbook.Sheets("PVT")
    Set ptHastane = wsPVT.PivotTables("hastanepvt") ' Pivot tablo ad�n� buraya yaz�n
    
    ' Yeni sayfadaki pivot tabloyu tan�mlay�n
    Set wsDepo = ThisWorkbook.Sheets("depo") ' Yeni sayfan�z�n ad�n� buraya yaz�n
    Set ptDepo = wsDepo.PivotTables("depopvt") ' Yeni pivot tablo ad�n� buraya yaz�n
    
    ' Pivot tablolar� yenileyin
    UserForm1.ListBox.AddItem "Pivot tablo g�ncellemeleri ba�lad�."
    ptHastane.RefreshTable
    ptDepo.RefreshTable
    UserForm1.ListBox.AddItem "Pivot tablo g�ncellemeleri tamamland�."
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
    Dim count As Integer
    
    ' Outlook uygulamas�n� ba�lat
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' �al��ma sayfas�n� belirle
    Set ws = ThisWorkbook.Sheets("PVT") ' Pivot tablonun bulundu�u sayfa ad�
    Set wsOrg = ThisWorkbook.Sheets("Org") ' Org sayfas�
    
    ' E s�tunundaki son dolu sat�r� bul
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).row
    
    ' Dinamik veri aral���n� belirle
    Set rng = ws.Range("C2:I" & lastRow)
    
' C s�tununu ge�ici olarak g�r�n�r yap
ws.Columns("C").Hidden = False

' Hastane ad�n� C3 h�cresinden ba�layarak t�m C s�tununda ara
Set searchRange = ws.Range("C3:C" & ws.Cells(ws.Rows.count, "C").End(xlUp).row)
hospitalName = ws.Range("C3").value

' C s�tununda farkl� hastane adlar� olup olmad���n� kontrol et
Dim cell As Range
For Each cell In searchRange
    If cell.value <> "" And cell.value <> hospitalName Then
        MsgBox "�htiya� fazlas� ila�lar� i�eren hastaneler s�tununda farkl� hastane adlar� tespit edildi." & vbCrLf & "L�tfen her i�lemde yaln�zca bir hastane se�iniz.", vbExclamation
        ws.Columns("C").Hidden = True
        Exit Sub
    End If
Next cell

' C s�tununu tekrar gizle
ws.Columns("C").Hidden = True
  
    ' K�saltma sayfas�nda hastane ad�n� bul
    Set findRow = wsOrg.Columns("B").Find(What:=hospitalName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not findRow Is Nothing Then
        ' Eczac�n�n ad�n� ve e-posta adresini al
        pharmacistName = findRow.Offset(0, 1).value
        emailAddress = findRow.Offset(0, 2).value
        
        ' E-posta adresi bo� de�ilse e-posta olu�tur
        If emailAddress <> "" Then
            ' E-posta olu�tur ve taslak olarak kaydet
            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                .To = emailAddress
                .Cc = "umit.yazir@mlpcare.com;ceyda.simsek@mlpcare.com"
                .Subject = "�la� �htiya� Fazlas� Talebi Hk."
                .Display ' E-postay� taslak olarak a�
                
                ' G�nderen e-posta adresini al
                senderEmail = .Session.Accounts.Item(1).SmtpAddress
                
                ' G�nderen e-posta adresini Org sayfas�nda bul ve hastane ad�n� al
                Set findRow = wsOrg.Columns("D").Find(What:=senderEmail, LookIn:=xlValues, LookAt:=xlWhole)
                If Not findRow Is Nothing Then
                    senderHospitalName = findRow.Offset(0, -2).value
                Else
                    senderHospitalName = "Bilinmiyor"
                End If
                
                ' Veri aral���n� HTML format�nda olu�tur
                Dim dataContent As String
                dataContent = "<table border='1' style='border-collapse:collapse;'>"
dataContent = dataContent & "<tr><td colspan='4' style='font-weight:bold; background-color:lightblue; text-align:center;'>" & hospitalName & "</td><td colspan='3' style='font-weight:bold; background-color:lightgreen; text-align:center;'>" & senderHospitalName & "</td><td colspan='2' style='font-weight:bold; background-color:lightblue; text-align:center;'>" & hospitalName & " taraf�ndan kar��lanacak miktarlar ve varsa A��klamalar</td></tr>"
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
    dataContent = dataContent & "<td style='word-wrap:break-word; width:1.6cm; font-weight:bold; text-align:center;background-color:" & IIf(cell.row Mod 2 = 0, "lightblue;", "white;") & "'>Kar�. Miktar (Kt)</td>" ' Kar�. Miktar (Kt) s�tunu
    dataContent = dataContent & "<td style='word-wrap:break-word; width:5cm; font-weight:bold; text-align:center;background-color:" & IIf(cell.row Mod 2 = 0, "lightblue;", "white;") & "'>A��klamalar</td>" ' A��klamalar s�tunu
Else
    dataContent = dataContent & "<td style='word-wrap:break-word; width:1.6cm; background-color:" & IIf(cell.row Mod 2 = 0, "lightgrey;", "white;") & "'></td>" ' Kar�. Miktar (Kt) s�tunu
    dataContent = dataContent & "<td style='word-wrap:break-word; width:5cm; background-color:" & IIf(cell.row Mod 2 = 0, "lightgrey;", "white;") & "'></td>" ' A��klamalar s�tunu
End If

    dataContent = dataContent & "</tr>"
Next cell
dataContent = dataContent & "</table>"


                
                ' E-posta i�eri�ini olu�tur
                emailBody = "<span style='font-size:12pt; font-family:Times New Roman;'>" & _
                            "Merhaba " & pharmacistName & "," & "<br><br>" & _
                            "A�a��daki tabloda sizin ihtiya� fazlan�z bizimse ihtiya� duydu�umuz ila�lar�n listesi ve ihtiya� miktarlar�m�z g�r�nmektedir." & "<br>" & _
                            "M�mk�nse ihtiyac�m�z kadar de�ilse sizin uygun g�rd���n�z miktarlarda yard�mc� olman�z� rica ediyoruz." & "<br><br>" & _
                            "Te�ekk�rler, iyi �al��malar." & "<br><br>" & _
                            dataContent & "<br><br>" & _
                            "* Bu mail Sat�n Alma �al��mas� Beta 5.1 taraf�ndan otomatik olarak olu�turulmu�tur. Yanl��l�k oldu�unu d���n�yorsan�z l�tfen Ecz. Harun Topal ile ileti�ime ge�iniz." & _
                            "</span>"
                
                .HTMLBody = emailBody & "<br><br>" & .HTMLBody ' Varsay�lan imzay� eklemek i�in mevcut HTMLBody'yi ekle
            End With
        End If
    End If
    
    ' Temizlik
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub



