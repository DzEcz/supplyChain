Attribute VB_Name = "Module1"
Sub AnaProsedur()
        ' Optimizasyonları kapat
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo HataYakalama ' Hata yakalama
    
    Dim currentSheet As Worksheet
    
    ' Mevcut aktif sayfayı belirle
    Set currentSheet = ActiveSheet
    
    ' UserForm'u göster
    UserForm1.Show vbModeless
    UserForm1.Caption = "İlerleme Durumu"
    DoEvents ' UserForm'un güncellenmesini sağlar
    
    ' Tüm düğmeleri pasif yap
    UserForm1.CommandButton1.Enabled = False
    UserForm1.CommandButton2.Enabled = False
    UserForm1.CommandButton3.Enabled = False
    
    ' Hesap sayfasının kilidini aç
    Sheets("Hesap").Unprotect Password:="8142" ' Şifreyi kendi belirlediğiniz şifre ile değiştirin
    Sheets("Pusula").Unprotect Password:="8142" ' Şifreyi kendi belirlediğiniz şifre ile değiştirin
    
    ' İşlemleri gerçekleştir
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
    
    ' İşlemler tamamlandığında bildirim ekle
    UserForm1.ListBox.AddItem "Tüm işlemler başarıyla gerçekleşti."
    
    ' Hesap sayfasını tekrar kilitle
    Sheets("Hesap").Protect Password:="8142" ' Şifreyi kendi belirlediğiniz şifre ile değiştirin
    Sheets("Pusula").Protect Password:="8142" ' Şifreyi kendi belirlediğiniz şifre ile değiştirin
    
    ' Başlatılan sayfaya geri dön
    currentSheet.Activate
    
    ' Kapatma butonunu aktif yap
    UserForm1.CommandButton1.Enabled = True
    UserForm1.CommandButton3.Enabled = True

    Exit Sub

HataYakalama:
    ' Hata durumunda UserForm'u gizle ve hata mesajını göster
    MsgBox "Bir hata oluştu: " & Err.Description & vbCrLf & _
           "Prosedür: " & Err.Source & vbCrLf & _
           "Satır: " & Erl, vbCritical
           
    ' Optimizasyonları aç
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


Sub Adshow()
    Application.StatusBar = "Ecz. Harun Topal"
End Sub

Sub PusulaSayfasiniGuncelle()

UserForm1.ListBox.AddItem "Pusula sayfası güncelleme işlemi başladı."
    Dim kaynakKitap As Workbook
    Dim hedefKitap As Workbook
    Dim kaynakSayfa As Worksheet
    Dim hedefSayfa As Worksheet
    Dim kaynakDosyaYolu As String
    
    ' Kaynak dosya yolunu belirleyin
    kaynakDosyaYolu = ThisWorkbook.Path & "\Pusula.xlsx"
    
    ' Kaynak çalışma kitabını açın
    Set kaynakKitap = Workbooks.Open(kaynakDosyaYolu)
    Set kaynakSayfa = kaynakKitap.Sheets("Sheet")
    
    ' Hedef çalışma kitabını ve sayfasını belirleyin
    Set hedefKitap = ThisWorkbook
    Set hedefSayfa = hedefKitap.Sheets("Pusula")
    
    ' Hedef sayfadaki mevcut verileri temizleyin
    hedefSayfa.Cells.Clear
    
    ' Kaynak sayfadaki verileri kopyalayın
    kaynakSayfa.UsedRange.Copy
    
    ' Verileri hedef sayfaya yapıştırın
    hedefSayfa.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    ' Kaynak çalışma kitabını kapatın
    kaynakKitap.Close False
    
    ' Kullanıcıya bildirimde bulunun
UserForm1.ListBox.AddItem "Pusula sayfasıgüncelleme işlemi tamamlandı."
End Sub

Sub VeriKopyala()

UserForm1.ListBox.AddItem "Pusula sayfasından veri kopyalama işlemi başladı."
   
    Dim wsPusula As Worksheet
    Dim wsHesap As Worksheet
    Dim lastRow As Long
    Dim kodCol As Long
    Dim adCol As Long
    Dim miktarCol As Long
    Dim kodData As Variant
    Dim i As Long
    
    ' Çalışma sayfalarını tanımla
    Set wsPusula = ThisWorkbook.Sheets("Pusula")
    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    
    ' Pusula sayfasındaki son satırı bul
    lastRow = wsPusula.Cells(wsPusula.Rows.count, "A").End(xlUp).row
    
    ' Pusula sayfasında veri olup olmadığını kontrol et
    If lastRow < 2 Then
        MsgBox "Lütfen Pusuladan çektiğiniz stok durum raporunu aynı klasöre kopyalayınız!", vbExclamation
        wsPusula.Activate
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Hesap sayfasındaki verileri kontrol et ve gerekirse sil
    If wsHesap.Cells(2, 1).value <> "" Then
        wsHesap.Rows("2:" & wsHesap.Rows.count).ClearContents
    End If
    
    ' Sütun numaralarını bul
    kodCol = wsPusula.Rows(1).Find("C. EMR Eşdeğer Ürün Grup Kodu").Column
    adCol = wsPusula.Rows(1).Find("Adı").Column
    miktarCol = wsPusula.Rows(1).Find("Miktar").Column
    
    ' Pusula sayfasındaki kod verilerini diziye al
    kodData = wsPusula.Range(wsPusula.Cells(2, kodCol), wsPusula.Cells(lastRow, kodCol)).value
    
    ' Kod verilerini sayıya dönüştür ve ondalık olmamasını sağla
    For i = 1 To UBound(kodData, 1)
        If IsNumeric(kodData(i, 1)) Then
            kodData(i, 1) = Round(CDbl(kodData(i, 1)), 0)
        End If
    Next i
    
    ' Hesap sayfasındaki başlıkları yaz
    wsHesap.Cells(1, 1).value = "EşdeğerKod"
    wsHesap.Cells(1, 2).value = "Müstahzar"
    wsHesap.Cells(1, 3).value = "Stok Miktar"
    
    ' Pusula sayfasındaki verileri Hesap sayfasına kopyala
    wsHesap.Range("A2:A" & lastRow).value = kodData
    wsHesap.Range("B2:B" & lastRow).value = wsPusula.Range(wsPusula.Cells(2, adCol), wsPusula.Cells(lastRow, adCol)).value
    wsHesap.Range("C2:C" & lastRow).value = wsPusula.Range(wsPusula.Cells(2, miktarCol), wsPusula.Cells(lastRow, miktarCol)).value
    
  
UserForm1.ListBox.AddItem "Pusula sayfasından veri kopyalama işlemi tamamlandı."
End Sub

'eşdeğerkodları üçe tamamla;

Sub KopyalaVeEkleHizli()

UserForm1.ListBox.AddItem "Müstahzar sayısının üçlemesi işlemi başladı."
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
    
    Application.ScreenUpdating = False ' Ekran güncellemelerini kapat
    Application.Calculation = xlCalculationManual ' Otomatik hesaplamayı kapat
    
    Set ws = ThisWorkbook.Sheets("Hesap") ' Çalışma sayfasını tanımla
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row ' Son satırı bul
    
    ' Sütun başlıklarını bul
    esdegerKodCol = Application.WorksheetFunction.Match("EşdeğerKod", ws.Rows(1), 0)
    mustahzarCol = Application.WorksheetFunction.Match("Müstahzar", ws.Rows(1), 0)
    stokMiktarCol = Application.WorksheetFunction.Match("Stok Miktar", ws.Rows(1), 0)
    
    ' Verileri diziye al
    data = ws.Range(ws.Cells(2, esdegerKodCol), ws.Cells(lastRow, stokMiktarCol)).value
    
    ' Sonuç dizisini başlat
    ReDim result(1 To (lastRow - 1) * 2, 1 To UBound(data, 2))
    resultIndex = 1
    
    ' Eşdeğer Kodları say
    For i = 1 To UBound(data, 1)
        kod = data(i, 1) ' Eşdeğer Kod sütunu
        If kodCount.exists(kod) Then
            kodCount(kod) = kodCount(kod) + 1
        Else
            kodCount.Add kod, 1
        End If
    Next i
    
    ' İki adet olan Eşdeğer Kodları kopyala
    For i = 1 To UBound(data, 1)
        kod = data(i, 1)
        If kodCount(kod) = 2 Then
            result(resultIndex, 1) = data(i, 1) ' EşdeğerKod
            result(resultIndex, 2) = data(i, 2) & "_kopya" & kodCount(kod) ' Müstahzar
            result(resultIndex, 3) = data(i, 3) ' Stok Miktar
            resultIndex = resultIndex + 1
            kodCount(kod) = kodCount(kod) + 1
        End If
    Next i
    
    ' Bir adet olan Eşdeğer Kodları kopyala
    For i = 1 To UBound(data, 1)
        kod = data(i, 1)
        If kodCount(kod) = 1 Then
            ' İki kopya ekle
            For j = 1 To 2
                If kodCount(kod) < 3 Then
                    result(resultIndex, 1) = data(i, 1) ' EşdeğerKod
                    result(resultIndex, 2) = data(i, 2) & "_kopya" & kodCount(kod) ' Müstahzar
                    result(resultIndex, 3) = data(i, 3) ' Stok Miktar
                    resultIndex = resultIndex + 1
                    kodCount(kod) = kodCount(kod) + 1
                End If
            Next j
        End If
    Next i
    
    ' Sonuçları çalışma sayfasına yaz
    ws.Range(ws.Cells(lastRow + 1, esdegerKodCol), ws.Cells(lastRow + resultIndex - 1, stokMiktarCol)).value = result
    
    ' EşdeğerKod verilerini alfabetik olarak sıralama
    ws.Range(ws.Cells(2, esdegerKodCol), ws.Cells(lastRow + resultIndex - 1, stokMiktarCol)).Sort Key1:=ws.Cells(2, esdegerKodCol), Order1:=xlAscending, header:=xlNo
    
    Application.ScreenUpdating = True ' Ekran güncellemelerini aç
    Application.Calculation = xlCalculationAutomatic ' Otomatik hesaplamayı aç

UserForm1.ListBox.AddItem "Müstahzar sayısının üçlemesi işlemi tamamlandı."
End Sub

Sub KutuMiktarKopyala()

UserForm1.ListBox.AddItem "Kutu içi miktarlarının kpyalanması işlemi başladı."
    Dim wsHesap As Worksheet
    Dim wsKutuiçi As Worksheet
    Dim rngHesap As Range
    Dim rngKutuiçi As Range
    Dim cell As Range
    Dim matchRow As Variant
    Dim colEsdegerKodHesap As Long
    Dim colKutuMiktarHesap As Long
    Dim colEsdegerKodKutuiçi As Long
    Dim colKutuIciKutuiçi As Long
    Dim hesapData As Variant
    Dim kutuiciData As Variant
    Dim i As Long
    Dim dict As Object
    
    ' Sayfaları tanımla
    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    Set wsKutuiçi = ThisWorkbook.Sheets("Kutuiçi")
    
    ' Sütun başlıklarının yerini bul
    colEsdegerKodHesap = Application.Match("EşdeğerKod", wsHesap.Rows(1), 0)
    colKutuMiktarHesap = Application.Match("Kutu Miktar", wsHesap.Rows(1), 0)
    colEsdegerKodKutuiçi = Application.Match("Eşdeğer", wsKutuiçi.Rows(1), 0)
    colKutuIciKutuiçi = Application.Match("Kutu İçi", wsKutuiçi.Rows(1), 0)
    
    ' Verileri diziye al
    hesapData = wsHesap.Range(wsHesap.Cells(2, colEsdegerKodHesap), wsHesap.Cells(wsHesap.Rows.count, colEsdegerKodHesap).End(xlUp)).Resize(, colKutuMiktarHesap - colEsdegerKodHesap + 1).value
    kutuiciData = wsKutuiçi.Range(wsKutuiçi.Cells(2, colEsdegerKodKutuiçi), wsKutuiçi.Cells(wsKutuiçi.Rows.count, colEsdegerKodKutuiçi).End(xlUp)).Resize(, colKutuIciKutuiçi - colEsdegerKodKutuiçi + 1).value
    
    ' Eşdeğer kodları ve kutu içi miktarlarını bir sözlükte sakla
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(kutuiciData, 1)
        dict(kutuiciData(i, 1)) = kutuiciData(i, 2)
    Next i
    
    ' Ekran güncellemelerini ve hesaplamaları kapat
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Hesap sayfasındaki her bir EşdeğerKod için
    For i = 1 To UBound(hesapData, 1)
        If dict.exists(hesapData(i, 1)) Then
            hesapData(i, colKutuMiktarHesap - colEsdegerKodHesap + 1) = dict(hesapData(i, 1))
        Else
            hesapData(i, colKutuMiktarHesap - colEsdegerKodHesap + 1) = 1
        End If
    Next i
    
    ' Sonuçları çalışma sayfasına yaz
    wsHesap.Range(wsHesap.Cells(2, colEsdegerKodHesap), wsHesap.Cells(UBound(hesapData, 1) + 1, colKutuMiktarHesap)).value = hesapData
    
    ' Ekran güncellemelerini ve hesaplamaları aç
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

UserForm1.ListBox.AddItem "Kutu içi miktarlarının kopyalanması işlemi tamamlandı."
End Sub
Sub EsdegerToplam()

UserForm1.ListBox.AddItem "Stok hesaplama işlemleri başladı."
    
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
    
    ' Sayfaları tanımla
    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    Set wsPusula = ThisWorkbook.Sheets("Pusula")
    
    ' Hesap sayfasındaki sütunları bul
    Set hesesdegerkodverisi = wsHesap.Rows(1).Find("EşdeğerKod")
    Set heskutumiktarverisi = wsHesap.Rows(1).Find("Kutu Miktar")
    Set hesesdmiktoplam = wsHesap.Rows(1).Find("Eşd.Mik. TOPLAM")
    Set heskrimiktoplam = wsHesap.Rows(1).Find("Kri.Mik. TOPLAM")
    Set hesmaxmiktartoplam = wsHesap.Rows(1).Find("Max.Mik TOPLAM")
    Set hesgopithmik = wsHesap.Rows(1).Find("İht. Mik.")
    
    ' Pusula sayfasındaki sütunları bul
    Set pusesdegerkodverisi = wsPusula.Rows(1).Find("C. EMR Eşdeğer Ürün Grup Kodu")
    Set pusmikverisi = wsPusula.Rows(1).Find("Miktar")
    Set puskrimikverisi = wsPusula.Rows(1).Find("Kritik Miktar")
    Set pusmaxmikverisi = wsPusula.Rows(1).Find("Max Miktar")
    
    ' Hesap sayfasındaki her bir EşdeğerKod icin işlemleri yap
    For Each cell In wsHesap.Range(hesesdegerkodverisi.Offset(1, 0), wsHesap.Cells(wsHesap.Rows.count, hesesdegerkodverisi.Column).End(xlUp))
        kod = Trim(UCase(cell.value))
        toplam = 0
        krimiktoplam = 0
        maxmiktartoplam = 0
        
        ' Pusula sayfasında eşleşen kodları bul ve miktarları topla
        For Each pCell In wsPusula.Range(pusesdegerkodverisi.Offset(1, 0), wsPusula.Cells(wsPusula.Rows.count, pusesdegerkodverisi.Column).End(xlUp))
            If Trim(UCase(pCell.value)) = kod Then
                toplam = toplam + CDbl(pCell.Offset(0, pusmikverisi.Column - pusesdegerkodverisi.Column).value)
                krimiktoplam = krimiktoplam + CDbl(pCell.Offset(0, puskrimikverisi.Column - pusesdegerkodverisi.Column).value)
                maxmiktartoplam = maxmiktartoplam + CDbl(pCell.Offset(0, pusmaxmikverisi.Column - pusesdegerkodverisi.Column).value)
            End If
        Next pCell
        
        ' Toplamı Kutu Miktar'a böl ve sonucu ilgili sütunlara yaz
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
        
        ' İht. Mik. sütununu hesapla
        If cell.Offset(0, hesesdmiktoplam.Column - hesesdegerkodverisi.Column).value < cell.Offset(0, heskrimiktoplam.Column - hesesdegerkodverisi.Column).value Then
            cell.Offset(0, hesgopithmik.Column - hesesdegerkodverisi.Column).value = Round(cell.Offset(0, hesmaxmiktartoplam.Column - hesesdegerkodverisi.Column).value - cell.Offset(0, hesesdmiktoplam.Column - hesesdegerkodverisi.Column).value, 0)
        Else
            cell.Offset(0, hesgopithmik.Column - hesesdegerkodverisi.Column).value = "Pass"
        End If
    Next cell
    
UserForm1.ListBox.AddItem "Stok hesaplama işlemleri tamamlandı."
End Sub
'Data sayfası ihtyiaç miktarları sıralaması, istediğim gibi değil ama sanırım iş görür
Sub DinamikSirala()

UserForm1.ListBox.AddItem "İhtiyaç fazlası sıralama işlemleri başladı."
    Dim ws As Worksheet
    Dim esdegerCol As Long
    Dim ihtiyacCol As Long
    Dim lastRow As Long
    Dim headerRow As Long
    Dim cell As Range

    ' Çalışma sayfasını belirle
    Set ws = ThisWorkbook.Sheets("Data") ' Sayfa adını ihtiyacınıza göre değiştirin

    ' Başlık satırını belirle
    headerRow = 1 ' Başlık satırının numarasını ihtiyacınıza göre değiştirin

    ' "Eşdeğer" ve "İhtiyaç" sütunlarını bul
    For Each cell In ws.Rows(headerRow).Cells
        If cell.value = "Eşdeğer" Then
            esdegerCol = cell.Column
        ElseIf cell.value = "İhtiyaç" Then
            ihtiyacCol = cell.Column
        End If
    Next cell

    ' Son satırı bul
    lastRow = ws.Cells(ws.Rows.count, esdegerCol).End(xlUp).row

    ' Sıralama işlemi
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Cells(headerRow + 1, ihtiyacCol), Order:=xlAscending
    ws.Sort.SortFields.Add key:=ws.Cells(headerRow + 1, esdegerCol), Order:=xlAscending
    With ws.Sort
        .SetRange ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column))
        .header = xlYes
        .Apply
    End With

UserForm1.ListBox.AddItem "İhtiyaç fazlası sıralama işlemleri tamamlandı."
End Sub

'ihtyiaç fazlası hastaneleri kopyalama
Sub KopyalaHastaneleri()

UserForm1.ListBox.AddItem "İhtiyaç fazlası bulunan hastane tespiti işlemleri başladı."
    ' Optimizasyonları kapat
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
    
    ' Hesap sayfasındaki sütun başlıklarını bul
    esdegerKodCol = Application.WorksheetFunction.Match("EşdeğerKod", wsHesap.Rows(1), 0)
    gopIhtMikCol = Application.WorksheetFunction.Match("İht. Mik.", wsHesap.Rows(1), 0)
    ihtFazHastAdCol = Application.WorksheetFunction.Match("İht. Faz. Hast AD", wsHesap.Rows(1), 0)
    ihtFazMiktarCol = Application.WorksheetFunction.Match("İht. Faz. Miktar", wsHesap.Rows(1), 0)
    
    ' Data sayfasındaki sütun başlıklarını bul
    hastaneAdiCol = Application.WorksheetFunction.Match("Hastane", wsData.Rows(1), 0)
    esdegerCol = Application.WorksheetFunction.Match("Eşdeğer", wsData.Rows(1), 0)
    ihtiyacCol = Application.WorksheetFunction.Match("İhtiyaç", wsData.Rows(1), 0)
    
    lastRow = wsData.Cells(wsData.Rows.count, esdegerCol).End(xlUp).row
    
    ' Data sayfasındaki her bir EşdeğerKod icin İhtiyaç ve Hastane Adı bilgilerini topla
    For ihtiyacRow = 2 To lastRow
        esdegerKod = wsData.Cells(ihtiyacRow, esdegerCol).value
        If Not ihtiyacDict.exists(esdegerKod) Then
            Set ihtiyacDict(esdegerKod) = New Collection
        End If
        ihtiyacDict(esdegerKod).Add Array(wsData.Cells(ihtiyacRow, ihtiyacCol).value, wsData.Cells(ihtiyacRow, hastaneAdiCol).value)
    Next ihtiyacRow
    
    ' Hesap sayfasındaki her bir EşdeğerKod icin işlemleri yap
    For i = 2 To wsHesap.Cells(wsHesap.Rows.count, esdegerKodCol).End(xlUp).row
        If wsHesap.Cells(i, gopIhtMikCol).value <> "Pass" Then
            esdegerKod = wsHesap.Cells(i, esdegerKodCol).value
            If ihtiyacDict.exists(esdegerKod) Then
                Set ihtiyacList = ihtiyacDict(esdegerKod)
                ' İhtiyaç miktarlarına göre küçükten büyüğe sırala
                ihtiyacArray = CollectionToArray(ihtiyacList)
                Call QuickSort(ihtiyacArray, LBound(ihtiyacArray, 2), UBound(ihtiyacArray, 2))
                
                ' İlk üç hastane ve ihtiyaç miktarını alt alta kopyala
                For j = 1 To Application.Min(3, UBound(ihtiyacArray, 2))
                    wsHesap.Cells(i, ihtFazMiktarCol).value = Round(ihtiyacArray(1, j), 0)
                    wsHesap.Cells(i, ihtFazHastAdCol).value = ihtiyacArray(2, j)
                    i = i + 1
                Next j
                ' Diğer satırları boş bırak
                For k = j To 3
                    wsHesap.Cells(i, ihtFazMiktarCol).value = ""
                    wsHesap.Cells(i, ihtFazHastAdCol).value = ""
                    i = i + 1
                Next k
                ' Aynı EşdeğerKod icin kopyalamayı durdur
                Do While wsHesap.Cells(i, esdegerKodCol).value = esdegerKod
                    wsHesap.Cells(i, ihtFazMiktarCol).value = ""
                    wsHesap.Cells(i, ihtFazHastAdCol).value = ""
                    i = i + 1
                Loop
                i = i - 1
            End If
        End If
    Next i
    
    ' Optimizasyonları aç
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    
UserForm1.ListBox.AddItem "İhtiyaç fazlası bulunan hastane tespiti işlemleri tamamlandı."
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

UserForm1.ListBox.AddItem "Tedarikçi ecza deposu tespiti işlemleri başladı."
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
    
    ' Çalışma sayfalarını tanımla
    Set wsHesap = ThisWorkbook.Sheets("Hesap")
    Set wsAnlMuad = ThisWorkbook.Sheets("AnlMuad")
    
    ' Son satırları bul
    lastRowHesap = wsHesap.Cells(wsHesap.Rows.count, 1).End(xlUp).row
    lastRowAnlMuad = wsAnlMuad.Cells(wsAnlMuad.Rows.count, 1).End(xlUp).row
    
    ' Sütun başlıklarının yerlerini bul
    ihtMikCol = wsHesap.Rows(1).Find("İht. Mik.").Column
    esdegerKodCol = wsHesap.Rows(1).Find("EşdeğerKod").Column
    depoDurumuCol = wsHesap.Rows(1).Find("Depo Adı & Durumu").Column
    esdegerCol = wsAnlMuad.Rows(1).Find("Eşdeğer").Column
    tedarikciCol = wsAnlMuad.Rows(1).Find("Tedarikçi").Column
    aciklamaCol = wsAnlMuad.Rows(1).Find("Açıklama").Column
    
    ' Eşdeğer kodlarının sayısını takip etmek için Scripting.Dictionary kullan
    Set esdegerCount = CreateObject("Scripting.Dictionary")
    
    ' Hesap sayfasında döngü
    For i = 2 To lastRowHesap
        ihtMik = wsHesap.Cells(i, ihtMikCol).value
        If ihtMik <> "Pass" Then
            esdegerKod = wsHesap.Cells(i, esdegerKodCol).value
            ' Eşdeğer kodunun sayısını artır
            If Not esdegerCount.exists(esdegerKod) Then
                esdegerCount(esdegerKod) = 1
            Else
                esdegerCount(esdegerKod) = esdegerCount(esdegerKod) + 1
            End If
            
            ' AnlMuad sayfasında eşdeğer kodu ara
            Dim foundCount As Long
            foundCount = 0
            For j = 2 To lastRowAnlMuad
                If wsAnlMuad.Cells(j, esdegerCol).value = esdegerKod Then
                    foundCount = foundCount + 1
                    If foundCount = esdegerCount(esdegerKod) Then
                        tedarikci = wsAnlMuad.Cells(j, tedarikciCol).value
                        aciklama = wsAnlMuad.Cells(j, aciklamaCol).value
                        ' Tedarikçi ve Açıklama bilgilerini birleştir ve yaz
                        wsHesap.Cells(i, depoDurumuCol).value = aciklama & " - " & tedarikci
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
    
UserForm1.ListBox.AddItem "Tedarikçi ecza deposu tespiti işlemleri tamamlandı."
End Sub

Sub PivotTabloyuYenile()
    Dim wsPVT As Worksheet
    Dim wsDepo As Worksheet
    Dim ptHastane As PivotTable
    Dim ptDepo As PivotTable
    
    ' PVT sayfasındaki pivot tabloyu tanımlayın
    Set wsPVT = ThisWorkbook.Sheets("PVT")
    Set ptHastane = wsPVT.PivotTables("hastanepvt") ' Pivot tablo adını buraya yazın
    
    ' Yeni sayfadaki pivot tabloyu tanımlayın
    Set wsDepo = ThisWorkbook.Sheets("depo") ' Yeni sayfanızın adını buraya yazın
    Set ptDepo = wsDepo.PivotTables("depopvt") ' Yeni pivot tablo adını buraya yazın
    
    ' Pivot tabloları yenileyin
    UserForm1.ListBox.AddItem "Pivot tablo güncellemeleri başladı."
    ptHastane.RefreshTable
    ptDepo.RefreshTable
    UserForm1.ListBox.AddItem "Pivot tablo güncellemeleri tamamlandı."
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
    
    ' Outlook uygulamasını başlat
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Çalışma sayfasını belirle
    Set ws = ThisWorkbook.Sheets("PVT") ' Pivot tablonun bulunduğu sayfa adı
    Set wsOrg = ThisWorkbook.Sheets("Org") ' Org sayfası
    
    ' E sütunundaki son dolu satırı bul
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).row
    
    ' Dinamik veri aralığını belirle
    Set rng = ws.Range("C2:I" & lastRow)
    
' C sütununu geçici olarak görünür yap
ws.Columns("C").Hidden = False

' Hastane adını C3 hücresinden başlayarak tüm C sütununda ara
Set searchRange = ws.Range("C3:C" & ws.Cells(ws.Rows.count, "C").End(xlUp).row)
hospitalName = ws.Range("C3").value

' C sütununda farklı hastane adları olup olmadığını kontrol et
Dim cell As Range
For Each cell In searchRange
    If cell.value <> "" And cell.value <> hospitalName Then
        MsgBox "İhtiyaç fazlası ilaçları içeren hastaneler sütununda farklı hastane adları tespit edildi." & vbCrLf & "Lütfen her işlemde yalnızca bir hastane seçiniz.", vbExclamation
        ws.Columns("C").Hidden = True
        Exit Sub
    End If
Next cell

' C sütununu tekrar gizle
ws.Columns("C").Hidden = True
  
    ' Kısaltma sayfasında hastane adını bul
    Set findRow = wsOrg.Columns("B").Find(What:=hospitalName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not findRow Is Nothing Then
        ' Eczacının adını ve e-posta adresini al
        pharmacistName = findRow.Offset(0, 1).value
        emailAddress = findRow.Offset(0, 2).value
        
        ' E-posta adresi boş değilse e-posta oluştur
        If emailAddress <> "" Then
            ' E-posta oluştur ve taslak olarak kaydet
            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                .To = emailAddress
                .Cc = "umit.yazir@mlpcare.com;ceyda.simsek@mlpcare.com"
                .Subject = "İlaç İhtiyaç Fazlası Talebi Hk."
                .Display ' E-postayı taslak olarak aç
                
                ' Gönderen e-posta adresini al
                senderEmail = .Session.Accounts.Item(1).SmtpAddress
                
                ' Gönderen e-posta adresini Org sayfasında bul ve hastane adını al
                Set findRow = wsOrg.Columns("D").Find(What:=senderEmail, LookIn:=xlValues, LookAt:=xlWhole)
                If Not findRow Is Nothing Then
                    senderHospitalName = findRow.Offset(0, -2).value
                Else
                    senderHospitalName = "Bilinmiyor"
                End If
                
                ' Veri aralığını HTML formatında oluştur
                Dim dataContent As String
                dataContent = "<table border='1' style='border-collapse:collapse;'>"
dataContent = dataContent & "<tr><td colspan='4' style='font-weight:bold; background-color:lightblue; text-align:center;'>" & hospitalName & "</td><td colspan='3' style='font-weight:bold; background-color:lightgreen; text-align:center;'>" & senderHospitalName & "</td><td colspan='2' style='font-weight:bold; background-color:lightblue; text-align:center;'>" & hospitalName & " tarafından karşılanacak miktarlar ve varsa Açıklamalar</td></tr>"
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
    dataContent = dataContent & "<td style='word-wrap:break-word; width:1.6cm; font-weight:bold; text-align:center;background-color:" & IIf(cell.row Mod 2 = 0, "lightblue;", "white;") & "'>Karş. Miktar (Kt)</td>" ' Karş. Miktar (Kt) sütunu
    dataContent = dataContent & "<td style='word-wrap:break-word; width:5cm; font-weight:bold; text-align:center;background-color:" & IIf(cell.row Mod 2 = 0, "lightblue;", "white;") & "'>Açıklamalar</td>" ' Açıklamalar sütunu
Else
    dataContent = dataContent & "<td style='word-wrap:break-word; width:1.6cm; background-color:" & IIf(cell.row Mod 2 = 0, "lightgrey;", "white;") & "'></td>" ' Karş. Miktar (Kt) sütunu
    dataContent = dataContent & "<td style='word-wrap:break-word; width:5cm; background-color:" & IIf(cell.row Mod 2 = 0, "lightgrey;", "white;") & "'></td>" ' Açıklamalar sütunu
End If

    dataContent = dataContent & "</tr>"
Next cell
dataContent = dataContent & "</table>"


                
                ' E-posta içeriğini oluştur
                emailBody = "<span style='font-size:12pt; font-family:Times New Roman;'>" & _
                            "Merhaba " & pharmacistName & "," & "<br><br>" & _
                            "Aşağıdaki tabloda sizin ihtiyaç fazlanız bizimse ihtiyaç duyduğumuz ilaçların listesi ve ihtiyaç miktarlarımız görünmektedir." & "<br>" & _
                            "Mümkünse ihtiyacımız kadar değilse sizin uygun gördüğünüz miktarlarda yardımcı olmanızı rica ediyoruz." & "<br><br>" & _
                            "Teşekkürler, iyi çalışmalar." & "<br><br>" & _
                            dataContent & "<br><br>" & _
                            "* Bu mail Satın Alma Çalışması Beta 5.1 tarafından otomatik olarak oluşturulmuştur. Yanlışlık olduğunu düşünüyorsanız lütfen Ecz. Harun Topal ile iletişime geçiniz." & _
                            "</span>"
                
                .HTMLBody = emailBody & "<br><br>" & .HTMLBody ' Varsayılan imzayı eklemek için mevcut HTMLBody'yi ekle
            End With
        End If
    End If
    
    ' Temizlik
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub



