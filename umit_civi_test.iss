Sub Main
	Call myDatabase()
End Sub

Function myDatabase

	Dim NewTable As Table
	Set NewTable = Client.NewTableDef
	
	Dim AddedField As Field
	Set myField = NewTable.NewField
	myField.Name = "ACIKLAMALAR"
	myField.Type = WI_CHAR_FIELD
	myField.Length = 200
	
	NewTable.AppendField myField
	
	Set myField = NewTable.NewField
	myField.Name = "CARI_DONEM"
	myField.Type = WI_NUM_FIELD
	NewTable.AppendField myField
	
	Set myField = NewTable.NewField
	myField.Name = "ONCEKI_DONEM"
	myField.Type = WI_NUM_FIELD
	NewTable.AppendField myField
	' Change the table settings to allow writing.
	NewTable.Protect = False
	
	' Create the database.
	Dim db As Database
	Set db = Client.NewDatabase("umit_civi_test.IMD", "", NewTable)
	
	' Obtain the recordset.
	Dim rs As RecordSet
	Set rs = db.RecordSet
	
	' Obtain a new record.
	Dim rec As Record
	Set rec = rs.NewRecord
	
	' Use the field name method to add data.
	rec.SetCharValue "ACIKLAMALAR", "Esas Faaliyetlerden Nakit Akislari [abstract]"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Dönem Kari (Zarari) (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Dönem Kari (Zarari) Mutabakati Ile Ilgili Düzeltmeler [abstract]"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Amortisman ve Itfa Gideriyle Ilgili Düzeltmeler"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Deger Düsüklügü (Iptali) Ile Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec

	rec.SetCharValue "ACIKLAMALAR", "Karsiliklarla Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Faiz Gelirleri ve Giderleriyle Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Gerçeklesmemis Kur Farklariyla Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Gerçege Uygun Deger Kayiplari (Kazançlari) Ile Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Stoklardaki Azalislar (Artislar) Ile Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Ticari Alacaklardaki Azalislar (Artislar) Ile Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Faaliyetler Ile Ilgili Diger Alacaklardaki Azalislar (Artislar) Ile Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Ticari Borçlardaki Artislar (Azalislar) Ile Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Faaliyetler Ile Ilgili Diger Borçlardaki Artislar (Azalislar) Ile Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Ertelenmis Gelirlerdeki Artislar (Azalislar) Ile Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Nakit Disi Kalemlere Iliskin Diger Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec

	rec.SetCharValue "ACIKLAMALAR", "Duran Varliklarin Elden Çikarilmasindan Kayiplar (Kazançlar) Ile Ilgili Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec

	rec.SetCharValue "ACIKLAMALAR", "Yatirim ya da Finansman Faaliyetlerinden Nakit Akislarina Neden Olan Diger Kalemlere Iliskin Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec

	rec.SetCharValue "ACIKLAMALAR", "Dönem Kari (Zarari) Mutabakati Ile Ilgili Diger Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec

	rec.SetCharValue "ACIKLAMALAR", "Dönem Kari (Zarari) Mutabakati Ile Ilgili Toplam Düzeltmeler (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec

	rec.SetCharValue "ACIKLAMALAR", "Faaliyetlerden Kaynaklanan Net Nakit Akisi (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Ödenen Kar Paylari (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Alinan Kar Paylari"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Ödenen Faiz(-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Alinan Faiz"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Vergi Iadeleri (Ödemeleri) (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec

	rec.SetCharValue "ACIKLAMALAR", "Diger Nakit Girisleri (Çikislari) (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Esas Faaliyetlerden Net Nakit Akisi (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Yatirim Faaliyetlerinden Nakit Akislari [abstract]"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Bagli Ortakliklardaki Paylarin Kontrol Kaybina Neden Olacak Sekilde Elden Çikarilmasindan Nakit Girisleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Bagli Ortaklik Ediniminden Nakit Çikislari (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Istirak ve Müsterek Girisimlerdeki Paylarin Elden Çikarilmasindan Nakit Girisleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Istirak ve Müsterek Girisim Paylarinin Ediniminden Nakit Çikislari (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Baska Isletme veya Fon Paylarinin veya Borçlanma Araçlarinin Elden Çikarilmasindan Nakit Girisleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Baska Isletme veya Fon Paylarinin veya Borçlanma Araçlarinin Ediniminden Nakit Çikislari (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Maddi ve Maddi Olmayan Duran Varliklarin Satisindan Nakit Girisleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Maddi ve Maddi Olmayan Duran Varlik Alimindan Nakit Çikislari (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Diger Uzun Vadeli Varliklarin Satisindan Nakit Girisleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Diger Uzun Vadeli Varlik Alimlarindan Nakit Çikislari (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Verilen Nakit Avans ve Borçlar (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Verilen Nakit Avans ve Borçlardan Geri Ödemeleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Türev Araçlardan Nakit Girisleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Türev Araçlardan Nakit Çikislari (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Devlet Tesviklerinden Nakit Girisleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Alinan Kar Paylari"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Ödenen Faiz(-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Alinan Faiz"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Vergi Iadeleri (Ödemeleri) (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Diger Nakit Girisleri (Çikislari) (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Yatirim Faaliyetlerinden Net Nakit Akisi (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Finansman Faaliyetlerinden Nakit Akislari [abstract]"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Bagli Ortakliklardaki Paylarin Kontrol Kaybina Neden Olmayacak Sekilde Elden Çikarilmasindan Nakit Girisleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Bagli Ortakliklarin Ilave Paylarinin Ediniminden Nakit Çikislari (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Özkaynak Araçlarinin Ihracindan veya Sermaye Artirimindan Nakit Girisleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Isletmenin Kendi Paylarini ve Diger Özkaynak Araçlarini Almasiyla veya Sermayenin Azaltilmasiyla Ilgili Nakit Çikislar (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Borçlanmadan Kaynaklanan Nakit Girisleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Borç Ödemelerinden Nakit Çikislari (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Finansal Kiralama Borçlarindan Nakit Çikislari (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Devlet Tesviklerinden Nakit Girisleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Ödenen Kar Paylari (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Ödenen Faiz (-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Vergi Iadeleri (Ödemeleri) (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Diger Nakit Girisleri (Çikislari) (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Finansman Faaliyetlerinden Net Nakit Akisi (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Kur Farklarinin Etkisinden Önce Nakit ve Nakit Benzerlerindeki Safi Artis (Azalis) (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Kur Farklarinin Nakit ve Nakit Benzerleri Üzerindeki Etkisi (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Nakit ve Nakit Benzerlerindeki Safi Artis (Azalis) (+/-)"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	rec.SetCharValue "ACIKLAMALAR", "Dönem Basi Nakit ve Nakit Benzerleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec	
	
	rec.SetCharValue "ACIKLAMALAR", "Dönem Sonu Nakit ve Nakit Benzerleri"
	rec.SetNumValue "CARI_DONEM", 0
	rec.SetNumValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec	
	' Protect the table before you commit it.
	NewTable.Protect = True
	
	' Commit the database.
	db.CommitDatabase
	' Open the database.
	Client.OpenDatabase "umit_civi_test.IMD"
	' Clear the memory.
	
	Set db = Nothing
	Set AddedField = Nothing
	Set NewTable = Nothing
	
End Function
	