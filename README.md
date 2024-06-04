.Net FrameWork

Bir web sayfası oluşturuldu, bu sayfa üzerinden kullanıcı siteye excel dosyası yükleyebiliyor.
Yapılan proje, bu excel dosyasında bulunan kolonlarla birebir aynı olacak şekilde mssql veritabanında otomatik olarak tablo
 oluşturuyor ve exceldeki datalar tabloya otomatik olarak insert oluyor.
Sitede bu veriler tablo halinde listeleniyor, veriler üzerinde düzenleme ve silme işlemi yapılıyor.
Daha sonra da bu veriler Export Excel diyerek yeni bir excel olarak indiriliyor.
İndirilen excel sisteme tekrar yüklendiğinde, tablo ve kolonlar daha önceden zaten oluşmuş olacağı için tablonun içeriği otomatik olarak 
silinip excel tabloya otomatik olarak yeniden insert oluyor.
Yeni yüklenen exceldeki veriler sitede yine listelenecek, veriler üzerinde düzenleme ve silme işlemi yapılabiliyor.
