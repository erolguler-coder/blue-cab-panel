# v1.8 / 21.09.2026

- MAHALLE boş kayıtlarda açık adresten onaylı mahalle/bölge çözümleme eklendi.
- Adres ipuçları yalnızca mahalleyi belirler; kurye ve Y.NO her zaman etkin MASTER'dan gelir.
- Nispetiye/Nisbetiye yazım farkı alias olarak desteklendi.
- 21.09.2026 FIT_CHEF testinde daha önce KONTROL/ATANMADI kalan 46 kayıt 46/46 çözüldü.

# v1.7 / 15.09.2026

- Normal atamada gömülü eski tablo yerine dışarıdan doğrulanan MASTER kullanılır.
- Ayrı ilçe ve mahalle sütunları iki işlem yolunda da okunur.
- Aynı mahalle adının başka ilçesine sıçrama ve eski kurye ile boşluk
  doldurma kaldırıldı; açık çelişki kontrol durumuna alınır.
- Balmumcu normal ataması doğrudan MASTER satırından gelir. Kanyon AVM
  özel adres ve Ahmet Aksu öncelik kuralları ayrı tutulur.
- KAPALI bölgede sadece MASTER'ın açıkça tanımladığı normal kurye korunur.
- MASTER yönetim ekranı: Excel yükleme, isteğe bağlı Drive yenileme,
  SHA sürümü, tarih, satır/durum sayıları, geçersiz dosyayı reddetme.
- Günlük raporda atama nedeni, kaynak satır ve MASTER sürümü yer alır.
- Mükerrer müşteriler korunur; paket adedi ve teslimat adedi ayrı sayılır.
- Firma sıralaması, kurye başlığı, Y.NO sayı değeri, firma alt özeti korunur.
- Excel metinleri formüle dönüştürülmez; üretilen formüllerin önbelleği yazılır.
- Gmail şablonu pakete dahildir. Canlı Google bağlantısı ayrı kabul testi ister.
