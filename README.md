# BLUE CAB Panel v1.8

Drive ilçe/mahalle/kurye MASTER kopyasıyla çalışan panel güncellemesi.

**Bu ZIP, günlük Excel listesi alanına yüklenmez.** Uygulama kaynaklarının
ve ek veri dosyalarının birlikte güncellenmesi gerekir.

Kurulum, sınırlar ve kabul kontrolleri: **KURULUM_OKU.txt**.
Test sonuçları: **TEST_RAPORU.json**.

## Kaynak ve öncelik
Normal atama: data/MASTER.xlsx içindeki İLÇE + MAHALLE / Kurye / Y.NO / Durum.
Kaynak: 08 - İstanbul İlçe Mahalle Kod MASTER, Drive'dan 15.09.2026'da alındı.
39 ilçe, 974 mahalle kaydı; 712 AKTİF, 262 KAPALI.

Konum ve paket doğrulama -> açık KAPALI kararı -> Ahmet Aksu önceliği ->
gerçek 09:00-11:00 -> özel adres -> MASTER ilçe-mahalle eşleşmesi.
Özel kurallar data/special_rules.json içindedir; Drive kural metni otomatik
koda çevrilmez. Kural belgeleri sources/ içinde referans olarak tutulur.
Çelişkili/eski metinlerin normal atamaları, yeni MASTER tablosunu ezmez.

MASTER yükleme ve Drive'dan yenileme, doğrulanmış değişmez sürümleri
saklar; etkin sürüm işaretçisi atomik olarak değişir. Dağıtımda tek
worker gerekir. Kaynak Drive dosyasının paylaşım ayarları değiştirilmez.


## v1.8 adres çözümleme
MAHALLE alanı boşsa panel önce açık adresteki mahalle adını, sonra yalnızca
onaylı ve ilçe-kapsamlı adres ipuçlarını dener. Bu ipuçları **kurye seçmez**;
yalnızca mahalleyi bulur. Kurye ve Y.NO her zaman etkin MASTER'dan gelir.
21.09.2026 FIT_CHEF örneğinde önceki 46 ATANMADI kayıt 46/46 çözüldü.

## Hızlı başlatma
```sh
pip install -r requirements.txt
python -m unittest discover -s tests -v
python app.py
```
Yerel geliştirme adresi: http://127.0.0.1:8080 .
Google OAuth için ayrı geçerli HTTPS callback ve ortam ayarları gerekir.
Üretim: `gunicorn app:app --workers 1 --threads 4 --timeout 180`.
.env.example açıklama dosyasıdır; uygulama .env dosyasını otomatik okumaz.
Ayarları gerçek işlem/sunucu ortamında tanımlayın.

## Güvenlik ve test kapsamı
MASTER değişiklikleri yönetici anahtarı ve CSRF kontrolü gerektirir.
Anahtarlar/Google tokenları pakete dahil değildir. Ayrı kurum erişim denetimi
kullanın; bu paket personel hesabı/rol sistemi değildir.
Yerel atama/Excel testleri gerçek MASTER + sentetik müşteri kayıtları kullanır.
Canlı HTTP/OAuth, Gmail, Sheets ve tarayıcı testleri bu ortamda yapılmadı.
