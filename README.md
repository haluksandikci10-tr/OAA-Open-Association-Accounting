# OAA v2
## Açık Kaynak Dernek Muhasebe ve Yönetim Sistemi

OAA (Open Association Accounting), dernek ve kulüplerin üye yönetimi ile işletme hesabı esaslı mali işlemlerini Microsoft Excel VBA üzerinde yürütebilmesi için geliştirilen açık kaynaklı bir uygulamadır.

## Platform ve uyumluluk

- Microsoft Excel 2010 ve üzeri
- Excel Macro-Enabled Workbook (`.xlsm`)
- VBA
- Git sürüm kontrolü

> Excel 365 zorunlu değildir. Yeni Excel sürümlerinde de çalışması hedeflenir; uyumluluk doğrulaması Excel 2010+ esas alınarak yapılır.

## Mimari

Uygulama mevcut katmanlı mimariyi korur:

```
UserForms
   |
Service
   |
Repository
   |
Excel ListObject tabloları
```

- UserForm'lar iş kurallarını doğrudan uygulamaz.
- Service katmanı iş kurallarını ve işlem bütünlüğünü yönetir.
- Repository katmanı Excel tablolarına veri erişimini sağlar.
- Mevcut çalışan mimari, gerekçeli inceleme ve onay olmadan değiştirilmez.

## Ana modüller

- Dernek ve sistem ayarları
- Üye yönetimi
- Üye borçları ve aidat tarifeleri
- Gelir ve gider
- Makbuz ve tahsilat
- İşletme defteri

Diğer yol haritası özellikleri tamamlanma durumlarına göre ayrıca duyurulur; README'de listelenmeleri tamamlandıkları anlamına gelmez.

## Kaynak kod düzeni

- `src/classes` — VBA sınıfları
- `src/forms` — UserForm kaynakları (`.frm` ve eşleşen `.frx`)
- `src/modules` — Standart modüller
- `tests` — Test kaynaklarının ayrıştırılması için ayrılmış alan
- `docs` — Proje belgeleri
- `application` — Geliştirme çalışma kitapları ve yedekler

`src/modules/modTests.bas` şu an testlerin ana kaynağıdır. Ayrı `tests` klasörüne taşınmış bağımsız test modülleri varmış gibi değerlendirilmemelidir.

## Test ve kalite

Test prosedürü:
1. Kaynak ve gerçek metot/alan adlarını doğrula.
2. VBA Editor'da **Debug > Compile VBAProject** çalıştır.
3. Compile başarılıysa testleri sırayla çalıştır.
4. Başarısız test düzeltilmeden sonraki teste geçme.
5. Test verilerini test sonunda temizle.

Son doğrulanan test: **TEST-144 — başarılı**.  
Sıradaki test: **TEST-145**.

Test durumu, kullanıcı tarafından doğrulanan çalıştırma sonuçlarına göre güncellenir. Kaynakta test prosedürünün bulunması tek başına testin geçtiği anlamına gelmez.

## Dokümantasyon

Belge envanteri ve eksik dosyalar için [docs/README.md](docs/README.md) sayfasına bakın. Onaylı PDF belgeleri `docs` altında sürümleri korunarak saklanmalıdır.

## Geliştirme durumu

OAA v2 aktif geliştirme aşamasındadır. Gelir, gider, işletme defteri ve üye işlemleri için testler yürütülmektedir. Proje henüz kararlı sürüm veya üretim kullanımı için tamamlanmış olarak sunulmamaktadır.

## Açık kaynak ve test verileri

Geliştirme çalışma kitaplarında test verileri bulunabilir. Yayın/merge öncesinde test verilerinin temiz başlangıç çalışma kitabından ayrılması ve kişisel/gerçek veri içermediğinin kontrol edilmesi gerekir.

## Katkı

Katkı kuralları için [CONTRIBUTING.md](CONTRIBUTING.md) dosyasına bakın. Mimari değişiklikler öncesinde gerekçe ve etki analizi sunulmalıdır.
