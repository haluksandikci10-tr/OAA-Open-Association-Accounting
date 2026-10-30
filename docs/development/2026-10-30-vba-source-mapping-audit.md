# OAA VBA Kaynak Eşleştirme Denetimi

Tarih: 2026-10-30  
Karşılaştırma: yerel teslim paketi `OAA-v2(1).zip` ile GitHub `develop-v2` dalı  
İnceleme dalı: `source-sync-2026-10-30`

## Sonuç

Yerel kaynak paketi ile GitHub deposunun mevcut kaynak ağacı birebir eşleşmiyor. Bu nedenle bu denetim kapsamında çalışan VBA kaynakları GitHub'ın mevcut `src` klasörlerine otomatik taşınmamış, mevcut dosyaların üzerine yazılmamış ve `develop-v2` dalına merge yapılmamıştır.

## Yerel teslim paketinin envanteri

- 31 sınıf dosyası (`.cls`); bu sayıya `ApplicationContext.cls` ve `Sayfa1.cls` dahildir.
- 13 standart modül (`.bas`); `Module1.bas` ve `Module2.bas` ayrı ayrı korunmuştur.
- 5 UserForm kaynak dosyası (`.frm`) ve bunlara ait 5 ikili kaynak dosyası (`.frx`).
- Test klasörleri mevcut olmakla birlikte `tests/fixtures`, `tests/integration` ve `tests/unit` klasörlerinde bu pakette dosya yoktur.
- Yerel dışa aktarımda `modTests.bas` yaklaşık 55.374 satırdır.

## GitHub develop-v2 mevcut kaynak ağacı

- `src/Forms`: 2 form çifti (`frmDues`, `frmMembers`).
- `src/Modules`: 4 modül (`modGlobals`, `modHelpers`, `modInstaller`, `modMain`).
- `src/bas`: 5 modül (`modGlobals`, `modHelpers`, `modInstaller`, `modMain`, `modTests`).
- `src/cls`: 3 sınıf (`clsLogger`, `clsRepository`, `clsSettings`).
- `tests`: `modDueTests.bas` ve `modTests.bas`.
- Depoda ayrıca iki Excel çalışma kitabı ve başka geliştirme belgeleri vardır.

## Belirlenen çakışma ve farklılıklar

1. `src/Modules` ve `src/bas` içinde aynı isimli modüller vardır. `modHelpers` ve `modMain` için SHA aynı görünürken `modGlobals` ve `modInstaller` dosyalarının SHA değerleri farklıdır. Bu dosyalar otomatik birleştirilmemelidir.
2. GitHub `tests/modTests.bas` ve `src/bas/modTests.bas` aynı içerik/SHA ile görünür; yerel dışa aktarımın `modTests.bas` dosyası çok daha büyüktür. Test kaynağı eşitlemesi ayrıca yapılmalıdır.
3. Yerel pakette `clsLogger` ve `clsSettings` adları GitHub'da da vardır; ancak yalnızca ada bakarak içerik veya uyumluluk eşitliği varsayılamaz.
4. Yerel UserForm adları (`frmAssociationWizard`, `frmIncome`, `frmMemberCard`, `frmMemberSearch`, `frmReceipts`) GitHub'daki iki formdan farklıdır. Formların işlevsel eşdeğerliği doğrulanmadan değiştirilmemelidir.
5. Yerel `Module1.bas` ve `Module2.bas` ayrı dosyalardır; içerikleri aynı görünse bile kaynak kimliği ve çağrı ilişkileri incelenmeden silinmemelidir.
6. `Sayfa1.cls` ve sayfa dışa aktarım metinleri Excel çalışma kitabı bileşenleriyle ilişkilidir; bunlar standart sınıf gibi körlemesine taşınmamalıdır.

## Güvenli sonraki adımlar

1. Yerel kaynakların Excel VBA projesinden alınmış son ve doğru dışa aktarım olduğu kullanıcı ortamında doğrulanmalı.
2. Sınıf/modül/form bazında bağımlılık ve ad çakışması haritası çıkarılmalı.
3. Kanonik klasör hedefleri, mevcut mimari dokümanlar ve gerçek proje yapısı esas alınarak belirlenmeli.
4. Eşleme önerisi ayrı bir PR içinde sunulmalı; kaynak dosyaları `develop-v2` üzerine doğrudan yazılmamalı.
5. PR içeriği yerel Excel'de VBA Compile ile doğrulanmadan birleştirilmemeli.
6. Compile başarılı olduktan sonra TEST-144'ün mevcut sonucu korunarak TEST-145 sırasıyla çalıştırılmalı.

## Test durumu

- Son başarılı test: TEST-144.
- Sıradaki test: TEST-145.
- Bu denetim TEST-145'i çalıştırmaz ve test durumunu değiştirmez.

## Güvenlik kararı

Bu commit yalnızca denetim belgesidir. VBA kaynakları, formlar, ikili form dosyaları, çalışma kitapları ve `develop-v2` dalı değiştirilmemiştir. Merge yapılmamıştır.
