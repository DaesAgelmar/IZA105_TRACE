# İZA105 TRACE — Revizyon 2

Bu sürüm dersin teorik çerçevesini, haftalık okumalarını ve ampirik uygulamalarını tek bir **M6 running case** etrafında bütünleştirir.

## Pedagojik mimari

Her içerik haftası üç katman içerir:

1. Veri problemi / teorik okuma
2. Epistemik sonuç çıpası
3. M6 Case Window

Ders içi akış:

`Reading → M6 case → Predict → Encounter → Probe → Explain → Decide → Verify → Consequence check`

M6 (`Before Validation: How Transition Operationalization Constructs Bibliometric Process Phases`) öğrencinin dış literatürde öğrendiği ilkeleri somut bir araştırma zincirinde sınadığı ortak vaka olarak kullanılır.

## Önemli: _variables.yml

Bu revizyon paketi **_variables.yml dosyasını içermez**.

Canlı Google Form bağlantılarınızı barındıran mevcut `_variables.yml` dosyanızı koruyun. Revizyon paketini proje köküne açtığınızda mevcut `_variables.yml` dosyasını silmeyin veya değiştirmeyin.

## Değişen içerikler

- `_quarto.yml`: M6 vaka sayfası navbar/sidebar'a eklendi.
- `references.bib`: data construction + specification uncertainty + M6 literatürü eklendi.
- `index.qmd`: üç katmanlı ders mimarisi.
- `course-guide.qmd`: iki dış okuma + M6 running case.
- `trace.qmd`: TRACE + Consequence Check.
- `m6-case.qmd`: dönem boyunca kullanılan ortak ampirik vaka.
- `participation.qmd`: READ/TASK yorum mantığı.
- `assessment.qmd`: vize ve finalde alternative branch.
- `final-project.qmd`: Decision Sensitivity Map + tested branch.
- `weeks/week01.qmd`–`week14.qmd`: tüm haftalar yeniden kurgulandı.
- `styles.css`: case/anchor/task kartları.
- `scripts/create_iza105_forms.gs`: en son revize edilmiş standalone-safe sürüm kullanıldı.

## Form sistemi

Mevcut 26 Google Form'un URL ve ID'leri değişmez. Haftalık sayfalar aynı `_variables.yml` anahtarlarını kullanır.

Formlar zaten genel `Claim–Evidence–Decision–Transfer` ve TRACE alanlarını içerdiği için yeni okuma mimarisiyle uyumludur.

## Önizleme

```bash
quarto preview
```

veya:

```bash
quarto render
```

Quarto bu çalışma ortamında kurulu olmadığı için bu paket üzerinde gerçek render çalıştırılmamıştır; buna karşılık dosya, citation-key ve shortcode tutarlılığı statik olarak doğrulanmıştır.
