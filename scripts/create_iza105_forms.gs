/**
 * İZA105 TRACE — Google Forms + Katılım Takip Sistemi
 *
 * BU SÜRÜM MEVCUT MASTER SHEET İÇİN HAZIRLANMIŞTIR.
 *
 * Master Sheet:
 * https://docs.google.com/spreadsheets/d/1lLysVcK5vaRVP8IZWKlCuG6fkvb-ll_cPIqmQJrh7CM/edit
 *
 * ÖNERİLEN İLK ÇALIŞTIRMA SIRASI
 * --------------------------------
 * 1) Bu kodun tamamını Apps Script'e yapıştırın ve Kaydet'e basın.
 * 2) validateSystem() çalıştırın.
 * 3) Master Sheet > ROSTER sekmesine:
 *      A = Öğrenci No
 *      B = Ad Soyad
 *      C = E-posta (BOŞ bırakılabilir)
 *    bilgilerini girin.
 * 4) refreshParticipationDashboard() çalıştırın.
 * 5) İsterseniz installAutoRefreshTrigger() çalıştırın.
 *
 * E-posta mantığı:
 * - Formlar Google hesabının e-posta adresini toplar.
 * - Öğrenci numarası eşleşirse e-posta ROSTER sütun C'ye otomatik yazılır.
 * - Katılım eşleştirmesinde ana anahtar Öğrenci No'dur; e-posta değildir.
 *
 * DİKKAT:
 * - setupIZA105() güvenli bir kurulum/doğrulama fonksiyonudur.
 * - CONFIG zaten doluysa yeni form üretmez.
 * - CONFIG'i elle silmeyin; aksi halde setupIZA105() yeni formlar üretir.
 */

const COURSE = "İZA105";

const MASTER_SHEET_ID =
  "1lLysVcK5vaRVP8IZWKlCuG6fkvb-ll_cPIqmQJrh7CM";

const SHEETS = {
  CONFIG: "CONFIG",
  ROSTER: "ROSTER",
  DASHBOARD: "KATILIM_OZET",
  VARIABLES: "VARIABLES_SNIPPET"
};

const WEEKS = [
  [1, true, "Veri bir dosya değil, analitik bir temsildir"],
  [2, true, "Tidy Data, gözlem birimi ve veri düzeni"],
  [3, true, "Veri edinimi: web, API ve yarı-yapısal kaynaklar"],
  [4, true, "Veri kalitesi: amaca uygunluk ve bağlam"],
  [5, true, "Veri temizleme bir ölçüm kararıdır"],
  [6, true, "Dönüştürme ve yeniden üretilebilir pipeline"],
  [7, false, "Vize — Veri Hazırlama Denetimi"],
  [8, true, "Veri entegrasyonu"],
  [9, true, "Veri modeli, depolama ve erişim"],
  [10, true, "Zamansal ve akış verisi"],
  [11, true, "Veri yönetişimi ve yeniden kullanım"],
  [12, true, "Provenance, lineage ve yeniden üretilebilirlik"],
  [13, true, "Aynı problem, farklı araç"],
  [14, false, "Final — Research Data Preparation Package"]
];


/* ========================================================================== */
/* 1. MASTER SHEET BAĞLANTISI                                                 */
/* ========================================================================== */

function getMasterSpreadsheet_() {
  if (!MASTER_SHEET_ID) {
    throw new Error(
      "MASTER_SHEET_ID boş. Kodun başındaki MASTER_SHEET_ID değerini kontrol edin."
    );
  }

  return SpreadsheetApp.openById(MASTER_SHEET_ID);
}


/* ========================================================================== */
/* 2. GÜVENLİ KURULUM                                                         */
/* ========================================================================== */

/**
 * Mevcut Master Sheet'i kontrol eder.
 *
 * - Eksik temel sekmeleri oluşturur.
 * - CONFIG boşsa formları ve CONFIG'i oluşturur.
 * - CONFIG doluysa mevcut formlara dokunmaz.
 * - VARIABLES_SNIPPET'i yeniler.
 * - Katılım dashboard'unu yeniler.
 *
 * Mevcut sistemde bu fonksiyonu çalıştırmak güvenlidir.
 */
function setupIZA105() {
  const ss = getMasterSpreadsheet_();

  ensureCoreSheets_(ss);

  const config = ss.getSheetByName(SHEETS.CONFIG);

  if (config.getLastRow() < 2) {
    createFormsAndConfig_(ss);
  } else {
    Logger.log(
      "CONFIG zaten dolu. Yeni form oluşturulmadı; mevcut formlar korunuyor."
    );
  }

  rebuildVariablesSnippet();
  refreshParticipationDashboard();

  Logger.log(`Master Sheet: ${ss.getUrl()}`);
  Logger.log("Kurulum/doğrulama tamamlandı.");
}


/**
 * Gerekli sekmeler eksikse oluşturur.
 * Var olan sayfalardaki verileri silmez.
 */
function ensureCoreSheets_(ss) {
  let config = ss.getSheetByName(SHEETS.CONFIG);
  if (!config) {
    config = ss.insertSheet(SHEETS.CONFIG);
  }

  if (config.getLastRow() === 0) {
    config.getRange(1, 1, 1, 7).setValues([[
      "Hafta",
      "Tür",
      "Başlık",
      "Form ID",
      "Öğrenci URL",
      "Edit URL",
      "Variable"
    ]]);
  }

  let roster = ss.getSheetByName(SHEETS.ROSTER);
  if (!roster) {
    roster = ss.insertSheet(SHEETS.ROSTER);
  }

  if (roster.getLastRow() === 0) {
    roster
      .getRange(1, 1, 1, 3)
      .setValues([["Öğrenci No", "Ad Soyad", "E-posta"]]);
  }

  if (!ss.getSheetByName(SHEETS.DASHBOARD)) {
    ss.insertSheet(SHEETS.DASHBOARD);
  }

  if (!ss.getSheetByName(SHEETS.VARIABLES)) {
    ss.insertSheet(SHEETS.VARIABLES);
  }
}


/* ========================================================================== */
/* 3. FORM OLUŞTURMA                                                          */
/* ========================================================================== */

/**
 * SADECE CONFIG boşsa çağrılır.
 * Yeni READ/TASK formları oluşturur ve CONFIG'e yazar.
 */
function createFormsAndConfig_(ss) {
  const config = ss.getSheetByName(SHEETS.CONFIG);

  config.clear();

  config.getRange(1, 1, 1, 7).setValues([[
    "Hafta",
    "Tür",
    "Başlık",
    "Form ID",
    "Öğrenci URL",
    "Edit URL",
    "Variable"
  ]]);

  const rows = [];

  WEEKS.forEach(([week, hasReading, title]) => {
    if (hasReading) {
      const readForm = createReadingForm_(week, title);

      rows.push([
        week,
        "READ",
        title,
        readForm.getId(),
        readForm.getPublishedUrl(),
        readForm.getEditUrl(),
        `w${pad2_(week)}_reading_form`
      ]);
    }

    const taskForm = createTaskForm_(week, title);

    rows.push([
      week,
      "TASK",
      title,
      taskForm.getId(),
      taskForm.getPublishedUrl(),
      taskForm.getEditUrl(),
      `w${pad2_(week)}_task_form`
    ]);
  });

  if (rows.length > 0) {
    config.getRange(2, 1, rows.length, 7).setValues(rows);
  }

  config.setFrozenRows(1);
  config.autoResizeColumns(1, 7);
}


/**
 * Ders öncesi akademik okuma formu.
 *
 * Hafta 7 ve 14 için READ formu oluşturulmaz.
 */
function createReadingForm_(week, title) {
  const form = FormApp.create(
    `${COURSE} W${pad2_(week)} — Ders Öncesi Okuma`
  );

  form.setDescription(
    `${title}\n\n` +
    "İki akademik okuma için Claim → Evidence → Decision → Transfer kaydı."
  );

  form.setCollectEmail(true);
  form.setAcceptingResponses(true);

  addStudentIdItem_(form);

  form
    .addParagraphTextItem()
    .setTitle("CLAIM — Okumalardaki temel veri/metodoloji iddiası nedir?")
    .setRequired(true);

  form
    .addParagraphTextItem()
    .setTitle("EVIDENCE — Bu iddia hangi veri veya kanıtla destekleniyor?")
    .setRequired(true);

  form
    .addParagraphTextItem()
    .setTitle("DECISION — Metinlerdeki önemli veri/metodoloji kararı nedir?")
    .setRequired(true);

  form
    .addParagraphTextItem()
    .setTitle("TRANSFER — Bu kararı bu haftaki veri problemine nasıl taşırsınız?")
    .setRequired(true);

  return form;
}


/**
 * Ders sonrası TRACE pekiştirme formu.
 *
 * Hafta 7: Vize teslimi
 * Hafta 14: Final teslimi
 */
function createTaskForm_(week, title) {
  const form = FormApp.create(
    `${COURSE} W${pad2_(week)} — TRACE Pekiştirme`
  );

  form.setDescription(
    `${title}\n\n` +
    "Ders sonrası TRACE kanıtı. " +
    "Araç karmaşıklığı değil, araç–problem uyumu önemlidir."
  );

  form.setCollectEmail(true);
  form.setAcceptingResponses(true);

  addStudentIdItem_(form);

  form
    .addCheckboxItem()
    .setTitle("Kullandığınız araç(lar)")
    .setChoiceValues([
      "Excel",
      "Power Query",
      "OpenRefine",
      "KNIME",
      "SQL / DuckDB",
      "Python",
      "R",
      "Diğer"
    ])
    .setRequired(true);

  [
    "PREDICT",
    "ENCOUNTER",
    "PROBE",
    "EXPLAIN",
    "DECIDE",
    "VERIFY"
  ].forEach(label => {
    form
      .addParagraphTextItem()
      .setTitle(`${label} — Kısa kanıt`)
      .setRequired(true);
  });

  form
    .addTextItem()
    .setTitle(
      "Üretilen dosya / GitHub / Drive / workflow bağlantısı"
    )
    .setRequired(true);

  const academicItem = form
    .addParagraphTextItem()
    .setTitle(
      "Akademik dayanak — Bu haftaki okumalar veri kararınızı nasıl etkiledi? " +
      "Alternatif bir karar mümkün müydü?"
    );

  // Vize ve final haftasında yeni zorunlu okuma yoktur.
  academicItem.setRequired(week !== 7 && week !== 14);

  return form;
}


/**
 * Öğrenci numarası alanı.
 * Regex çok katı tutulmuyor; mevcut lisansüstü numaralarında
 * baş harf + rakam yapısını kabul ediyor.
 */
function addStudentIdItem_(form) {
  const validation = FormApp
    .createTextValidation()
    .requireTextMatchesPattern("^[A-Za-z][0-9]+$")
    .setHelpText(
      "Öğrenci numaranızı boşluksuz girin. Örnek: Y261393001"
    )
    .build();

  form
    .addTextItem()
    .setTitle("Öğrenci No")
    .setHelpText("Örnek: Y261393001")
    .setValidation(validation)
    .setRequired(true);
}


/* ========================================================================== */
/* 4. E-POSTA SENKRONİZASYONU                                                 */
/* ========================================================================== */

/**
 * Manuel çağrı:
 * ROSTER'da e-posta sütunu boş olan öğrencilerin e-postalarını
 * form yanıtlarından tamamlar.
 */
function syncRosterEmails() {
  const ss = getMasterSpreadsheet_();
  syncRosterEmails_(ss);
}


/**
 * Formlardaki öğrenci numarasını ROSTER ile eşleştirir.
 *
 * E-posta yalnız ROSTER hücresi boşsa yazılır.
 * Mevcut e-posta değiştirilmez.
 */
function syncRosterEmails_(ss) {
  const roster = ss.getSheetByName(SHEETS.ROSTER);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  if (!roster || !config) {
    throw new Error(
      "ROSTER veya CONFIG sayfası bulunamadı."
    );
  }

  const rosterLastRow = roster.getLastRow();

  if (rosterLastRow < 2) {
    Logger.log("ROSTER boş. E-posta senkronizasyonu atlandı.");
    return;
  }

  const rosterData = roster
    .getRange(2, 1, rosterLastRow - 1, 3)
    .getValues();

  const studentRowMap = new Map();

  rosterData.forEach((row, index) => {
    const studentId = normalizeStudentId_(row[0]);

    if (studentId) {
      studentRowMap.set(studentId, index + 2);
    }
  });

  const configLastRow = config.getLastRow();

  if (configLastRow < 2) {
    Logger.log("CONFIG boş. E-posta senkronizasyonu atlandı.");
    return;
  }

  const configRows = config
    .getRange(2, 1, configLastRow - 1, 7)
    .getValues();

  let addedCount = 0;

  configRows.forEach(row => {
    const formId = String(row[3]).trim();

    if (!formId) {
      return;
    }

    let form;

    try {
      form = FormApp.openById(formId);
    } catch (error) {
      Logger.log(
        `Form açılamadı: ${formId} — ${error.message}`
      );
      return;
    }

    form.getResponses().forEach(response => {
      const email = String(
        response.getRespondentEmail() || ""
      ).trim();

      if (!email) {
        return;
      }

      let studentId = "";

      response.getItemResponses().forEach(itemResponse => {
        if (
          itemResponse.getItem().getTitle() === "Öğrenci No"
        ) {
          studentId = normalizeStudentId_(
            itemResponse.getResponse()
          );
        }
      });

      if (!studentId || !studentRowMap.has(studentId)) {
        return;
      }

      const rosterRow = studentRowMap.get(studentId);
      const emailCell = roster.getRange(rosterRow, 3);
      const currentEmail = String(
        emailCell.getValue() || ""
      ).trim();

      if (!currentEmail) {
        emailCell.setValue(email);
        addedCount += 1;
      }
    });
  });

  Logger.log(
    `E-posta senkronizasyonu tamamlandı. Yeni eklenen: ${addedCount}`
  );
}


/* ========================================================================== */
/* 5. KATILIM DASHBOARD                                                       */
/* ========================================================================== */

/**
 * ANA GÜNCELLEME FONKSİYONU
 *
 * Her çalıştırıldığında:
 * 1) Form yanıtlarından e-postaları ROSTER'a senkronize eder.
 * 2) READ/TASK tamamlanma durumlarını hesaplar.
 * 3) KATILIM_OZET sayfasını yeniden oluşturur.
 */
function refreshParticipationDashboard() {
  const ss = getMasterSpreadsheet_();

  syncRosterEmails_(ss);
  refreshParticipationDashboard_(ss);

  Logger.log(
    `Katılım özeti güncellendi: ${ss.getUrl()}`
  );
}


function refreshParticipationDashboard_(ss) {
  const roster = ss.getSheetByName(SHEETS.ROSTER);
  const config = ss.getSheetByName(SHEETS.CONFIG);
  const out = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!roster || !config || !out) {
    throw new Error(
      "ROSTER, CONFIG veya KATILIM_OZET sayfası bulunamadı."
    );
  }

  const rosterLastRow = roster.getLastRow();

  const students = rosterLastRow >= 2
    ? roster
        .getRange(2, 1, rosterLastRow - 1, 3)
        .getValues()
        .filter(row => normalizeStudentId_(row[0]) !== "")
    : [];

  const configLastRow = config.getLastRow();

  const cfg = configLastRow >= 2
    ? config
        .getRange(2, 1, configLastRow - 1, 7)
        .getValues()
        .filter(row => String(row[3]).trim() !== "")
    : [];

  out.clear();

  const headers = [
    "Öğrenci No",
    "Ad Soyad",
    "E-posta"
  ]
    .concat(
      cfg.map(row => {
        const week = row[0];
        const type = row[1];

        return `W${pad2_(week)}-${type}`;
      })
    )
    .concat([
      "Tamamlanan",
      "Toplam",
      "Oran"
    ]);

  out
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  const responseSets = cfg.map(row => {
    const formId = String(row[3]).trim();
    return getStudentIds_(formId);
  });

  const body = students.map(student => {
    const studentId = normalizeStudentId_(student[0]);
    const studentName = student[1];
    const studentEmail = student[2];

    const marks = responseSets.map(responseSet => {
      return responseSet.has(studentId) ? 1 : 0;
    });

    const done = marks.reduce(
      (sum, value) => sum + value,
      0
    );

    const total = marks.length;
    const rate = total ? done / total : 0;

    return [
      studentId,
      studentName,
      studentEmail,
      ...marks,
      done,
      total,
      rate
    ];
  });

  if (body.length > 0) {
    out
      .getRange(2, 1, body.length, body[0].length)
      .setValues(body);

    out
      .getRange(2, headers.length, body.length, 1)
      .setNumberFormat("0%");
  }

  out.setFrozenRows(1);
  out.autoResizeColumns(
    1,
    Math.min(headers.length, 15)
  );
}


/**
 * Belirli bir formda yanıt veren öğrenci numaralarını döndürür.
 */
function getStudentIds_(formId) {
  const ids = new Set();

  if (!formId) {
    return ids;
  }

  let form;

  try {
    form = FormApp.openById(formId);
  } catch (error) {
    Logger.log(
      `Katılım okunamadı; form açılamadı: ${formId} — ${error.message}`
    );
    return ids;
  }

  form.getResponses().forEach(response => {
    response.getItemResponses().forEach(itemResponse => {
      if (
        itemResponse.getItem().getTitle() === "Öğrenci No"
      ) {
        const studentId = normalizeStudentId_(
          itemResponse.getResponse()
        );

        if (studentId) {
          ids.add(studentId);
        }
      }
    });
  });

  return ids;
}


/* ========================================================================== */
/* 6. QUARTO VARIABLES_SNIPPET                                                */
/* ========================================================================== */

/**
 * CONFIG'deki öğrenci URL'lerinden _variables.yml snippet'ini yeniden üretir.
 */
function rebuildVariablesSnippet() {
  const ss = getMasterSpreadsheet_();

  ensureCoreSheets_(ss);

  const config = ss.getSheetByName(SHEETS.CONFIG);
  const sh = ss.getSheetByName(SHEETS.VARIABLES);

  sh.clear();

  sh
    .getRange("A1")
    .setValue(
      "# Aşağıdaki satırları Quarto projesindeki _variables.yml dosyasına kopyalayın."
    );

  const lastRow = config.getLastRow();

  if (lastRow < 2) {
    sh.getRange("A2").setValue(
      "# CONFIG boş; form bağlantısı bulunamadı."
    );
    return;
  }

  const rows = config
    .getRange(2, 1, lastRow - 1, 7)
    .getValues();

  const lines = rows
    .filter(row => row[4] && row[6])
    .map(row => {
      const studentUrl = String(row[4]).trim();
      const variableName = String(row[6]).trim();

      return `${variableName}: "${studentUrl}"`;
    });

  if (lines.length > 0) {
    sh
      .getRange(2, 1, lines.length, 1)
      .setValues(lines.map(line => [line]));
  }

  sh.autoResizeColumn(1);

  Logger.log(
    "VARIABLES_SNIPPET yeniden oluşturuldu."
  );
}


/* ========================================================================== */
/* 7. SİSTEM DOĞRULAMA                                                        */
/* ========================================================================== */

/**
 * Hiçbir veriyi değiştirmeden temel sistemi kontrol eder.
 *
 * İlk olarak bunu çalıştırmanız önerilir.
 */
function validateSystem() {
  const ss = getMasterSpreadsheet_();

  Logger.log(`Master Sheet bulundu: ${ss.getUrl()}`);

  const missingSheets = Object.values(SHEETS)
    .filter(name => !ss.getSheetByName(name));

  if (missingSheets.length > 0) {
    throw new Error(
      `Eksik sekmeler: ${missingSheets.join(", ")}`
    );
  }

  const config = ss.getSheetByName(SHEETS.CONFIG);
  const roster = ss.getSheetByName(SHEETS.ROSTER);

  const configRows = Math.max(
    config.getLastRow() - 1,
    0
  );

  const rosterRows = Math.max(
    roster.getLastRow() - 1,
    0
  );

  Logger.log(`CONFIG kayıt sayısı: ${configRows}`);
  Logger.log(`ROSTER öğrenci sayısı: ${rosterRows}`);

  if (configRows === 0) {
    Logger.log(
      "UYARI: CONFIG boş. setupIZA105() çalıştırılırsa yeni formlar oluşturulur."
    );
  } else {
    const formIds = config
      .getRange(2, 4, configRows, 1)
      .getValues()
      .flat()
      .filter(Boolean);

    let validForms = 0;

    formIds.forEach(formId => {
      try {
        FormApp.openById(String(formId).trim());
        validForms += 1;
      } catch (error) {
        Logger.log(
          `UYARI: Form açılamadı: ${formId}`
        );
      }
    });

    Logger.log(
      `Erişilebilen form sayısı: ${validForms}/${formIds.length}`
    );
  }

  Logger.log("Sistem doğrulaması tamamlandı.");
}


/* ========================================================================== */
/* 8. OTOMATİK YENİLEME — İSTEĞE BAĞLI                                       */
/* ========================================================================== */

/**
 * Katılım tablosunu 15 dakikada bir otomatik yenilemek için
 * zaman tabanlı trigger kurar.
 *
 * Aynı trigger zaten varsa ikinci kez oluşturmaz.
 */
function installAutoRefreshTrigger() {
  const functionName = "refreshParticipationDashboard";

  const exists = ScriptApp
    .getProjectTriggers()
    .some(trigger => {
      return trigger.getHandlerFunction() === functionName;
    });

  if (exists) {
    Logger.log(
      "Otomatik yenileme trigger'ı zaten mevcut."
    );
    return;
  }

  ScriptApp
    .newTrigger(functionName)
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log(
    "15 dakikalık otomatik katılım yenileme trigger'ı kuruldu."
  );
}


/**
 * Sadece refreshParticipationDashboard için oluşturulmuş
 * proje trigger'larını kaldırır.
 */
function removeAutoRefreshTriggers() {
  const functionName = "refreshParticipationDashboard";

  let removed = 0;

  ScriptApp
    .getProjectTriggers()
    .forEach(trigger => {
      if (
        trigger.getHandlerFunction() === functionName
      ) {
        ScriptApp.deleteTrigger(trigger);
        removed += 1;
      }
    });

  Logger.log(
    `Kaldırılan otomatik yenileme trigger sayısı: ${removed}`
  );
}


/* ========================================================================== */
/* 9. YARDIMCI FONKSİYONLAR                                                   */
/* ========================================================================== */

function normalizeStudentId_(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, "")
    .toUpperCase();
}


function pad2_(value) {
  return String(value).padStart(2, "0");
}
