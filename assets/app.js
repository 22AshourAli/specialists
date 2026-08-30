/* =================== Mokhtasoon Dashboard — v2 =================== */
(function () {
  "use strict";

  var LS_LANG = "dm_lang", LS_THEME = "dm_theme", LS_XLSX = "dm_xlsx_b64";
  // Light gate against accidental writes; MUST match PUBLISH_PIN in api/update.js.
  var PUBLISH_PIN = "mokh-2026!ash";

  /* =========================================================
     i18n (Arabic / English)
  ========================================================= */
  var STR = {
    ar: {
      docTitle: "داشبورد أعمال المختصون",
      title: "داشبورد أعمال المختصون",
      subtitle: "تحليل تفاعلي — UDS + SAP + الفواتير + الإنتاجية",
      chpUpdated: "البيانات محدّثة:",
      chpSource: "المصدر:",
      chpCurr: "العملة:",
      slxOffice: "المكتب",
      slxYear: "السنة",
      slxMonth: "الشهر",
      slxRange: "الفترة الزمنية",
      slxQism: "القسم",
      mthAll: "كل الشهور",
      rngFull: "السنة كاملة", rngYtd: "حتى {0}", rngLast6: "آخر 6 أشهر", rngLast3: "آخر 3 أشهر",
      yearhint: "اختر 'الكل' لعرض إجمالي كل الفترات — أو سنة معيّنة للمقارنة",
      all: "الكل",
      allYears: "الكل (الجمع بين السنوات)",
      yearPre: "سنة ",
      btnUpdate: "تحديث البيانات",
      btnDl: "نسخة احتياطية (Excel)",
      uploadAsk: "ارفع ملف الإكسل (تقرير اعمال المختصون …xlsx) وسيتم إعادة بناء اللوحة فوراً",
      cardEsn: "إجمالي الإسناد",
      cardEnt: "إجمالي الإنتاجية",
      cardFou: "إجمالي الفوترة",
      cardInv: "الفواتير شامل الضريبة (15٪)",
      cardPaid: "إجمالي ما تم صرفه",
      cardUnb: "إجمالي غير المفوتر",
      cardSap: "إجمالي SAP",
      cardME: "الإسناد شهرياً",
      cardMP: "الإنتاجية شهرياً",
      subEsn: "قيمة الأعمال المسندة | {0}",
      subEnt: "قيمة الأعمال المنتجة | {0}",
      subFou: "قيمة الأعمال المفوترة | {0}",
      subInv: "فواتير شاملة ضريبة 15٪ | {0}",
      subPaid: "صافي الصرف النهائي | {0}",
      subUnb: "إيرادات غير مفوترة (كلي — ورقة UNBILLED)",
      subSap: "مفوتر: {0} • الضريبة 15%: {1}",
      subME: "آخر شهر ({0}) • الفترة ({1}) • المكاتب: {2}",
      subMP: "آخر شهر ({0}) • الفترة ({1}) • المكاتب: {2}",
      deltaAll: "كل السنوات",
      deltaUp: "▲ {0}٪ عن {1}",
      deltaDown: "▼ {0}٪ عن {1}",
      deltaFlat: "≈ {1}",
      chTrend: "الاتجاه الشهري — الإسناد / الإنتاجية / الفوترة",
      chDonut: "حصة المكاتب من الإسناد",
      chOffice: "الإسناد حسب المكتب (شهرياً)",
      chTypeSplit: "الإسناد السنوي — UDS مقابل SAP",
      chUnbilled: "أعلى مشاريع غير المفوترة",
      chPaid: "ما تم صرفه شهرياً (UDS / SAP)",
      chMap: "الإسناد — الشهر × السنة",
      psTypeSplit: "اضغط على سنة لفلترة الداشبورد بها",
      psUnbilled: "إيرادات غير مفوترة على أعلى 8 مشاريع",
      sEsnad: "الإسناد", sEntajia: "الإنتاجية", sFoutra: "الفوترة",
      trAgg: "الإجمالي عبر جميع السنوات • ",
      trYear: "سنة {0} • ",
      trOffices: "المكاتب: ",
      dnBy: "حسب سنة ",
      dnAll: "الإجمالي",
      paidSub: "فترات: {0} • النطاق {1}←{2} • القسم: {3}",
      heatEmpty: "لا توجد بيانات",
      unitK: "ألف", unitM: "مليون", unitB: "مليار",
      curr: "ر.س",
      toastOk: "تم تحديث البيانات بنجاح من الملف المرفوع ✓",
      toastErr: "تعذّرت قراءة الملف — تأكد أنه ملف إكسل صحيح (xlsx)",
      toastErrDetail: "تعذّر معالجة الملف — السبب الفني:",
      toastMissing: "الملف مضغوط/غير كامل — الأوراق الناقصة:",
      toastQuota: "تم التحديث، لكن المتصفح لن يحفظ الملف محلياً (المساحة ممتلئة)",
      toastDiff: "تم تحديث الأرقام ✓ {0}",
      toastSame: "تمت قراءة الملف بنجاح، والأرقام متطابقة تماماً مع ما في اللوحة — لا يوجد تغيير في القيم أو عدد الصفوف",
      toastDup: "تم تجاهل {0} صف مكرر مطابق حرفياً لصف محسوب مسبقاً (حتى لا يُحتسب الإجمالي مرتين) — مثال:{1}",
      toastSaved: "تم تصدير نسخة احتياطية Excel احترافية (ملخص + بيانات كاملة) — يفتح في أي برنامج جداول",
      toastNew: "البيانات مأخوذة من ملف مرفوع محفوظ محلياً",
      toastSynced: "تم مزامنة أحدث البيانات المنشورة من الخادم ✓",
      badgeLocal: "مرفوع",
      badgeRemote: "منشور عالمياً",
      btnPublish: "نشر التحديث",
      publishOk: "تم نشر التحديث بنجاح ✓ — سيطّلعه الجميع خلال دقيقة",
      publishErr: "تعذّر النشر على الخادم — التحديث تطبّق محلياً فقط",
      publishBusy: "جارٍ نشر التحديث…",
      offShort: {
        "مكتب خريص": "خريص", "مكتب الشمال": "الشمال", "مكتب الجنوب": "الجنوب", "طوارئ SAP": "طوارئ",
        "مكتب الدرعية": "الدرعية", "مكتب الشرق": "الشرق",
      },
      mSoft: "", // Reserve
      footNote: "لوحة تحكم تفاعلية مبنية على ملف <b>تقرير اعمال المختصون</b> · جميع القيم بالريال السعودي (ر.س) شاملة الضريبة حيث تكون معرّفة · لتحديث البيانات: اضغط زر <b>تحديث البيانات</b> وارفع ملف الإكسل، أو شغّل ملف <b>تحديث الداشبورد.bat</b>",
      footInvN: "فواتير:",
      footUnbN: "مشروع غير مفوتر:",
      footDup: "صف مكرر مستبعد:",
      credit: "صمم وطور بواسطة",
      creditName: "عاشور علي",
      authorName: "احمد مسعود",
      waTip: "تواصل عبر واتساب: +966 539220371",
      callTip: "اتصال: 0539220371",
    },
    en: {
      docTitle: "Mokhtasoon Business Dashboard",
      title: "Mokhtasoon Business Dashboard",
      subtitle: "Interactive analytics — UDS + SAP + Invoices + Productivity",
      chpUpdated: "Data updated:",
      chpSource: "Source:",
      chpCurr: "Currency:",
slxOffice: "Office",
      slxYear: "Year",
      slxMonth: "Month",
      slxRange: "Time range",
      slxQism: "Section",
      mthAll: "All months",
      rngFull: "Full year", rngYtd: "Up to {0}", rngLast6: "Last 6 months", rngLast3: "Last 3 months",
      yearhint: "Pick 'All' for the combined total — or an exact year to compare",
      all: "All",
      allYears: "All (combined years)",
      yearPre: "Year ",
      btnUpdate: "Update Data",
      btnDl: "Backup (Excel)",
      uploadAsk: "Upload the Excel file (تقرير اعمال المختصون …xlsx) and the dashboard will rebuild instantly",
      cardEsn: "Total Attribution",
      cardEnt: "Total Productivity",
      cardFou: "Total Billing",
      cardInv: "Invoices incl. VAT (15%)",
      cardPaid: "Total Paid",
      cardUnb: "Total Unbilled",
      cardSap: "Total SAP",
      cardME: "Monthly Attribution",
      cardMP: "Monthly Productivity",
      subEsn: "Value of attributed work | {0}",
      subEnt: "Value of produced work | {0}",
      subFou: "Value of billed work | {0}",
      subInv: "Invoices incl. 15% VAT | {0}",
      subPaid: "Final net paid | {0}",
      subUnb: "Unbilled revenue (total — UNBILLED sheet)",
      subSap: "Billed: {0} • VAT 15%: {1}",
      subME: "Last month ({0}) • range ({1}) • offices: {2}",
      subMP: "Last month ({0}) • range ({1}) • offices: {2}",
      deltaAll: "All years",
      deltaUp: "▲ {0}% vs {1}",
      deltaDown: "▼ {0}% vs {1}",
      deltaFlat: "≈ {1}",
      chTrend: "Monthly trend — Attribution / Productivity / Billing",
      chDonut: "Offices share of attribution",
      chOffice: "Attribution by office (monthly)",
      chTypeSplit: "Annual attribution — UDS vs SAP",
      chUnbilled: "Top unbilled projects",
      chPaid: "Monthly paid (UDS / SAP)",
      chMap: "Attribution heatmap — month × year",
      psTypeSplit: "Click a bar year to filter the dashboard",
      psUnbilled: "Unbilled revenue on top 8 projects",
      sEsnad: "Attribution", sEntajia: "Productivity", sFoutra: "Billing",
      trAgg: "Total across all years • ",
      trYear: "Year {0} • ",
      trOffices: "Offices: ",
      dnBy: "By year ",
      dnAll: "Total",
      paidSub: "Periods: {0} • range {1}→{2} • section: {3}",
      heatEmpty: "No data",
      unitK: "K", unitM: "M", unitB: "B",
      curr: "SAR",
      toastOk: "Data updated successfully from the uploaded file ✓",
      toastErr: "Could not read the file — make sure it's a valid Excel (xlsx) file",
      toastErrDetail: "Could not process the file — technical reason:",
      toastMissing: "The file is compressed/incomplete — missing sheets:",
      toastQuota: "Updated, but the browser couldn't store the file locally (storage full)",
      toastDiff: "Numbers updated ✓ {0}",
      toastSame: "File read successfully; numbers match the dashboard exactly — no change in values or row count",
      toastDup: "Ignored {0} duplicate row(s) matching a row already counted (to avoid double counting) — e.g.:{1}",
      toastSaved: "Exported a professional Excel backup (summary + full data) — opens in any spreadsheet app",
      toastNew: "Data loaded from a locally saved uploaded file",
      toastSynced: "Synced the latest published data from the server ✓",
      badgeLocal: "uploaded",
      badgeRemote: "published globally",
      btnPublish: "Publish",
      publishOk: "Update published ✓ — everyone will see it within a minute",
      publishErr: "Could not publish to the server — updated locally only",
      publishBusy: "Publishing…",
      offShort: {
        "مكتب خريص": "Khurays", "مكتب الشمال": "Al-Shamal", "مكتب الجنوب": "Al-Junoub", "طوارئ SAP": "SAP Emerg.",
        "مكتب الدرعية": "Al-Diriyah", "مكتب الشرق": "Al-Sharq",
      },
      mSoft: "", mSoft2: "",
      footNote: "Interactive dashboard built from the file <b>تقرير اعمال المختصون</b> · all amounts in SAR incl. VAT where defined · to update: press <b>Update Data</b> and upload the Excel file, or run <b>تحديث الداشبورد.bat</b>",
      footInvN: "Invoices:",
      footUnbN: "unbilled projects:",
      footDup: "duplicates excluded:",
      credit: "Designed & developed by",
      creditName: "Ashour Ali",
      authorName: "Ahmed Masoud",
      waTip: "Chat via WhatsApp: +966 539220371",
      callTip: "Call: 0539220371",
    }
  };

  var MONTHS_ARR = {
    ar: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"],
    en: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
  };

  var lang = (localStorage.getItem(LS_LANG) === "en") ? "en" : "ar";
  var theme = (localStorage.getItem(LS_THEME) === "light") ? "light" : "dark";

  function MONTHS() { return MONTHS_ARR[lang]; }
  function T(k, a) {
    var s = (STR[lang] && STR[lang][k] !== undefined) ? STR[lang][k] : (STR.ar[k] !== undefined ? STR.ar[k] : k);
    if (a) for (var i = 0; i < a.length; i++) s = s.split("{" + i + "}").join(a[i]);
    return s;
  }
  function curr() { return STR[lang].curr; }
  function money(n) {
    var v = (isFinite(n) ? n : 0);
    return v.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function compact(n) {
    n = n || 0;
    if (n >= 1e9) return (n / 1e9).toFixed(2) + " " + T("unitB");
    if (n >= 1e6) return (n / 1e6).toFixed(2) + " " + T("unitM");
    if (n >= 1e3) return (n / 1e3).toFixed(2) + " " + T("unitK");
    return money(n);
  }

  document.documentElement.setAttribute("data-theme", theme);
  document.documentElement.lang = lang;
  document.documentElement.dir = (lang === "ar") ? "rtl" : "ltr";

  /* =========================================================
     Data loading (static or from a locally saved upload)
  ========================================================= */
  var D = null;
  var DATA_FROM_UPLOAD = false;
  var wb = null;

  function isBlank(v) {
    return v === null || v === undefined || String(v).trim() === "";
  }
  function num(v) {
    if (isBlank(v)) return 0;
    var f = parseFloat(v);
    return isNaN(f) ? 0 : f;
  }
  function clean(v) {
    if (isBlank(v)) return null;
    if (typeof v === "number") return Math.round(v * 10000) / 10000;
    return String(v).trim() || null;
  }
  function rowsOf(wb, name) {
    var n = null;
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var sn = wb.SheetNames[i];
      if (sn === name || sn.trim() === name) { n = sn; break; }
    }
    if (!n) return null;
    return XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null, blankrows: true });
  }
  function pm(row, mc, yc) {
    return { m: Math.round(num(row[mc])), y: Math.round(num(row[yc])) };
  }

  /* ---- mirrors reader.py (direct in-browser update from Excel) ---- */
  var OFFICE_BLOCKS = {
    "مكتب خريص": [["اسناد", 1], ["انتاجيه", 2], ["فوتره", 3]],
    "مكتب الشمال": [["اسناد", 4], ["انتاجيه", 5], ["فوتره", 6]],
    "مكتب الجنوب": [["اسناد", 7], ["انتاجيه", 8], ["فوتره", 9]],
    "طوارئ SAP": [
      ["اسناد", 10], ["اسناد", 13], ["اسناد", 16],
      ["انتاجيه", 11], ["انتاجيه", 14], ["انتاجيه", 17],
      ["فوتره", 12], ["فوتره", 15], ["فوتره", 18],
    ],
  };
  var METRIC_KEYS = ["اسناد", "انتاجيه", "فوتره"];
  var ADM_ORDER = ["agg", "2023", "2024", "2025", "2026_قديم", "2026_جديد"];
  var YEAR_MAP = { agg: "agg", "2023": 2023, "2024": 2024, "2025": 2025, "2026_قديم": 2026, "2026_جديد": 2026 };
  // Base monthly offices (from the admin report) — the month math below stays on these.
  var ALL_OFF = Object.keys(OFFICE_BLOCKS);
  // Display list = base offices + any extra standalone office sheets in the source.
  function allOffices() {
    return (D && D.meta && D.meta.offices && D.meta.offices.length) ? D.meta.offices : ALL_OFF;
  }

  function findAdmBlocks(adm) {
    function cellStr(r, c) {
      var v = adm[r] ? adm[r][c] : null;
      return (v === null || v === undefined) ? "" : String(v).trim();
    }
    var blocks = [];
    for (var r = 0; r < adm.length; r++) {
      if (cellStr(r, 0) !== "الشهر") continue;
      var j = r + 2, lastM = null, tot = null;
      for (; j < adm.length; j++) {
        var s = cellStr(j, 0);
        if (s === "الإجمالي") { tot = j; break; }
        if (s === "") continue;
        var n = Math.round(Number(s));
        if (n >= 1 && n <= 12 && /^\d+(\.\d+)?$/.test(s)) lastM = j;
        else break;
      }
      if (lastM !== null && tot !== null) blocks.push({ r0: r + 2, r1: lastM, rtot: tot });
    }
    return blocks;
  }

  function readPunchcard(adm) {
    var punch = {};
    var blocks = findAdmBlocks(adm);
    blocks.forEach(function (b, bi) {
      var label = ADM_ORDER[bi];
      if (!label) return;
      var ylab = YEAR_MAP[label];
      if (!punch[ylab]) punch[ylab] = {};
      ALL_OFF.forEach(function (o) { if (!punch[ylab][o]) punch[ylab][o] = {}; });
      for (var r = b.r0; r <= b.r1; r++) {
        var m = Math.round(num(adm[r] ? adm[r][0] : 0));
        if (m < 1 || m > 12) continue;
        ALL_OFF.forEach(function (off) {
          var vals = { اسناد: 0, انتاجيه: 0, فوتره: 0 };
          OFFICE_BLOCKS[off].forEach(function (spec) { vals[spec[0]] += num(adm[r] ? adm[r][spec[1]] : 0); });
          var prev = punch[ylab][off][m] || [0, 0, 0];
          punch[ylab][off][m] = [prev[0] + vals["اسناد"], prev[1] + vals["انتاجيه"], prev[2] + vals["فوتره"]];
        });
      }
      ALL_OFF.forEach(function (off) {
        var vals = { اسناد: 0, انتاجيه: 0, فوتره: 0 };
        OFFICE_BLOCKS[off].forEach(function (spec) { vals[spec[0]] += num(adm[b.rtot] ? adm[b.rtot][spec[1]] : 0); });
        var tot = punch[ylab][off]["_total"] || [0, 0, 0];
        punch[ylab][off]["_total"] = [tot[0] + vals["اسناد"], tot[1] + vals["انتاجيه"], tot[2] + vals["فوتره"]];
      });
    });
    return punch;
  }

  function readInvoices(inv) {
    var summary = {};
    var ycols = { 2023: 4, 2024: 7, 2025: 10, 2026: 13 };
    [2023, 2024, 2025, 2026].forEach(function (y) {
      var c = ycols[y];
      summary[String(y)] = {
        "جزئية": Math.round(num(inv[3][c]) * 10000) / 10000,
        "نهائية": Math.round(num(inv[4][c]) * 10000) / 10000,
        "مباشر": Math.round(num(inv[5][c]) * 10000) / 10000,
        "إجمالي": Math.round(num(inv[6][c]) * 10000) / 10000,
      };
    });
    var hasAmt = function (row) {
      for (var c = 11; c <= 17; c++) if (num(row[c])) return true;
      return false;
    };
    var dateYx = function (v) {
      var d = null;
      if (v instanceof Date && !isNaN(v.getTime())) d = v;
      else if (typeof v === "number" && v > 1000) {
        var dd = new Date(Math.round((v - 25569) * 86400000));
        if (!isNaN(dd.getTime())) d = dd;
      } else if (typeof v === "string") {
        var s = v.trim();
        var m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.]\d{1,2}/);
        if (m) return { m: +m[2], y: +m[1] };
        m = s.match(/^(\d{1,2})[-\/.]\d{1,2}[-\/.](\d{4})/);
        if (m) return { m: +m[1], y: +m[2] };
      }
      return d ? { m: d.getMonth() + 1, y: d.getFullYear() } : null;
    };
    var detail = [];
    var last = null;
    var maxYear = 2026;
    for (var i = 9; i < inv.length; i++) {
      if (isBlank(inv[i][0]) && isBlank(inv[i][2])) continue;
      if (!hasAmt(inv[i])) {
        if (!/^\d+$/.test(String(inv[i][0]).trim())) continue;
        if (!num(inv[i][9])) continue;
      }
      var yx = pm(inv[i], 8, 9);
      if (!yx.y || yx.m < 1 || yx.m > 12) {
        yx = dateYx(inv[i][10]) || (last ? { m: last.m, y: last.y } : null);
        if (!yx) yx = { m: 12, y: maxYear };
      }
      last = { m: yx.m, y: yx.y };
      maxYear = Math.max(maxYear, yx.y);
      var row = {
        "عقد": clean(inv[i][1]), "قسم": clean(inv[i][2]), "شهر": yx.m, "سنه": yx.y,
        "قيمة": Math.round(num(inv[i][11]) * 10000) / 10000,
        "ضريبة": Math.round(num(inv[i][12]) * 10000) / 10000,
        "shamel": Math.round(num(inv[i][13]) * 10000) / 10000,
        "غرامات": Math.round(num(inv[i][14]) * 10000) / 10000,
        "سلامه": Math.round(num(inv[i][15]) * 10000) / 10000,
        "safy": Math.round(num(inv[i][16]) * 10000) / 10000,
        "paid": Math.round(num(inv[i][17]) * 10000) / 10000,
        "مكتب": clean(inv[i][18]),
        "_sig": [clean(inv[i][1]), clean(inv[i][2]), clean(inv[i][3]),
          clean(inv[i][4]), clean(inv[i][5]), clean(inv[i][6]), clean(inv[i][7]),
          yx.m, yx.y, String(inv[i][10]),
          num(inv[i][11]), num(inv[i][12]), num(inv[i][13]),
          num(inv[i][14]), num(inv[i][15]), num(inv[i][16]), num(inv[i][17]),
          clean(inv[i][18])].join("\u0001"),
      };
      detail.push(row);
    }
    var dropped = [];
    var seen = {};
    var unique = [];
    for (var j = 0; j < detail.length; j++) {
      var row = detail[j];
      if (seen[row._sig]) { dropped.push(row); continue; }
      seen[row._sig] = 1;
      unique.push(row);
    }
    detail = unique;
    detail.forEach(function (r) { delete r._sig; });
    var sum = function (f) { return Math.round(detail.reduce(function (a, d) { return a + d[f]; }, 0) * 10000) / 10000; };
    return {
      "by_year": summary, "paid_total": sum("paid"),
      "foutra_bt": sum("قيمة"), "foutra_darib": sum("ضريبة"), "foutra_shamel": sum("shamel"),
      "detail": detail, "deduped_drops": dropped,
      "_check": sum("paid"),
    };
  }

  function readSap(sap) {
    var rows = [];
    for (var i = 1; i < sap.length; i++) {
      if (!/^\d+$/.test(String(sap[i][0]).trim())) continue;
      var yx = pm(sap[i], 7, 8);
      if (!yx.y) continue;
      rows.push({
        "مكتب": clean(sap[i][1]), "عقد": clean(sap[i][2]), "انوع": clean(sap[i][5]),
        "موقف": clean(sap[i][6]), "شهر": yx.m, "سنه": yx.y,
        "اسناد": num(sap[i][17]), "مفوتر": num(sap[i][18]), "ضريبه": num(sap[i][19]),
        "shamel": num(sap[i][20]), "صافي": num(sap[i][23]),
      });
    }
    var totals = { اسناد: 0, مفوتر: 0, ضريبه: 0, shamel: 0, صافي: 0 };
    rows.forEach(function (r) {
      totals["اسناد"] += r["اسناد"]; totals["مفوتر"] += r["مفوتر"];
      totals["ضريبه"] += r["ضريبه"]; totals["shamel"] += r["shamel"]; totals["صافي"] += r["صافي"];
    });
    METRIC_KEYS.concat(["shamel", "صافي"]).forEach(function (k) { totals[k] = Math.round(totals[k] * 10000) / 10000; });
    return { totals: totals, detail: rows };
  }

  function readUds(uds) {
    var rows = [];
    for (var i = 1; i < uds.length; i++) {
      if (!/^\d+$/.test(String(uds[i][0]).trim())) continue;
      var yx = pm(uds[i], 27, 28);
      if (!yx.y) yx = pm(uds[i], 13, 14);
      rows.push({
        "مكتب": clean(uds[i][1]), "اسناد": num(uds[i][24]), "مفوتر": num(uds[i][35]),
        "shamel": num(uds[i][31]), "صرف": num(uds[i][36]), "متبقي": num(uds[i][37]),
        "شهر": yx.m, "سنه": yx.y, "افادة_فوتره": clean(uds[i][17]),
        "افادة_تنفيذ": clean(uds[i][15]), "موقف": clean(uds[i][18]),
      });
    }
    var totals = { اسناد: 0, مفوتر: 0, shamel: 0, صرف: 0, متبقي: 0 };
    rows.forEach(function (r) {
      totals["اسناد"] += r["اسناد"]; totals["مفوتر"] += r["مفوتر"];
      totals["shamel"] += r["shamel"]; totals["صرف"] += r["صرف"]; totals["متبقي"] += r["متبقي"];
    });
    Object.keys(totals).forEach(function (k) { totals[k] = Math.round(totals[k] * 10000) / 10000; });
    return { totals: totals, detail: rows };
  }

  function readUnbilled(unb) {
    var projects = [];
    for (var i = 3; i < unb.length; i++) {
      var name = clean(unb[i][1]);
      if (!name) continue;
      projects.push({
        "name": name, "مبدئيه": Math.round(num(unb[i][4]) * 10000) / 10000,
        "نهائيه": Math.round(num(unb[i][5]) * 10000) / 10000,
        "رصيد_مفوتر": Math.round(num(unb[i][8]) * 10000) / 10000,
      });
    }
    var total = Math.round(projects.reduce(function (a, p) { return a + p["رصيد_مفوتر"]; }, 0) * 10000) / 10000;
    return { total: total, projects: projects };
  }

  function readTypeSplit(tf) {
    var out = {};
    var base = tf[4] || [];
    [2023, 2024, 2025, 2026].forEach(function (y, k) {
      var c = 1 + k * 2;
      out[y] = { UDS: num(base[c]), SAP: num(base[c + 1]) };
    });
    if (tf.length > 13) {
      var b2 = tf[13] || [];
      if (num(b2[1]) || num(b2[2])) {
        out[2026].UDS += num(b2[1]);
        out[2026].SAP += num(b2[2]);
      }
    }
    [2023, 2024, 2025, 2026].forEach(function (y) {
      out[y].UDS = Math.round(out[y].UDS * 10000) / 10000;
      out[y].SAP = Math.round(out[y].SAP * 10000) / 10000;
    });
    return out;
  }

  // Reads annual assignment totals from a standalone office sheet ("خريص", "الشرق", ...).
  // Blocks are titled "تفاصيل أوامر العمل المسنده ... عقد قديم/جديد", followed by
  // السنه / القسم / القيمه / الإجمالي rows. Old contract: yearly UDS columns plus a
  // single SAP total; new contract: alternating UDS/SAP pairs per year.
  function readOfficeSheet(rows) {
    var cs = function (r, c) {
      var v = rows[r] ? rows[r][c] : null;
      return (v === null || v === undefined) ? "" : String(v).trim();
    };
    var agg = 0, years = {};
    for (var r = 0; r < rows.length; r++) {
      var title = cs(r, 0);
      if (title.indexOf("تفاصيل أوامر العمل المسنده") !== 0) continue;
      var oldBlock = title.indexOf("عقد قديم") > -1;
      var hdr = r + 1, val = r + 3;
      if (cs(hdr, 0) !== "السنه" || cs(val, 0) !== "القيمه") continue;
      var ycols = {};
      for (var c = 1; c < (rows[hdr] || []).length; c++) {
        if (/^20\d\d$/.test(cs(hdr, c))) ycols[Number(cs(hdr, c))] = c;
      }
      agg += num((rows[val] || [])[9]) + num((rows[val] || [])[10]);
      for (var y in ycols) {
        var c2 = ycols[y];
        var v = num((rows[val] || [])[c2]) + (oldBlock ? 0 : num((rows[val] || [])[c2 + 1]));
        years[y] = (years[y] || 0) + v;
      }
    }
    Object.keys(years).forEach(function (y) { years[y] = r4(years[y]); });
    return { agg: r4(agg), years: years };
  }

  // Detects standalone office sheets (anything outside the standard set) and exposes
  // them as extra offices carrying yearly/aggregate assignment totals.
  function scanExtraOffices(srcWb) {
    var EXCLUDED = {};
    ["UNBILLED", "تقرير الاداره", "تقرير الفواتير", "SAP", "UDS", "تقرير تفصيلي"].forEach(function (n) {
      EXCLUDED[n.trim()] = 1;
    });
    var NAME_FIX = { "الدرعيه": "الدرعية" };
    var out = {};
    srcWb.SheetNames.forEach(function (sn) {
      var st = String(sn || "").trim();
      if (!st || EXCLUDED[st]) return;
      var label = "مكتب " + (NAME_FIX[st] || st);
      if (OFFICE_BLOCKS[label]) return;
      var rows = rowsOf(srcWb, sn);
      if (!rows) return;
      out[label] = readOfficeSheet(rows);
    });
    return out;
  }

  // Injects each extra office's assignment totals into the punchcard (agg + each year).
  function addOfficeTotals(punch, extras) {
    ["agg", 2023, 2024, 2025, 2026].forEach(function (yk) {
      Object.keys(extras).forEach(function (label) {
        if (!punch[yk]) punch[yk] = {};
        if (!punch[yk][label]) punch[yk][label] = {};
        var v = (yk === "agg") ? extras[label].agg : ((extras[label].years || {})[Number(yk)] || 0);
        punch[yk][label]["_total"] = [r4(v), 0, 0];
      });
    });
  }

  function buildFromWorkbook(srcWb, fname) {
    var adm = rowsOf(srcWb, "تقرير الاداره");
    var inv = rowsOf(srcWb, "تقرير الفواتير");
    var sap = rowsOf(srcWb, "SAP");
    var uds = rowsOf(srcWb, "UDS");
    var unb = rowsOf(srcWb, "UNBILLED");
    var tf = rowsOf(srcWb, "تقرير تفصيلي");
    if (!adm || !inv || !sap || !uds || !unb || !tf) {
      throw new Error("missing sheet");
    }
    var invData = readInvoices(inv);
    invData.detail.sort(function (a, b) { return a["سنه"] - b["سنه"] || a["شهر"] - b["شهر"]; });
    assertTotals(sap, uds, unb);
    var punch = readPunchcard(adm);
    var extras = scanExtraOffices(srcWb);
    addOfficeTotals(punch, extras);
    var extraNames = Object.keys(extras);
    var seenD = {};
    invData["deduped_drops"].forEach(function (drop) {
      if (seenD[drop._sig]) return;
      seenD[drop._sig] = 1;
      var off = drop["مكتب"] || "طوارئ SAP";
      [drop["سنه"], "agg"].forEach(function (yb) {
        var o = punch[yb] || {};
        var pc = o[off] || {};
        if (pc[drop["شهر"]]) pc[drop["شهر"]][2] -= drop["قيمة"];
        if (pc["_total"]) pc["_total"][2] -= drop["قيمة"];
      });
    });
    var unbData = readUnbilled(unb);
    return {
      meta: {
        title: "ملخص أعمال المختصون",
        source: fname,
        generated_at: new Date().toISOString(),
        years: ["agg", 2023, 2024, 2025, 2026],
        offices: ALL_OFF.slice().concat(extraNames),
        months: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
        counts: {
          invoices: invData.detail.length,
          unbilled: unbData.projects.length,
          dups: (invData.deduped_drops || []).filter(function (r) { return r["قيمة"] || r["paid"]; }).length,
        },
      },
      punchcard: punch,
      invoices: invData,
      sap: readSap(sap),
      uds: readUds(uds),
      unbilled: unbData,
      type_split: readTypeSplit(tf),
    };
  }

  function assertTotals(sap, uds, unb) { /* تُسرّب هنا للتصحيح فقط */ }

  function decodeFromLS() {
    var b64 = localStorage.getItem(LS_XLSX);
    if (!b64) return null;
    try {
      var srcWb = XLSX.read(b64, { type: "base64", cellDates: true });
      return srcWb;
    } catch (e) { return null; }
  }

  function loadData() {
    var srcWb = null;
    if (typeof XLSX !== "undefined") {
      srcWb = decodeFromLS();
      if (srcWb) {
        try {
          D = buildFromWorkbook(srcWb, "ملف مرفوع (محفوظ)");
          DATA_FROM_UPLOAD = true;
          return;
        } catch (e) {
          localStorage.removeItem(LS_XLSX);
        }
      }
    }
    if (!window.DASHBOARD_DATA) {
      document.body.innerHTML = "<p style='color:#f87171;padding:40px'>تعذّر تحميل البيانات — شغّل ملف تحديث الداشبورد.bat أو ارفع ملف الإكسل.</p>";
      return;
    }
    D = window.DASHBOARD_DATA;
  }

  /* =========================================================
     State & filters
  ========================================================= */
  var S = { offices: [], year: "agg", month: null, m1: 1, m2: 12, qism: "الكل" };

  function yearsSel() {
    return S.year === "agg" ? [2023, 2024, 2025, 2026] : [Number(S.year)];
  }

  /* ---------- extrapolation helpers ---------- */
  function punchTotals(year, offices, m1, m2) {
    var pc = D.punchcard[String(year)] || {};
    var es = 0, en = 0, fo = 0;
    offices.forEach(function (off) {
      var o = pc[off] || {};
      for (var m = m1; m <= m2; m++) { if (o[m]) { es += o[m][0]; en += o[m][1]; fo += o[m][2]; } }
    });
    return { esnad: es, entajia: en, foutra: fo };
  }

  function monthValue(year, offices, m, idx) {
    var pc = D.punchcard[String(year)] || {};
    var v = 0;
    offices.forEach(function (off) { var o = pc[off] || {}; if (o[m]) v += o[m][idx]; });
    return v;
  }

  function officeEsnad(year) {
    var pc = D.punchcard[String(year)] || {};
    var out = {};
    allOffices().forEach(function (off) { var o = pc[off] || {}; out[off] = o._total ? o._total[0] : 0; });
    return out;
  }

  function monthlySeries(year, offices, m1, m2) {
    var pc = D.punchcard[String(year)] || {};
    var arr = [];
    for (var m = m1; m <= m2; m++) {
      var e = 0, a = 0, f = 0;
      offices.forEach(function (off) { var v = (pc[off] || {})[m]; if (v) { e += v[0]; a += v[1]; f += v[2]; } });
      arr.push({ m: m, esnad: e, entajia: a, foutra: f });
    }
    return arr;
  }

  function invFiltered(applyQism) {
    var ys = yearsSel();
    var q = applyQism === undefined ? S.qism : applyQism;
    return D.invoices.detail.filter(function (d) {
      var off = d["مكتب"] || "طوارئ SAP";
      return ys.indexOf(d["سنه"]) > -1 && d["شهر"] >= S.m1 && d["شهر"] <= S.m2 &&
        (q === "الكل" || d["قسم"] === q) && S.offices.indexOf(off) > -1;
    });
  }

  /* =========================================================
     KPI cards
  ========================================================= */
  var ICONS = {
    clip: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2"/><rect x="9" y="3" width="6" height="4" rx="1"/></svg>',
    prod: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 12h-4l-3 8-6-16-3 8H2"/></svg>',
    inv: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M16 13H8M16 17H8M10 9H8"/></svg>',
    doc: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1z"/><path d="M8 7h8M8 11h8M8 15h5"/></svg>',
    cash: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="5" width="20" height="14" rx="2"/><circle cx="12" cy="12" r="2.5"/></svg>',
    clock: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
    grid: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/></svg>',
    up: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M23 6l-9.5 9.5-5-5L1 18"/><path d="M17 6h6v6"/></svg>',
    gauge: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6z"/><path d="M4.2 9.6a8 8 0 1 1 15.6 0"/><path d="M2 12h2M20 12h2"/></svg>',
  };

  var cardsCfg = [
    { id: "cardE", icon: "clip", lblKey: "cardEsn", cc: "#22d3ee", card: "esn" },
    { id: "cardP", icon: "prod", lblKey: "cardEnt", cc: "#6366f1", card: "ent" },
    { id: "cardF", icon: "inv", lblKey: "cardFou", cc: "#34d399", card: "fou" },
    { id: "cardInv", icon: "doc", lblKey: "cardInv", cc: "#fbbf24", card: "invShamel" },
    { id: "cardPaid", icon: "cash", lblKey: "cardPaid", cc: "#fb7185", card: "paid" },
    { id: "cardUnb", icon: "clock", lblKey: "cardUnb", cc: "#f87171", card: "unb" },
    { id: "cardSap", icon: "grid", lblKey: "cardSap", cc: "#a78bfa", card: "sap" },
    { id: "cardME", icon: "up", lblKey: "cardME", cc: "#60a5fa", card: "mEsnad" },
    { id: "cardMP", icon: "gauge", lblKey: "cardMP", cc: "#38bdf8", card: "mEntajia" },
  ];

  function renderCards() {
    var t = punchTotals(S.year, S.offices, S.m1, S.m2);
    var inv = invFiltered();
    var invShamel = 0, paid = 0;
    inv.forEach(function (r) { invShamel += r.shamel; paid += r.paid; });

    var monthN = Math.max(1, S.m2 - S.m1 + 1);
    var rangeTxt = MONTHS()[S.m1 - 1] + " ← " + MONTHS()[S.m2 - 1];
    var lastMV = S.m2;
    while (lastMV > S.m1) {
      var eT = monthValue(S.year, S.offices, lastMV, 0);
      var aT = monthValue(S.year, S.offices, lastMV, 1);
      var fT = monthValue(S.year, S.offices, lastMV, 2);
      if (eT || aT || fT) break;
      lastMV--;
    }
    var lastMtxt = MONTHS()[lastMV - 1];
    var short = function (o) { return STR[lang].offShort[o] || o; };
    var mEsnadV = monthValue(S.year, S.offices, lastMV, 0);
    var mEntajiaV = monthValue(S.year, S.offices, lastMV, 1);
    var offTxt = S.offices.map(short).join("، ");

    function deltaFor(metricTotal) {
      if (S.year === "agg") return null;
      var yr = Number(S.year);
      var prev = yr - 1;
      if (prev < 2023) return null;
      var old = punchTotals(prev, S.offices, S.m1, S.m2);
      var base = metricTotal === "esn" ? old.esnad : metricTotal === "ent" ? old.entajia : old.foutra;
      if (!base) return null;
      var cur = metricTotal === "esn" ? t.esnad : metricTotal === "ent" ? t.entajia : t.foutra;
      return ((cur - base) / base) * 100;
    }

    var vals = {
      esn: { v: t.esnad, sub: T("subEsn", [rangeTxt]) },
      ent: { v: t.entajia, sub: T("subEnt", [rangeTxt]) },
      fou: { v: t.foutra, sub: T("subFou", [rangeTxt]) },
      invShamel: { v: invShamel, sub: T("subInv", [rangeTxt]) },
      paid: { v: paid, sub: T("subPaid", [rangeTxt]) },
      unb: { v: D.unbilled.total, sub: T("subUnb") },
      sap: {
        v: D.sap.totals["shamel"],
        sub: T("subSap", [money(D.sap.totals["مفوتر"]), money(D.sap.totals["ضريبه"])])
          + " • " + T("sEsnad") + ": " + money(D.sap.totals["اسناد"]),
      },
      mEsnad: { v: mEsnadV, sub: T("subME", [lastMtxt, rangeTxt, offTxt]) },
      mEntajia: { v: mEntajiaV, sub: T("subMP", [lastMtxt, rangeTxt, offTxt]) },
    };

    var html = "";
    cardsCfg.forEach(function (c, i) {
      var v = vals[c.card];
      var isMetric = c.card === "esn" || c.card === "ent" || c.card === "fou";
      var delta = isMetric ? deltaFor(c.card) : null;
      var deltaHtml = "";
      if (isMetric && delta === null) deltaHtml = '<span class="delta flat">' + T("deltaAll") + "</span>";
      else if (delta > 0.5) deltaHtml = '<span class="delta up">' + T("deltaUp", [delta.toFixed(1), Number(S.year) - 1]) + "</span>";
      else if (delta < -0.5) deltaHtml = '<span class="delta down">' + T("deltaDown", [Math.abs(delta).toFixed(1), Number(S.year) - 1]) + "</span>";
      else if (isMetric) deltaHtml = '<span class="delta flat">' + T("deltaFlat", [0, Number(S.year) - 1]) + "</span>";
      html +=
        '<div class="card" style="--cc:' + c.cc + ';animation-delay:' + (i * 45) + 'ms">' +
        '<div class="row"><div class="icn">' + ICONS[c.icon] + '</div><div class="ctitle">' + T(c.lblKey) + '</div></div>' +
        '<div class="cval">' + money(v.v) + ' <small>' + curr() + '</small></div>' +
        '<div class="csub">' + v.sub + '</div>' + deltaHtml +
        '</div>';
    });
    document.getElementById("kpiCards").innerHTML = html;
  }

  /* =========================================================
     Charts — theme-aware colors
  ========================================================= */
  var TH = {
    dark: {
      text: "#c6d4ea", label: "#8ba0bd", split: "rgba(148,163,184,.12)",
      axis: "rgba(148,163,184,.2)", tipBg: "rgba(13,20,48,.94)", tipBd: "rgba(148,163,184,.25)",
      tipTxt: "#e8eef8", mapArea: "#15203a", mapBd: "#3b4d73", mapEm: "#1a2746", ringBd: "#0d1430",
    },
    light: {
      text: "#475569", label: "#64748b", split: "rgba(100,116,139,.2)",
      axis: "rgba(100,116,139,.28)", tipBg: "rgba(255,255,255,.97)", tipBd: "rgba(100,116,139,.28)",
      tipTxt: "#0f172a", mapArea: "#dbeafe", mapBd: "#93c5fd", mapEm: "#bfdbfe", ringBd: "#ffffff",
    },
  };
  function th() { return TH[theme]; }
  function cur(n) { return money(n) + " " + curr(); }

  var PALETTE = ["#22d3ee", "#6366f1", "#a78bfa", "#fbbf24", "#fb7185", "#60a5fa", "#34d399", "#f87171"];

  var charts = {};
  function initChart(id) {
    var el = document.getElementById(id);
    if (!el) return null;
    if (!charts[id]) charts[id] = echarts.init(el);
    return charts[id];
  }
  var AXIS_LABEL = null;
  var SPLIT = null;

  function baseTip() {
    return {
      trigger: "axis", axisPointer: { type: "shadow" },
      backgroundColor: th().tipBg, borderColor: th().tipBd,
      textStyle: { color: th().tipTxt, fontSize: 12 },
      valueFormatter: function (v) { return (v == null ? "—" : cur(v)); },
    };
  }

  function renderTrend() {
    var data = monthlySeries(S.year, S.offices, S.m1, S.m2);
    var xd = data.map(function (r) { return MONTHS()[r.m - 1]; });
    var short = function (o) { return STR[lang].offShort[o] || o; };
    var t = th();
    initChart("chartTrend").setOption({
      tooltip: {
        trigger: "axis", axisPointer: { type: "shadow" },
        backgroundColor: t.tipBg, borderColor: t.tipBd, textStyle: { color: t.tipTxt },
        formatter: function (ps) {
          var h = "<b>" + ps[0].axisValue + "</b>";
          ps.forEach(function (p) {
            h += "<br><span style='display:inline-block;width:9px;height:9px;border-radius:50%;background:" + p.color + ";margin-inline-end:5px'></span>" +
              p.seriesName + ": " + cur(p.value);
          });
          return h;
        },
      },
      legend: { top: 0, textStyle: { color: t.text }, itemWidth: 14, itemHeight: 10 },
      grid: { left: 8, right: 14, top: 34, bottom: 6, containLabel: true },
      xAxis: { type: "category", data: xd, axisLine: { lineStyle: { color: t.axis } }, axisLabel: { color: t.label, fontSize: 11 } },
      yAxis: { type: "value", splitLine: { lineStyle: { color: t.split } }, axisLabel: { color: t.label, fontSize: 10, formatter: compact } },
      series: [
        { name: T("sEsnad"), type: "line", smooth: true, symbol: "circle", symbolSize: 6, data: data.map(function (r) { return r.esnad; }), color: "#22d3ee", lineStyle: { width: 3 }, areaStyle: { opacity: 0.18 }, animationDelay: function (i) { return i * 50; } },
        { name: T("sEntajia"), type: "line", smooth: true, symbol: "circle", symbolSize: 6, data: data.map(function (r) { return r.entajia; }), color: "#6366f1", lineStyle: { width: 3 }, areaStyle: { opacity: 0.14 }, animationDelay: function (i) { return i * 60; } },
        { name: T("sFoutra"), type: "line", smooth: true, symbol: "circle", symbolSize: 6, data: data.map(function (r) { return r.foutra; }), color: "#34d399", lineStyle: { width: 3 }, areaStyle: { opacity: 0.12 }, animationDelay: function (i) { return i * 70; } },
      ],
      animationDuration: 700, animationEasing: "cubicOut",
    });
    var pc = D.punchcard[String(S.year)] || {};
    var sub = "";
    if (S.year === "agg") sub = T("trAgg");
    else sub = T("trYear", [S.year]);
    sub += T("trOffices") + S.offices.map(short).join("، ");
    document.getElementById("subTrend").textContent = sub;
  }

  function renderDonut() {
    var data = officeEsnad(S.year);
    var slices = S.offices.map(function (off) { return { name: off, value: data[off] }; }).filter(function (s) { return s.value > 0; });
    var t = th();
    var c = initChart("chartDonut");
    c.setOption({
      tooltip: {
        trigger: "item", backgroundColor: t.tipBg, borderColor: t.tipBd, textStyle: { color: t.tipTxt },
        formatter: function (p) { return p.name + "<br>" + cur(p.value) + " (" + p.percent + "٪)"; },
      },
      legend: { bottom: 0, textStyle: { color: t.text, fontSize: 10.5 }, itemWidth: 12, itemHeight: 8, type: "scroll" },
      color: PALETTE,
      series: [{
        type: "pie", radius: ["52%", "76%"], center: ["50%", "44%"], avoidLabelOverlap: true,
        itemStyle: { borderColor: t.ringBd, borderWidth: 2 },
        label: { color: t.text, fontSize: 11, formatter: "{b}\n{d}٪" },
        labelLine: { lineStyle: { color: th().mapBd } },
        emphasis: { scaleSize: 6, label: { fontWeight: "700" } },
        animationDuration: 800,
        data: slices.map(function (s, i) { return Object.assign({}, s, { itemStyle: { color: PALETTE[i % PALETTE.length] } }); }),
      }],
    });
    c.off("click");
    c.on("click", function (p) { setOffice([p.name]); });
    document.getElementById("subDonut").textContent =
      T("dnBy") + (S.year === "agg" ? T("dnAll") : S.year);
  }

  function renderOffice() {
    var pc = D.punchcard[String(S.year)] || {};
    var t = th();
    // Offices with monthly data only; keep all series when none has data (empty chart).
    var shown = S.offices.filter(function (off) {
      var o = pc[off] || {};
      for (var m = S.m1; m <= S.m2; m++) if (o[m] && o[m][0] > 0) return true;
      return false;
    });
    if (!shown.length) shown = S.offices.slice();
    var series = shown.map(function (off, i) {
      var o = pc[off] || {};
      var arr = [];
      for (var m = S.m1; m <= S.m2; m++) arr.push(o[m] ? o[m][0] : 0);
      return {
        name: off, type: "bar", barMaxWidth: 16,
        itemStyle: { color: PALETTE[i % PALETTE.length], borderRadius: [3, 3, 0, 0] },
        data: arr, animationDelay: function (j) { return j * 60; },
      };
    });
    var xd = [];
    for (var m = S.m1; m <= S.m2; m++) xd.push(MONTHS()[m - 1]);
    initChart("chartOffice").setOption({
      tooltip: Object.assign(baseTip(), {
        formatter: function (ps) {
          var h = "<b>" + ps[0].axisValue + "</b>";
          ps.forEach(function (p) { h += "<br>" + p.seriesName + ": " + cur(p.value); });
          return h;
        },
      }),
      legend: { top: 0, textStyle: { color: t.text, fontSize: 10.5 }, itemWidth: 12, itemHeight: 8, type: "scroll" },
      grid: { left: 8, right: 8, top: 34, bottom: 6, containLabel: true },
      xAxis: { type: "category", data: xd, axisLabel: { color: t.label, fontSize: 10 }, axisLine: { lineStyle: { color: t.axis } } },
      yAxis: { type: "value", splitLine: { lineStyle: { color: t.split } }, axisLabel: { color: t.label, fontSize: 10, formatter: compact } },
      series: series,
      animationDuration: 650, animationEasing: "cubicOut",
    });
    document.getElementById("subOffice").textContent =
      T("sEsnad") + " " + (S.year === "agg" ? T("dnAll") : S.year) + " — " + T("chOffice");
  }

  function renderTypeSplit() {
    var years = [2023, 2024, 2025, 2026];
    var cats = years.slice().reverse().map(String);
    var uds = years.slice().reverse().map(function (y) { return D.type_split[y].UDS; });
    var sap = years.slice().reverse().map(function (y) { return D.type_split[y].SAP; });
    var t = th();
    initChart("chartTypeSplit").setOption({
      tooltip: Object.assign(baseTip(), {
        formatter: function (ps) {
          var h = "<b>" + ps[0].axisValue + "</b>";
          var tot = 0;
          ps.forEach(function (p) { tot += p.value; });
          ps.forEach(function (p) { h += "<br>" + p.seriesName + ": " + cur(p.value); });
          h += "<br><b style='color:#22d3ee'>" + T("sEsnad") + ": " + cur(tot) + "</b>";
          return h;
        },
      }),
      legend: { top: 0, textStyle: { color: t.text }, itemWidth: 12, itemHeight: 8 },
      grid: { left: 6, right: 10, top: 34, bottom: 6, containLabel: true },
      xAxis: { type: "value", splitLine: { lineStyle: { color: t.split } }, axisLabel: { color: t.label, fontSize: 10, formatter: compact } },
      yAxis: { type: "category", data: cats, axisLine: { lineStyle: { color: t.axis } }, axisLabel: { color: t.label, fontSize: 11 } },
      series: [
        { name: "UDS", type: "bar", stack: "t", barMaxWidth: 30, itemStyle: { color: "#22d3ee" }, data: uds, animationDelay: function (i) { return i * 80; } },
        { name: "SAP", type: "bar", stack: "t", barMaxWidth: 30, itemStyle: { color: "#a78bfa" }, data: sap, animationDelay: function (i) { return i * 100; } },
      ],
      animationDuration: 700,
    });
    var c = initChart("chartTypeSplit");
    c.off("click");
    c.on("click", function (p) {
      var y = Number(p.name);
      if (y >= 2023 && y <= 2026) { S.year = y; document.getElementById("slicerYear").value = String(y); refresh(); }
    });
  }

  function renderUnbilled() {
    var proj = D.unbilled.projects.slice().sort(function (a, b) { return b["رصيد_مفوتر"] - a["رصيد_مفوتر"]; });
    var top = proj.slice(0, 8);
    var short = function (s) {
      return s.replace(/^مشروع العقد الموحد لشركة الكهرباء\s*-\s*/, "").trim();
    };
    var data = top.map(function (p) { return { name: short(p.name), value: p["رصيد_مفوتر"], full: p.name }; });
    var t = th();
    initChart("chartUnbilled").setOption({
      tooltip: {
        trigger: "item", backgroundColor: t.tipBg, borderColor: t.tipBd, textStyle: { color: t.tipTxt },
        formatter: function (p) { return "<b>" + p.data.full + "</b><br>" + cur(p.value) + " (" + p.percent + "٪)"; },
      },
      legend: { type: "scroll", bottom: 0, textStyle: { color: t.text, fontSize: 10 }, itemWidth: 10, itemHeight: 8 },
      color: ["#fbbf24", "#fb7185", "#a78bfa", "#60a5fa", "#34d399", "#f87171", "#22d3ee", "#f97316"],
      series: [{
        type: "pie", radius: "62%", center: ["50%", "43%"],
        itemStyle: { borderColor: t.ringBd, borderWidth: 2 },
        label: { color: t.text, fontSize: 10, formatter: "{b}\n{d}٪" },
        labelLine: { lineStyle: { color: th().mapBd } },
        animationDuration: 750,
        data: data,
      }],
    });
  }

  function renderPaid() {
    var rows = invFiltered();
    var map = {};
    rows.forEach(function (r) {
      var k = r["سنه"] + "-" + r["شهر"];
      if (!map[k]) map[k] = { UDS: 0, SAP: 0 };
      if (r["قسم"] === "UDS" || r["قسم"] === "SAP") map[k][r["قسم"]] += r.paid;
    });
    var keys = Object.keys(map).sort(function (a, b) {
      return a.split("-")[0] - b.split("-")[0] || a.split("-")[1] - b.split("-")[1];
    });
    var labels = keys.map(function (k) { var p = k.split("-"); return p[0] + " / " + MONTHS()[Number(p[1]) - 1]; });
    var uds = keys.map(function (k) { return map[k].UDS; });
    var sap = keys.map(function (k) { return map[k].SAP; });
    var t = th();
    initChart("chartPaid").setOption({
      tooltip: Object.assign(baseTip(), {
        formatter: function (ps) {
          var h = "<b>" + ps[0].axisValue + "</b>";
          var tot = 0;
          ps.forEach(function (p) { tot += p.value; });
          ps.forEach(function (p) { h += "<br>" + p.seriesName + ": " + cur(p.value); });
          h += "<br><b style='color:#22d3ee'>" + T("sEsnad") + ": " + cur(tot) + "</b>";
          return h;
        },
      }),
      legend: { top: 0, textStyle: { color: t.text }, itemWidth: 12, itemHeight: 8 },
      grid: { left: 8, right: 8, top: 34, bottom: 6, containLabel: true },
      xAxis: { type: "category", data: labels, axisLabel: { color: t.label, fontSize: 9.5 }, axisLine: { lineStyle: { color: t.axis } } },
      yAxis: { type: "value", splitLine: { lineStyle: { color: t.split } }, axisLabel: { color: t.label, fontSize: 10, formatter: compact } },
      series: [
        { name: "UDS", type: "bar", stack: "p", itemStyle: { color: "#22d3ee" }, data: uds },
        { name: "SAP", type: "bar", stack: "p", itemStyle: { color: "#a78bfa" }, data: sap },
      ],
      animationDuration: 650,
    });
    document.getElementById("subPaid").textContent = T("paidSub", [
      (S.year === "agg" ? T("dnAll") : S.year), MONTHS()[S.m1 - 1], MONTHS()[S.m2 - 1], S.qism,
    ]);
  }

  function renderHeat() {
    var yl = (S.year === "agg") ? [2023, 2024, 2025, 2026] : [Number(S.year)];
    var m1 = Number(S.m1), m2 = Number(S.m2);
    var t = th();
    var mLabs = [];
    for (var m = m1; m <= m2; m++) mLabs.push(MONTHS()[m - 1]);
    var cells = [], maxV = 0;
    yl.forEach(function (y, yi) {
      for (var mi = 0; mi < mLabs.length; mi++) {
        var mt = m1 + mi;
        var v = monthValue(String(y), S.offices, mt, 0);
        if (v > maxV) maxV = v;
        cells.push([mi, yi, Math.round(v)]);
      }
    });
    if (maxV > 0) maxV *= 1.05;
    var subEl = document.getElementById("subMap");
    var offTxt = S.offices.map(function (o) { return STR[lang].offShort[o] || o; }).join("، ");
    if (subEl) subEl.textContent = (S.year === "agg" ? T("trAgg") : T("trYear", [S.year])) + T("trOffices") + offTxt;
    initChart("chartHeat").setOption({
      tooltip: {
        backgroundColor: t.tipBg, borderColor: t.tipBd, textStyle: { color: t.tipTxt },
        formatter: function (p) {
          return "<b>" + yl[p.value[1]] + " / " + mLabs[p.value[0]] + "</b><br>" +
            T("sEsnad") + ": " + cur(p.value[2]);
        },
      },
      grid: { top: 8, left: 10, right: 10, bottom: 30, containLabel: true },
      xAxis: { type: "category", data: mLabs, splitArea: { show: true }, axisLabel: { color: t.label, fontSize: 10 }, axisLine: { lineStyle: { color: t.axis } } },
      yAxis: { type: "category", data: yl, splitArea: { show: true }, axisLabel: { color: t.label, fontSize: 11 }, axisLine: { lineStyle: { color: t.axis } } },
      visualMap: {
        min: 0, max: Math.max(1, Math.round(maxV)),
        calculable: true, orient: "horizontal", left: "center", bottom: 2,
        text: [compact(maxV), compact(0)],
        textStyle: { color: t.text },
        formatter: compact,
        inRange: { color: (theme === "dark") ? ["#0e1626", "#0e7490", "#0ea5e9", "#22d3ee"] : ["#f1f5f9", "#99f6e4", "#0e7490", "#0369a1"] },
      },
      series: [{
        type: "heatmap", data: cells,
        label: { show: false },
        itemStyle: { borderColor: t.ringBd, borderWidth: 2, borderRadius: 4 },
        emphasis: { itemStyle: { shadowBlur: 10, shadowColor: "rgba(34,211,238,.4)" } },
      }],
      animationDuration: 700,
    });
  }

  /* =========================================================
     Filters
  ========================================================= */
  function buildOfficeSlicer() {
    var el = document.getElementById("slicerOffice");
    var offs = allOffices();
    var html = '<span class="pill' + (S.offices.length === offs.length ? " on" : "") + '" data-k="all">' + T("all") + "</span>";
    offs.forEach(function (o) {
      var on = S.offices.length === 1 && S.offices[0] === o;
      html += '<span class="pill' + (on ? " on" : "") + '" data-k="' + o + '" title="' + o + '">' + STR[lang].offShort[o] + "</span>";
    });
    el.innerHTML = html;
    el.querySelectorAll(".pill").forEach(function (p) {
      p.addEventListener("click", function () {
        var k = p.getAttribute("data-k");
        if (k === "all") S.offices = offs.slice();
        else {
          if (S.offices.length === 1 && S.offices[0] === k) S.offices = offs.slice();
          else S.offices = [k];
        }
        refresh();
      });
    });
  }

  var yearBound = false;
  function buildYearSlicer() {
    var sel = document.getElementById("slicerYear");
    var html = '<option value="agg">' + T("allYears") + "</option>";
    D.meta.years.forEach(function (y) {
      if (y === "agg") return;
      html += '<option value="' + y + '">' + T("yearPre") + y + "</option>";
    });
    sel.innerHTML = html;
    sel.value = String(S.year);
    if (!yearBound) {
      yearBound = true;
      sel.addEventListener("change", function () {
        S.year = sel.value;
        S.month = null; // a year change resets any pinned month
        refresh();
      });
    }
  }

  var monthBound = false;
  // Pins the dashboard to a single month (m1 == m2). Empty value = "all months"
  // (back to the time-range pills).
  function buildMonthSlicer() {
    var sel = document.getElementById("slicerMonth");
    if (!sel) return;
    var html = '<option value="">' + T("mthAll") + "</option>";
    MONTHS().forEach(function (m, i) {
      html += '<option value="' + (i + 1) + '">' + m + "</option>";
    });
    sel.innerHTML = html;
    sel.value = S.month === null ? "" : String(S.month);
    if (!monthBound) {
      monthBound = true;
      sel.addEventListener("change", function () {
        if (sel.value === "") {
          // Back to a time range: restore the last range that was active before pinning.
          var prev = S._range || [1, Math.max(1, S.m2)];
          S.m1 = prev[0];
          S.m2 = prev[1];
          S.month = null;
        } else {
          S._range = [S.m1, S.m2]; // remember the range we are overriding
          S.month = Number(sel.value);
          S.m1 = S.month;
          S.m2 = S.month;
        }
        refresh();
      });
    }
  }

  function monthDataMax(year) {
    var pc = D.punchcard[String(year)] || {};
    var max = 0;
    ALL_OFF.forEach(function (off) {
      var o = pc[off] || {};
      for (var m = 1; m <= 12; m++) {
        if (o[m] && (o[m][0] > 0 || o[m][1] > 0 || o[m][2] > 0)) max = Math.max(max, m);
      }
    });
    return max || 12;
  }

  function buildRangeSlicer() {
    var el = document.getElementById("slicerRange");
    if (!el) return;
    var lastM = monthDataMax(S.year);
    var defs = [
      { k: "full", lbl: T("rngFull"), r: [1, 12] },
      { k: "ytd", lbl: T("rngYtd", [MONTHS()[lastM - 1]]), r: [1, lastM], show: lastM < 12 },
      { k: "last6", lbl: T("rngLast6"), r: [Math.max(1, lastM - 5), lastM], show: lastM > 6 },
      { k: "last3", lbl: T("rngLast3"), r: [Math.max(1, lastM - 2), lastM], show: lastM > 3 },
    ].filter(function (d) { return d.show !== false; });
    var html = defs.map(function (d) {
      var on = S.month === null && S.m1 === d.r[0] && S.m2 === d.r[1];
      return '<span class="pill' + (on ? " on" : "") + '" data-m1="' + d.r[0] + '" data-m2="' + d.r[1] + '">' + d.lbl + "</span>";
    }).join("");
    el.innerHTML = html;
    el.querySelectorAll(".pill").forEach(function (p) {
      p.addEventListener("click", function () {
        S.month = null; // picking a range clears any pinned month
        S.m1 = Number(p.getAttribute("data-m1"));
        S.m2 = Number(p.getAttribute("data-m2"));
        refresh();
      });
    });
  }

  function buildQismSlicer() {
    var el = document.getElementById("slicerQism");
    var html = ["الكل", "UDS", "SAP"].map(function (q) {
      var lbl = q === "الكل" ? T("all") : q;
      return '<span class="pill' + (S.qism === q ? " on" : "") + '" data-k="' + q + '">' + lbl + "</span>";
    }).join("");
    el.innerHTML = html;
    el.querySelectorAll(".pill").forEach(function (p) {
      p.addEventListener("click", function () {
        S.qism = p.getAttribute("data-k");
        buildQismSlicer();
        refresh();
      });
    });
  }

  function setOffice(arr) { S.offices = arr; refresh(); }

  function refresh() {
    if (!S.offices.length) S.offices = allOffices();
    document.getElementById("slicerYear").value = String(S.year);
    var ms = document.getElementById("slicerMonth");
    if (ms) ms.value = S.month === null ? "" : String(S.month);
    buildOfficeSlicer();
    buildRangeSlicer();
    renderCards();
    renderTrend();
    renderDonut();
    renderOffice();
    renderTypeSplit();
    renderUnbilled();
    renderPaid();
    renderHeat();
  }

  /* =========================================================
     Toolbar: language / theme / update / download / publish
  ========================================================= */
  function applyStaticI18n() {
    document.querySelectorAll("[data-i18n]").forEach(function (el) {
      var k = el.getAttribute("data-i18n");
      el.textContent = T(k);
    });
    document.title = T("docTitle");
  }

  function fmtSaudi12(iso) {
    if (!iso) return "—";
    var d = new Date(iso);
    if (isNaN(d.getTime())) return String(iso);
    try {
      var g = {};
      new Intl.DateTimeFormat("en-CA", {
        timeZone: "Asia/Riyadh", hour12: false, year: "numeric", month: "2-digit",
        day: "2-digit", hour: "2-digit", minute: "2-digit",
      }).formatToParts(d).forEach(function (x) { if (x.type !== "literal") g[x.type] = x.value; });
      var h = parseInt(g.hour, 10), mm = parseInt(g.minute, 10);
      var h12 = h % 12 || 12;
      var pad = function (n) { return (n < 10 ? "0" : "") + n; };
      return g.year + "-" + g.month + "-" + g.day + " " + pad(h12) + ":" + pad(mm) + (h >= 12 ? " م" : " ص");
    } catch (e) { return String(iso); }
  }

  function setThemeUI() {
    var ic = document.getElementById("themeIcon");
    document.getElementById("themeLbl").textContent =
      (theme === "dark") ? (lang === "ar" ? "فاتح" : "Light") : (lang === "ar" ? "داكن" : "Dark");
    if (theme === "dark") {
      ic.innerHTML = '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/>';
    } else {
      ic.innerHTML = '<path d="M21 12.8A9 9 0 1 1 11.2 3 7 7 0 0 0 21 12.8z"/>';
    }
  }

  function setLangUI() {
    applyStaticI18n();
    document.documentElement.lang = lang;
    document.documentElement.dir = (lang === "ar") ? "rtl" : "ltr";
    document.getElementById("langLbl").textContent = (lang === "ar") ? "English" : "العربية";
    document.getElementById("currChip").innerHTML = "<span>" + T("chpCurr") + "</span> " + curr();
    document.getElementById("creditCall").title = T("callTip");
    document.getElementById("fabWa").title = T("waTip");
    setThemeUI();
    buildYearSlicer();
    buildMonthSlicer();
    refresh();
  }

  function setTheme(t) {
    theme = t;
    document.documentElement.setAttribute("data-theme", t);
    try { localStorage.setItem(LS_THEME, t); } catch (e) { }
    setThemeUI();
    refresh();
  }

  /* ---------- toasts ---------- */
  function toast(msg, kind) {
    var wrap = document.getElementById("toastWrap");
    if (!wrap) return;
    var el = document.createElement("div");
    el.className = "toast " + (kind || "");
    el.textContent = msg;
    wrap.appendChild(el);
    setTimeout(function () {
      el.classList.add("out");
      setTimeout(function () { el.remove(); }, 320);
    }, 3600);
  }

  function setLoading(btn, on) { if (btn) btn.classList.toggle("loading", on); }

  function figuresOf(d0) {
    if (!d0 || !d0.punchcard || !d0.punchcard.agg) return null;
    var pc = d0.punchcard.agg;
    var es = 0, en = 0, fo = 0;
    ALL_OFF.forEach(function (off) {
      var t = pc[off] && pc[off]._total;
      if (t) { es += t[0]; en += t[1]; fo += t[2]; }
    });
    var sh = 0, pd = 0;
    (d0.invoices.detail || []).forEach(function (r) { sh += r.shamel; pd += r.paid; });
    return {
      esn: Math.round(es * 100) / 100,
      ent: Math.round(en * 100) / 100,
      fou: Math.round(fo * 100) / 100,
      sh: Math.round(sh * 100) / 100,
      pd: Math.round(pd * 100) / 100,
      unb: Math.round(d0.unbilled.total * 100) / 100,
      invN: (d0.invoices.detail || []).length,
      dups: (d0.invoices.deduped_drops || []).filter(function (r) { return r["قيمة"] || r["paid"]; }).length,
    };
  }

  function dupSample(d0) {
    var drops = (d0.invoices.deduped_drops || []).filter(function (r) { return r["قيمة"] || r["paid"]; });
    var d = drops[drops.length - 1];
    if (!d) return "";
    var parts = [];
    if (d["قسم"]) parts.push(d["قسم"]);
    if (d["سنه"]) parts.push(d["سنه"] + "/" + String(d["شهر"] || "?"));
    return " " + money(Number(d["قيمة"]) || Number(d["paid"]) || 0) + (parts.length ? " (" + parts.join(" ") + ")" : "");
  }

  function diffFigures(oldF, newF) {
    if (!oldF || !newF) return "";
    var map = {
      esn: "cardEsn", ent: "cardEnt", fou: "cardFou",
      sh: "cardInv", pd: "cardPaid", unb: "cardUnb",
    };
    var parts = [];
    Object.keys(map).forEach(function (k) {
      if (oldF[k] !== newF[k]) parts.push(T(map[k]) + " " + cur(oldF[k]) + " ← " + cur(newF[k]));
    });
    if (oldF.invN !== newF.invN) parts.push(T("footInvN") + " " + oldF.invN + " ← " + newF.invN);
    return parts.slice(0, 3).join("  •  ");
  }

  function updateFoot() {
    var el = document.getElementById("footNote");
    if (!el || !D) return;
    var c = D.meta.counts || {};
    el.innerHTML = T("footNote") + '<span class="footmeta"> &nbsp;|&nbsp; ' +
      T("footInvN") + " " + (c.invoices || 0) + " &nbsp;|&nbsp; " +
      T("footUnbN") + " " + (c.unbilled || 0) +
      ((c.dups || 0) > 0 ? ' &nbsp;|&nbsp; ' + T("footDup") + " " + c.dups : "") + "</span>";
  }

  function handleFile(ev) {
    var f = ev.target.files && ev.target.files[0];
    ev.target.value = "";
    if (!f) return;
    var btn = document.getElementById("updateBtn");
    setLoading(btn, true);
    var rd = new FileReader();
    rd.onerror = function () { setLoading(btn, false); toast(T("toastErr"), "err"); };
    rd.onload = function () {
      try {
        var b64 = String(rd.result).split(",")[1] || rd.result;
        var srcWb = XLSX.read(b64, { type: "base64", cellDates: true });
        var NEEDED = ["تقرير الاداره", "تقرير الفواتير", "UNBILLED", "SAP", "UDS", "تقرير تفصيلي"];
        var missing = NEEDED.filter(function (n) {
          return !srcWb.SheetNames.some(function (s) { return s === n || s.trim() === n; });
        });
        if (missing.length) {
          setLoading(btn, false);
          toast(T("toastMissing") + " " + missing.join("، "), "err");
          return;
        }
        var prev = figuresOf(D);
        D = buildFromWorkbook(srcWb, f.name);
        DATA_FROM_UPLOAD = true;
        try { localStorage.setItem(LS_XLSX, b64); } catch (e) { toast(T("toastQuota")); }
        document.getElementById("genAt").textContent = fmtSaudi12(D.meta.generated_at);
        document.getElementById("srcName").textContent = D.meta.source;
        var badge = document.getElementById("srcBadge");
        badge.style.display = "inline-block";
        badge.textContent = T("badgeLocal");
        refresh();
        updateFoot();
        setLoading(btn, false);
        var nf = figuresOf(D);
        var df = diffFigures(prev, nf);
        if (df) toast(T("toastDiff", [df]), "ok");
        else if (nf.dups > prev.dups) toast(T("toastDup", [nf.dups - prev.dups, dupSample(D)]), "");
        else toast(T("toastSame"), "");
        publishData();
      } catch (e) {
        setLoading(btn, false);
        toast(T("toastErrDetail") + " " + (e && e.message ? e.message : String(e)), "err");
      }
    };
    rd.readAsDataURL(f);
  }

  /* =========================================================
     Excel backup export — a structured, styled workbook
  ========================================================= */
  function r4(x) { return Math.round((x || 0) * 10000) / 10000; }

  function aoaSheet(rows, widths) {
    var ws = XLSX.utils.aoa_to_sheet(rows);
    if (widths) ws["!cols"] = widths.map(function (w) { return { wch: w }; });
    return ws;
  }

  // Bolds the first row and gives it a navy fill (readable header).
  function styleFirstRow(ws) {
    if (!ws || !ws["!ref"]) return ws;
    var rng = XLSX.utils.decode_range(ws["!ref"]);
    for (var c = rng.s.c; c <= rng.e.c; c++) {
      var addr = XLSX.utils.encode_cell({ r: 0, c: c });
      if (ws[addr]) ws[addr].s = {
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { patternType: "solid", fgColor: { rgb: "1F3A5F" } },
        alignment: { vertical: "center", horizontal: "center" },
        border: { bottom: { style: "thin", color: { rgb: "0F2440" } } },
      };
    }
    return ws;
  }

  // Headline figures used by the summary sheet (agnostic to the current filters).
  function kpiSummary() {
    var t = punchTotals("agg", allOffices(), 1, 12);
    return [
      [T("cardEsn"), t.esnad],
      [T("cardEnt"), t.entajia],
      [T("cardFou"), t.foutra],
      [T("cardInv"), D.invoices.foutra_shamel],
      [T("cardPaid"), D.invoices.paid_total],
      [T("cardUnb"), D.unbilled.total],
      [T("cardSap"), D.sap.totals.shamel],
      [T("footInvN"), D.meta.counts.invoices],
      [T("footUnbN"), D.meta.counts.unbilled],
      [T("chpUpdated"), String(D.meta.generated_at)],
      [T("chpSource"), String(D.meta.source)],
    ];
  }

  function downloadData() {
    if (!D || typeof XLSX === "undefined") return;
    var wb = XLSX.utils.book_new();
    var lbl = lang === "ar" ? ["البيان", "القيمة"] : ["Item", "Value"];

    // Summary: headline figures + metadata.
    var sum = [lbl].concat(kpiSummary());
    XLSX.utils.book_append_sheet(wb, styleFirstRow(aoaSheet(sum, [40, 24])), "الملخص");

    // Invoices detail.
    var invRows = [["عقد", "قسم", "شهر", "سنه", "قيمة", "ضريبة", "شامل", "غرامات", "سلامه", "صافي", "مصروف", "مكتب"]];
    D.invoices.detail.forEach(function (r) {
      invRows.push([r["عقد"], r["قسم"], r["شهر"], r["سنه"], r["قيمة"], r["ضريبة"], r["shamel"], r["غرامات"], r["سلامه"], r["safy"], r["paid"], r["مكتب"]]);
    });
    XLSX.utils.book_append_sheet(wb, styleFirstRow(aoaSheet(invRows, [12, 10, 6, 6, 13, 13, 13, 13, 13, 13, 13, 12])), "الفواتير");

    // Monthly matrix for the currently selected year.
    var pc = D.punchcard[String(S.year)] || {};
    var mths = [];
    allOffices().forEach(function (off) {
      var o = pc[off] || {};
      for (var m = 1; m <= 12; m++) {
        if (o[m] && (o[m][0] || o[m][1] || o[m][2])) mths.push([off, m, r4(o[m][0]), r4(o[m][1]), r4(o[m][2])]);
      }
    });
    mths.unshift(["مكتب", "رقم الشهر", "إسناد", "إنتاجية", "فوترة"]);
    XLSX.utils.book_append_sheet(wb, styleFirstRow(aoaSheet(mths, [16, 10, 14, 14, 14])), "الشهري");

    // SAP detail.
    var sapRows = [["مكتب", "عقد", "نوع", "موقف", "شهر", "سنه", "اسناد", "مفوتر", "ضريبه", "شامل"]];
    D.sap.detail.forEach(function (r) { sapRows.push([r["مكتب"], r["عقد"], r["انوع"], r["موقف"], r["شهر"], r["سنه"], r["اسناد"], r["مفوتر"], r["ضريبه"], r["shamel"]]); });
    XLSX.utils.book_append_sheet(wb, styleFirstRow(aoaSheet(sapRows, [14, 10, 12, 14, 6, 6, 14, 12, 12, 14])), "SAP");

    // UDS detail.
    var udsRows = [["مكتب", "شهر", "سنه", "اسناد", "مفوتر", "شامل", "صرف", "متبقي", "موقف", "افادة_فوتره", "افادة_تنفيذ"]];
    D.uds.detail.forEach(function (r) {
      udsRows.push([r["مكتب"], r["شهر"], r["سنه"], r["اسناد"], r["مفوتر"], r["shamel"], r["صرف"], r["متبقي"], r["موقف"], r["افادة_فوتره"], r["افادة_تنفيذ"]]);
    });
    XLSX.utils.book_append_sheet(wb, styleFirstRow(aoaSheet(udsRows, [14, 6, 6, 14, 12, 14, 12, 12, 10, 16, 16])), "UDS");

    // Unbilled projects.
    var unbRows = [["المشروع", "مبدئيه", "نهائيه", "رصيد_مفوتر"]];
    D.unbilled.projects.forEach(function (p) { unbRows.push([p["name"], p["مبدئيه"], p["نهائيه"], p["رصيد_مفوتر"]]); });
    XLSX.utils.book_append_sheet(wb, styleFirstRow(aoaSheet(unbRows, [44, 14, 14, 14])), "UNBILLED");

    var out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    var blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    var d = new Date();
    var pad = function (n) { return (n < 10 ? "0" : "") + n; };
    a.download = "Mokhtasoon-backup-" + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + ".xlsx";
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 400);
    toast(T("toastSaved"), "ok");
  }

  // Publishes the current data to the hosted server (Vercel api/update) so everyone
  // sees it. Best effort: locals (file:// / http dev) simply no-op.
  var publishing = false;
  function publishData() {
    if (!D || publishing) return;
    if (window.location.protocol !== "https:") return;
    publishing = true;
    var btn = document.getElementById("publishBtn");
    setLoading(btn, true);
    toast(T("publishBusy"), "");
    fetch("api/update", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pin: PUBLISH_PIN, data: JSON.stringify(D) }),
    }).then(function (res) {
      return res.json().catch(function () { return {}; }).then(function (j) {
        publishing = false;
        setLoading(btn, false);
        if (res.ok && j && j.ok) toast(T("publishOk"), "ok");
        else toast(T("publishErr"), "err");
      });
    }).catch(function () {
      publishing = false;
      setLoading(btn, false);
      toast(T("publishErr"), "err");
    });
  }

  /* =========================================================
     Global sync: adopt the latest published dataset (api/data)
  ========================================================= */
  // Reloads window.DASHBOARD_DATA from a remote script and returns true when the
  // published data is newer than the locally loaded one.
  function adoptRemote(script) {
    try {
      var before = window.DASHBOARD_DATA;
      new Function(script)();
      var nd = window.DASHBOARD_DATA;
      if (!nd || !nd.meta || !nd.meta.generated_at) { window.DASHBOARD_DATA = before; return false; }
      var cur = D && D.meta && D.meta.generated_at;
      if (cur && cur >= nd.meta.generated_at) return false; // keep the newer local copy
      D = nd;
      DATA_FROM_UPLOAD = false;
      return true;
    } catch (e) { return false; }
  }

  // Reflects the current dataset's metadata in the header chips.
  function showSourceUI() {
    document.getElementById("genAt").textContent = fmtSaudi12(D.meta.generated_at);
    document.getElementById("srcName").textContent = D.meta.source;
  }

  function showRemoteSource() {
    showSourceUI();
    var badge = document.getElementById("srcBadge");
    badge.style.display = "inline-block";
    badge.textContent = T("badgeRemote");
  }

  // Best-effort pull of the published dataset once per page load (https only).
  // Vercel/static hosts: the api/data route serves the repo's latest data.js.
  function syncPublished() {
    if (window.location.protocol !== "https:") return;
    fetch("api/data?ts=" + Date.now(), { cache: "no-store" })
      .then(function (r) { return r.ok ? r.text() : null; })
      .then(function (txt) {
        if (!txt || !adoptRemote(txt)) return;
        showRemoteSource();
        refresh();
        updateFoot();
        toast(T("toastSynced"), "ok");
      })
      .catch(function () { /* offline server is fine — fall back to bundled data */ });
  }

  function applyStagger() {
    var delays = [0, 60, 120, 180, 240, 300, 360, 420];
    document.querySelectorAll(".panel").forEach(function (el, i) {
      el.style.animationDelay = (delays[i % delays.length]) + "ms";
    });
  }

  function initControls() {
    document.getElementById("langBtn").addEventListener("click", function () {
      lang = (lang === "ar") ? "en" : "ar";
      try { localStorage.setItem(LS_LANG, lang); } catch (e) { }
      setLangUI();
    });
    document.getElementById("themeBtn").addEventListener("click", function () {
      setTheme(theme === "dark" ? "light" : "dark");
    });
    document.getElementById("updateBtn").addEventListener("click", function () {
      toast(T("uploadAsk"));
      document.getElementById("fileXlsx").click();
    });
    document.getElementById("dlBtn").addEventListener("click", downloadData);
    document.getElementById("publishBtn").addEventListener("click", publishData);
    document.getElementById("fileXlsx").addEventListener("change", handleFile);
  }

  /* =========================================================
     Boot
  ========================================================= */
  loadData();
  if (!D) return;

  initControls();
  S.offices = allOffices();

  document.getElementById("genAt").textContent = fmtSaudi12(D.meta.generated_at);
  document.getElementById("srcName").textContent = D.meta.source;
  if (DATA_FROM_UPLOAD) {
    document.getElementById("srcBadge").style.display = "inline-block";
    document.getElementById("srcBadge").textContent = T("badgeLocal");
  }

  updateFoot();
  buildYearSlicer();
  buildMonthSlicer();
  buildRangeSlicer();
  buildQismSlicer();
  applyStaticI18n();
  document.getElementById("langLbl").textContent = (lang === "ar") ? "English" : "العربية";
  document.getElementById("currChip").innerHTML = "<span>" + T("chpCurr") + "</span> " + curr();
  document.getElementById("creditCall").title = T("callTip");
  document.getElementById("fabWa").title = T("waTip");
  setThemeUI();
  refresh();
  updateFoot();
  applyStagger();
  syncPublished();

  window.addEventListener("resize", function () {
    Object.keys(charts).forEach(function (k) { if (charts[k]) charts[k].resize(); });
  });

  if (DATA_FROM_UPLOAD) toast(T("toastNew"), "ok");
})();