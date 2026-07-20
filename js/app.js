/* =========================================================
   مُطابق — منطق أداة التسوية البنكية
   كل المعالجة تتم محليًا داخل المتصفح (لا رفع بيانات لأي خادم)
   ========================================================= */

(function () {
  "use strict";

  // ---------------------------------------------------------
  // الحالة العامة
  // ---------------------------------------------------------
  const state = {
    bank: { headers: [], rows: [], mapping: {} },
    ledger: { headers: [], rows: [], mapping: {} },
    records: { bank: [], ledger: [] },   // بعد التطبيع
    matches: [],                          // { bankId, ledgerId, manual, diff }
    bankOnly: [],
    ledgerOnly: [],
    selection: { bank: null, ledger: null }
  };

  const KEYWORDS = {
    date:   ["تاريخ", "التاريخ", "date", "تارىخ"],
    desc:   ["بيان", "البيان", "وصف", "الوصف", "description", "details", "narration", "ملاحظات", "شرح"],
    amount: ["مبلغ", "المبلغ", "amount", "قيمة", "القيمة"],
    debit:  ["مدين", "سحب", "السحب", "debit", "withdrawal"],
    credit: ["دائن", "ايداع", "إيداع", "الايداع", "credit", "deposit"]
  };

  const ARABIC_DIGITS = { "٠":"0","١":"1","٢":"2","٣":"3","٤":"4","٥":"5","٦":"6","٧":"7","٨":"8","٩":"9","٬":"","،":"" };

  // ---------------------------------------------------------
  // أدوات مساعدة عامة
  // ---------------------------------------------------------
  function normalizeDigits(str) {
    return String(str).replace(/[٠-٩٬،]/g, ch => ARABIC_DIGITS[ch] ?? ch);
  }

  function parseAmount(value) {
    if (value === null || value === undefined || value === "") return null;
    if (typeof value === "number") return value;
    let s = normalizeDigits(String(value)).trim();
    if (s === "") return null;
    let negative = false;
    if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1); }
    s = s.replace(/[^\d.\-]/g, "");
    if (s === "" || s === "-" || s === ".") return null;
    let n = parseFloat(s);
    if (isNaN(n)) return null;
    if (negative) n = -Math.abs(n);
    return n;
  }

  function excelSerialToDate(serial) {
    // Excel epoch (with 1900 leap-year bug offset handled by SheetJS convention)
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    return new Date(utcValue * 1000);
  }

  function parseDateValue(value) {
    if (value === null || value === undefined || value === "") return null;
    if (value instanceof Date && !isNaN(value)) return value;
    if (typeof value === "number") {
      const d = excelSerialToDate(value);
      return isNaN(d) ? null : d;
    }
    const s = normalizeDigits(String(value)).trim();
    if (!s) return null;
    // Try common formats: yyyy-mm-dd, dd/mm/yyyy, dd-mm-yyyy
    let m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }

  function fmtDate(d) {
    if (!d) return "—";
    return d.toLocaleDateString("en-GB");
  }

  function fmtAmount(n) {
    if (n === null || n === undefined || isNaN(n)) return "—";
    return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function daysBetween(a, b) {
    if (!a || !b) return Infinity;
    return Math.abs((a - b) / 86400000);
  }

  function guessColumn(headers, kind) {
    const words = KEYWORDS[kind];
    const found = headers.find(h => words.some(w => String(h).toLowerCase().includes(w.toLowerCase())));
    return found || "";
  }

  // ---------------------------------------------------------
  // قراءة الملفات (Excel / CSV) عبر SheetJS
  // ---------------------------------------------------------
  function readFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array", cellDates: true });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          const headers = rows.length ? Object.keys(rows[0]) : [];
          resolve({ headers, rows });
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  // ---------------------------------------------------------
  // واجهة رفع الملفات
  // ---------------------------------------------------------
  const hint = document.getElementById("uploadHint");

  function bindUpload(inputId, labelId, key) {
    const input = document.getElementById(inputId);
    const label = document.getElementById(labelId);
    const drop = label; // label acts as drop zone

    input.addEventListener("change", () => handleFile(input.files[0]));

    ["dragover", "dragenter"].forEach(evt =>
      drop.addEventListener(evt, (e) => { e.preventDefault(); drop.classList.add("is-filled"); })
    );
    drop.addEventListener("drop", (e) => {
      e.preventDefault();
      if (e.dataTransfer.files.length) {
        input.files = e.dataTransfer.files;
        handleFile(e.dataTransfer.files[0]);
      }
    });

    async function handleFile(file) {
      if (!file) return;
      label.textContent = "⏳ جاري القراءة: " + file.name;
      try {
        const { headers, rows } = await readFile(file);
        if (!headers.length) throw new Error("الملف فارغ أو غير مقروء");
        state[key].headers = headers;
        state[key].rows = rows;
        label.textContent = "✓ " + file.name + " — " + rows.length + " صف";
        label.classList.add("is-filled");
        hint.textContent = "";
        maybeShowMapping();
      } catch (err) {
        hint.textContent = "تعذّرت قراءة الملف (" + file.name + "): " + err.message;
        label.textContent = "حدث خطأ — أعد الاختيار";
      }
    }
  }

  bindUpload("file-bank", "label-bank", "bank");
  bindUpload("file-ledger", "label-ledger", "ledger");

  function maybeShowMapping() {
    if (state.bank.rows.length && state.ledger.rows.length) {
      document.getElementById("panel-mapping").classList.remove("hidden");
      renderMapping("bank", "الكشف البنكي");
      renderMapping("ledger", "الكشف الدفتري");
      document.getElementById("panel-mapping").scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  // ---------------------------------------------------------
  // واجهة تحديد الأعمدة (Mapping)
  // ---------------------------------------------------------
  const ROLE_OPTIONS = [
    ["", "— تجاهل —"],
    ["date", "التاريخ"],
    ["desc", "البيان"],
    ["amount", "المبلغ (بقيمة موجبة/سالبة)"],
    ["debit", "مدين / سحب"],
    ["credit", "دائن / إيداع"]
  ];

  function renderMapping(key, title) {
    const data = state[key];
    const fieldWrap = document.getElementById("mapping-" + key);
    fieldWrap.innerHTML = "";

    const auto = {
      date: guessColumn(data.headers, "date"),
      desc: guessColumn(data.headers, "desc"),
      amount: guessColumn(data.headers, "amount"),
      debit: guessColumn(data.headers, "debit"),
      credit: guessColumn(data.headers, "credit")
    };

    data.headers.forEach((h) => {
      const field = document.createElement("div");
      field.className = "mapping-field";
      const roleGuess = Object.keys(auto).find(r => auto[r] === h) || "";
      field.innerHTML = `
        <span>${escapeHtml(h)}</span>
        <select data-col="${escapeHtml(h)}">
          ${ROLE_OPTIONS.map(([v, label]) => `<option value="${v}" ${v === roleGuess ? "selected" : ""}>${label}</option>`).join("")}
        </select>`;
      fieldWrap.appendChild(field);
    });

    fieldWrap.querySelectorAll("select").forEach(sel => {
      sel.addEventListener("change", () => collectMapping(key));
    });
    collectMapping(key);
    renderPreview(key);
  }

  function collectMapping(key) {
    const fieldWrap = document.getElementById("mapping-" + key);
    const mapping = {};
    fieldWrap.querySelectorAll("select").forEach(sel => {
      if (sel.value) mapping[sel.value] = sel.dataset.col;
    });
    state[key].mapping = mapping;
  }

  function renderPreview(key) {
    const data = state[key];
    const wrap = document.getElementById("preview-" + key);
    const rows = data.rows.slice(0, 5);
    let html = "<table><thead><tr>" + data.headers.map(h => `<th>${escapeHtml(h)}</th>`).join("") + "</tr></thead><tbody>";
    rows.forEach(r => {
      html += "<tr>" + data.headers.map(h => `<td>${escapeHtml(r[h])}</td>`).join("") + "</tr>";
    });
    html += "</tbody></table>";
    wrap.innerHTML = html;
  }

  function escapeHtml(v) {
    return String(v ?? "").replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[c]));
  }

  // ---------------------------------------------------------
  // تطبيع البيانات إلى سجلات موحّدة
  // ---------------------------------------------------------
  function buildRecords(key) {
    const data = state[key];
    const m = data.mapping;
    if (!m.amount && !m.debit && !m.credit) {
      throw new Error(`حدد عمود "المبلغ" أو "مدين/دائن" في ${key === "bank" ? "الكشف البنكي" : "الكشف الدفتري"}`);
    }
    return data.rows.map((r, idx) => {
      let amount = null;
      if (m.amount) {
        amount = parseAmount(r[m.amount]);
      } else {
        const debit = m.debit ? parseAmount(r[m.debit]) || 0 : 0;
        const credit = m.credit ? parseAmount(r[m.credit]) || 0 : 0;
        amount = credit - debit;
      }
      const date = m.date ? parseDateValue(r[m.date]) : null;
      const desc = m.desc ? String(r[m.desc] ?? "") : "";
      return {
        id: key + "_" + idx,
        source: key,
        date,
        desc,
        amount: amount === null ? 0 : amount,
        absAmount: Math.round(Math.abs(amount === null ? 0 : amount) * 100) / 100,
        matched: false
      };
    }).filter(rec => rec.absAmount > 0 || rec.desc); // تجاهل الصفوف الفارغة تمامًا
  }

  // ---------------------------------------------------------
  // محرك المطابقة
  // ---------------------------------------------------------
  function runMatching() {
    const dateTolerance = parseFloat(document.getElementById("dateTolerance").value) || 0;
    const amountTolerance = parseFloat(document.getElementById("amountTolerance").value) || 0;

    let bankRecs, ledgerRecs;
    try {
      bankRecs = buildRecords("bank");
      ledgerRecs = buildRecords("ledger");
    } catch (err) {
      alert(err.message);
      return;
    }

    state.records.bank = bankRecs;
    state.records.ledger = ledgerRecs;
    state.matches = [];

    const ledgerPool = ledgerRecs.slice();

    bankRecs.forEach(bankRec => {
      let candidates = ledgerPool.filter(l =>
        Math.abs(l.absAmount - bankRec.absAmount) <= amountTolerance + 0.001
      );

      if (candidates.length > 1 && (bankRec.date || dateTolerance > 0)) {
        const withinDate = candidates.filter(c => daysBetween(c.date, bankRec.date) <= dateTolerance);
        if (withinDate.length) candidates = withinDate;
      }

      if (candidates.length > 1) {
        candidates.sort((a, b) => daysBetween(a.date, bankRec.date) - daysBetween(b.date, bankRec.date));
      }

      if (candidates.length) {
        const chosen = candidates[0];
        bankRec.matched = true;
        chosen.matched = true;
        const idx = ledgerPool.indexOf(chosen);
        ledgerPool.splice(idx, 1);
        state.matches.push({
          bankId: bankRec.id, ledgerId: chosen.id, manual: false,
          diff: Math.round((bankRec.absAmount - chosen.absAmount) * 100) / 100
        });
      }
    });

    state.bankOnly = bankRecs.filter(r => !r.matched);
    state.ledgerOnly = ledgerRecs.filter(r => !r.matched);

    renderResults();
    document.getElementById("panel-results").classList.remove("hidden");
    document.getElementById("stateStamp").textContent = "تمت المطابقة ⎔";
    document.getElementById("stateStamp").classList.add("is-active");
    document.getElementById("panel-results").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function findRec(id) {
    return state.records.bank.find(r => r.id === id) || state.records.ledger.find(r => r.id === id);
  }

  // ---------------------------------------------------------
  // عرض النتائج
  // ---------------------------------------------------------
  function renderResults() {
    renderStats();
    renderMatchedTab();
    renderOnlyTab("bank");
    renderOnlyTab("ledger");
  }

  function renderStats() {
    const bankTotal = state.records.bank.length;
    const ledgerTotal = state.records.ledger.length;
    const matched = state.matches.length;
    const bankOnlySum = state.bankOnly.reduce((s, r) => s + r.absAmount, 0);
    const ledgerOnlySum = state.ledgerOnly.reduce((s, r) => s + r.absAmount, 0);
    const netDiff = Math.round((bankOnlySum - ledgerOnlySum) * 100) / 100;

    const cards = [
      { label: "حركات البنكي", num: bankTotal, cls: "" },
      { label: "حركات الدفتري", num: ledgerTotal, cls: "" },
      { label: "متطابقة", num: matched, cls: "is-match" },
      { label: "بالبنكي فقط", num: state.bankOnly.length, cls: "is-flag" },
      { label: "بالدفتري فقط", num: state.ledgerOnly.length, cls: "is-flag" },
      { label: "صافي الفرق", num: fmtAmount(netDiff), cls: "is-diff" }
    ];
    document.getElementById("statsStrip").innerHTML = cards.map(c => `
      <div class="stat-card ${c.cls}">
        <div class="stat-num">${c.num}</div>
        <div class="stat-label">${c.label}</div>
      </div>`).join("");
  }

  function renderMatchedTab() {
    const wrap = document.getElementById("tab-matched");
    if (!state.matches.length) {
      wrap.innerHTML = `<div class="empty-state">لا توجد حركات متطابقة بعد.</div>`;
      return;
    }
    let html = `<table class="rtable"><thead><tr>
        <th>تاريخ (بنكي)</th><th>بيان (بنكي)</th><th>مبلغ (بنكي)</th>
        <th>تاريخ (دفتري)</th><th>بيان (دفتري)</th><th>مبلغ (دفتري)</th>
        <th>الفرق</th><th>الحالة</th>
      </tr></thead><tbody>`;
    state.matches.forEach(m => {
      const b = findRec(m.bankId), l = findRec(m.ledgerId);
      const diffOk = Math.abs(m.diff) < 0.005;
      html += `<tr>
        <td>${fmtDate(b.date)}</td><td>${escapeHtml(b.desc)}</td><td class="amount">${fmtAmount(b.amount)}</td>
        <td>${fmtDate(l.date)}</td><td>${escapeHtml(l.desc)}</td><td class="amount">${fmtAmount(l.amount)}</td>
        <td><span class="diff-tag ${diffOk ? "ok" : "warn"}">${fmtAmount(m.diff)}</span></td>
        <td><span class="match-stamp ${m.manual ? "is-manual" : ""}">${m.manual ? "✓ يدوي" : "✓ تلقائي"}</span></td>
      </tr>`;
    });
    html += "</tbody></table>";
    wrap.innerHTML = html;
  }

  function renderOnlyTab(source) {
    const list = source === "bank" ? state.bankOnly : state.ledgerOnly;
    const wrap = document.getElementById("tab-" + (source === "bank" ? "bankonly" : "ledgeronly"));
    if (!list.length) {
      wrap.innerHTML = `<div class="empty-state">لا توجد حركات غير متطابقة هنا. 👏</div>`;
      return;
    }
    let html = `<table class="rtable"><thead><tr>
        <th>اختيار</th><th>التاريخ</th><th>البيان</th><th>المبلغ</th>
      </tr></thead><tbody>`;
    list.forEach(r => {
      const isSelected = state.selection[source] === r.id;
      html += `<tr class="row-selectable ${isSelected ? "row-selected" : ""}" data-id="${r.id}" data-source="${source}">
        <td>${isSelected ? "✓" : "○"}</td>
        <td>${fmtDate(r.date)}</td><td>${escapeHtml(r.desc)}</td><td class="amount">${fmtAmount(r.amount)}</td>
      </tr>`;
    });
    html += "</tbody></table>";
    wrap.innerHTML = html;

    wrap.querySelectorAll("tr.row-selectable").forEach(tr => {
      tr.addEventListener("click", () => {
        const id = tr.dataset.id, src = tr.dataset.source;
        state.selection[src] = state.selection[src] === id ? null : id;
        renderOnlyTab(src);
        updateManualBar();
      });
    });
  }

  function updateManualBar() {
    const bar = document.getElementById("manualBar");
    if (state.selection.bank && state.selection.ledger) bar.classList.remove("hidden");
    else bar.classList.add("hidden");
  }

  document.getElementById("btnManualMatch").addEventListener("click", () => {
    const bankId = state.selection.bank, ledgerId = state.selection.ledger;
    if (!bankId || !ledgerId) return;
    const b = findRec(bankId), l = findRec(ledgerId);
    b.matched = true; l.matched = true;
    state.matches.push({ bankId, ledgerId, manual: true, diff: Math.round((b.absAmount - l.absAmount) * 100) / 100 });
    state.bankOnly = state.bankOnly.filter(r => r.id !== bankId);
    state.ledgerOnly = state.ledgerOnly.filter(r => r.id !== ledgerId);
    state.selection = { bank: null, ledger: null };
    updateManualBar();
    renderResults();
  });

  document.getElementById("btnClearSelection").addEventListener("click", () => {
    state.selection = { bank: null, ledger: null };
    renderOnlyTab("bank"); renderOnlyTab("ledger");
    updateManualBar();
  });

  // ---------------------------------------------------------
  // التبويبات
  // ---------------------------------------------------------
  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach(p => p.classList.add("hidden"));
      btn.classList.add("active");
      const map = { matched: "tab-matched", bankonly: "tab-bankonly", ledgeronly: "tab-ledgeronly" };
      document.getElementById(map[btn.dataset.tab]).classList.remove("hidden");
    });
  });

  // ---------------------------------------------------------
  // التصدير إلى Excel
  // ---------------------------------------------------------
  function exportReport() {
    const wb = XLSX.utils.book_new();

    const bankTotal = state.records.bank.length;
    const ledgerTotal = state.records.ledger.length;
    const bankOnlySum = state.bankOnly.reduce((s, r) => s + r.absAmount, 0);
    const ledgerOnlySum = state.ledgerOnly.reduce((s, r) => s + r.absAmount, 0);

    const summary = [
      ["تقرير التسوية البنكية"],
      ["تاريخ الإنشاء", new Date().toLocaleString("ar-EG")],
      [],
      ["عدد حركات الكشف البنكي", bankTotal],
      ["عدد حركات الكشف الدفتري", ledgerTotal],
      ["عدد الحركات المتطابقة", state.matches.length],
      ["حركات موجودة بالبنكي فقط", state.bankOnly.length, "بمجموع", bankOnlySum.toFixed(2)],
      ["حركات موجودة بالدفتري فقط", state.ledgerOnly.length, "بمجموع", ledgerOnlySum.toFixed(2)],
      ["صافي الفرق", (bankOnlySum - ledgerOnlySum).toFixed(2)]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), "الملخص");

    const matchedRows = [["تاريخ البنكي","بيان البنكي","مبلغ البنكي","تاريخ الدفتري","بيان الدفتري","مبلغ الدفتري","الفرق","نوع المطابقة"]];
    state.matches.forEach(m => {
      const b = findRec(m.bankId), l = findRec(m.ledgerId);
      matchedRows.push([fmtDate(b.date), b.desc, b.amount, fmtDate(l.date), l.desc, l.amount, m.diff, m.manual ? "يدوي" : "تلقائي"]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(matchedRows), "المتطابقة");

    const bankOnlyRows = [["التاريخ","البيان","المبلغ"]].concat(
      state.bankOnly.map(r => [fmtDate(r.date), r.desc, r.amount])
    );
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bankOnlyRows), "بالبنكي فقط");

    const ledgerOnlyRows = [["التاريخ","البيان","المبلغ"]].concat(
      state.ledgerOnly.map(r => [fmtDate(r.date), r.desc, r.amount])
    );
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ledgerOnlyRows), "بالدفتري فقط");

    XLSX.writeFile(wb, "تقرير_التسوية_البنكية.xlsx");
  }

  document.getElementById("btnExport").addEventListener("click", exportReport);

  document.getElementById("btnRestart").addEventListener("click", () => location.reload());

  document.getElementById("btnRun").addEventListener("click", runMatching);

})();
