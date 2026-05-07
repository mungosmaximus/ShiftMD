// ShiftMD v3.0 - Dvojezična frontend aplikacija
// @ts-nocheck
document.addEventListener("DOMContentLoaded", function () {
  // ===== JEZIČKI SISTEM =====
  let currentLang = localStorage.getItem("shiftmd_lang") || "sr";
  
  function applyLanguage(lang) {
    currentLang = lang;
    localStorage.setItem("shiftmd_lang", lang);
    
    // Ažuriraj sve elemente sa data-key
    document.querySelectorAll("[data-key]").forEach(el => {
      const key = el.getAttribute("data-key");
      const text = t(lang, key);
      if (text) {
        if (el.tagName === "INPUT" && el.getAttribute("data-key-placeholder")) {
          el.placeholder = text;
        } else if (el.tagName === "INPUT" || el.tagName === "SELECT" || el.tagName === "TEXTAREA") {
          // Ne menjaj value za inpute
        } else {
          el.textContent = text;
        }
      }
    });
    
    // Ažuriraj placeholder-e
    document.querySelectorAll("[data-key-placeholder]").forEach(el => {
      const key = el.getAttribute("data-key-placeholder");
      const text = t(lang, key);
      if (text) el.placeholder = text;
    });
    
    // Ažuriraj hint tekst
    const allowLessCheckbox = document.getElementById("allowLessSpecialists");
    const allowLessHint = document.getElementById("allowLessHint");
    if (allowLessHint && allowLessCheckbox) {
      const hintKey = allowLessCheckbox.checked ? "form_allow_less_hint_checked" : "form_allow_less_hint_unchecked";
      allowLessHint.textContent = t(lang, hintKey);
    }
    
    // Ažuriraj mesece u select-u
    updateMonthNames(lang);
    
    // Ažuriraj dane u kalendaru
    updateCalendarHeader(lang);
    
    // Ažuriraj aktivni jezik u navigaciji
    document.querySelectorAll(".lang-switch").forEach(link => {
      link.classList.toggle("active-lang", link.getAttribute("data-lang") === lang);
    });
    
    // Ažuriraj HTML lang atribut
    document.documentElement.lang = lang;
    
    // Ažuriraj title
    document.title = t(lang, "page_title");
    
    // Ponovo renderuj kalendar ako su selektovani mesec i godina
    if (monthEl.value && yearEl.value) {
      renderCalendar(yearEl.value, monthEl.value);
    }
  }
  
  function updateMonthNames(lang) {
    const monthSelect = document.getElementById("month");
    if (!monthSelect) return;
    
    const currentValue = monthSelect.value;
    const months = [
      "", "month_1", "month_2", "month_3", "month_4", "month_5", "month_6",
      "month_7", "month_8", "month_9", "month_10", "month_11", "month_12"
    ];
    
    // Prva opcija je placeholder
    monthSelect.options[0].textContent = t(lang, "form_month_placeholder");
    
    // Meseci 1-12
    for (let i = 1; i <= 12; i++) {
      if (monthSelect.options[i]) {
        monthSelect.options[i].textContent = t(lang, months[i]);
      }
    }
    
    monthSelect.value = currentValue;
  }
  
  function updateCalendarHeader(lang) {
    const header = document.getElementById("calendarHeader");
    if (!header) return;
    
    const days = ["day_mon", "day_tue", "day_wed", "day_thu", "day_fri", "day_sat", "day_sun"];
    const divs = header.querySelectorAll("div");
    
    divs.forEach((div, i) => {
      if (days[i]) div.textContent = t(lang, days[i]);
    });
  }
  
  // Language switcher
  document.querySelectorAll(".lang-switch").forEach(link => {
    link.addEventListener("click", function(e) {
      e.preventDefault();
      const lang = this.getAttribute("data-lang");
      applyLanguage(lang);
    });
  });
  
  // ===== KALENDAR I FORMA =====
  const calendarEl = document.getElementById("calendar");
  const monthEl = document.getElementById("month");
  const yearEl = document.getElementById("year");
  const form = document.getElementById("form");
  const resultEl = document.getElementById("result");
  const downloadBtn = document.getElementById("downloadBtn");
  const statusEl = document.getElementById("status");
  const submitBtn = document.getElementById("submitBtn");
  const excelInput = document.getElementById("excel");
  const dutyStaffCountEl = document.getElementById("dutyStaffCount");
  const allowLessSpecialistsEl = document.getElementById("allowLessSpecialists");

  let selectedDates = new Set();
  let lastGeneratedFile = null;

  function setStatus(type, text) {
    statusEl.className = "status " + type;
    statusEl.textContent = text;
    statusEl.classList.remove("hidden");
  }

  function clearStatus() {
    statusEl.classList.add("hidden");
    statusEl.textContent = "";
  }

  function populateYears() {
    const currentYear = new Date().getFullYear();
    while (yearEl.options.length > 1) yearEl.remove(1);
    for (let i = 0; i < 3; i++) {
      const year = currentYear + i;
      const option = document.createElement("option");
      option.value = year;
      option.textContent = year;
      yearEl.appendChild(option);
    }
  }

  populateYears();

  function renderCalendar(year, month) {
    calendarEl.innerHTML = "";
    selectedDates.clear();

    if (!year || !month) {
      calendarEl.innerHTML = `<p style="grid-column: 1 / -1; text-align: center; color: #6b7280;">${t(currentLang, "calendar_placeholder")}</p>`;
      return;
    }

    const y = parseInt(year);
    const m = parseInt(month);
    
    if (isNaN(y) || isNaN(m) || m < 1 || m > 12) {
      calendarEl.innerHTML = `<p style="grid-column: 1 / -1; text-align: center; color: #dc2626;">${t(currentLang, "calendar_error")}</p>`;
      return;
    }
    
    const daysInMonth = new Date(y, m, 0).getDate();
    let firstDay = new Date(y, m - 1, 1).getDay();
    firstDay = firstDay === 0 ? 6 : firstDay - 1;

    for (let i = 0; i < firstDay; i++) {
      const empty = document.createElement("div");
      empty.className = "empty";
      empty.style.visibility = "hidden";
      calendarEl.appendChild(empty);
    }

    for (let day = 1; day <= daysInMonth; day++) {
      const dateObj = new Date(y, m - 1, day);
      const dateStr = String(day).padStart(2, "0") + "-" + String(m).padStart(2, "0") + "-" + y;
      const div = document.createElement("div");
      div.className = "day";
      div.textContent = day;
      div.dataset.date = dateStr;
      const dayOfWeek = dateObj.getDay();
      if (dayOfWeek === 0 || dayOfWeek === 6) div.classList.add("weekend");
      
      div.addEventListener("click", function () {
        if (selectedDates.has(dateStr)) {
          selectedDates.delete(dateStr);
          div.classList.remove("selected");
        } else {
          selectedDates.add(dateStr);
          div.classList.add("selected");
        }
      });
      calendarEl.appendChild(div);
    }

    const totalCells = firstDay + daysInMonth;
    const remainingCells = totalCells % 7;
    if (remainingCells > 0) {
      for (let i = 0; i < (7 - remainingCells); i++) {
        const empty = document.createElement("div");
        empty.className = "empty";
        empty.style.visibility = "hidden";
        calendarEl.appendChild(empty);
      }
    }
  }

  monthEl.addEventListener("change", () => renderCalendar(yearEl.value, monthEl.value));
  yearEl.addEventListener("change", () => renderCalendar(yearEl.value, monthEl.value));

  if (monthEl.value && yearEl.value) renderCalendar(yearEl.value, monthEl.value);

  // Update hint on checkbox change
  allowLessSpecialistsEl.addEventListener("change", function() {
    const hintKey = this.checked ? "form_allow_less_hint_checked" : "form_allow_less_hint_unchecked";
    document.getElementById("allowLessHint").textContent = t(currentLang, hintKey);
  });

  form.addEventListener("submit", async function (e) {
    e.preventDefault();
    clearStatus();
    resultEl.textContent = "";
    downloadBtn.classList.add("hidden");
    downloadBtn.textContent = t(currentLang, "download_btn");
    lastGeneratedFile = null;

    if (!excelInput.files || excelInput.files.length === 0) {
      setStatus("error", t(currentLang, "status_error_no_excel"));
      return;
    }
    if (selectedDates.size === 0) {
      setStatus("error", t(currentLang, "status_error_no_dates"));
      return;
    }
    const dutyStaffCount = parseInt(dutyStaffCountEl.value);
    if (isNaN(dutyStaffCount) || dutyStaffCount < 1) {
      setStatus("error", t(currentLang, "status_error_no_count"));
      return;
    }

    submitBtn.disabled = true;
    downloadBtn.classList.add("hidden");
    setStatus("loading", t(currentLang, "status_loading"));

    try {
      const fd = new FormData();
      fd.append("hospital", document.getElementById("hospital").value);
      fd.append("month", monthEl.value);
      fd.append("year", yearEl.value);
      fd.append("excel", excelInput.files[0]);
      fd.append("dutyDates", JSON.stringify(Array.from(selectedDates)));
      fd.append("dutyStaffCount", dutyStaffCount);
      fd.append("allowLessSpecialists", allowLessSpecialistsEl.checked ? "true" : "false");

      const startTime = performance.now();
      const res = await fetch("/generate", { method: "POST", body: fd });
      const endTime = performance.now();
      const duration = ((endTime - startTime) / 1000).toFixed(2);

      if (!res.ok) {
        let errorMsg = "Greška u generisanju rasporeda.";
        try { const errorData = await res.json(); errorMsg = errorData.error || errorMsg; } catch (e) {}
        throw new Error(errorMsg);
      }

      const data = await res.json();

      let formattedSchedule = `${t(currentLang, "result_title")}\n`;
      formattedSchedule += "═".repeat(50) + "\n";
      formattedSchedule += `ShiftMD v3.0\n`;
      formattedSchedule += `${t(currentLang, "result_hospital")}: ${data.hospital}\n`;
      formattedSchedule += `${t(currentLang, "result_month")}: ${data.month}/${data.year}\n`;
      formattedSchedule += `${t(currentLang, "result_duty_count")}: ${data.dutyStaffCount}\n`;
      formattedSchedule += `${t(currentLang, "result_rule")}: ${data.allowLessSpecialists ? t(currentLang, "result_rule_relaxed") : t(currentLang, "result_rule_strict")}\n`;
      formattedSchedule += `${t(currentLang, "result_time")}: ${duration} ${t(currentLang, "result_seconds")}\n`;
      formattedSchedule += "═".repeat(50) + "\n\n";

      if (data.schedule && Object.keys(data.schedule).length > 0) {
        for (const [date, slots] of Object.entries(data.schedule)) {
          formattedSchedule += `📅 ${date}\n`;
          formattedSchedule += "─".repeat(40) + "\n";
          if (slots && slots.length > 0) {
            slots.forEach((slot, index) => {
              formattedSchedule += `\n  ${t(currentLang, "excel_duty_slot")} ${index + 1}:\n`;
              if (slot.persons && slot.persons.length > 0) {
                slot.persons.forEach((person) => {
                  const prefix = person.isChief ? "★ " : "• ";
                  formattedSchedule += `  ${prefix}${person.name} - ${person.role} (${person.share})\n`;
                });
              }
            });
          }
          formattedSchedule += "\n";
        }
      } else {
        formattedSchedule += t(currentLang, "result_no_data") + "\n";
      }

      resultEl.textContent = formattedSchedule;
      
      if (data.excelFile) {
        lastGeneratedFile = data.excelFile;
        downloadBtn.classList.remove("hidden");
        downloadBtn.textContent = t(currentLang, "download_btn_with_name").replace("{filename}", data.excelFile);
      }

      clearStatus();
      setTimeout(() => {
        setStatus("success", t(currentLang, "status_success").replace("{duration}", duration));
      }, 100);

    } catch (err) {
      console.error("ShiftMD greška:", err);
      setStatus("error", err.message);
    } finally {
      submitBtn.disabled = false;
    }
  });

  downloadBtn.addEventListener("click", function () {
    if (lastGeneratedFile) {
      window.open(`/download/${encodeURIComponent(lastGeneratedFile)}`, "_blank");
    } else {
      setStatus("error", t(currentLang, "status_error_no_file"));
    }
  });

  // Inicijalizuj jezik
  applyLanguage(currentLang);
});