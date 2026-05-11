// ShiftMD v3.6 - Dvojezična frontend aplikacija sa dva tipa dežurstava
// @ts-nocheck
document.addEventListener("DOMContentLoaded", function () {
  let currentLang = localStorage.getItem("shiftmd_lang") || "sr";
  
  function applyLanguage(lang) {
    currentLang = lang;
    localStorage.setItem("shiftmd_lang", lang);
    
    document.querySelectorAll("[data-key]").forEach(el => {
      const key = el.getAttribute("data-key");
      const text = t(lang, key);
      if (text) {
        if (el.tagName === "INPUT" && el.getAttribute("data-key-placeholder")) {
          el.placeholder = text;
        } else if (el.tagName !== "INPUT" && el.tagName !== "SELECT" && el.tagName !== "TEXTAREA") {
          el.textContent = text;
        }
      }
    });
    
    document.querySelectorAll("[data-key-placeholder]").forEach(el => {
      const key = el.getAttribute("data-key-placeholder");
      const text = t(lang, key);
      if (text) el.placeholder = text;
    });
    
    document.querySelectorAll(".allowLessHint").forEach(hint => {
      const checkbox = document.getElementById(hint.id.replace("Hint", ""));
      if (checkbox) {
        hint.textContent = t(lang, checkbox.checked ? "form_allow_less_hint_checked" : "form_allow_less_hint_unchecked");
      }
    });
    
    document.querySelectorAll(".rankedHint").forEach(hint => {
      const checkbox = document.getElementById(hint.id.replace("Hint", ""));
      if (checkbox) {
        hint.textContent = t(lang, checkbox.checked ? "form_ranked_hint_checked" : "form_ranked_hint_unchecked");
      }
    });
    
    updateMonthNames(lang);
    updateCalendarHeader("calendarHeader1", lang);
    updateCalendarHeader("calendarHeader2", lang);
    
    document.querySelectorAll(".lang-switch").forEach(link => {
      link.classList.toggle("active-lang", link.getAttribute("data-lang") === lang);
    });
    
    document.documentElement.lang = lang;
    document.title = t(lang, "page_title");
    
    if (monthEl.value && yearEl.value) {
      renderCalendar(yearEl.value, monthEl.value, calendar1El, selectedDates1, "type1");
      if (enableType2El.checked) {
        renderCalendar(yearEl.value, monthEl.value, calendar2El, selectedDates2, "type2");
      }
    }
  }
  
  function updateMonthNames(lang) {
    const monthSelect = document.getElementById("month");
    if (!monthSelect) return;
    const currentValue = monthSelect.value;
    const months = ["", "month_1", "month_2", "month_3", "month_4", "month_5", "month_6", "month_7", "month_8", "month_9", "month_10", "month_11", "month_12"];
    monthSelect.options[0].textContent = t(lang, "form_month_placeholder");
    for (let i = 1; i <= 12; i++) {
      if (monthSelect.options[i]) monthSelect.options[i].textContent = t(lang, months[i]);
    }
    monthSelect.value = currentValue;
  }
  
  function updateCalendarHeader(headerId, lang) {
    const header = document.getElementById(headerId);
    if (!header) return;
    const days = ["day_mon", "day_tue", "day_wed", "day_thu", "day_fri", "day_sat", "day_sun"];
    const divs = header.querySelectorAll("div");
    divs.forEach((div, i) => { if (days[i]) div.textContent = t(lang, days[i]); });
  }
  
  document.querySelectorAll(".lang-switch").forEach(link => {
    link.addEventListener("click", function(e) {
      e.preventDefault();
      applyLanguage(this.getAttribute("data-lang"));
    });
  });
  
  // ===== DOM ELEMENTI =====
  const calendar1El = document.getElementById("calendar1");
  const calendar2El = document.getElementById("calendar2");
  const monthEl = document.getElementById("month");
  const yearEl = document.getElementById("year");
  const form = document.getElementById("form");
  const resultEl = document.getElementById("result");
  const downloadBtn = document.getElementById("downloadBtn");
  const statusEl = document.getElementById("status");
  const submitBtn = document.getElementById("submitBtn");
  const excelInput = document.getElementById("excel");
  
  const dutyStaffCount1El = document.getElementById("dutyStaffCount1");
  const allowLessSpecialists1El = document.getElementById("allowLessSpecialists1");
  const useRankedDuties1El = document.getElementById("useRankedDuties1");
  
  const enableType2El = document.getElementById("enableType2");
  const type2Section = document.getElementById("type2Section");
  const customName2El = document.getElementById("customName2");
  
  const dutyStaffCount2El = document.getElementById("dutyStaffCount2");
  const allowLessSpecialists2El = document.getElementById("allowLessSpecialists2");
  const useRankedDuties2El = document.getElementById("useRankedDuties2");

  let selectedDates1 = new Set();
  let selectedDates2 = new Set();
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

  function renderCalendar(year, month, calendarEl, selectedDatesSet, type) {
    calendarEl.innerHTML = "";
    selectedDatesSet.clear();

    if (!year || !month) {
      const placeholderKey = type === "type1" ? "calendar_placeholder_type1" : "calendar_placeholder_type2";
      calendarEl.innerHTML = `<p style="grid-column:1/-1;text-align:center;color:#6b7280">${t(currentLang, placeholderKey)}</p>`;
      return;
    }

    const y = parseInt(year);
    const m = parseInt(month);
    if (isNaN(y) || isNaN(m) || m < 1 || m > 12) {
      calendarEl.innerHTML = `<p style="grid-column:1/-1;text-align:center;color:#dc2626">${t(currentLang, "calendar_error")}</p>`;
      return;
    }

    const daysInMonth = new Date(y, m, 0).getDate();
    let firstDay = new Date(y, m - 1, 1).getDay();
    firstDay = firstDay === 0 ? 6 : firstDay - 1;

    for (let i = 0; i < firstDay; i++) {
      const empty = document.createElement("div");
      empty.className = "empty"; empty.style.visibility = "hidden";
      calendarEl.appendChild(empty);
    }

    for (let day = 1; day <= daysInMonth; day++) {
      const dateObj = new Date(y, m - 1, day);
      const dateStr = String(day).padStart(2, "0") + "-" + String(m).padStart(2, "0") + "-" + y;
      const div = document.createElement("div");
      div.className = "day"; div.textContent = day; div.dataset.date = dateStr;
      if (dateObj.getDay() === 0 || dateObj.getDay() === 6) div.classList.add("weekend");
      
      div.addEventListener("click", function () {
        const otherSet = type === "type1" ? selectedDates2 : selectedDates1;
        if (otherSet.has(dateStr)) {
          setStatus("error", t(currentLang, "status_error_overlap"));
          return;
        }
        
        if (selectedDatesSet.has(dateStr)) {
          selectedDatesSet.delete(dateStr);
          div.classList.remove("selected");
        } else {
          selectedDatesSet.add(dateStr);
          div.classList.add("selected");
        }
        clearStatus();
      });
      
      calendarEl.appendChild(div);
    }

    const totalCells = firstDay + daysInMonth;
    const remainingCells = totalCells % 7;
    if (remainingCells > 0) {
      for (let i = 0; i < (7 - remainingCells); i++) {
        const empty = document.createElement("div");
        empty.className = "empty"; empty.style.visibility = "hidden";
        calendarEl.appendChild(empty);
      }
    }
  }

  monthEl.addEventListener("change", () => {
    renderCalendar(yearEl.value, monthEl.value, calendar1El, selectedDates1, "type1");
    if (enableType2El.checked) {
      renderCalendar(yearEl.value, monthEl.value, calendar2El, selectedDates2, "type2");
    }
  });
  
  yearEl.addEventListener("change", () => {
    renderCalendar(yearEl.value, monthEl.value, calendar1El, selectedDates1, "type1");
    if (enableType2El.checked) {
      renderCalendar(yearEl.value, monthEl.value, calendar2El, selectedDates2, "type2");
    }
  });

  if (monthEl.value && yearEl.value) {
    renderCalendar(yearEl.value, monthEl.value, calendar1El, selectedDates1, "type1");
  }

  enableType2El.addEventListener("change", function() {
    type2Section.style.display = this.checked ? "block" : "none";
    customName2El.disabled = !this.checked;
    if (this.checked && monthEl.value && yearEl.value) {
      renderCalendar(yearEl.value, monthEl.value, calendar2El, selectedDates2, "type2");
    }
  });

  allowLessSpecialists1El.addEventListener("change", function() {
    document.getElementById("allowLessHint1").textContent = t(currentLang, this.checked ? "form_allow_less_hint_checked" : "form_allow_less_hint_unchecked");
  });
  useRankedDuties1El.addEventListener("change", function() {
    document.getElementById("rankedHint1").textContent = t(currentLang, this.checked ? "form_ranked_hint_checked" : "form_ranked_hint_unchecked");
  });
  allowLessSpecialists2El.addEventListener("change", function() {
    document.getElementById("allowLessHint2").textContent = t(currentLang, this.checked ? "form_allow_less_hint_checked" : "form_allow_less_hint_unchecked");
  });
  useRankedDuties2El.addEventListener("change", function() {
    document.getElementById("rankedHint2").textContent = t(currentLang, this.checked ? "form_ranked_hint_checked" : "form_ranked_hint_unchecked");
  });

  // ===== FORM SUBMIT =====
  form.addEventListener("submit", async function (e) {
    e.preventDefault();
    clearStatus();
    resultEl.textContent = "";
    downloadBtn.classList.add("hidden");
    downloadBtn.textContent = t(currentLang, "download_btn");
    lastGeneratedFile = null;

    if (!excelInput.files || excelInput.files.length === 0) { setStatus("error", t(currentLang, "status_error_no_excel")); return; }
    if (selectedDates1.size === 0) { setStatus("error", t(currentLang, "status_error_no_dates1")); return; }
    if (enableType2El.checked && selectedDates2.size === 0) { setStatus("error", t(currentLang, "status_error_no_dates2")); return; }
    
    if (enableType2El.checked) {
      for (const date of selectedDates1) {
        if (selectedDates2.has(date)) {
          setStatus("error", t(currentLang, "status_error_overlap"));
          return;
        }
      }
    }

    const dutyStaffCount1 = parseInt(dutyStaffCount1El.value);
    if (isNaN(dutyStaffCount1) || dutyStaffCount1 < 1) { setStatus("error", t(currentLang, "status_error_no_count")); return; }

    submitBtn.disabled = true;
    downloadBtn.classList.add("hidden");
    setStatus("loading", t(currentLang, "status_loading"));

    try {
      const fd = new FormData();
      fd.append("hospital", document.getElementById("hospital").value);
      fd.append("month", monthEl.value);
      fd.append("year", yearEl.value);
      fd.append("excel", excelInput.files[0]);
      
      fd.append("dutyDates1", JSON.stringify(Array.from(selectedDates1)));
      fd.append("dutyStaffCount1", dutyStaffCount1);
      fd.append("allowLessSpecialists1", allowLessSpecialists1El.checked ? "true" : "false");
      fd.append("useRankedDuties1", useRankedDuties1El.checked ? "true" : "false");
      
      fd.append("enableType2", enableType2El.checked ? "true" : "false");
      
      if (enableType2El.checked) {
        fd.append("dutyDates2", JSON.stringify(Array.from(selectedDates2)));
        fd.append("dutyStaffCount2", dutyStaffCount2El.value || "1");
        fd.append("allowLessSpecialists2", allowLessSpecialists2El.checked ? "true" : "false");
        fd.append("useRankedDuties2", useRankedDuties2El.checked ? "true" : "false");
      }
      
      fd.append("customName1", document.getElementById("customName1").value || "");
      fd.append("customName2", document.getElementById("customName2").value || "");

      const startTime = performance.now();
      const res = await fetch("/generate", { method: "POST", body: fd });
      const endTime = performance.now();
      const duration = ((endTime - startTime) / 1000).toFixed(2);

      if (!res.ok) {
        let errorMsg = "Greška u generisanju rasporeda.";
        try { const errorData = await res.json(); errorMsg = errorData.error || errorMsg; } catch (ex) {}
        throw new Error(errorMsg);
      }

      const data = await res.json();

      let formattedSchedule = `${t(currentLang, "result_title")}\n`;
      formattedSchedule += "═".repeat(50) + "\n";
      formattedSchedule += `ShiftMD v3.6\n`;
      formattedSchedule += `${t(currentLang, "result_hospital")}: ${data.hospital}\n`;
      formattedSchedule += `${t(currentLang, "result_month")}: ${data.month}/${data.year}\n`;
      
      // Prikaži info o tipovima
      if (data.type1Name) {
        formattedSchedule += `\n▸ ${data.type1Name}: ${data.type1Count || 0} dana, ${data.staffCount1 || 0} dežurstava/dan`;
        if (data.useRanked1) formattedSchedule += ` (rangirani timovi)`;
        formattedSchedule += `\n`;
      }
      if (data.type2Name && data.hasType2) {
        formattedSchedule += `▸ ${data.type2Name}: ${data.type2Count || 0} dana, ${data.staffCount2 || 0} dežurstava/dan`;
        if (data.useRanked2) formattedSchedule += ` (rangirani timovi)`;
        formattedSchedule += `\n`;
      }
      
      formattedSchedule += `\n${t(currentLang, "result_time")}: ${duration} ${t(currentLang, "result_seconds")}\n`;
      formattedSchedule += "═".repeat(50) + "\n\n";

      if (data.schedule && Object.keys(data.schedule).length > 0) {
        // Grupiši datume po tipu
        const datesByType = {};
        if (data.dutyTypes) {
          for (const [date, typeLabel] of Object.entries(data.dutyTypes)) {
            if (!datesByType[typeLabel]) datesByType[typeLabel] = [];
            datesByType[typeLabel].push(date);
          }
        } else {
          datesByType["Dežurstva"] = Object.keys(data.schedule);
        }
        
        for (const [typeLabel, typeDates] of Object.entries(datesByType)) {
          formattedSchedule += `\n▸ ${typeLabel}\n`;
          
          typeDates.sort((a, b) => {
            const [da, ma, ya] = a.split("-").map(Number);
            const [db, mb, yb] = b.split("-").map(Number);
            return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
          });
          
          for (const date of typeDates) {
            formattedSchedule += `\n📅 ${date}\n`;
            formattedSchedule += "─".repeat(40) + "\n";
            const slots = data.schedule[date];
            if (slots && slots.length > 0) {
              slots.forEach((slot, index) => {
                formattedSchedule += `\n  Dežurstvo ${index + 1}:\n`;
                if (slot.persons && slot.persons.length > 0) {
                  slot.persons.forEach((person) => {
                    const prefix = person.isChief ? "★ " : "• ";
                    const rankInfo = person.rank ? ` [Rang ${person.rank}]` : "";
                    formattedSchedule += `  ${prefix}${person.name} - ${person.role} (${person.share})${rankInfo}\n`;
                  });
                }
              });
            }
          }
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
    if (lastGeneratedFile) window.open(`/download/${encodeURIComponent(lastGeneratedFile)}`, "_blank");
    else setStatus("error", t(currentLang, "status_error_no_file"));
  });

  applyLanguage(currentLang);
});