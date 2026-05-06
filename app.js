// ShiftMD v2.0 - Frontend aplikacija
// @ts-nocheck
document.addEventListener("DOMContentLoaded", function () {
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

  // ===== DINAMIČKO POPUNJAVANJE GODINA =====
  function populateYears() {
    const currentYear = new Date().getFullYear();
    
    // Počisti postojeće opcije osim prve (placeholder)
    while (yearEl.options.length > 1) {
      yearEl.remove(1);
    }
    
    // Dodaj tekuću godinu i naredne 2 godine
    for (let i = 0; i < 3; i++) {
      const year = currentYear + i;
      const option = document.createElement("option");
      option.value = year;
      option.textContent = year;
      yearEl.appendChild(option);
    }
  }
  
  // Popuni godine pri učitavanju
  populateYears();

  function setStatus(type, text) {
    statusEl.className = "status " + type;
    statusEl.textContent = text;
    statusEl.classList.remove("hidden");
  }

  function clearStatus() {
    statusEl.classList.add("hidden");
    statusEl.textContent = "";
  }

  function renderCalendar(year, month) {
    calendarEl.innerHTML = "";
    selectedDates.clear();

    if (!year || !month) {
      calendarEl.innerHTML = '<p style="grid-column: 1 / -1; text-align: center; color: #6b7280;">Izaberite mesec i godinu da biste videli kalendar</p>';
      return;
    }

    const y = parseInt(year);
    const m = parseInt(month);
    
    if (isNaN(y) || isNaN(m) || m < 1 || m > 12) {
      calendarEl.innerHTML = '<p style="grid-column: 1 / -1; text-align: center; color: #dc2626;">Neispravan mesec ili godina.</p>';
      return;
    }
    
    const daysInMonth = new Date(y, m, 0).getDate();
    
    let firstDay = new Date(y, m - 1, 1).getDay();
    if (firstDay === 0) {
      firstDay = 6;
    } else {
      firstDay = firstDay - 1;
    }

    for (let i = 0; i < firstDay; i++) {
      const empty = document.createElement("div");
      empty.className = "empty";
      empty.style.visibility = "hidden";
      calendarEl.appendChild(empty);
    }

    for (let day = 1; day <= daysInMonth; day++) {
      const dateObj = new Date(y, m - 1, day);
      const dateStr =
        String(day).padStart(2, "0") + "-" +
        String(m).padStart(2, "0") + "-" +
        y;

      const div = document.createElement("div");
      div.className = "day";
      div.textContent = day;
      div.dataset.date = dateStr;

      const dayOfWeek = dateObj.getDay();
      if (dayOfWeek === 0 || dayOfWeek === 6) {
        div.classList.add("weekend");
      }

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

  monthEl.addEventListener("change", function () {
    renderCalendar(yearEl.value, monthEl.value);
  });

  yearEl.addEventListener("change", function () {
    renderCalendar(yearEl.value, monthEl.value);
  });

  if (monthEl.value && yearEl.value) {
    renderCalendar(yearEl.value, monthEl.value);
  }

  form.addEventListener("submit", async function (e) {
    e.preventDefault();
    clearStatus();
    resultEl.textContent = "";
    downloadBtn.classList.add("hidden");
    downloadBtn.textContent = "📥 Preuzmi Excel raspored";
    lastGeneratedFile = null;

    if (!excelInput.files || excelInput.files.length === 0) {
      setStatus("error", "Molimo otpremite Excel fajl.");
      return;
    }

    if (selectedDates.size === 0) {
      setStatus("error", "Izaberite bar jedan dan dežurstva.");
      return;
    }

    const dutyStaffCount = parseInt(dutyStaffCountEl.value);
    if (isNaN(dutyStaffCount) || dutyStaffCount < 1) {
      setStatus("error", "Broj dežurnih po danu mora biti najmanje 1.");
      return;
    }

    submitBtn.disabled = true;
    downloadBtn.classList.add("hidden");
    setStatus("loading", "ShiftMD računa raspored i generiše Excel fajl...");

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
      
      const res = await fetch("/generate", {
        method: "POST",
        body: fd
      });

      const endTime = performance.now();
      const duration = ((endTime - startTime) / 1000).toFixed(2);

      if (!res.ok) {
        let errorMsg = "Greška u generisanju rasporeda.";
        try {
          const errorData = await res.json();
          errorMsg = errorData.error || errorMsg;
        } catch (parseErr) {
          errorMsg = await res.text() || errorMsg;
        }
        throw new Error(errorMsg);
      }

      const data = await res.json();

      // Prikaži tekstualni raspored
      let formattedSchedule = "RASPORED DEŽURSTAVA\n";
      formattedSchedule += "═".repeat(50) + "\n";
      formattedSchedule += `ShiftMD v2.0\n`;
      formattedSchedule += `Ustanova: ${data.hospital}\n`;
      formattedSchedule += `Mesec: ${data.month}/${data.year}\n`;
      formattedSchedule += `Broj dežurstava po danu: ${data.dutyStaffCount}\n`;
      formattedSchedule += `Pravilo: ${data.allowLessSpecialists ? 'Obavezan glavni specijalista' : 'Minimum 50% specijalista'}\n`;
      formattedSchedule += `Vreme: ${duration} sekundi\n`;
      formattedSchedule += "═".repeat(50) + "\n\n";

      if (data.schedule && Object.keys(data.schedule).length > 0) {
        for (const [date, slots] of Object.entries(data.schedule)) {
          formattedSchedule += `📅 ${date}\n`;
          formattedSchedule += "─".repeat(40) + "\n";
          
          if (slots && slots.length > 0) {
            slots.forEach((slot, index) => {
              formattedSchedule += `\n  Dežurstvo ${index + 1}:\n`;
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
        formattedSchedule += "⚠️ Nema generisanih dežurstava.\n";
      }

      resultEl.textContent = formattedSchedule;
      
      if (data.excelFile) {
        lastGeneratedFile = data.excelFile;
        downloadBtn.classList.remove("hidden");
        downloadBtn.textContent = `📥 Preuzmi Excel raspored (${data.excelFile})`;
        console.log("ShiftMD: Excel fajl dostupan:", data.excelFile);
      }

      clearStatus();
      
      setTimeout(() => {
        setStatus("success", `✅ ShiftMD: Raspored generisan za ${duration}s!`);
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
      console.log("ShiftMD: Preuzimanje fajla:", lastGeneratedFile);
      window.open(`/download/${encodeURIComponent(lastGeneratedFile)}`, "_blank");
    } else {
      setStatus("error", "Nema generisanog fajla za preuzimanje.");
    }
  });
});