// @ts-nocheck
document.addEventListener("DOMContentLoaded", function () {
  const calendarEl = document.getElementById("calendar");
  const monthEl = document.getElementById("month");
  const yearEl = document.getElementById("year");
  const form = document.getElementById("form");
  const resultEl = document.getElementById("result");
  const pdfBtn = document.getElementById("pdfBtn");
  const statusEl = document.getElementById("status");
  const submitBtn = document.getElementById("submitBtn");

  let selectedDates = new Set();

  function setStatus(type, text) {
    statusEl.className = "status " + type;
    statusEl.textContent = text;
    statusEl.classList.remove("hidden");
  }

  function clearStatus() {
    statusEl.classList.add("hidden");
  }

  function renderCalendar(year, month) {
    calendarEl.innerHTML = "";
    selectedDates.clear();

    if (!year || !month) return;

    const daysInMonth = new Date(year, month, 0).getDate();

    let firstDay = new Date(year, month - 1, 1).getDay();
    firstDay = firstDay === 0 ? 6 : firstDay - 1;

    for (let i = 0; i < firstDay; i++) {
      const empty = document.createElement("div");
      empty.className = "empty";
      calendarEl.appendChild(empty);
    }

    for (let day = 1; day <= daysInMonth; day++) {
      const date = new Date(year, month - 1, day);
      const dateStr =
        String(day).padStart(2, "0") + "-" +
        String(month).padStart(2, "0") + "-" +
        year;

      const div = document.createElement("div");
      div.className = "day";
      div.textContent = day;

      if (date.getDay() === 0 || date.getDay() === 6) {
        div.classList.add("weekend");
      }

      div.onclick = function () {
        if (selectedDates.has(dateStr)) {
          selectedDates.delete(dateStr);
          div.classList.remove("selected");
        } else {
          selectedDates.add(dateStr);
          div.classList.add("selected");
        }
      };

      calendarEl.appendChild(div);
    }
  }

  monthEl.onchange = function () {
    renderCalendar(yearEl.value, monthEl.value);
  };

  yearEl.onchange = function () {
    renderCalendar(yearEl.value, monthEl.value);
  };

  form.onsubmit = async function (e) {
    e.preventDefault();
    clearStatus();
    resultEl.textContent = "";
    pdfBtn.classList.add("hidden");

    if (selectedDates.size === 0) {
      setStatus("error", "Izaberite bar jedan dan dezurstva.");
      return;
    }

    submitBtn.disabled = true;
    setStatus("loading", "Racunanje rasporeda...");

    try {
      const fd = new FormData(form);
      fd.append("dutyDates", JSON.stringify(Array.from(selectedDates)));

      const res = await fetch("/generate", {
        method: "POST",
        body: fd
      });

      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.error || "Greska u generisanju.");
      }

      resultEl.textContent =
        JSON.stringify(data.schedule, null, 2);

      clearStatus();
      pdfBtn.classList.remove("hidden");

    } catch (err) {
      setStatus("error", err.message);
    } finally {
      submitBtn.disabled = false;
    }
  };

  pdfBtn.onclick = function () {
    window.open("/pdf");
  };
});