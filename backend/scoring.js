const { isWeekend } = require("./utils");

module.exports = function score(e, date) {
  let s = 0;
  s -= e.assignedDates.length * 5;
  if (isWeekend(date)) s -= e.weekendCount * 3;
  s -= e.seniority;
  return s;
};
``