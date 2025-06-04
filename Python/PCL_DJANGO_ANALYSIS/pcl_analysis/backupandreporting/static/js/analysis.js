document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll("table").forEach((table) => {
    table.classList.add("table", "table-bordered", "table-hover");
  });
});