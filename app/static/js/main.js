document.addEventListener("DOMContentLoaded", () => {
    const rows = Array.from(document.querySelectorAll(".student-row"));
    const showMoreBtn = document.getElementById("showMoreBtn");
    const rowsInfo = document.getElementById("rowsInfo");

    if (!rows.length || !showMoreBtn) {
        return;
    }

    const step = 50;
    let visibleCount = step;

    function renderRows() {
        rows.forEach((row, index) => {
            row.style.display = index < visibleCount ? "" : "none";
        });

        rowsInfo.textContent = `Показано ${Math.min(visibleCount, rows.length)} из ${rows.length}`;

        if (visibleCount >= rows.length) {
            showMoreBtn.style.display = "none";
        }
    }

    showMoreBtn.addEventListener("click", () => {
        visibleCount += step;
        renderRows();
    });

    renderRows();
});
