const form = document.getElementById("reportForm");
const progress = document.getElementById("progress");
const fileInputs = document.querySelectorAll('input[type="file"]');

function updateFileLabel(input) {
    const label = document.querySelector(`[data-file-label="${input.id}"]`);
    if (!label) return;
    label.textContent = input.files && input.files.length ? input.files[0].name : "No file selected";
}

fileInputs.forEach((input) => {
    const zone = input.closest(".drop-zone");

    input.addEventListener("change", () => updateFileLabel(input));

    zone.addEventListener("dragover", (event) => {
        event.preventDefault();
        zone.classList.add("drag-over");
    });

    zone.addEventListener("dragleave", () => {
        zone.classList.remove("drag-over");
    });

    zone.addEventListener("drop", (event) => {
        event.preventDefault();
        zone.classList.remove("drag-over");
        if (event.dataTransfer.files.length) {
            input.files = event.dataTransfer.files;
            updateFileLabel(input);
        }
    });
});

form.addEventListener("reset", () => {
    window.setTimeout(() => fileInputs.forEach(updateFileLabel), 0);
});

form.addEventListener("submit", (event) => {
    const missing = Array.from(fileInputs).some((input) => !input.files || !input.files.length);
    const invalid = Array.from(fileInputs).some((input) => {
        if (!input.files || !input.files.length) return false;
        return !input.files[0].name.toLowerCase().endsWith(".csv");
    });

    if (missing || invalid) {
        event.preventDefault();
        alert(missing ? "Please choose both CSV files." : "Only CSV files are allowed.");
        return;
    }

    progress.hidden = false;
    form.querySelector(".primary-btn").disabled = true;
    form.querySelector(".primary-btn").textContent = "Generating...";
});
