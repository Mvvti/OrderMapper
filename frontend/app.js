let orderInput;
let cennikInput;
let placowkiInput;
let templateInput;
let outputInput;
let dateInput;
let statusLabel;
let logsArea;
let generateButton;
let progressBar;
let buttonStateTimer;
let pdfPdfInput;
let pdfOutputInput;
let pdfStatusLabel;
let pdfLogsArea;
let pdfGenerateButton;
let pdfProgressBar;
let pdfSelectedFilesInfo;
let pdfButtonStateTimer;
let pdfPdfPaths = [];
let startScreen;
let viewMain;
let viewInstrukcja;
let viewInstrukcjaPdf;
let pdfPanel;
let activeView = "start";

function toggleView(viewName) {
    viewMain.style.display = viewName === "main" ? "" : "none";
    viewInstrukcja.style.display = viewName === "instrukcja" ? "" : "none";
    viewInstrukcjaPdf.style.display = viewName === "instrukcja-pdf" ? "" : "none";
    pdfPanel.style.display = viewName === "pdf" ? "" : "none";
    activeView = viewName;
}

function hideStartScreen() {
    startScreen.classList.add("start-screen--hidden");
}

function showStartScreen() {
    startScreen.classList.remove("start-screen--hidden");
    toggleView("start");
}

function clearButtonStateClasses() {
    generateButton.classList.remove("btn--loading", "btn--success", "btn--error");
}

function clearPdfButtonStateClasses() {
    pdfGenerateButton.classList.remove("btn--loading", "btn--success", "btn--error");
}

function getFileName(path) {
    if (!path) return "";
    const normalized = String(path).replace(/\\/g, "/");
    const parts = normalized.split("/");
    return parts[parts.length - 1] || normalized;
}

function updatePdfSelectedFilesInfo() {
    if (!pdfSelectedFilesInfo) return;
    if (!pdfPdfPaths.length) {
        pdfSelectedFilesInfo.textContent = "Brak wybranych plików.";
        return;
    }
    pdfSelectedFilesInfo.textContent = pdfPdfPaths.map(getFileName).join(", ");
}

function updatePdfGenerateAvailability() {
    const hasPdfs = pdfPdfPaths.length > 0;
    const hasOutput = Boolean((pdfOutputInput && pdfOutputInput.value || "").trim());
    pdfGenerateButton.disabled = !(hasPdfs && hasOutput);
}

function setButtonState(state) {
    if (buttonStateTimer) {
        clearTimeout(buttonStateTimer);
        buttonStateTimer = null;
    }

    clearButtonStateClasses();

    if (state === "loading") {
        generateButton.disabled = true;
        generateButton.classList.add("btn--loading");
        generateButton.textContent = "Przetwarzanie...";
        progressBar.classList.add("progress-bar--active");
        return;
    }

    if (state === "success") {
        generateButton.disabled = false;
        generateButton.classList.add("btn--success");
        generateButton.textContent = "Zakończono ✅";
        progressBar.classList.remove("progress-bar--active");
        buttonStateTimer = setTimeout(() => setButtonState("default"), 2000);
        return;
    }

    if (state === "error") {
        generateButton.disabled = false;
        generateButton.classList.add("btn--error");
        generateButton.textContent = "Błąd ❌";
        progressBar.classList.remove("progress-bar--active");
        buttonStateTimer = setTimeout(() => setButtonState("default"), 2000);
        return;
    }

    generateButton.disabled = false;
    generateButton.textContent = "Generuj plik";
    progressBar.classList.remove("progress-bar--active");
}

function setPdfButtonState(state) {
    if (pdfButtonStateTimer) {
        clearTimeout(pdfButtonStateTimer);
        pdfButtonStateTimer = null;
    }

    clearPdfButtonStateClasses();

    if (state === "loading") {
        pdfGenerateButton.disabled = true;
        pdfGenerateButton.classList.add("btn--loading");
        pdfGenerateButton.textContent = "Przetwarzanie...";
        pdfProgressBar.classList.add("progress-bar--active");
        return;
    }

    if (state === "success") {
        pdfGenerateButton.disabled = false;
        pdfGenerateButton.classList.add("btn--success");
        pdfGenerateButton.textContent = "Zakończono ✅";
        pdfProgressBar.classList.remove("progress-bar--active");
        pdfButtonStateTimer = setTimeout(() => setPdfButtonState("default"), 2000);
        return;
    }

    if (state === "error") {
        pdfGenerateButton.disabled = false;
        pdfGenerateButton.classList.add("btn--error");
        pdfGenerateButton.textContent = "Błąd ❌";
        pdfProgressBar.classList.remove("progress-bar--active");
        pdfButtonStateTimer = setTimeout(() => setPdfButtonState("default"), 2000);
        return;
    }

    pdfGenerateButton.disabled = false;
    pdfGenerateButton.textContent = "Generuj plik";
    pdfProgressBar.classList.remove("progress-bar--active");
    updatePdfGenerateAvailability();
}

function appendLog(msg) {
    if (!logsArea) {
        logsArea = document.querySelector("textarea");
    }
    const line = msg == null ? "" : String(msg);
    logsArea.value += `${line}\n`;
    logsArea.scrollTop = logsArea.scrollHeight;
}

function appendPdfLog(msg) {
    const line = msg == null ? "" : String(msg);
    pdfLogsArea.value += `${line}\n`;
    pdfLogsArea.scrollTop = pdfLogsArea.scrollHeight;
}

function onPipelineDone(result) {
    setButtonState("success");
    statusLabel.textContent = "Status: Zakończono ✅";

    const total = result && result.records_count != null ? result.records_count : 0;
    const noPrice = result && result.records_without_price != null ? result.records_without_price : 0;
    const noFacility = result && result.records_without_facility != null ? result.records_without_facility : 0;
    const outputPath = result && result.output_file_path ? result.output_file_path : "";

    appendLog(`Rekordów: ${total} | Bez ceny: ${noPrice} | Bez placówki: ${noFacility} | Plik: ${outputPath}`);
}

function onPipelineError(err) {
    setButtonState("error");
    statusLabel.textContent = "Status: Błąd ❌";
    appendLog("BŁĄD PIPELINE:");
    appendLog(err);
}

function clearPwrInputErrors(fieldMap) {
    Object.values(fieldMap).forEach(input => input.classList.remove("input--error"));
}

function markPwrInputErrors(errors, fieldMap) {
    errors.forEach(error => {
        const field = error && error.field ? String(error.field) : "";
        const input = fieldMap[field];
        if (input) {
            input.classList.add("input--error");
        }
    });
}

function onPdfPipelineDone(result) {
    setPdfButtonState("success");
    pdfStatusLabel.textContent = "Status: Zakończono ✅";

    const total = result && result.records_count != null ? result.records_count : 0;
    const outputPath = result && result.output_file_path ? result.output_file_path : "";
    appendPdfLog(`Rekordów: ${total} | Plik: ${outputPath}`);
}

function onPdfPipelineError(err) {
    setPdfButtonState("error");
    pdfStatusLabel.textContent = "Status: Błąd ❌";
    appendPdfLog("BŁĄD PIPELINE PDF:");
    appendPdfLog(err);
}

window.appendLog = appendLog;
window.onPipelineDone = onPipelineDone;
window.onPipelineError = onPipelineError;
window.appendPdfLog = appendPdfLog;
window.onPdfPipelineDone = onPdfPipelineDone;
window.onPdfPipelineError = onPdfPipelineError;

let _initialized = false;
window.addEventListener("pywebviewready", () => {
    if (_initialized) return;
    _initialized = true;
    const rows = document.querySelectorAll(".form-row");

    startScreen = document.getElementById("start-screen");
    viewMain = document.getElementById("view-main");
    viewInstrukcja = document.getElementById("view-instrukcja");
    viewInstrukcjaPdf = document.getElementById("view-instrukcja-pdf");
    pdfPanel = document.getElementById("pdf-panel");

    orderInput = document.getElementById("order-file");
    cennikInput = document.getElementById("price-file");
    placowkiInput = document.getElementById("facility-file");
    templateInput = document.getElementById("template-file");
    outputInput = document.getElementById("output-folder");
    dateInput = document.getElementById("delivery-date");
    dateInput.value = new Date().toISOString().split("T")[0];
    statusLabel = document.querySelector(".status");
    logsArea = document.querySelector("textarea");
    generateButton = document.querySelector(".btn-primary");
    progressBar = document.getElementById("progress-bar");

    pdfPdfInput = document.getElementById("pdf-pdf-file");
    pdfOutputInput = document.getElementById("pdf-output-folder");
    pdfStatusLabel = document.getElementById("pdf-status");
    pdfLogsArea = document.getElementById("pdf-logs");
    pdfGenerateButton = document.getElementById("pdf-generate-btn");
    pdfProgressBar = document.getElementById("pdf-progress-bar");
    pdfSelectedFilesInfo = document.getElementById("pdf-selected-files");

    const pwrFieldMap = {
        order: orderInput,
        cennik: cennikInput,
        placowki: placowkiInput,
        template: templateInput,
        output: outputInput,
        date: dateInput,
    };

    setButtonState("default");
    setPdfButtonState("default");
    updatePdfSelectedFilesInfo();

    Object.values(pwrFieldMap).forEach(input => {
        input.addEventListener("input", () => {
            input.classList.remove("input--error");
        });
    });

    window.pywebview.api.get_defaults().then(defaults => {
        if (defaults.cennik) cennikInput.value = defaults.cennik;
        if (defaults.placowki) placowkiInput.value = defaults.placowki;
        if (defaults.template) templateInput.value = defaults.template;
    });

    const orderPickButton = rows[0].querySelector("button");
    const cennikPickButton = rows[1].querySelector("button");
    const placowkiPickButton = rows[2].querySelector("button");
    const templatePickButton = rows[3].querySelector("button");

    const instrukcjaButton = document.getElementById("btn-instrukcja");
    const pomocButton = document.getElementById("btn-pomoc");
    const logoButton = document.getElementById("btn-logo");
    const startCardPwr = document.getElementById("start-card-pwr");
    const startCardPdf = document.getElementById("start-card-pdf");
    const backMainButton = document.getElementById("btn-back-main");
    const backPdfButton = document.getElementById("btn-back-pdf");
    const pdfPickPdfButton = document.getElementById("pdf-btn-pick-pdf");
    const pdfPickOutputButton = document.getElementById("pdf-btn-pick-output");

    instrukcjaButton.addEventListener("click", () => {
        if (activeView === "pdf" || activeView === "instrukcja-pdf") {
            toggleView("instrukcja-pdf");
            return;
        }
        toggleView("instrukcja");
    });

    pomocButton.addEventListener("click", async () => {
        await window.pywebview.api.open_url("mailto:higmautomate@higma-service.pl");
    });

    logoButton.addEventListener("click", async () => {
        await window.pywebview.api.open_url("https://higma-service.pl");
    });

    document.getElementById("btn-wróć").addEventListener("click", () => toggleView("main"));
    document.getElementById("btn-wróć-pdf").addEventListener("click", () => toggleView("pdf"));

    startCardPwr.addEventListener("click", () => {
        hideStartScreen();
        toggleView("main");
    });

    startCardPdf.addEventListener("click", () => {
        hideStartScreen();
        toggleView("pdf");
    });

    backMainButton.addEventListener("click", () => {
        showStartScreen();
    });

    backPdfButton.addEventListener("click", () => {
        showStartScreen();
    });

    const outputPickButton = rows[4].querySelector("button");

    orderPickButton.addEventListener("click", async () => {
        const path = await window.pywebview.api.pick_file(["Excel Files (*.xlsx;*.xls)", "*.*"]);
        if (path) {
            orderInput.value = path;
        }
    });

    cennikPickButton.addEventListener("click", async () => {
        const path = await window.pywebview.api.pick_file(["Excel Files (*.xlsx;*.xls)", "*.*"]);
        if (path) {
            cennikInput.value = path;
        }
    });

    placowkiPickButton.addEventListener("click", async () => {
        const path = await window.pywebview.api.pick_file(["Excel Files (*.xlsx;*.xls)", "*.*"]);
        if (path) {
            placowkiInput.value = path;
        }
    });

    templatePickButton.addEventListener("click", async () => {
        const path = await window.pywebview.api.pick_file(["Excel Files (*.xlsx;*.xls)", "*.*"]);
        if (path) {
            templateInput.value = path;
        }
    });

    outputPickButton.addEventListener("click", async () => {
        const path = await window.pywebview.api.pick_folder();
        if (path) {
            outputInput.value = path;
        }
    });

    pdfPickPdfButton.addEventListener("click", async () => {
        const picked = await window.pywebview.api.pick_pdf_files();
        const selected = Array.isArray(picked) ? picked.filter(Boolean) : (picked ? [picked] : []);
        if (!selected.length) {
            return;
        }
        pdfPdfPaths = selected.map(String);
        pdfPdfInput.value = `Wybrano plików: ${pdfPdfPaths.length}`;
        updatePdfSelectedFilesInfo();
        updatePdfGenerateAvailability();
    });

    pdfPickOutputButton.addEventListener("click", async () => {
        const path = await window.pywebview.api.pick_folder();
        if (path) {
            pdfOutputInput.value = path;
            updatePdfGenerateAvailability();
        }
    });

    pdfOutputInput.addEventListener("input", updatePdfGenerateAvailability);

    generateButton.addEventListener("click", async () => {
        const order = orderInput.value.trim();
        const cennik = cennikInput.value.trim();
        const placowki = placowkiInput.value.trim();
        const template = templateInput.value.trim();
        const outputDir = outputInput.value.trim();
        const date = dateInput.value.trim();

        clearPwrInputErrors(pwrFieldMap);

        const validationErrors = await window.pywebview.api.validate_inputs(
            order,
            cennik,
            placowki,
            template,
            outputDir,
            date
        );

        if (Array.isArray(validationErrors) && validationErrors.length > 0) {
            validationErrors.forEach(error => {
                const message = error && error.message ? String(error.message) : "Nieznany błąd walidacji.";
                appendLog(`❌ ${message}`);
            });
            markPwrInputErrors(validationErrors, pwrFieldMap);
            return;
        }

        statusLabel.textContent = "Status: Przetwarzanie...";
        setButtonState("loading");
        appendLog("Start pipeline...");

        try {
            await window.pywebview.api.run_pipeline(order, cennik, placowki, template, outputDir, date);
        } catch (err) {
            onPipelineError(String(err));
        }
    });

    pdfGenerateButton.addEventListener("click", async () => {
        const pdfPaths = pdfPdfPaths.slice();
        const outputDir = pdfOutputInput.value.trim();

        if (!pdfPaths.length || !outputDir) {
            appendPdfLog("BŁĄD: Uzupełnij wszystkie pola");
            return;
        }

        pdfStatusLabel.textContent = "Status: Przetwarzanie...";
        setPdfButtonState("loading");
        appendPdfLog("Start pipeline PDF...");

        try {
            await window.pywebview.api.run_pdf_pipeline(pdfPaths, outputDir);
        } catch (err) {
            onPdfPipelineError(String(err));
        }
    });
});
