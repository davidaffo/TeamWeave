(() => {
  const fileInput = document.getElementById("file-input");
  const importCodeInput = document.getElementById("import-code");
  const importCodeBtn = document.getElementById("import-code-btn");
  const datasetBox = document.getElementById("dataset-box");
  const mainTabs = Array.from(document.querySelectorAll(".main-tab"));
  const analysisSection = document.getElementById("analysis-section");
  const generatorSection = document.getElementById("generator-section");
  const twFormTitle = document.getElementById("tw-form-title");
  const twFormLimit = document.getElementById("tw-form-limit");
  const twFormPlayers = document.getElementById("tw-form-players");
  const twFormQuestions = document.getElementById("tw-form-questions");
  const twQuestionSetSelect = document.getElementById("tw-question-set-select");
  const twQuestionSetLoad = document.getElementById("tw-question-set-load");
  const twQuestionFileInput = document.getElementById("tw-question-file-input");
  const twQuestionSetStatus = document.getElementById("tw-question-set-status");
  const twFormCategories = document.getElementById("tw-form-categories");
  const twFormConvert = document.getElementById("tw-form-convert");
  const twFormCopy = document.getElementById("tw-form-copy");
  const twFormDownload = document.getElementById("tw-form-download");
  const twFormDownloadQuestions = document.getElementById("tw-form-download-questions");
  const twFormOutput = document.getElementById("tw-form-output");
  const twFormPreview = document.getElementById("tw-form-preview");
  const datasetSelect = document.getElementById("dataset-select");
  const datasetManager = document.getElementById("dataset-manager");
  const deleteDatasetBtn = document.getElementById("delete-dataset-btn");
  const compareRow = document.getElementById("compare-row");
  const compareNote = document.getElementById("compare-note");
  const compareSelectA = document.getElementById("compare-a");
  const compareSelectB = document.getElementById("compare-b");
  const statusEl = document.getElementById("status");
  const summaryEl = document.getElementById("summary");
  const viewsEl = document.getElementById("views");
  const viewContainer = document.getElementById("view-container");
  const exportLinkBtn = document.getElementById("export-link-btn");
  const exportCsvBtn = document.getElementById("export-csv-btn");
  const exportHtmlBtn = document.getElementById("export-html-btn");
  const exportBox = document.getElementById("export-box");
  const exportLinkInput = document.getElementById("export-link");
  const copyLinkBtn = document.getElementById("copy-link-btn");
  const tabs = Array.from(document.querySelectorAll(".tab"));

  const CATEGORY_MAP = new Map([
    ["Immagina di andare in trasferta. Chi vorresti in camera con te? (Seleziona 3)", "Sociali"],
    ["Immagina di andare in trasferta. Chi non vorresti in camera con te? (Seleziona 3)", "Sociali"],
    ["Chi preferiresti che l'allenatore avesse scelto come capitano della squadra? (Seleziona 3)", "Attitudinali"],
    ["Chi non sarebbe adatta a fare il capitano? (Seleziona 3)", "Attitudinali"],
    ["Immagina di giocare una partita decisiva. Chi vorresti con te in campo? (Seleziona 3)", "Tecniche"],
    ["Chi non vorresti con te in campo in una partita decisiva? (Seleziona 3)", "Tecniche"],
    ["Sei al tie-break (set dello spareggio). Chi secondo te dovrebbe giocare? (Seleziona 3)", "Tecniche"],
    ["Chi non dovrebbe giocare al tie-break? (Seleziona 3)", "Tecniche"],
    ["Durante l'allenamento, in un esercizio a coppie, chi vorresti con te? (Seleziona 3)", "Attitudinali"],
    ["Durante l'allenamento, chi non vorresti come compagna in un esercizio a coppie? (Seleziona 3)", "Attitudinali"],
    ["Durante l'allenamento, chi secondo te dà sempre il massimo? (Seleziona 3)", "Attitudinali"],
    ["Chi secondo te non dà sempre il massimo durante l'allenamento? (Seleziona 3)", "Attitudinali"],
    ["In un momento decisivo della partita, chi vorresti andasse a battere? (Seleziona 3)", "Tecniche"],
    ["In un momento decisivo, chi preferiresti non mandare a battere? (Seleziona 3)", "Tecniche"],
    ["Sei l'alzatore, sul 24-23 per la tua squadra. A chi alzi il pallone? (Seleziona 3)", "Tecniche"],
    ["Chi non vorresti che attaccasse sul 24-23? (Seleziona 3)", "Tecniche"],
    ["Se una sera sei sola, con chi ti piacerebbe uscire? (Seleziona 3)", "Sociali"],
    ["Con chi non ti andrebbe di uscire una sera? (Seleziona 3)", "Sociali"],
    ["Devi creare una squadra forte per vincere con altre 3 persone. Chi scegli? (Seleziona 3)", "Tecniche"],
    ["Chi non sceglieresti per creare una squadra forte? (Seleziona 3)", "Tecniche"],
    ["Chi tra le tue compagne può dare il maggior contributo in un momento critico della partita? (Seleziona 3)", "Attitudinali"],
    ["Chi non pensi possa contribuire in un momento critico della partita? (Seleziona 3)", "Attitudinali"],
    ["Hai un problema e vuoi confrontarti con delle compagne di squadra. Con chi ne parli? (Seleziona 3)", "Sociali"],
    ["Con chi non ti confronteresti per parlare di un problema? (Seleziona 3)", "Sociali"],
    ["La squadra ha dei problemi. Chi potrebbe farsi portavoce presso l’allenatore o i dirigenti? (Seleziona 3)", "Attitudinali"],
    ["Chi non sarebbe adatta a farsi portavoce della squadra? (Seleziona 3)", "Attitudinali"],
    ["Chi sono secondo te le più forti della squadra? (Seleziona 3)", "Tecniche"],
    ["Chi sono secondo te le più scarse della squadra? (Seleziona 3)", "Tecniche"],
    ["Chi sono le compagne più simpatiche? (Seleziona 3)", "Sociali"],
    ["Chi sono le compagne più antipatiche? (Seleziona 3)", "Sociali"]
  ]);

  const CATEGORIES = ["Tecniche", "Attitudinali", "Sociali"];
  const CATEGORY_MAP_NORMALIZED = new Map(
    Array.from(CATEGORY_MAP.entries()).map(([question, category]) => [
      normalizeQuestion(question),
      category
    ])
  );

  const playerSelections = new WeakMap();
  const renderers = new Map();
  const datasets = [];
  let currentAnalysis = null;
  let currentDatasetId = null;
  let datasetCounter = 1;
  let visibleDatasetIds = new Set();
  let hasVisibleState = false;
  let latestTeamweaveScript = "";
  let isRestoringDatasets = false;
  const DATASETS_KEY = "teamweave:datasets";
  const DATASET_STATE_KEY = "teamweave:datasetState";
  const QUESTION_SET_CHOICE_KEY = "teamweave:generatorQuestionSet";
  const VIEW_KEY = "teamweave:lastView";
  const LINK_PARAM = "data";
  const QUESTION_SETS_DIR = "question-sets/";
  const DEFAULT_QUESTION_SETS = [
    { name: "Set Domande 1", url: "question-sets/Set%20Domande%201.txt" }
  ];
  const PALETTE = {
    posMatrix: ["#ffffff", "#00a933"],
    posTotals: ["#ffffff", "#81d41a"],
    negMatrix: ["#ffffff", "#ff0000"],
    negTotals: ["#ffffff", "#800080"],
    nonRic: ["#ffffff", "#ff8000"],
    network: ["#d73027", "#fee08b", "#1a9850"]
  };
  const IS_EXPORT_MODE = Boolean(window.TEAMWEAVE_EXPORT_MODE);
  const EXPORT_ORDER = Array.isArray(window.TEAMWEAVE_EXPORT_ORDER)
    ? window.TEAMWEAVE_EXPORT_ORDER.slice()
    : null;
  const INLINE_STYLES = "__INLINE_STYLES__";
  const INLINE_APP = "__INLINE_APP__";


  if ("serviceWorker" in navigator && !window.TEAMWEAVE_DISABLE_SW) {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("sw.js").then((registration) => {
        registration.update().catch(() => undefined);
        registration.addEventListener("updatefound", () => {
          const newWorker = registration.installing;
          if (!newWorker) {
            return;
          }
          newWorker.addEventListener("statechange", () => {
            if (newWorker.state === "installed" && navigator.serviceWorker.controller) {
              newWorker.postMessage({ type: "SKIP_WAITING" });
            }
          });
        });
      }).catch(() => undefined);

      navigator.serviceWorker.addEventListener("controllerchange", () => {
        window.location.reload();
      });
    });
  }

  fileInput.addEventListener("change", (event) => {
    const files = Array.from(event.target.files || []);
    if (!files.length) {
      return;
    }
    readFiles(files);
    event.target.value = "";
  });

  datasetSelect.addEventListener("change", () => {
    setCurrentDataset(datasetSelect.value);
  });

  deleteDatasetBtn.addEventListener("click", () => {
    const dataset = getCurrentDataset();
    if (!dataset) {
      return;
    }
    removeDataset(dataset.id);
  });

  compareSelectA.addEventListener("change", () => {
    ensureCompareSelection();
    const active = getActiveView() || "responses";
    renderView(active);
    saveDatasetState();
    updateDatasetManager();
  });

  compareSelectB.addEventListener("change", () => {
    ensureCompareSelection();
    const active = getActiveView() || "responses";
    renderView(active);
    saveDatasetState();
    updateDatasetManager();
  });

  importCodeBtn.addEventListener("click", async () => {
    const value = importCodeInput.value;
    if (!value || !value.trim()) {
      setStatus("Incolla un codice prima di importare.", true);
      return;
    }
    const decoded = await decodeLinkPayload(extractCode(value));
    let csv = null;
    if (decoded && decoded.kind === "bytes") {
      csv = unpackCsvFromBinary(decoded.value);
    } else if (decoded && decoded.kind === "text") {
      csv = decodePackedCsv(decoded.value);
    }
    if (!csv) {
      setStatus("Codice non valido.", true);
      return;
    }
    loadCsv(csv, { name: `Codice ${datasetCounter}`, persist: true });
    datasetCounter += 1;
  });

  exportLinkBtn.addEventListener("click", async () => {
    const dataset = getCurrentDataset();
    if (!dataset) {
      setStatus("Carica un CSV prima di creare il codice.", true);
      return;
    }
    let payload = dataset.csv;
    try {
      payload = packCsvForLink(dataset.csv);
    } catch (error) {
      payload = dataset.csv;
    }
    const encoded = await encodeLinkPayload(payload);
    exportLinkInput.value = encoded;
    exportBox.classList.remove("hidden");
  });

  copyLinkBtn.addEventListener("click", async () => {
    const value = exportLinkInput.value;
    if (!value) {
      return;
    }
    try {
      await navigator.clipboard.writeText(value);
      setStatus("Codice copiato negli appunti.", false);
    } catch (error) {
      setStatus("Impossibile copiare il codice.", true);
    }
  });

  exportCsvBtn.addEventListener("click", () => {
    const dataset = getCurrentDataset();
    if (!dataset) {
      setStatus("Nessun CSV salvato in memoria.", true);
      return;
    }
    downloadCsv(dataset.csv);
  });

  exportHtmlBtn.addEventListener("click", async () => {
    const dataset = getCurrentDataset();
    if (!dataset) {
      setStatus("Carica un CSV prima di esportare l'HTML.", true);
      return;
    }
    try {
      const html = await buildStaticHtml(dataset);
      downloadHtml(html, dataset.name);
    } catch (error) {
      setStatus("Impossibile esportare l'HTML.", true);
    }
  });


  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      if (!currentAnalysis) {
        return;
      }
      tabs.forEach((btn) => btn.classList.remove("active"));
      tab.classList.add("active");
      const view = tab.dataset.view;
      saveView(view);
      renderView(view);
    });
  });

  mainTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      mainTabs.forEach((btn) => btn.classList.remove("active"));
      tab.classList.add("active");
      const target = tab.dataset.main;
      const isGenerator = target === "generator";
      if (analysisSection) {
        analysisSection.classList.toggle("hidden", isGenerator);
      }
      if (generatorSection) {
        generatorSection.classList.toggle("hidden", !isGenerator);
      }
    });
  });

  initTeamweaveFormGenerator();
  initializeFromStorage();

  function readFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error("Impossibile leggere il file."));
      reader.readAsText(file);
    });
  }

  async function readFiles(files) {
    for (const file of files) {
      try {
        const content = await readFile(file);
        loadCsv(content, { name: file.name, persist: true });
      } catch (error) {
        setStatus("Impossibile leggere il file.", true);
      }
    }
  }

  function loadCsv(text, options = {}) {
    const { name = `Dataset ${datasets.length + 1}`, persist = true, replaceAll = false } = options;
    try {
      const dataset = createDatasetFromCsv(text, name);
      if (replaceAll) {
        datasets.length = 0;
        visibleDatasetIds.clear();
        hasVisibleState = false;
      }
      datasets.push(dataset);
      visibleDatasetIds.add(dataset.id);
      datasetCounter = Math.max(datasetCounter, datasets.length + 1);
      saveDatasets();
      setCurrentDataset(dataset.id, { persist });
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  function createDatasetFromCsv(text, name) {
    const sanitized = text.replace(/^\uFEFF/, "");
    const { headers, rows } = parseCsv(sanitized);
    const analysis = analyzeData(headers, rows);
    return {
      id: `ds-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
      name,
      csv: sanitized,
      analysis
    };
  }

  function setStatus(message, isError) {
    statusEl.textContent = message;
    statusEl.classList.toggle("error", isError);
  }

  function saveView(view) {
    try {
      localStorage.setItem(VIEW_KEY, view);
    } catch (error) {
      return;
    }
  }

  function restoreView() {
    try {
      const cached = localStorage.getItem(VIEW_KEY);
      if (!cached) {
        return null;
      }
      const valid = tabs.some((tab) => tab.dataset.view === cached);
      return valid ? cached : null;
    } catch (error) {
      return null;
    }
  }

  function saveDatasets() {
    try {
      const payload = datasets
        .filter((dataset) => dataset.name !== "Ultimo CSV")
        .map((dataset) => ({
          id: dataset.id,
          name: dataset.name,
          csv: dataset.csv
        }));
      localStorage.setItem(DATASETS_KEY, JSON.stringify(payload));
    } catch (error) {
      setStatus("Memoria del browser piena: impossibile salvare i dataset.", true);
    }
  }

  function pruneLegacyDatasets() {
    let removed = false;
    for (let i = datasets.length - 1; i >= 0; i -= 1) {
      if (datasets[i].name === "Ultimo CSV") {
        datasets.splice(i, 1);
        removed = true;
      }
    }
    if (removed) {
      if (!datasets.some((dataset) => dataset.id === currentDatasetId)) {
        currentDatasetId = datasets[0]?.id || null;
      }
      saveDatasets();
    }
    return removed;
  }

  function saveDatasetState() {
    try {
      const state = {
        currentDatasetId,
        compareA: compareSelectA?.value || null,
        compareB: compareSelectB?.value || null,
        visibleIds: Array.from(visibleDatasetIds)
      };
      localStorage.setItem(DATASET_STATE_KEY, JSON.stringify(state));
    } catch (error) {
      return;
    }
  }

  function restoreDatasets() {
    try {
      const cached = localStorage.getItem(DATASETS_KEY);
      if (!cached) {
        return false;
      }
      const entries = JSON.parse(cached);
      if (!Array.isArray(entries) || entries.length === 0) {
        return false;
      }
      datasets.length = 0;
      let removedLegacy = false;
      entries.forEach((entry) => {
        if (!entry || !entry.csv) {
          return;
        }
        if (entry.name === "Ultimo CSV") {
          removedLegacy = true;
          return;
        }
        try {
          const dataset = createDatasetFromCsv(entry.csv, entry.name || `Dataset ${datasets.length + 1}`);
          dataset.id = entry.id || dataset.id;
          datasets.push(dataset);
        } catch (error) {
          return;
        }
      });
      pruneLegacyDatasets();
      if (removedLegacy) {
        saveDatasets();
      }
      if (!datasets.length) {
        return false;
      }
      isRestoringDatasets = true;
      datasetCounter = Math.max(datasetCounter, datasets.length + 1);
      updateDatasetControls();
      const stateRaw = localStorage.getItem(DATASET_STATE_KEY);
      let nextId = datasets[0].id;
      if (stateRaw) {
        try {
          const state = JSON.parse(stateRaw);
          if (Array.isArray(state?.visibleIds)) {
            hasVisibleState = true;
            visibleDatasetIds = new Set(state.visibleIds);
          }
          if (state?.currentDatasetId && datasets.some((d) => d.id === state.currentDatasetId)) {
            nextId = state.currentDatasetId;
          }
          if (state?.compareA) {
            compareSelectA.value = state.compareA;
          }
          if (state?.compareB) {
            compareSelectB.value = state.compareB;
          }
          ensureCompareSelection();
        } catch (error) {
          return false;
        }
      }
      setCurrentDataset(nextId, { persist: false });
      isRestoringDatasets = false;
      saveDatasetState();
      return true;
    } catch (error) {
      isRestoringDatasets = false;
      return false;
    }
  }

  async function restoreFromLink() {
    const hash = location.hash.replace(/^#/, "");
    if (!hash.startsWith(`${LINK_PARAM}=`)) {
      return false;
    }
    const encoded = hash.slice(LINK_PARAM.length + 1);
    const decoded = await decodeLinkPayload(encoded);
    let csv = null;
    if (decoded && decoded.kind === "bytes") {
      csv = unpackCsvFromBinary(decoded.value);
    } else if (decoded && decoded.kind === "text") {
      csv = decodePackedCsv(decoded.value);
    }
    if (csv) {
      loadCsv(csv, { name: "Link dati", persist: true, replaceAll: true });
      return true;
    }
    return false;
  }

  function parseCsv(text) {
    const rows = [];
    let current = [];
    let value = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i += 1) {
      const char = text[i];
      const next = text[i + 1];
      if (char === '"') {
        if (inQuotes && next === '"') {
          value += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === "," && !inQuotes) {
        current.push(value);
        value = "";
      } else if ((char === "\n" || char === "\r") && !inQuotes) {
        if (char === "\r" && next === "\n") {
          i += 1;
        }
        current.push(value);
        value = "";
        if (current.some((cell) => cell.trim() !== "")) {
          rows.push(current);
        }
        current = [];
      } else {
        value += char;
      }
    }
    if (value.length || current.length) {
      current.push(value);
      if (current.some((cell) => cell.trim() !== "")) {
        rows.push(current);
      }
    }

    if (rows.length < 2) {
      throw new Error("CSV non valido o vuoto.");
    }

    const headers = rows[0].map((cell) => cell.trim());
    const dataRows = rows.slice(1);

    return { headers, rows: dataRows };
  }

  function analyzeData(headers, rows) {
    const nameIndex = findNameColumn(headers);
    const firstQuestionIndex = nameIndex + 1;
    if (firstQuestionIndex >= headers.length) {
      throw new Error("CSV senza colonne di domande riconoscibili.");
    }

    const questionHeaders = headers.slice(firstQuestionIndex);
    const questionMeta = questionHeaders.map((question, idx) => {
      const tagged = extractCategoryTag(question);
      const normalized = normalizeQuestion(tagged.question);
      const category = tagged.category || CATEGORY_MAP_NORMALIZED.get(normalized) || null;
      return {
        question: tagged.question,
        category,
        positive: idx % 2 === 0
      };
    });
    const posQuestionLabels = [];
    const negQuestionLabels = [];
    const posIndexByMeta = [];
    const negIndexByMeta = [];
    questionMeta.forEach((meta, idx) => {
      if (meta.positive) {
        posIndexByMeta[idx] = posQuestionLabels.length;
        posQuestionLabels.push(meta.question);
        return;
      }
      negIndexByMeta[idx] = negQuestionLabels.length;
      negQuestionLabels.push(meta.question);
    });

    const names = [];
    const nameIndexMap = new Map();
    const normalizedMap = new Map();
    const rowsNormalized = rows.map((row) => row.map((cell) => cell.trim()));

    rowsNormalized.forEach((row) => {
      const rawName = row[nameIndex] || "";
      const normalized = normalizeName(rawName);
      if (!normalized) {
        return;
      }
      if (!nameIndexMap.has(normalized)) {
        nameIndexMap.set(normalized, names.length);
        normalizedMap.set(normalized, rawName.trim());
        names.push(rawName.trim());
      }
    });

    if (names.length === 0) {
      throw new Error("Nessun nome trovato nella colonna 'Seleziona il tuo nome'.");
    }

    const size = names.length;
    const posMatrix = createMatrix(size, 0);
    const negMatrix = createMatrix(size, 0);
    const leadershipPos = createMatrix(size, 0);
    const leadershipNeg = createMatrix(size, 0);
    const receivedPosByCat = initCategoryMap(names, CATEGORIES);
    const receivedNegByCat = initCategoryMap(names, CATEGORIES);
    const posReceivedByQuestion = createRectMatrix(size, posQuestionLabels.length, 0);
    const negReceivedByQuestion = createRectMatrix(size, negQuestionLabels.length, 0);
    const unknownSelections = new Set();
    const missingCategories = new Set();

    rowsNormalized.forEach((row) => {
      const chooserRaw = row[nameIndex] || "";
      const chooserKey = normalizeName(chooserRaw);
      if (!chooserKey || !nameIndexMap.has(chooserKey)) {
        return;
      }
      const chooserIndex = nameIndexMap.get(chooserKey);

      questionMeta.forEach((meta, idx) => {
        const cell = row[firstQuestionIndex + idx] || "";
        const picks = splitNames(cell);
        const seen = new Set();
        if (!meta.category) {
          if (cell.trim() !== "") {
            missingCategories.add(meta.question);
          }
          return;
        }

        picks.forEach((pick) => {
          const pickKey = normalizeName(pick);
          if (!pickKey || seen.has(pickKey)) {
            return;
          }
          seen.add(pickKey);
          if (!nameIndexMap.has(pickKey)) {
            unknownSelections.add(pick.trim());
            return;
          }
          const pickIndex = nameIndexMap.get(pickKey);

          if (meta.positive) {
            if (chooserIndex !== pickIndex) {
              posMatrix[chooserIndex][pickIndex] += 1;
              if (meta.category === "Attitudinali" || meta.category === "Sociali") {
                leadershipPos[chooserIndex][pickIndex] += 1;
              }
            }
            receivedPosByCat.get(pickKey)[meta.category] += 1;
            const qIndex = posIndexByMeta[idx];
            if (qIndex != null) {
              posReceivedByQuestion[pickIndex][qIndex] += 1;
            }
          } else {
            if (chooserIndex !== pickIndex) {
              negMatrix[chooserIndex][pickIndex] += 1;
              if (meta.category === "Attitudinali" || meta.category === "Sociali") {
                leadershipNeg[chooserIndex][pickIndex] += 1;
              }
            }
            receivedNegByCat.get(pickKey)[meta.category] += 1;
            const qIndex = negIndexByMeta[idx];
            if (qIndex != null) {
              negReceivedByQuestion[pickIndex][qIndex] += 1;
            }
          }
        });
      });
    });

    const posRowSum = sumRows(posMatrix);
    const posColSum = sumColumns(posMatrix);
    const negRowSum = sumRows(negMatrix);
    const negColSum = sumColumns(negMatrix);

    const reciprociPos = buildReciprociMatrix(posMatrix);
    const reciprociNeg = buildReciprociMatrix(negMatrix);

    const forzaLegami = buildForzaMatrix(posMatrix);
    const forzaAntagonismo = buildForzaMatrix(negMatrix);

    const nonRicambiatePos = buildNonRicambiateMatrix(posMatrix, reciprociPos);

    const reciprociPosRowSum = sumRows(reciprociPos);
    const reciprociNegRowSum = sumRows(reciprociNeg);

    const forzaLegamiRowSum = sumRows(forzaLegami);
    const forzaAntagonismoRowSum = sumRows(forzaAntagonismo);
    const nonRicambiateRowSum = sumRows(nonRicambiatePos);
    const nonRicambiateColSum = sumColumns(nonRicambiatePos);

    const summaryRows = names.map((name, index) => {
      const key = normalizeName(name);
      const posCats = receivedPosByCat.get(key);
      const negCats = receivedNegByCat.get(key);
      return {
        name,
        positiveReceived: {
          Tecniche: posCats.Tecniche,
          Attitudinali: posCats.Attitudinali,
          Sociali: posCats.Sociali,
          Totali: posColSum[index]
        },
        negativeReceived: {
          Tecniche: negCats.Tecniche,
          Attitudinali: negCats.Attitudinali,
          Sociali: negCats.Sociali,
          Totali: negColSum[index]
        },
        positiveGiven: posRowSum[index],
        negativeGiven: reciprociNegRowSum[index],
        reciprociPos: reciprociPosRowSum[index],
        reciprociNeg: reciprociNegRowSum[index],
        forzaLegami: forzaLegamiRowSum[index],
        forzaAntagonismo: forzaAntagonismoRowSum[index],
        nonRicambiateDate: nonRicambiateRowSum[index],
        nonRicambiateRicevute: nonRicambiateColSum[index]
      };
    });

    const averages = computeAverages(summaryRows);
    const classifications = summaryRows.map((row) => {
      const inflPos = classifyInfluence(row.positiveReceived.Totali, averages.positiveReceivedTotal);
      const inflNeg = classifyInfluence(row.negativeReceived.Totali, averages.negativeReceivedTotal);
      const byCategory = Object.fromEntries(CATEGORIES.map((category) => {
        const positiveAverage = averages[`positiveReceived${category}`];
        const negativeAverage = averages[`negativeReceived${category}`];
        const positive = positiveAverage > 0
          ? classifyInfluence(row.positiveReceived[category], positiveAverage) : "Assente";
        const negative = negativeAverage > 0
          ? classifyInfluence(row.negativeReceived[category], negativeAverage) : "Assente";
        return [category, {
          inflPos: positive,
          inflNeg: negative,
          label: synthesizeCategoryLabel(category, positive, negative)
        }];
      }));
      const equilibrio = classifyBalance(
        row.nonRicambiateDate,
        row.nonRicambiateRicevute,
        averages.nonRicambiateDate,
        averages.nonRicambiateRicevute
      );
      return {
        name: row.name,
        inflPos,
        inflNeg,
        equilibrio,
        byCategory,
        label: CATEGORIES.map((category) => byCategory[category].label).join(" · ")
      };
    });

    return {
      headers,
      rows: rowsNormalized,
      names,
      matrices: {
        posMatrix,
        negMatrix,
        leadershipPos,
        leadershipNeg,
        reciprociPos,
        reciprociNeg,
        forzaLegami,
        forzaAntagonismo,
        nonRicambiatePos
      },
      totals: {
        posRowSum,
        posColSum,
        negRowSum,
        negColSum,
        reciprociPosRowSum,
        reciprociNegRowSum,
        forzaLegamiRowSum,
        forzaAntagonismoRowSum,
        nonRicambiateRowSum,
        nonRicambiateColSum
      },
      questionMeta,
      summaryRows,
      averages,
      classifications,
      questionStats: {
        posLabels: posQuestionLabels,
        negLabels: negQuestionLabels,
        posReceivedByQuestion,
        negReceivedByQuestion
      },
      warnings: {
        unknownSelections: Array.from(unknownSelections).sort(),
        missingCategories: Array.from(missingCategories).sort()
      }
    };
  }

  function findNameColumn(headers) {
    const idx = headers.findIndex((header) =>
      header.toLowerCase().includes("seleziona il tuo nome")
    );
    if (idx !== -1) {
      return idx;
    }
    return 1;
  }

  function normalizeName(name) {
    return name.trim().replace(/\s+/g, " ").toLowerCase();
  }

  function extractCategoryTag(question) {
    const raw = String(question || "");
    const match = raw.match(/^\s*\[(Tecniche|Attitudinali|Sociali)\]\s*/i);
    if (!match) {
      return { question: raw, category: null };
    }
    const cleaned = raw.slice(match[0].length).trim();
    const lowered = match[1].toLowerCase();
    const category = lowered.startsWith("tec")
      ? "Tecniche"
      : lowered.startsWith("att")
        ? "Attitudinali"
        : "Sociali";
    return { question: cleaned || raw, category };
  }

  function normalizeQuestion(question) {
    return question
      .trim()
      .replace(/[’‘]/g, "'")
      .replace(/\s+/g, " ")
      .toLowerCase();
  }

  function escapeQuotes(text) {
    return String(text || "")
      .replace(/\\/g, "\\\\")
      .replace(/"/g, '\\"')
      .replace(/\r?\n/g, "\\n");
  }

  function splitNames(value) {
    if (!value) {
      return [];
    }
    return value
      .split(",")
      .map((part) => part.trim())
      .filter(Boolean);
  }

  function createMatrix(size, fillValue) {
    return Array.from({ length: size }, () => Array(size).fill(fillValue));
  }

  function createRectMatrix(rows, cols, fillValue) {
    return Array.from({ length: rows }, () => Array(cols).fill(fillValue));
  }

  function initCategoryMap(names, categories) {
    const map = new Map();
    names.forEach((name) => {
      const entry = {};
      categories.forEach((cat) => {
        entry[cat] = 0;
      });
      map.set(normalizeName(name), entry);
    });
    return map;
  }

  function sumRows(matrix) {
    return matrix.map((row) => row.reduce((acc, value) => acc + value, 0));
  }

  function sumColumns(matrix) {
    const size = matrix.length;
    const sums = Array(size).fill(0);
    matrix.forEach((row) => {
      row.forEach((value, index) => {
        sums[index] += value;
      });
    });
    return sums;
  }

  function buildReciprociMatrix(matrix) {
    const size = matrix.length;
    const result = createMatrix(size, 0);
    for (let i = 0; i < size; i += 1) {
      for (let j = 0; j < size; j += 1) {
        if (i === j) {
          result[i][j] = 0;
        } else if (matrix[i][j] > 0 && matrix[j][i] > 0) {
          result[i][j] = 1;
        }
      }
    }
    return result;
  }

  function buildForzaMatrix(matrix) {
    const size = matrix.length;
    const result = createMatrix(size, 0);
    for (let i = 0; i < size; i += 1) {
      for (let j = 0; j < size; j += 1) {
        if (i === j) {
          result[i][j] = 0;
        } else {
          result[i][j] = Math.min(matrix[i][j], matrix[j][i]);
        }
      }
    }
    return result;
  }

  function buildNonRicambiateMatrix(matrix, reciproci) {
    const size = matrix.length;
    const result = createMatrix(size, 0);
    for (let i = 0; i < size; i += 1) {
      for (let j = 0; j < size; j += 1) {
        if (matrix[i][j] > 0 && reciproci[i][j] === 0) {
          result[i][j] = matrix[i][j];
        }
      }
    }
    return result;
  }

  function computeAverages(rows) {
    const count = rows.length || 1;
    const totals = {
      positiveReceivedTecniche: 0,
      positiveReceivedAttitudinali: 0,
      positiveReceivedSociali: 0,
      positiveReceivedTotal: 0,
      negativeReceivedTecniche: 0,
      negativeReceivedAttitudinali: 0,
      negativeReceivedSociali: 0,
      negativeReceivedTotal: 0,
      positiveGiven: 0,
      negativeGiven: 0,
      reciprociPos: 0,
      reciprociNeg: 0,
      forzaLegami: 0,
      forzaAntagonismo: 0,
      nonRicambiateDate: 0,
      nonRicambiateRicevute: 0
    };

    rows.forEach((row) => {
      totals.positiveReceivedTecniche += row.positiveReceived.Tecniche;
      totals.positiveReceivedAttitudinali += row.positiveReceived.Attitudinali;
      totals.positiveReceivedSociali += row.positiveReceived.Sociali;
      totals.positiveReceivedTotal += row.positiveReceived.Totali;
      totals.negativeReceivedTecniche += row.negativeReceived.Tecniche;
      totals.negativeReceivedAttitudinali += row.negativeReceived.Attitudinali;
      totals.negativeReceivedSociali += row.negativeReceived.Sociali;
      totals.negativeReceivedTotal += row.negativeReceived.Totali;
      totals.positiveGiven += row.positiveGiven;
      totals.negativeGiven += row.negativeGiven;
      totals.reciprociPos += row.reciprociPos;
      totals.reciprociNeg += row.reciprociNeg;
      totals.forzaLegami += row.forzaLegami;
      totals.forzaAntagonismo += row.forzaAntagonismo;
      totals.nonRicambiateDate += row.nonRicambiateDate;
      totals.nonRicambiateRicevute += row.nonRicambiateRicevute;
    });

    const averages = {};
    Object.keys(totals).forEach((key) => {
      averages[key] = totals[key] / count;
    });
    return averages;
  }

  function classifyInfluence(value, average) {
    if (!Number.isFinite(value)) {
      return "";
    }
    if (value >= average) {
      return "Alta";
    }
    if (value >= average / 2) {
      return "Media";
    }
    return "Bassa";
  }

  function synthesizeCategoryLabel(category, positive, negative) {
    const labels = {
      Tecniche: ["Leader tecnica", "Leader tecnica controversa", "Poco apprezzata tecnicamente", "Apprezzata tecnicamente", "Valutazione tecnica mista", "Poco considerata tecnicamente"],
      Sociali: ["Benvoluta nel gruppo", "Rapporti contrastanti", "Poco accolta nel gruppo", "Benvoluta nel gruppo", "Rapporti contrastanti", "Poco coinvolta socialmente"],
      Attitudinali: ["Impegno riconosciuto", "Impegno discusso", "Scarso impegno percepito", "Impegno apprezzato", "Impegno percepito in modo misto", "Impegno poco riconosciuto"]
    };
    if (positive === "Assente" && negative === "Assente") {
      return `${category}: nessuna scelta`;
    }
    const options = labels[category];
    if (positive === "Alta") {
      return negative === "Alta" ? options[1] : options[0];
    }
    if (negative === "Alta") return options[2];
    if (positive === "Media") {
      return negative === "Media" ? options[4] : options[3];
    }
    return negative === "Media" ? options[4] : options[5];
  }

  function classifyBalance(nonRicDate, nonRicRecv, avgDate, avgRecv) {
    if (!Number.isFinite(nonRicDate) || !Number.isFinite(nonRicRecv)) {
      return "";
    }
    if (nonRicRecv - nonRicDate >= avgDate) {
      return "Selettiva";
    }
    if (nonRicDate - nonRicRecv >= avgRecv) {
      return "Ignorata";
    }
    return "Bilanciata";
  }

  function synthesizeLabel(inflPos, inflNeg, equilibrio) {
    if (inflPos === "Alta") {
      if (inflNeg === "Alta") {
        if (equilibrio === "Selettiva") {
          return "Leader controversa";
        }
        if (equilibrio === "Bilanciata") {
          return "Figura di riferimento";
        }
        return "Figura controversa";
      }
      if (inflNeg === "Media") {
        if (equilibrio === "Selettiva") {
          return "Leader silenziosa";
        }
        if (equilibrio === "Bilanciata") {
          return "Benvoluta";
        }
        return "Stimata";
      }
      if (equilibrio === "Selettiva") {
        return "Leader positiva";
      }
      if (equilibrio === "Bilanciata") {
        return "Stimata";
      }
      return "Apprezzata";
    }

    if (inflPos === "Media") {
      if (inflNeg === "Alta") {
        if (equilibrio === "Ignorata") {
          return "Invisibile e controversa";
        }
        if (equilibrio === "Bilanciata") {
          return "Controversa";
        }
        return "Marginale";
      }
      if (inflNeg === "Media") {
        if (equilibrio === "Selettiva") {
          return "Bilanciata";
        }
        if (equilibrio === "Bilanciata") {
          return "Neutrale";
        }
        return "Presenza discreta";
      }
      if (equilibrio === "Selettiva") {
        return "Apprezzata";
      }
      if (equilibrio === "Bilanciata") {
        return "Stimata";
      }
      return "Presenza discreta";
    }

    if (inflNeg === "Alta") {
      if (equilibrio === "Ignorata") {
        return "Esclusa";
      }
      if (equilibrio === "Bilanciata") {
        return "Altruista";
      }
      return "Dispersa";
    }
    if (inflNeg === "Media") {
      if (equilibrio === "Ignorata") {
        return "Trascurata";
      }
      return "Non considerata";
    }
    return "Invisibile";
  }

  function updateSummary(analysis) {
    document.getElementById("summary-athletes").textContent = analysis.names.length;
    document.getElementById("summary-reciproci-pos").textContent = formatPercent(
      computeReciprocityIndex(analysis.matrices.reciprociPos)
    );
    document.getElementById("summary-reciproci-neg").textContent = formatPercent(
      computeReciprocityIndex(analysis.matrices.reciprociNeg)
    );

    const warnings = [];
    if (analysis.warnings.unknownSelections.length) {
      warnings.push(`Nomi fuori lista: ${analysis.warnings.unknownSelections.join(", ")}`);
    }
    if (analysis.warnings.missingCategories.length) {
      warnings.push("Alcune domande non sono mappate alle categorie.");
    }
    if (warnings.length) {
      setStatus(warnings.join(" "), true);
    }
  }

  function getCurrentDataset() {
    return datasets.find((dataset) => dataset.id === currentDatasetId) || null;
  }

  function removeDataset(id) {
    const index = datasets.findIndex((dataset) => dataset.id === id);
    if (index === -1) {
      return;
    }
    datasets.splice(index, 1);
    visibleDatasetIds.delete(id);
    saveDatasets();
    if (!datasets.length) {
      currentDatasetId = null;
      currentAnalysis = null;
      summaryEl.classList.add("hidden");
      viewsEl.classList.add("hidden");
      viewContainer.innerHTML = "";
      datasetBox.classList.add("hidden");
      saveDatasetState();
      setStatus("Dataset rimosso.", false);
      return;
    }
    const nextId = datasets[0].id;
    setCurrentDataset(nextId, { persist: false });
    setStatus("Dataset rimosso.", false);
  }

  function ensureVisibleDatasetIds() {
    if (IS_EXPORT_MODE) {
      return false;
    }
    const validIds = new Set(datasets.map((dataset) => dataset.id));
    let changed = false;
    visibleDatasetIds.forEach((id) => {
      if (!validIds.has(id)) {
        visibleDatasetIds.delete(id);
        changed = true;
      }
    });
    if (!visibleDatasetIds.size && !hasVisibleState) {
      datasets.forEach((dataset) => visibleDatasetIds.add(dataset.id));
      changed = true;
    }
    return changed;
  }

  function getVisibleDatasetIds() {
    if (IS_EXPORT_MODE) {
      return new Set(datasets.map((dataset) => dataset.id));
    }
    ensureVisibleDatasetIds();
    return new Set(visibleDatasetIds);
  }

  function getVisibleDatasets() {
    const visibleIds = getVisibleDatasetIds();
    return datasets.filter((dataset) => visibleIds.has(dataset.id));
  }

  function updateDatasetManager() {
    if (!datasetManager) {
      return;
    }
    datasetManager.textContent = "";
    if (!datasets.length) {
      return;
    }
    const visibleIds = getVisibleDatasetIds();
    const list = document.createElement("div");
    list.className = "dataset-manager-grid";
    datasets.forEach((dataset) => {
      const row = document.createElement("div");
      row.className = "dataset-manager-item";

      const toggleLabel = document.createElement("label");
      toggleLabel.className = "dataset-manager-toggle";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = visibleIds.has(dataset.id);
      checkbox.addEventListener("change", () => {
        if (checkbox.checked) {
          visibleDatasetIds.add(dataset.id);
        } else {
          visibleDatasetIds.delete(dataset.id);
        }
        hasVisibleState = true;
        const active = getActiveView() || "responses";
        updateDatasetControls();
        renderView(active);
      });
      const name = document.createElement("span");
      name.textContent = dataset.name;
      toggleLabel.appendChild(checkbox);
      toggleLabel.appendChild(name);

      const orderGroup = document.createElement("div");
      orderGroup.className = "dataset-manager-order";
      const radioA = document.createElement("input");
      radioA.type = "radio";
      radioA.name = "dataset-order-a";
      radioA.value = dataset.id;
      radioA.checked = compareSelectA.value === dataset.id;
      radioA.addEventListener("change", () => {
        if (!visibleDatasetIds.has(dataset.id)) {
          visibleDatasetIds.add(dataset.id);
          checkbox.checked = true;
          hasVisibleState = true;
        }
        compareSelectA.value = dataset.id;
        ensureCompareSelection();
        const active = getActiveView() || "responses";
        renderView(active);
        saveDatasetState();
        updateDatasetManager();
      });
      const labelA = document.createElement("label");
      labelA.appendChild(radioA);
      labelA.appendChild(document.createTextNode("A"));

      const radioB = document.createElement("input");
      radioB.type = "radio";
      radioB.name = "dataset-order-b";
      radioB.value = dataset.id;
      radioB.checked = compareSelectB.value === dataset.id;
      radioB.addEventListener("change", () => {
        if (!visibleDatasetIds.has(dataset.id)) {
          visibleDatasetIds.add(dataset.id);
          checkbox.checked = true;
          hasVisibleState = true;
        }
        compareSelectB.value = dataset.id;
        ensureCompareSelection();
        const active = getActiveView() || "responses";
        renderView(active);
        saveDatasetState();
        updateDatasetManager();
      });
      const labelB = document.createElement("label");
      labelB.appendChild(radioB);
      labelB.appendChild(document.createTextNode("B"));

      orderGroup.appendChild(labelA);
      orderGroup.appendChild(labelB);

      row.appendChild(toggleLabel);
      row.appendChild(orderGroup);
      list.appendChild(row);
    });
    datasetManager.appendChild(list);
  }

  function setCurrentDataset(id, options = {}) {
    const { persist = true } = options;
    const visibleIds = getVisibleDatasetIds();
    let targetId = id;
    if (visibleIds.size && !visibleIds.has(targetId)) {
      const fallback = datasets.find((entry) => visibleIds.has(entry.id));
      targetId = fallback ? fallback.id : id;
    }
    const dataset = datasets.find((entry) => entry.id === targetId);
    if (!dataset) {
      return;
    }
    currentDatasetId = dataset.id;
    currentAnalysis = dataset.analysis;
    updateDatasetControls();
    if (!isRestoringDatasets) {
      saveDatasetState();
    }
    updateSummary(dataset.analysis);
    summaryEl.classList.remove("hidden");
    viewsEl.classList.remove("hidden");
    const selectedView = getActiveView() || restoreView() || "responses";
    renderView(selectedView);
    tabs.forEach((btn) => btn.classList.remove("active"));
    const activeTab = tabs.find((btn) => btn.dataset.view === selectedView) || tabs[0];
    activeTab.classList.add("active");
    setStatus(
      `Caricato: ${dataset.analysis.rows.length} risposte, ${dataset.analysis.names.length} atlete. (${dataset.name})`,
      false
    );
    applyExportModeUI();
  }

  function updateDatasetControls() {
    if (!datasets.length) {
      datasetBox.classList.add("hidden");
      return;
    }
    datasetBox.classList.remove("hidden");
    const visibleIds = getVisibleDatasetIds();
    const visibleDatasets = getVisibleDatasets();
    updateDatasetManager();
    if (!visibleDatasets.length) {
      currentDatasetId = null;
      currentAnalysis = null;
      summaryEl.classList.add("hidden");
      viewsEl.classList.add("hidden");
      viewContainer.classList.add("hidden");
      compareRow.classList.add("hidden");
      compareNote.classList.add("hidden");
      compareSelectA.value = "";
      compareSelectB.value = "";
      compareSelectA.selectedIndex = -1;
      compareSelectB.selectedIndex = -1;
      if (!isRestoringDatasets) {
        saveDatasetState();
      }
      return;
    }
    if (currentDatasetId && !visibleIds.has(currentDatasetId)) {
      setCurrentDataset(visibleDatasets[0].id, { persist: false });
      return;
    }
    viewContainer.classList.remove("hidden");
    populateDatasetSelect(datasetSelect, currentDatasetId, visibleDatasets);
    populateDatasetSelect(compareSelectA, compareSelectA.value, visibleDatasets);
    populateDatasetSelect(compareSelectB, compareSelectB.value, visibleDatasets);
    ensureCompareSelection();
    const showCompare = visibleDatasets.length >= 2;
    compareRow.classList.toggle("hidden", !showCompare);
    compareNote.classList.toggle("hidden", !showCompare);
    compareSelectA.disabled = !showCompare;
    compareSelectB.disabled = !showCompare;
    if (!isRestoringDatasets) {
      saveDatasetState();
    }
    applyExportModeUI();
  }

  function applyExportModeUI() {
    if (!IS_EXPORT_MODE) {
      return;
    }
    const panel = document.querySelector(".upload-panel");
    if (!panel) {
      return;
    }
    const title = panel.querySelector("h2");
    if (title) {
      title.textContent = "Dataset caricati";
    }
    const description = panel.querySelector("p");
    if (description) {
      description.textContent = "Seleziona uno o due dataset da visualizzare.";
    }
    if (analysisSection) {
      analysisSection.classList.remove("hidden");
    }
    if (generatorSection) {
      generatorSection.classList.add("hidden");
    }
    if (mainTabs.length) {
      mainTabs.forEach((btn) => {
        btn.classList.toggle("active", btn.dataset.main === "analysis");
        btn.classList.add("hidden");
      });
    }
    panel.querySelectorAll(
      "button, .export-box, .import-box, .tw-form-box, .upload-row, .dataset-row, #compare-row, #compare-note, #status, #file-input, #import-code, #export-link"
    ).forEach((el) => {
      el.classList.add("hidden");
    });
    let picker = panel.querySelector(".dataset-picker");
    if (!picker) {
      picker = document.createElement("div");
      picker.className = "dataset-picker";
      panel.appendChild(picker);
    }
    picker.innerHTML = "";
    if (!datasets.length) {
      const empty = document.createElement("p");
      empty.className = "note";
      empty.textContent = "Nessun dataset.";
      picker.appendChild(empty);
      return;
    }
    const selectedIds = new Set(getExportSelectedIds());
    datasets.forEach((dataset) => {
      const label = document.createElement("label");
      label.className = "dataset-option";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = dataset.id;
      checkbox.checked = selectedIds.has(dataset.id);
      checkbox.addEventListener("change", () => {
        handleExportCheckboxChange(picker, checkbox);
      });
      const name = document.createElement("span");
      name.textContent = dataset.name;
      label.appendChild(checkbox);
      label.appendChild(name);
      picker.appendChild(label);
    });
    if (selectedIds.size) {
      summaryEl.classList.remove("hidden");
      viewsEl.classList.remove("hidden");
      viewContainer.classList.remove("hidden");
    }
  }

  function getExportSelectedIds() {
    const ids = [];
    if (compareSelectA?.value) {
      ids.push(compareSelectA.value);
    }
    if (compareSelectB?.value && compareSelectB.value !== compareSelectA.value) {
      ids.push(compareSelectB.value);
    }
    if (!ids.length && currentDatasetId) {
      ids.push(currentDatasetId);
    }
    return ids;
  }

  function handleExportCheckboxChange(container, changed) {
    const inputs = Array.from(container.querySelectorAll("input[type=\"checkbox\"]"));
    const selected = inputs.filter((input) => input.checked).map((input) => input.value);
    applyExportSelection(selected);
  }

  function orderByExport(ids) {
    if (!IS_EXPORT_MODE || !EXPORT_ORDER || !EXPORT_ORDER.length) {
      return ids.slice();
    }
    const orderMap = new Map(EXPORT_ORDER.map((id, index) => [id, index]));
    return ids
      .map((id, index) => ({
        id,
        index,
        order: orderMap.has(id) ? orderMap.get(id) : Number.MAX_SAFE_INTEGER
      }))
      .sort((a, b) => (a.order === b.order ? a.index - b.index : a.order - b.order))
      .map((entry) => entry.id);
  }

  function applyExportSelection(ids) {
    const unique = Array.from(new Set(ids));
    if (!unique.length) {
      currentDatasetId = null;
      currentAnalysis = null;
      compareSelectA.value = "";
      compareSelectB.value = "";
      compareSelectA.selectedIndex = -1;
      compareSelectB.selectedIndex = -1;
      summaryEl.classList.add("hidden");
      viewsEl.classList.add("hidden");
      viewContainer.classList.add("hidden");
      return;
    }
    summaryEl.classList.remove("hidden");
    viewsEl.classList.remove("hidden");
    viewContainer.classList.remove("hidden");
    if (unique.length === 1) {
      compareSelectA.value = unique[0];
      compareSelectB.value = unique[0];
      setCurrentDataset(unique[0], { persist: false });
      return;
    }
    const ordered = orderByExport(unique);
    compareSelectA.value = ordered[0];
    compareSelectB.value = ordered[1];
    setCurrentDataset(ordered[0], { persist: false });
  }

  function populateDatasetSelect(select, selectedId, source = datasets) {
    select.textContent = "";
    source.forEach((dataset) => {
      const option = document.createElement("option");
      option.value = dataset.id;
      option.textContent = dataset.name;
      select.appendChild(option);
    });
    const targetId = source.some((dataset) => dataset.id === selectedId)
      ? selectedId
      : source[0]?.id;
    if (targetId) {
      select.value = targetId;
    }
  }

  function ensureCompareSelection() {
    if (IS_EXPORT_MODE) {
      return;
    }
    const visibleDatasets = getVisibleDatasets();
    if (visibleDatasets.length < 2) {
      if (visibleDatasets.length === 1) {
        compareSelectA.value = visibleDatasets[0].id;
        compareSelectB.value = visibleDatasets[0].id;
      }
      return;
    }
    const visibleIds = visibleDatasets.map((dataset) => dataset.id);
    if (!visibleIds.includes(compareSelectA.value)) {
      compareSelectA.value = visibleIds[0];
    }
    if (!visibleIds.includes(compareSelectB.value)) {
      compareSelectB.value = visibleIds[1] || visibleIds[0];
    }
    if (compareSelectA.value === compareSelectB.value) {
      const alternative = visibleIds.find((id) => id !== compareSelectA.value);
      if (alternative) {
        compareSelectB.value = alternative;
      }
    }
  }

  function getActiveView() {
    const active = tabs.find((tab) => tab.classList.contains("active"));
    return active ? active.dataset.view : null;
  }

  function renderView(view) {
    viewContainer.innerHTML = "";
    const compareDatasets = getCompareDatasets();
    if (!currentAnalysis && !compareDatasets) {
      return;
    }

    if (view === "confronto") {
      viewContainer.classList.remove("dual-view");
      viewContainer.appendChild(renderCompareSection());
      return;
    }

    if (compareDatasets) {
      renderDualView(view, compareDatasets);
      return;
    }

    viewContainer.classList.remove("dual-view");
    if (!renderers.has(view)) {
      renderers.set(view, createRenderer(view));
    }
      renderers.get(view)(currentAnalysis, viewContainer);
  }

  function computeReciprocityIndex(matrix) {
    const rowTotals = sumRows(matrix);
    const denom = matrix.length - 1;
    if (denom <= 0) {
      return 0;
    }
    const total = rowTotals.reduce((acc, value) => acc + value / denom, 0);
    return total / rowTotals.length;
  }

  function createRenderer(view) {
    switch (view) {
      case "matrix-pos":
        return (analysis, container) => {
          container.appendChild(renderMatrixSection(
            "Matrice positiva",
            "Conta quante volte ogni atleta ha scelto positivamente un'altra.",
            analysis.names,
            analysis.matrices.posMatrix,
            analysis.totals.posRowSum,
            analysis.totals.posColSum,
            "Totale scelte fatte",
            {
              matrix: { palette: PALETTE.posMatrix, scope: "matrix" },
              totals: { palette: PALETTE.posTotals, scope: "totals" }
            },
            true,
            "Totale scelte ricevute"
          ));
          appendNetworkForView(analysis, container, view);
        };
      case "capitano":
        return (analysis, container) => {
          container.appendChild(renderCaptainSection(analysis));
        };
      case "giocatrice":
        return (analysis, container) => {
          container.appendChild(renderPlayerSection(analysis));
        };
      case "responses":
        return (analysis, container) => {
          container.appendChild(renderResponsesSection(analysis));
        };
      case "top-domande":
        return (analysis, container) => {
          container.appendChild(renderTopQuestionsSection(analysis));
        };
      case "top-ricevute":
        return (analysis, container) => {
          container.appendChild(renderTopRecipientsSection(analysis));
        };
      case "matrix-neg":
        return (analysis, container) => {
          container.appendChild(renderMatrixSection(
            "Matrice negativa",
            "Conta quante volte ogni atleta ha indicato negativamente un'altra.",
            analysis.names,
            analysis.matrices.negMatrix,
            analysis.totals.negRowSum,
            analysis.totals.negColSum,
            "Totale scelte fatte",
            {
              matrix: { palette: PALETTE.negMatrix, scope: "matrix" },
              totals: { palette: PALETTE.negTotals, scope: "totals" }
            },
            true,
            "Totale scelte ricevute"
          ));
          appendNetworkForView(analysis, container, view);
        };
      case "reciproci-pos":
        return (analysis, container) => {
          container.appendChild(renderReciprociSection(
            "Reciproci positivi",
            "1 indica una scelta reciproca positiva.",
            analysis.names,
            analysis.matrices.reciprociPos,
            {
              matrix: { palette: PALETTE.posMatrix, scope: "matrix" },
              totals: { palette: PALETTE.posTotals, scope: "totals" }
            }
          ));
          container.appendChild(renderNetworkSection({
            title: "Rete dei reciproci positivi",
            note: "Il grafo mostra solo i legami reciproci positivi.",
            names: analysis.names,
            matrix: analysis.matrices.reciprociPos,
            strengthMatrix: analysis.matrices.forzaLegami,
            nodePalette: PALETTE.network,
            edgePalette: ["#d5d8dc", "#1a9850"],
            toggleId: "link-strength-toggle-pos"
          }));
        };
      case "reciproci-neg":
        return (analysis, container) => {
          container.appendChild(renderReciprociSection(
            "Reciproci negativi",
            "1 indica un reciproco negativo.",
            analysis.names,
            analysis.matrices.reciprociNeg,
            {
              matrix: { palette: PALETTE.negMatrix, scope: "matrix" },
              totals: { palette: PALETTE.negTotals, scope: "totals" }
            }
          ));
          container.appendChild(renderNetworkSection({
            title: "Rete dei reciproci negativi",
            note: "Il grafo mostra solo i legami reciproci negativi.",
            names: analysis.names,
            matrix: analysis.matrices.reciprociNeg,
            strengthMatrix: analysis.matrices.forzaAntagonismo,
            nodePalette: [...PALETTE.network].reverse(),
            edgePalette: ["#d5d8dc", "#d73027"],
            toggleId: "link-strength-toggle-neg"
          }));
        };
      case "forza-legami":
        return (analysis, container) => {
          container.appendChild(renderMatrixSection(
            "Forza legami",
            "Minimo tra scelte reciproche positive.",
            analysis.names,
            analysis.matrices.forzaLegami,
            analysis.totals.forzaLegamiRowSum,
            null,
            "Totale forza",
            {
              matrix: { palette: PALETTE.posMatrix, scope: "combined" },
              totals: { palette: PALETTE.posMatrix, scope: "combined" }
            },
            false
          ));
        };
      case "forza-antagonismo":
        return (analysis, container) => {
          container.appendChild(renderMatrixSection(
            "Forza antagonismo",
            "Minimo tra scelte reciproche negative.",
            analysis.names,
            analysis.matrices.forzaAntagonismo,
            analysis.totals.forzaAntagonismoRowSum,
            null,
            "Totale forza",
            {
              matrix: { palette: PALETTE.negMatrix, scope: "combined" },
              totals: { palette: PALETTE.negMatrix, scope: "combined" }
            },
            false
          ));
        };
      case "non-ricambiate":
        return (analysis, container) => {
          container.appendChild(renderMatrixSection(
            "Positive non ricambiate",
            "Scelte positive non ricambiate.",
            analysis.names,
            analysis.matrices.nonRicambiatePos,
            analysis.totals.nonRicambiateRowSum,
            analysis.totals.nonRicambiateColSum,
            "Positive date e non ricambiate",
            {
              matrix: { palette: PALETTE.nonRic, scope: "combined" },
              totals: { palette: PALETTE.nonRic, scope: "combined" }
            },
            true,
            "Positive ricevute non ricambiate"
          ));
        };
      case "riassunto":
        return (analysis, container) => {
          container.appendChild(renderSummarySection(analysis));
        };
      case "confronto":
        return (_, container) => {
          container.appendChild(renderCompareSection());
        };
      default:
        return () => undefined;
    }
  }

  function appendNetworkForView(analysis, container, view) {
    const config = getNetworkConfigForView(view, analysis);
    if (!config) {
      return;
    }
    container.appendChild(renderNetworkSection(config));
  }

  function getNetworkConfigForView(view, analysis) {
    const mapping = {
      "matrix-pos": {
        title: "Grafo scelte positive",
        note: "Il grafo mostra le scelte positive ricevute.",
        matrix: transposeMatrix(analysis.matrices.posMatrix),
        strengthMatrix: transposeMatrix(analysis.matrices.posMatrix),
        nodePalette: PALETTE.network,
        edgePalette: ["#d5d8dc", "#1a9850"]
      },
      "matrix-neg": {
        title: "Grafo scelte negative",
        note: "Il grafo mostra le scelte negative ricevute.",
        matrix: transposeMatrix(analysis.matrices.negMatrix),
        strengthMatrix: transposeMatrix(analysis.matrices.negMatrix),
        nodePalette: [...PALETTE.network].reverse(),
        edgePalette: ["#d5d8dc", "#d73027"]
      },
      "reciproci-pos": {
        title: "Rete dei reciproci positivi",
        note: "Il grafo mostra solo i legami reciproci positivi.",
        matrix: analysis.matrices.reciprociPos,
        strengthMatrix: analysis.matrices.forzaLegami,
        nodePalette: PALETTE.network,
        edgePalette: ["#d5d8dc", "#1a9850"]
      },
      "reciproci-neg": {
        title: "Rete dei reciproci negativi",
        note: "Il grafo mostra solo i legami reciproci negativi.",
        matrix: analysis.matrices.reciprociNeg,
        strengthMatrix: analysis.matrices.forzaAntagonismo,
        nodePalette: [...PALETTE.network].reverse(),
        edgePalette: ["#d5d8dc", "#d73027"]
      },
      "forza-legami": {
        title: "Grafo forza legami",
        note: "Il grafo usa la forza dei legami positivi.",
        matrix: analysis.matrices.forzaLegami,
        strengthMatrix: analysis.matrices.forzaLegami,
        nodePalette: PALETTE.network,
        edgePalette: ["#d5d8dc", "#1a9850"]
      },
      "forza-antagonismo": {
        title: "Grafo forza antagonismo",
        note: "Il grafo usa la forza delle relazioni negative.",
        matrix: analysis.matrices.forzaAntagonismo,
        strengthMatrix: analysis.matrices.forzaAntagonismo,
        nodePalette: [...PALETTE.network].reverse(),
        edgePalette: ["#d5d8dc", "#d73027"]
      },
      "non-ricambiate": {
        title: "Grafo positive non ricambiate",
        note: "Il grafo mostra le scelte positive non ricambiate.",
        matrix: analysis.matrices.nonRicambiatePos,
        strengthMatrix: analysis.matrices.nonRicambiatePos,
        nodePalette: PALETTE.network,
        edgePalette: ["#d5d8dc", "#1a9850"]
      },
      "responses": {
        title: "Grafo scelte positive",
        note: "Il grafo mostra le scelte positive tra atlete.",
        matrix: analysis.matrices.posMatrix,
        strengthMatrix: analysis.matrices.posMatrix,
        nodePalette: PALETTE.network,
        edgePalette: ["#d5d8dc", "#1a9850"]
      },
      "riassunto": {
        title: "Grafo scelte positive",
        note: "Il grafo mostra le scelte positive tra atlete.",
        matrix: analysis.matrices.posMatrix,
        strengthMatrix: analysis.matrices.posMatrix,
        nodePalette: PALETTE.network,
        edgePalette: ["#d5d8dc", "#1a9850"]
      }
    };

    const config = mapping[view];
    if (!config) {
      return null;
    }
    return {
      ...config,
      names: analysis.names,
      toggleId: `link-strength-toggle-${view}-${Math.random().toString(16).slice(2, 8)}`
    };
  }

  function transposeMatrix(matrix) {
    const size = matrix.length;
    if (!size) {
      return [];
    }
    const result = createMatrix(size, 0);
    for (let i = 0; i < size; i += 1) {
      for (let j = 0; j < size; j += 1) {
        result[j][i] = matrix[i][j];
      }
    }
    return result;
  }

  function getCompareDatasets() {
    const visibleDatasets = getVisibleDatasets();
    if (visibleDatasets.length < 2) {
      return null;
    }
    const datasetA = visibleDatasets.find((dataset) => dataset.id === compareSelectA.value);
    const datasetB = visibleDatasets.find((dataset) => dataset.id === compareSelectB.value);
    if (!datasetA || !datasetB || datasetA.id === datasetB.id) {
      return null;
    }
    return { datasetA, datasetB };
  }

  function renderDualView(view, { datasetA, datasetB }) {
    viewContainer.classList.add("dual-view");
    if (!renderers.has(view)) {
      renderers.set(view, createRenderer(view));
    }
    const renderer = renderers.get(view);
    const panels = [];
    [datasetA, datasetB].forEach((dataset) => {
      const panel = document.createElement("div");
      panel.className = "view-panel";
      const label = document.createElement("p");
      label.className = "note";
      label.textContent = dataset.name;
      panel.appendChild(label);
      renderer(dataset.analysis, panel);
      viewContainer.appendChild(panel);
      panels.push(panel);
    });
    if (panels.length === 2 && view !== "giocatrice" && view !== "capitano") {
      highlightDifferences(panels[0], panels[1]);
      syncDualScroll(panels[0], panels[1]);
    }
  }

  function highlightDifferences(panelA, panelB) {
    const tablesA = Array.from(panelA.querySelectorAll("table"));
    const tablesB = Array.from(panelB.querySelectorAll("table"));
    const count = Math.min(tablesA.length, tablesB.length);
    for (let i = 0; i < count; i += 1) {
      compareTables(tablesA[i], tablesB[i]);
    }
  }

  function syncDualScroll(panelA, panelB) {
    const wrapsA = Array.from(panelA.querySelectorAll(".table-wrap"));
    const wrapsB = Array.from(panelB.querySelectorAll(".table-wrap"));
    const count = Math.min(wrapsA.length, wrapsB.length);
    for (let i = 0; i < count; i += 1) {
      attachScrollSync(wrapsA[i], wrapsB[i]);
    }
  }

  function attachScrollSync(elA, elB) {
    const key = "data-sync-attached";
    if (elA.getAttribute(key) && elB.getAttribute(key)) {
      return;
    }
    let isSyncing = false;
    const sync = (source, target) => {
      if (isSyncing) {
        return;
      }
      isSyncing = true;
      target.scrollLeft = source.scrollLeft;
      target.scrollTop = source.scrollTop;
      requestAnimationFrame(() => {
        isSyncing = false;
      });
    };
    elA.addEventListener("scroll", () => sync(elA, elB));
    elB.addEventListener("scroll", () => sync(elB, elA));
    elA.setAttribute(key, "true");
    elB.setAttribute(key, "true");
  }

  function compareTables(tableA, tableB) {
    const rowsA = getTableDataRows(tableA);
    const rowsB = getTableDataRows(tableB);
    const rowCount = Math.min(rowsA.length, rowsB.length);
    for (let r = 0; r < rowCount; r += 1) {
      const cellsA = rowsA[r];
      const cellsB = rowsB[r];
      const cellCount = Math.min(cellsA.length, cellsB.length);
      for (let c = 0; c < cellCount; c += 1) {
        const cellA = cellsA[c];
        const cellB = cellsB[c];
        const valueA = parseCellValue(cellA);
        const valueB = parseCellValue(cellB);
        if (valueA !== null && valueB !== null) {
          if (valueA !== valueB) {
            markDiffCell(cellA, 0);
            markDiffCell(cellB, valueB - valueA);
          }
        } else if (cellA.textContent.trim() !== cellB.textContent.trim()) {
          markDiffCell(cellA, 0);
          markDiffCell(cellB, 0);
        }
      }
    }
  }

  function getTableDataRows(table) {
    const rows = Array.from(table.querySelectorAll("tbody tr, tfoot tr"));
    return rows.map((row) => Array.from(row.querySelectorAll("td")));
  }

  function parseCellValue(cell) {
    if (!cell) {
      return null;
    }
    if (cell.dataset && cell.dataset.rank) {
      const rank = Number(cell.dataset.rank);
      return Number.isFinite(rank) ? rank : null;
    }
    const text = cell.textContent;
    if (!text) {
      return null;
    }
    const cleaned = text.trim().replace(",", ".");
    if (cleaned === "—") {
      return null;
    }
    if (cleaned.endsWith("%")) {
      const num = Number.parseFloat(cleaned.slice(0, -1));
      return Number.isFinite(num) ? num / 100 : null;
    }
    const value = Number.parseFloat(cleaned.replace(/[^0-9.-]/g, ""));
    return Number.isFinite(value) ? value : null;
  }

  function markDiffCell(cell, delta) {
    if (!Number.isFinite(delta) || delta === 0) {
      return;
    }
    const arrow = document.createElement("span");
    arrow.className = delta > 0 ? "diff-arrow diff-up" : "diff-arrow diff-down";
    arrow.textContent = delta > 0 ? "▲" : "▼";
    cell.appendChild(arrow);
  }

  function renderMatrixSection(title, description, names, matrix, rowTotals, colTotals, totalLabel, colorConfig, includeTotalsRow, totalsRowLabel = "Totali") {
    const section = document.createElement("div");
    const heading = document.createElement("h2");
    heading.textContent = title;
    const note = document.createElement("p");
    note.className = "note";
    note.textContent = description;
    section.appendChild(heading);
    section.appendChild(note);

    const table = document.createElement("table");
    table.className = "matrix-table";
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    const firstTh = document.createElement("th");
    firstTh.innerHTML = "Chi viene scelto →<br>Chi sceglie ↓";
    headerRow.appendChild(firstTh);

    names.forEach((name) => {
      const th = document.createElement("th");
      setHeaderText(th, name);
      headerRow.appendChild(th);
    });

    const totalTh = document.createElement("th");
    setHeaderText(totalTh, totalLabel);
    totalTh.classList.add("sticky-right");
    headerRow.appendChild(totalTh);

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    const minMax = getMinMaxConfig(matrix, rowTotals, colTotals, colorConfig);
    matrix.forEach((row, i) => {
      const tr = document.createElement("tr");
      const rowHeader = document.createElement("th");
      rowHeader.textContent = names[i];
      tr.appendChild(rowHeader);

      row.forEach((cell) => {
        const td = document.createElement("td");
        td.textContent = cell;
        if (colorConfig?.matrix) {
          const range = minMax.matrix;
          applyHeatmap(td, cell, range.min, range.max, colorConfig.matrix.palette);
        }
        tr.appendChild(td);
      });

      const total = document.createElement("td");
      total.textContent = rowTotals[i];
      total.classList.add("sticky-right");
      if (colorConfig?.totals) {
        const range = minMax.totals;
        applyHeatmap(total, rowTotals[i], range.min, range.max, colorConfig.totals.palette);
      }
      tr.appendChild(total);

      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    if (includeTotalsRow && colTotals) {
      const tfoot = document.createElement("tfoot");
      const tr = document.createElement("tr");
      const label = document.createElement("th");
      label.textContent = totalsRowLabel;
      tr.appendChild(label);

      colTotals.forEach((value) => {
        const td = document.createElement("td");
        td.textContent = value;
        if (colorConfig?.totals) {
          const range = minMax.totals;
          applyHeatmap(td, value, range.min, range.max, colorConfig.totals.palette);
        }
        tr.appendChild(td);
      });

      const total = document.createElement("td");
      total.textContent = colTotals.reduce((acc, value) => acc + value, 0);
      total.classList.add("sticky-right");
      tr.appendChild(total);

      tfoot.appendChild(tr);
      table.appendChild(tfoot);
    }

    const wrapper = document.createElement("div");
    wrapper.className = "table-wrap";
    wrapper.appendChild(table);
    section.appendChild(wrapper);

    return section;
  }

  function renderReciprociSection(title, description, names, matrix, colorConfig) {
    const section = document.createElement("div");
    const heading = document.createElement("h2");
    heading.textContent = title;
    const note = document.createElement("p");
    note.className = "note";
    note.textContent = description;
    section.appendChild(heading);
    section.appendChild(note);

    const table = document.createElement("table");
    table.className = "matrix-table";
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    const firstTh = document.createElement("th");
    firstTh.innerHTML = "Chi viene scelto →<br>Chi sceglie ↓";
    headerRow.appendChild(firstTh);

    names.forEach((name) => {
      const th = document.createElement("th");
      setHeaderText(th, name);
      headerRow.appendChild(th);
    });

    const totalTh = document.createElement("th");
    setHeaderText(totalTh, "Totale reciprocità");
    totalTh.classList.add("sticky-right");
    headerRow.appendChild(totalTh);

    const idxTh = document.createElement("th");
    setHeaderText(idxTh, "Indice reciprocità");
    headerRow.appendChild(idxTh);

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    const rowTotals = sumRows(matrix);
    const colTotals = sumColumns(matrix);
    const reciprocalIndex = rowTotals.map((total, idx) => {
      const denom = matrix.length - 1;
      return denom > 0 ? total / denom : 0;
    });

    const minMax = getMinMaxConfig(matrix, rowTotals, colTotals, colorConfig);
    matrix.forEach((row, i) => {
      const tr = document.createElement("tr");
      const rowHeader = document.createElement("th");
      rowHeader.textContent = names[i];
      tr.appendChild(rowHeader);

      row.forEach((cell) => {
        const td = document.createElement("td");
        td.textContent = cell;
        if (colorConfig?.matrix) {
          const range = minMax.matrix;
          applyHeatmap(td, cell, range.min, range.max, colorConfig.matrix.palette);
        }
        tr.appendChild(td);
      });

      const total = document.createElement("td");
      total.textContent = rowTotals[i];
      total.classList.add("sticky-right");
      if (colorConfig?.totals) {
        const range = minMax.totals;
        applyHeatmap(total, rowTotals[i], range.min, range.max, colorConfig.totals.palette);
      }
      tr.appendChild(total);

      const idxCell = document.createElement("td");
      idxCell.textContent = formatPercent(reciprocalIndex[i]);
      tr.appendChild(idxCell);

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);

    const tfoot = document.createElement("tfoot");
    const tr = document.createElement("tr");
    const label = document.createElement("th");
    label.textContent = "Totali";
    tr.appendChild(label);

    colTotals.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      if (colorConfig?.totals) {
        const range = minMax.totals;
        applyHeatmap(td, value, range.min, range.max, colorConfig.totals.palette);
      }
      tr.appendChild(td);
    });

    const totalSum = rowTotals.reduce((acc, value) => acc + value, 0);
    const totalCell = document.createElement("td");
    totalCell.textContent = totalSum;
    totalCell.classList.add("sticky-right");
    tr.appendChild(totalCell);

    const avgIdx = reciprocalIndex.reduce((acc, value) => acc + value, 0) / reciprocalIndex.length;
    const avgCell = document.createElement("td");
    avgCell.textContent = formatPercent(avgIdx);
    tr.appendChild(avgCell);

    tfoot.appendChild(tr);
    table.appendChild(tfoot);

    const wrapper = document.createElement("div");
    wrapper.className = "table-wrap";
    wrapper.appendChild(table);
    section.appendChild(wrapper);

    return section;
  }

  function setHeaderText(cell, text) {
    const value = text == null ? "" : String(text);
    const words = value.trim().split(/\s+/).filter(Boolean);
    cell.textContent = "";
    if (words.length <= 1) {
      cell.textContent = value;
      return;
    }
    const lines = [];
    words.forEach((word) => {
      if (!lines.length) {
        lines.push(word);
        return;
      }
      if (word.length <= 1) {
        lines[lines.length - 1] = `${lines[lines.length - 1]} ${word}`;
        return;
      }
      lines.push(word);
    });
    lines.forEach((line) => {
      const span = document.createElement("span");
      span.textContent = line;
      span.style.display = "block";
      cell.appendChild(span);
    });
  }

  async function buildStaticHtml(dataset) {
    const baseHtml = document.documentElement.outerHTML;
    const [styles, appJs] = await Promise.all([
      getAssetText("styles.css", ".css", "Seleziona styles.css"),
      getAssetText("app.js", ".js", "Seleziona app.js")
    ]);

    const exportedDatasets = datasets.length ? datasets : [dataset];
    const state = {
      currentDatasetId,
      compareA: compareSelectA?.value || null,
      compareB: compareSelectB?.value || null
    };
    const activeView = getActiveView() || restoreView() || "responses";
    const baseOrder = [
      state.compareA,
      state.compareB,
      ...exportedDatasets.map((entry) => entry.id)
    ].filter(Boolean);
    const exportOrder = Array.from(new Set(baseOrder));

    const datasetsJson = escapeScriptTag(JSON.stringify(exportedDatasets.map((entry) => ({
      id: entry.id,
      name: entry.name,
      csv: entry.csv
    }))));
    const exportOrderJson = escapeScriptTag(JSON.stringify(exportOrder));
    const stateJson = escapeScriptTag(JSON.stringify(state));
    const viewJson = escapeScriptTag(JSON.stringify(activeView));

    const preload = `
      window.TEAMWEAVE_DISABLE_SW = true;
      window.TEAMWEAVE_EXPORT_MODE = true;
      window.TEAMWEAVE_EXPORT_ORDER = ${exportOrderJson};
      (function() {
        const datasets = ${datasetsJson};
        const state = ${stateJson};
        const view = ${viewJson};
        localStorage.setItem(${JSON.stringify(DATASETS_KEY)}, JSON.stringify(datasets));
        localStorage.setItem(${JSON.stringify(DATASET_STATE_KEY)}, JSON.stringify(state));
        localStorage.setItem(${JSON.stringify(VIEW_KEY)}, view);
      })();
    `;

    const safeAppJs = escapeScriptTag(appJs);
    const inlineScript = `<script>${preload}\n${safeAppJs}</script>`;
    const inlineStyles = `<style>\n${styles}\n</style>`;

    const manifestPattern = new RegExp("<link[^>]*manifest\\.json[^>]*>", "i");
    const stylesPattern = new RegExp("<link[^>]*styles\\.css[^>]*>", "i");
    const scriptPattern = new RegExp("<script[^>]*app\\.js[^>]*><" + "/script>", "i");

    let html = `<!doctype html>\n${baseHtml}`
      .replace(manifestPattern, "")
      .replace(stylesPattern, inlineStyles)
      .replace(scriptPattern, inlineScript);

    if (!html.includes(inlineScript)) {
      html = html.replace("</body>", `${inlineScript}\n</body>`);
    }
    if (!html.includes(inlineStyles)) {
      html = html.replace("</head>", `${inlineStyles}\n</head>`);
    }

    return html;
  }

  async function getAssetText(url, accept, label) {
    try {
      const response = await fetch(url, { cache: "no-store" });
      if (response.ok) {
        return await response.text();
      }
    } catch (error) {
      // Fallback to inline content for blocked fetch.
    }
    if (url.endsWith("styles.css") && INLINE_STYLES !== "__INLINE_STYLES__") {
      return INLINE_STYLES;
    }
    if (url.endsWith("app.js") && INLINE_APP !== "__INLINE_APP__") {
      return INLINE_APP;
    }
    throw new Error(`Impossibile recuperare: ${label}`);
  }

  function escapeScriptTag(value) {
    const pattern = new RegExp("<" + "/script", "gi");
    return String(value || "").replace(pattern, "<\\/script");
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function renderSummarySection(analysis) {
    const section = document.createElement("div");
    const heading = document.createElement("h2");
    heading.textContent = "Dati riassuntivi";
    section.appendChild(heading);

    const note = document.createElement("p");
    note.className = "note";
    note.textContent = "Indicatori principali per atleta + classificazione automatica.";
    section.appendChild(note);

    const table = document.createElement("table");
    table.className = "matrix-table";

    const thead = document.createElement("thead");
    const row1 = document.createElement("tr");
    [
      "Atleta",
      "Scelte ricevute + (Tecniche)",
      "Scelte ricevute + (Attitudinali)",
      "Scelte ricevute + (Sociali)",
      "Scelte ricevute + (Totali)",
      "Scelte ricevute - (Tecniche)",
      "Scelte ricevute - (Attitudinali)",
      "Scelte ricevute - (Sociali)",
      "Scelte ricevute - (Totali)",
      "Scelte date +",
      "Scelte date -",
      "Reciproci +",
      "Reciproci -",
      "Forza legami",
      "Forza antagonismo",
      "Positive date e non ricambiate",
      "Positive ricevute e non ricambiate"
    ].forEach((label) => {
      const th = document.createElement("th");
      setHeaderText(th, label);
      row1.appendChild(th);
    });
    thead.appendChild(row1);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    analysis.summaryRows.forEach((row) => {
      const tr = document.createElement("tr");
      const cells = [
        row.name,
        row.positiveReceived.Tecniche,
        row.positiveReceived.Attitudinali,
        row.positiveReceived.Sociali,
        row.positiveReceived.Totali,
        row.negativeReceived.Tecniche,
        row.negativeReceived.Attitudinali,
        row.negativeReceived.Sociali,
        row.negativeReceived.Totali,
        row.positiveGiven,
        row.negativeGiven,
        row.reciprociPos,
        row.reciprociNeg,
        row.forzaLegami,
        row.forzaAntagonismo,
        row.nonRicambiateDate,
        row.nonRicambiateRicevute
      ];

      cells.forEach((value, index) => {
        const cell = index === 0 ? document.createElement("th") : document.createElement("td");
        cell.textContent = value;
        tr.appendChild(cell);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    const tfoot = document.createElement("tfoot");
    const tr = document.createElement("tr");
    const averageLabel = document.createElement("th");
    averageLabel.textContent = "Media di squadra";
    tr.appendChild(averageLabel);

    const avgCells = [
      analysis.averages.positiveReceivedTecniche,
      analysis.averages.positiveReceivedAttitudinali,
      analysis.averages.positiveReceivedSociali,
      analysis.averages.positiveReceivedTotal,
      analysis.averages.negativeReceivedTecniche,
      analysis.averages.negativeReceivedAttitudinali,
      analysis.averages.negativeReceivedSociali,
      analysis.averages.negativeReceivedTotal,
      analysis.averages.positiveGiven,
      analysis.averages.negativeGiven,
      analysis.averages.reciprociPos,
      analysis.averages.reciprociNeg,
      analysis.averages.forzaLegami,
      analysis.averages.forzaAntagonismo,
      analysis.averages.nonRicambiateDate,
      analysis.averages.nonRicambiateRicevute
    ];

    avgCells.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = formatNumber(value);
      tr.appendChild(td);
    });

    tfoot.appendChild(tr);
    table.appendChild(tfoot);

    const wrapper = document.createElement("div");
    wrapper.className = "table-wrap";
    wrapper.appendChild(table);
    section.appendChild(wrapper);

    applySummaryColors(table, analysis);

    const classHeading = document.createElement("h3");
    classHeading.textContent = "Classificazione";
    classHeading.style.marginTop = "2rem";
    section.appendChild(classHeading);

    const classNote = document.createElement("p");
    classNote.className = "note";
    classNote.textContent = "Le influenze per categoria sono confrontate con la rispettiva media di squadra: alta dalla media in su, media da metà media, bassa sotto metà media. Assente indica che non ci sono scelte di quel segno nella categoria. Le etichette descrivono le percezioni espresse nelle risposte; l’ambito attitudinale comprende anche fiducia e leadership, oltre all’impegno.";
    section.appendChild(classNote);

    const classTable = document.createElement("table");
    classTable.className = "matrix-table";
    const classHead = document.createElement("thead");
    const classRow = document.createElement("tr");
    [
      "Atleta", "Influenza positiva (Totale)", "Influenza negativa (Totale)",
      ...CATEGORIES.flatMap((category) => [`Influenza positiva (${category})`, `Influenza negativa (${category})`]),
      "Equilibrio relazionale", "Etichetta sintetica"
    ].forEach((label) => {
      const th = document.createElement("th");
      setHeaderText(th, label);
      classRow.appendChild(th);
    });
    classHead.appendChild(classRow);
    classTable.appendChild(classHead);

    const classBody = document.createElement("tbody");
    const influenceRank = { Bassa: 1, Media: 2, Alta: 3 };
    const equilibrioRank = { Ignorata: 1, Bilanciata: 2, Selettiva: 3 };
    analysis.classifications.forEach((row) => {
      const tr = document.createElement("tr");
      const cells = [
        { value: row.name },
        { value: row.inflPos, mode: "positive" },
        { value: row.inflNeg, mode: "negative" },
        ...CATEGORIES.flatMap((category) => [
          { value: row.byCategory[category].inflPos, mode: "positive" },
          { value: row.byCategory[category].inflNeg, mode: "negative" }
        ]),
        { value: row.equilibrio, mode: "balance" },
        { value: row.label, mode: "label" }
      ];
      cells.forEach(({ value, mode }, idx) => {
        const cell = idx === 0 ? document.createElement("th") : document.createElement("td");
        cell.textContent = value;
        if (mode === "positive" || mode === "negative") {
          applyLabelFill(cell, value, mode);
          cell.dataset.rank = influenceRank[value] || 0;
        } else if (mode === "balance") {
          applyEquilibrioFill(cell, value);
          cell.dataset.rank = equilibrioRank[value] || 0;
        } else if (mode === "label") {
          cell.classList.add("classification-label");
        }
        tr.appendChild(cell);
      });
      classBody.appendChild(tr);
    });
    classTable.appendChild(classBody);

    const classWrap = document.createElement("div");
    classWrap.className = "table-wrap classification-wrap";
    classWrap.appendChild(classTable);
    section.appendChild(classWrap);

    return section;
  }

  function renderCompareSection() {
    const section = document.createElement("div");
    const heading = document.createElement("h2");
    heading.textContent = "Confronto test";
    section.appendChild(heading);

    const note = document.createElement("p");
    note.className = "note";
    note.textContent = "Le differenze sono calcolate come B - A.";
    section.appendChild(note);

    if (datasets.length < 2) {
      const empty = document.createElement("p");
      empty.textContent = "Carica almeno due dataset per vedere il confronto.";
      section.appendChild(empty);
      return section;
    }

    const datasetA = datasets.find((dataset) => dataset.id === compareSelectA.value);
    const datasetB = datasets.find((dataset) => dataset.id === compareSelectB.value);
    if (!datasetA || !datasetB) {
      const empty = document.createElement("p");
      empty.textContent = "Seleziona due dataset validi.";
      section.appendChild(empty);
      return section;
    }

    const title = document.createElement("p");
    title.className = "note";
    title.textContent = `A: ${datasetA.name} | B: ${datasetB.name}`;
    section.appendChild(title);

    const summaryTable = document.createElement("table");
    summaryTable.className = "matrix-table";
    const summaryHead = document.createElement("thead");
    const summaryRow = document.createElement("tr");
    ["Metrica", "A", "B", "Delta"].forEach((label) => {
      const th = document.createElement("th");
      setHeaderText(th, label);
      summaryRow.appendChild(th);
    });
    summaryHead.appendChild(summaryRow);
    summaryTable.appendChild(summaryHead);

    const summaryBody = document.createElement("tbody");
    const summaryMetrics = [
      {
        label: "Risposte",
        get: (analysis) => analysis.rows.length,
        polarity: "neutral"
      },
      {
        label: "Atlete",
        get: (analysis) => analysis.names.length,
        polarity: "neutral"
      },
      {
        label: "Indice reciprocita +",
        get: (analysis) => computeReciprocityIndex(analysis.matrices.reciprociPos),
        polarity: "higher"
      },
      {
        label: "Indice reciprocita -",
        get: (analysis) => computeReciprocityIndex(analysis.matrices.reciprociNeg),
        polarity: "lower"
      }
    ];

    summaryMetrics.forEach((metric) => {
      const tr = document.createElement("tr");
      const label = document.createElement("th");
      label.textContent = metric.label;
      tr.appendChild(label);
      const aValue = metric.get(datasetA.analysis);
      const bValue = metric.get(datasetB.analysis);
      const delta = Number(bValue) - Number(aValue);
      const cells = [formatCompareValue(aValue), formatCompareValue(bValue), formatCompareDelta(delta, metric.label)];
      cells.forEach((value, idx) => {
        const td = document.createElement("td");
        td.textContent = value;
        if (idx === 2) {
          applyDeltaClass(td, delta, metric.polarity);
        }
        tr.appendChild(td);
      });
      summaryBody.appendChild(tr);
    });
    summaryTable.appendChild(summaryBody);

    const summaryWrap = document.createElement("div");
    summaryWrap.className = "table-wrap";
    summaryWrap.appendChild(summaryTable);
    section.appendChild(summaryWrap);

    const detailHeading = document.createElement("h3");
    detailHeading.textContent = "Differenze per atleta";
    detailHeading.style.marginTop = "2rem";
    section.appendChild(detailHeading);

    const compareTable = document.createElement("table");
    compareTable.className = "matrix-table";
    const head = document.createElement("thead");
    const headerRow = document.createElement("tr");
    [
      "Atleta",
      "Delta scelte ricevute + (Totali)",
      "Delta scelte ricevute - (Totali)",
      "Delta reciproci +",
      "Delta reciproci -",
      "Delta positive ricevute non ricambiate",
      "Delta positive date non ricambiate"
    ].forEach((label) => {
      const th = document.createElement("th");
      setHeaderText(th, label);
      headerRow.appendChild(th);
    });
    head.appendChild(headerRow);
    compareTable.appendChild(head);

    const body = document.createElement("tbody");
    const mapA = buildSummaryMap(datasetA.analysis.summaryRows);
    const mapB = buildSummaryMap(datasetB.analysis.summaryRows);
    const allNames = Array.from(
      new Set([...datasetA.analysis.names, ...datasetB.analysis.names].map((name) => normalizeName(name)))
    );
    const nameLookup = new Map();
    datasetA.analysis.names.forEach((name) => nameLookup.set(normalizeName(name), name));
    datasetB.analysis.names.forEach((name) => {
      if (!nameLookup.has(normalizeName(name))) {
        nameLookup.set(normalizeName(name), name);
      }
    });
    allNames.forEach((key) => {
      const rowA = mapA.get(key);
      const rowB = mapB.get(key);
      const tr = document.createElement("tr");
      const nameCell = document.createElement("th");
      nameCell.textContent = nameLookup.get(key) || key;
      tr.appendChild(nameCell);
      const metrics = [
        { value: diffValue(rowA?.positiveReceived?.Totali, rowB?.positiveReceived?.Totali), polarity: "higher" },
        { value: diffValue(rowA?.negativeReceived?.Totali, rowB?.negativeReceived?.Totali), polarity: "lower" },
        { value: diffValue(rowA?.reciprociPos, rowB?.reciprociPos), polarity: "higher" },
        { value: diffValue(rowA?.reciprociNeg, rowB?.reciprociNeg), polarity: "lower" },
        { value: diffValue(rowA?.nonRicambiateRicevute, rowB?.nonRicambiateRicevute), polarity: "lower" },
        { value: diffValue(rowA?.nonRicambiateDate, rowB?.nonRicambiateDate), polarity: "lower" }
      ];
      metrics.forEach((metric) => {
        const td = document.createElement("td");
        td.textContent = formatCompareDelta(metric.value);
        applyDeltaClass(td, metric.value, metric.polarity);
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });
    compareTable.appendChild(body);

    const compareWrap = document.createElement("div");
    compareWrap.className = "table-wrap";
    compareWrap.appendChild(compareTable);
    section.appendChild(compareWrap);

    return section;
  }

  function buildSummaryMap(rows) {
    const map = new Map();
    rows.forEach((row) => {
      map.set(normalizeName(row.name), row);
    });
    return map;
  }

  function diffValue(valueA, valueB) {
    if (!Number.isFinite(valueA) || !Number.isFinite(valueB)) {
      return null;
    }
    return valueB - valueA;
  }

  function formatCompareValue(value) {
    if (!Number.isFinite(value)) {
      return "—";
    }
    if (value % 1 !== 0 && value <= 1) {
      return formatPercent(value);
    }
    return formatNumber(value);
  }

  function formatCompareDelta(value, label) {
    if (!Number.isFinite(value)) {
      return "—";
    }
    const sign = value > 0 ? "+" : "";
    if (label && label.includes("Indice")) {
      return `${sign}${formatPercent(value)}`;
    }
    return `${sign}${formatNumber(value)}`;
  }

  function applyDeltaClass(cell, value, polarity) {
    if (!Number.isFinite(value) || value === 0 || polarity === "neutral") {
      return;
    }
    const isPositive = value > 0;
    if (polarity === "higher") {
      cell.classList.add(isPositive ? "delta-good" : "delta-bad");
      return;
    }
    if (polarity === "lower") {
      cell.classList.add(isPositive ? "delta-bad" : "delta-good");
    }
  }

  function analyzeCaptainPairs(analysis) {
    const positive = buildReciprociMatrix(analysis.matrices.leadershipPos);
    const influenceRanks = { Assente: 0, Bassa: 1, Media: 2, Alta: 3 };
    const profiles = analysis.classifications.map((row) => ({
      attitude: influenceRanks[row.byCategory.Attitudinali.inflPos] || 0,
      social: influenceRanks[row.byCategory.Sociali.inflPos] || 0
    }));
    const pairs = [];
    for (let first = 0; first < analysis.names.length; first += 1) {
      for (let second = first + 1; second < analysis.names.length; second += 1) {
        const onlyFirst = [], onlySecond = [], shared = [], uncovered = [];
        analysis.names.forEach((name, index) => {
          if (index === first || index === second) return;
          const fromFirst = positive[first][index] > 0;
          const fromSecond = positive[second][index] > 0;
          if (fromFirst && fromSecond) shared.push(index);
          else if (fromFirst) onlyFirst.push(index);
          else if (fromSecond) onlySecond.push(index);
          else uncovered.push(index);
        });
        const covered = onlyFirst.length + onlySecond.length + shared.length;
        pairs.push({
          first, second, onlyFirst, onlySecond, shared, uncovered, covered,
          totalPeers: Math.max(0, analysis.names.length - 2),
          degreeFirst: onlyFirst.length + shared.length, degreeSecond: onlySecond.length + shared.length,
          attitudeFloor: Math.min(profiles[first].attitude, profiles[second].attitude),
          socialFloor: Math.min(profiles[first].social, profiles[second].social),
          profileSum: profiles[first].attitude + profiles[second].attitude + profiles[first].social + profiles[second].social,
        });
      }
    }
    return pairs;
  }

  function rankCaptainPairs(pairs) {
    return [...pairs].sort((a, b) => b.attitudeFloor - a.attitudeFloor
      || b.socialFloor - a.socialFloor
      || b.profileSum - a.profileSum
      || b.covered - a.covered);
  }

  function renderCaptainPairSection(analysis) {
    const section = document.createElement("div");
    section.className = "captain-section";
    const addText = (parent, tag, text, className) => {
      const element = document.createElement(tag);
      element.textContent = text;
      if (className) element.className = className;
      parent.appendChild(element);
      return element;
    };
    addText(section, "h2", "Capitano e vicecapitano");
    addText(section, "p", "Prima il profilo positivo attitudinale e sociale di entrambe le candidate, poi la copertura delle compagne attraverso reciproci positivi. Le scelte tecniche sono escluse da questa analisi.", "note");
    const pairs = analyzeCaptainPairs(analysis);
    if (!pairs.length) {
      addText(section, "p", "Servono almeno due giocatrici per confrontare le coppie.", "note");
      return section;
    }
    const controls = document.createElement("div");
    controls.className = "captain-controls";
    section.appendChild(controls);
    const makeSelect = (label, options) => {
      const control = document.createElement("label");
      control.className = "player-control";
      control.appendChild(document.createTextNode(label));
      const select = document.createElement("select");
      options.forEach(([value, text]) => select.add(new Option(text, value)));
      control.appendChild(select);
      controls.appendChild(control);
      return select;
    };
    const options = analysis.names.map((name, index) => [String(index), name]);
    const captain = makeSelect("Capitano", options);
    const deputy = makeSelect("Vicecapitano", options);
    const best = rankCaptainPairs(pairs)[0];
    captain.value = String(best.first);
    deputy.value = String(best.second);
    const swap = document.createElement("button");
    swap.type = "button";
    swap.textContent = "Inverti i ruoli";
    controls.appendChild(swap);
    addText(section, "p", "Copertura: altre compagne raggiunte almeno da una candidata, senza duplicati. Si usano solo i reciproci sociali e attitudinali; il rapporto tra capitano e vice non conta nella copertura né come criterio di scelta.", "note");
    addText(section, "p", "Primo criterio: si confronta il livello positivo attitudinale più basso della coppia, poi quello sociale più basso, infine la somma dei quattro livelli (Assente = 0, Bassa = 1, Media = 2, Alta = 3). Così una candidata forte non compensa un livello debole dell’altra. Secondo criterio, a parità di profilo: copertura senza duplicati. A parità di profilo e copertura le coppie sono equivalenti, senza altri criteri di spareggio. I ruoli restano a tua scelta e non modificano il punteggio.", "note");
    const recommendation = addText(section, "p", "", "captain-recommendation");
    const detail = document.createElement("div");
    section.appendChild(detail);
    addText(section, "h3", "Classifica delle coppie");
    const ranking = document.createElement("div");
    section.appendChild(ranking);
    const renderDetail = () => {
      const first = Number(captain.value), second = Number(deputy.value);
      Array.from(captain.options).forEach((option) => { option.disabled = Number(option.value) === second; });
      Array.from(deputy.options).forEach((option) => { option.disabled = Number(option.value) === first; });
      const pair = pairs.find((item) => item.first === Math.min(first, second) && item.second === Math.max(first, second));
      detail.replaceChildren();
      if (!pair) return;
      addText(detail, "h3", `${analysis.names[first]} · ${analysis.names[second]}`);
      const stats = document.createElement("div");
      stats.className = "player-stats";
      [
        ["Minimo attitudinale +", ["Assente", "Bassa", "Media", "Alta"][pair.attitudeFloor]],
        ["Minimo sociale +", ["Assente", "Bassa", "Media", "Alta"][pair.socialFloor]],
        ["Compagne coperte", `${pair.covered} / ${pair.totalPeers}`],
        ["Copertura", pair.totalPeers ? formatPercent(pair.covered / pair.totalPeers) : "Non applicabile"],
        ["Compagne in comune", pair.shared.length],
        ["Compagne non coperte", pair.uncovered.length]
      ].forEach(([label, value]) => {
        const card = document.createElement("div");
        addText(card, "span", label);
        addText(card, "strong", String(value));
        stats.appendChild(card);
      });
      detail.appendChild(stats);
      const onlyCaptain = first === pair.first ? pair.onlyFirst : pair.onlySecond;
      const onlyDeputy = first === pair.first ? pair.onlySecond : pair.onlyFirst;
      const groups = [
        [`Solo ${analysis.names[first]}`, onlyCaptain, "captain-only"],
        [`Entrambe`, pair.shared, "captain-shared"],
        [`Solo ${analysis.names[second]}`, onlyDeputy, "deputy-only"],
        ["Non coperte", pair.uncovered, "captain-uncovered"]
      ];
      addText(detail, "h4", "Come si distribuisce la copertura");
      const chart = document.createElement("div");
      chart.className = "captain-coverage";
      chart.setAttribute("role", "img");
      chart.setAttribute("aria-label", groups.map(([label, members]) => `${label}: ${members.length}`).join("; "));
      groups.forEach(([label, members, className]) => {
        if (!members.length) return;
        const segment = document.createElement("div");
        segment.className = className;
        segment.style.flexGrow = members.length;
        segment.title = `${label}: ${members.length}`;
        segment.textContent = String(members.length);
        chart.appendChild(segment);
      });
      if (pair.totalPeers) detail.appendChild(chart);
      else addText(detail, "p", "Non ci sono altre compagne su cui calcolare la copertura.", "note");
      const groupList = document.createElement("div");
      groupList.className = "captain-groups";
      groups.forEach(([label, members, className]) => {
        const card = document.createElement("div");
        card.className = className;
        addText(card, "h4", `${label} (${members.length})`);
        addText(card, "p", members.map((index) => analysis.names[index]).join(", ") || "Nessuna");
        groupList.appendChild(card);
      });
      detail.appendChild(groupList);
      addText(detail, "h4", "Indicatori per assegnare i ruoli");
      detail.appendChild(renderPlayerTable(
        ["Giocatrice", "Ruolo scelto", "Influenza attitudinale +", "Influenza sociale +", "Altre compagne coperte", "Attitudinali ricevute +", "Attitudinali ricevute −", "Sociali ricevute +", "Sociali ricevute −", "Profilo attitudinale"],
        [first, second].map((index, position) => {
          const row = analysis.summaryRows[index];
          return [analysis.names[index], position ? "Vicecapitano" : "Capitano", analysis.classifications[index].byCategory.Attitudinali.inflPos, analysis.classifications[index].byCategory.Sociali.inflPos, index === pair.first ? pair.degreeFirst : pair.degreeSecond, row.positiveReceived.Attitudinali, row.negativeReceived.Attitudinali, row.positiveReceived.Sociali, row.negativeReceived.Sociali, analysis.classifications[index].byCategory.Attitudinali.label];
        })
      ));
      Array.from(ranking.querySelectorAll("tbody tr")).forEach((row) => {
        const active = row.dataset.pair === `${pair.first}-${pair.second}`;
        row.classList.toggle("captain-selected", active);
        row.querySelector("button")?.setAttribute("aria-pressed", String(active));
      });
    };
    const renderRanking = () => {
      const ordered = rankCaptainPairs(pairs);
      const top = ordered[0];
      const tied = ordered.filter((pair) => pair.attitudeFloor === top.attitudeFloor
        && pair.socialFloor === top.socialFloor && pair.profileSum === top.profileSum && pair.covered === top.covered).length;
      recommendation.textContent = `${tied === 1 ? "Migliore coppia" : "Una delle migliori coppie"} per profilo attitudinale e sociale, poi copertura: ${analysis.names[top.first]} e ${analysis.names[top.second]}. Copre ${top.covered} su ${top.totalPeers} compagne.`;
      if (tied > 1) recommendation.textContent += ` ${tied} coppie sono a pari merito.`;
      if (!ordered.some((pair) => pair.attitudeFloor >= 2 && pair.socialFloor >= 2)) {
        recommendation.textContent += " Nessuna coppia ha entrambe le influenze positive almeno medie per tutte e due le candidate: verifica i profili prima della scelta.";
      }
      if (!ordered.some((pair) => pair.covered > 0)) {
        recommendation.textContent += " Nessuna coppia copre altre compagne con reciproci sociali o attitudinali: la copertura non distingue le coppie.";
      }
      const table = renderPlayerTable(
        ["Coppia", "Minimo attitudinale +", "Minimo sociale +", "Somma livelli +", "Compagne coperte", "Copertura", "Coperte dalla prima", "Coperte dalla seconda", "In comune", "Dettaglio"],
        ordered.map((pair) => [
          `${analysis.names[pair.first]} + ${analysis.names[pair.second]}`,
          ["Assente", "Bassa", "Media", "Alta"][pair.attitudeFloor],
          ["Assente", "Bassa", "Media", "Alta"][pair.socialFloor], pair.profileSum,
          `${pair.covered} / ${pair.totalPeers}`, pair.totalPeers ? formatPercent(pair.covered / pair.totalPeers) : "—",
          pair.degreeFirst, pair.degreeSecond, pair.shared.length, ""
        ])
      );
      table.classList.add("captain-ranking");
      table.querySelectorAll("tbody tr").forEach((row, index) => {
        const pair = ordered[index];
        row.dataset.pair = `${pair.first}-${pair.second}`;
        const button = document.createElement("button");
        button.type = "button";
        button.textContent = "Esamina";
        button.setAttribute("aria-label", `Esamina ${analysis.names[pair.first]} e ${analysis.names[pair.second]}`);
        button.addEventListener("click", () => {
          captain.value = String(pair.first);
          deputy.value = String(pair.second);
          renderDetail();
        });
        row.lastElementChild.appendChild(button);
      });
      ranking.replaceChildren(table);
      renderDetail();
    };
    captain.addEventListener("change", renderDetail);
    deputy.addEventListener("change", renderDetail);
    swap.addEventListener("click", () => {
      const previous = captain.value;
      captain.value = deputy.value;
      deputy.value = previous;
      renderDetail();
    });
    renderRanking();
    return section;
  }

  function analyzeCouncil(analysis, members) {
    const unique = [...new Set(members)].filter((index) => Number.isInteger(index) && index >= 0 && index < analysis.names.length);
    const selected = new Set(unique);
    const ranks = { Assente: 0, Bassa: 1, Media: 2, Alta: 3 };
    const matrix = analysis.matrices.leadershipPos;
    const peers = analysis.names.flatMap((name, index) => selected.has(index) ? [] : [{
      index,
      reachedBy: unique.filter((member) => matrix[member][index] > 0 && matrix[index][member] > 0)
    }]);
    const attitudes = unique.map((index) => ranks[analysis.classifications[index].byCategory.Attitudinali.inflPos] || 0);
    const socials = unique.map((index) => ranks[analysis.classifications[index].byCategory.Sociali.inflPos] || 0);
    return {
      members: unique, peers,
      covered: peers.filter((peer) => peer.reachedBy.length > 0).length,
      totalPeers: peers.length,
      attitudeFloor: attitudes.length ? Math.min(...attitudes) : 0,
      socialFloor: socials.length ? Math.min(...socials) : 0,
      profileSum: [...attitudes, ...socials].reduce((sum, value) => sum + value, 0)
    };
  }

  function* councilCandidates(analysis, size, fixedCaptain = null) {
    const count = analysis.names.length;
    if (!Number.isInteger(size) || size < 1 || size > count) return;
    if (fixedCaptain !== null && (!Number.isInteger(fixedCaptain) || fixedCaptain < 0 || fixedCaptain >= count)) return;
    const pool = analysis.names.map((_, index) => index).filter((index) => index !== fixedCaptain);
    const selected = fixedCaptain === null ? [] : [fixedCaptain];
    function* visit(start) {
      if (selected.length === size) {
        yield analyzeCouncil(analysis, selected);
        return;
      }
      const remaining = size - selected.length;
      for (let index = start; index <= pool.length - remaining; index += 1) {
        selected.push(pool[index]);
        yield* visit(index + 1);
        selected.pop();
      }
    }
    yield* visit(0);
  }

  function renderCouncilSection(analysis) {
    const section = document.createElement("div");
    section.className = "council-section";
    const add = (parent, tag, text, className) => {
      const element = document.createElement(tag);
      element.textContent = text;
      if (className) element.className = className;
      parent.appendChild(element);
      return element;
    };
    add(section, "h2", "Consiglio della squadra");
    add(section, "p", "Imposta il numero totale di consiglieri, capitano incluso: vengono confrontate le composizioni complete e selezionata la migliore. Puoi applicare un’alternativa o modificare i vice manualmente. Il profilo attitudinale e sociale resta il primo criterio; la copertura delle altre compagne il secondo. Tecnica e rapporti interni al consiglio non entrano nella valutazione.", "note");
    if (!analysis.names.length) {
      add(section, "p", "Non ci sono giocatrici da selezionare.", "note");
      return section;
    }
    const initial = rankCaptainPairs(analyzeCaptainPairs(analysis))[0];
    let captainIndex = initial?.first ?? 0;
    const deputies = new Set(initial ? [initial.second] : []);
    const control = document.createElement("label");
    control.className = "player-control";
    control.appendChild(document.createTextNode("Capitano del consiglio"));
    const captain = document.createElement("select");
    analysis.names.forEach((name, index) => captain.add(new Option(name, index)));
    captain.value = String(captainIndex);
    control.appendChild(captain);
    section.appendChild(control);
    const sizeControl = add(section, "label", "Numero totale di consiglieri (capitano incluso)", "player-control");
    const sizeSelect = document.createElement("select");
    sizeSelect.className = "council-size";
    analysis.names.forEach((_, index) => sizeSelect.add(new Option(`${index + 1} — 1 capitano e ${index} vice`, index + 1)));
    sizeSelect.value = String(Math.min(2, analysis.names.length));
    sizeControl.appendChild(sizeSelect);
    const lockLabel = add(section, "label", "", "council-option");
    const lockCaptain = document.createElement("input");
    lockCaptain.type = "checkbox";
    lockCaptain.checked = true;
    lockLabel.appendChild(lockCaptain);
    add(lockLabel, "span", "Mantieni il capitano selezionato nei consigli suggeriti");
    add(section, "p", "Ordine: profilo attitudinale e sociale di tutti i membri, poi copertura. A parità di questi criteri i consigli sono equivalenti. Cambiando numero o capitano viene applicata la migliore composizione completa.", "note");
    const suggestions = add(section, "div", "", "council-suggestions");
    suggestions.setAttribute("aria-live", "polite");
    const fieldset = document.createElement("fieldset");
    fieldset.className = "council-candidates";
    add(fieldset, "legend", "Vicecapitani — modifica la composizione suggerita");
    const candidates = document.createElement("div");
    candidates.className = "council-options";
    fieldset.appendChild(candidates);
    section.appendChild(fieldset);
    const detail = document.createElement("div");
    detail.setAttribute("aria-live", "polite");
    section.appendChild(detail);
    const render = () => {
      const members = [captainIndex, ...deputies];
      const council = analyzeCouncil(analysis, members);
      candidates.replaceChildren();
      analysis.names.forEach((name, index) => {
        const label = document.createElement("label");
        label.className = "council-option";
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = String(index);
        checkbox.checked = deputies.has(index);
        checkbox.disabled = index === captainIndex || (!deputies.has(index) && deputies.size >= Number(sizeSelect.value) - 1);
        label.appendChild(checkbox);
        const caption = document.createElement("span");
        add(caption, "strong", `${name}${index === captainIndex ? " · Capitano" : ""}`);
        const profile = analysis.classifications[index].byCategory;
        add(caption, "small", `Attitudinale +: ${profile.Attitudinali.inflPos} · Sociale +: ${profile.Sociali.inflPos}`);
        label.appendChild(caption);
        candidates.appendChild(label);
        checkbox.addEventListener("change", () => {
          if (checkbox.checked) deputies.add(index);
          else deputies.delete(index);
          render();
          candidates.querySelector(`input[value="${index}"]`)?.focus();
        });
      });
      detail.replaceChildren();
      add(detail, "h3", `Consiglio selezionato: ${members.length} / ${sizeSelect.value} membri — 1 capitano e ${deputies.size} vice`);
      const stats = document.createElement("div");
      stats.className = "player-stats";
      [
        ["Minimo attitudinale +", ["Assente", "Bassa", "Media", "Alta"][council.attitudeFloor]],
        ["Minimo sociale +", ["Assente", "Bassa", "Media", "Alta"][council.socialFloor]],
        ["Compagne coperte", `${council.covered} / ${council.totalPeers}`],
        ["Copertura", council.totalPeers ? formatPercent(council.covered / council.totalPeers) : "Non applicabile"]
      ].forEach(([label, value]) => {
        const card = document.createElement("div");
        add(card, "span", label);
        add(card, "strong", String(value));
        stats.appendChild(card);
      });
      detail.appendChild(stats);
      add(detail, "p", "Ogni compagna esterna al consiglio conta una sola volta, anche se ha reciproci positivi con più membri. Aggiungendo un vice cambia anche il numero di compagne esterne su cui si calcola la percentuale.", "note");
      if (!council.totalPeers) add(detail, "p", "Tutte le giocatrici fanno parte del consiglio: non ci sono compagne esterne su cui calcolare la copertura.", "note");
      add(detail, "h3", "Membri del consiglio");
      detail.appendChild(renderPlayerTable(
        ["Giocatrice", "Ruolo", "Influenza attitudinale +", "Influenza sociale +", "Compagne esterne coperte", "Coperte solo da lei"],
        members.map((index) => [analysis.names[index], index === captainIndex ? "Capitano" : "Vicecapitano", analysis.classifications[index].byCategory.Attitudinali.inflPos, analysis.classifications[index].byCategory.Sociali.inflPos,
          council.peers.filter((peer) => peer.reachedBy.includes(index)).length,
          council.peers.filter((peer) => peer.reachedBy.length === 1 && peer.reachedBy[0] === index).length])
      ));
      add(detail, "h3", "Copertura della squadra");
      const groups = [
        ["Raggiunte da un membro", council.peers.filter((peer) => peer.reachedBy.length === 1), "captain-only"],
        ["Raggiunte da più membri", council.peers.filter((peer) => peer.reachedBy.length > 1), "captain-shared"],
        ["Non coperte", council.peers.filter((peer) => !peer.reachedBy.length), "captain-uncovered"]
      ];
      const chart = document.createElement("div");
      chart.className = "captain-coverage";
      chart.setAttribute("role", "img");
      chart.setAttribute("aria-label", groups.map(([label, peers]) => `${label}: ${peers.length}`).join("; "));
      groups.forEach(([label, peers, tone]) => {
        if (!peers.length) return;
        const segment = add(chart, "div", String(peers.length), tone);
        segment.style.flexGrow = peers.length;
        segment.title = `${label}: ${peers.length}`;
      });
      if (council.totalPeers) detail.appendChild(chart);
      const lists = document.createElement("div");
      lists.className = "captain-groups";
      groups.forEach(([label, peers, tone]) => {
        const group = add(lists, "div", "", tone);
        add(group, "h4", `${label} (${peers.length})`);
        add(group, "p", peers.map((peer) => analysis.names[peer.index]).join(", ") || "Nessuna");
      });
      detail.appendChild(lists);
      if (council.totalPeers) detail.appendChild(renderPlayerTable(["Compagna esterna", "Raggiunta da"], council.peers.map((peer) => [analysis.names[peer.index], peer.reachedBy.map((index) => analysis.names[index]).join(", ") || "Nessun membro"])));
    };
    let calculationVersion = 0;
    const suggest = async () => {
      const version = ++calculationVersion;
      fieldset.disabled = true;
      const size = Number(sizeSelect.value);
      const fixed = lockCaptain.checked ? captainIndex : null;
      suggestions.replaceChildren();
      const status = add(suggestions, "p", "Confronto dei consigli in corso…", "note");
      let best = [], examined = 0, ties = 0;
      const equalScore = (a, b) => a.attitudeFloor === b.attitudeFloor && a.socialFloor === b.socialFloor && a.profileSum === b.profileSum && a.covered === b.covered;
      for (const candidate of councilCandidates(analysis, size, fixed)) {
        const previous = best[0];
        best = rankCaptainPairs([...best, candidate]).slice(0, 10);
        if (!previous || !equalScore(previous, best[0])) ties = 1;
        else if (equalScore(candidate, best[0])) ties += 1;
        examined += 1;
        if (examined % 200 === 0) {
          status.textContent = `Confronto in corso: ${examined} consigli valutati…`;
          await new Promise((resolve) => setTimeout(resolve, 0));
          if (version !== calculationVersion) return;
        }
      }
      if (version !== calculationVersion) return;
      fieldset.disabled = false;
      if (!best.length) {
        status.textContent = "Nessun consiglio disponibile con questa composizione.";
        return;
      }
      const apply = (council) => {
        if (!council.members.includes(captainIndex)) captainIndex = council.members[0];
        captain.value = String(captainIndex);
        deputies.clear();
        council.members.filter((index) => index !== captainIndex).forEach((index) => deputies.add(index));
        render();
      };
      status.textContent = `${examined} consigli completi confrontati, ${ties} a pari merito al primo posto. Mostrati i migliori ${best.length}. ${fixed === null ? "Capitano modificabile tra i membri: il ruolo non cambia la valutazione." : `Capitano mantenuto: ${analysis.names[fixed]}.`}`;
      add(suggestions, "h3", `Consigli suggeriti: ${size} membri (1 capitano e ${size - 1} vice)`);
      const table = renderPlayerTable(
        ["Consiglio", "Minimo attitudinale +", "Minimo sociale +", "Somma livelli +", "Compagne coperte", "Copertura", "Azione"],
        best.map((council) => [council.members.map((index) => analysis.names[index]).join(", "), ["Assente", "Bassa", "Media", "Alta"][council.attitudeFloor], ["Assente", "Bassa", "Media", "Alta"][council.socialFloor], council.profileSum, `${council.covered} / ${council.totalPeers}`, council.totalPeers ? formatPercent(council.covered / council.totalPeers) : "Non applicabile", ""])
      );
      table.classList.add("council-ranking");
      table.querySelectorAll("tbody tr").forEach((row, index) => {
        const button = add(row.lastElementChild, "button", "Usa questo consiglio");
        button.type = "button";
        button.addEventListener("click", () => apply(best[index]));
      });
      suggestions.appendChild(table);
      apply(best[0]);
    };
    sizeSelect.addEventListener("change", () => { suggest(); });
    lockCaptain.addEventListener("change", () => { suggest(); });
    captain.addEventListener("change", () => {
      const next = Number(captain.value);
      if (deputies.delete(next)) deputies.add(captainIndex);
      captainIndex = next;
      suggest();
    });
    render();
    suggest();
    return section;
  }

  function renderCaptainSection(analysis) {
    const section = document.createElement("div");
    const control = document.createElement("label");
    control.className = "player-control";
    control.appendChild(document.createTextNode("Composizione"));
    const mode = document.createElement("select");
    mode.add(new Option("Capitano e un vice", "pair"));
    mode.add(new Option("Consiglio: capitano e più vice", "council"));
    control.appendChild(mode);
    const content = document.createElement("div");
    section.append(control, content);
    const panels = new Map();
    const render = () => {
      if (!panels.has(mode.value)) panels.set(mode.value, mode.value === "council" ? renderCouncilSection(analysis) : renderCaptainPairSection(analysis));
      content.replaceChildren(panels.get(mode.value));
    };
    mode.addEventListener("change", render);
    render();
    return section;
  }

  function getPlayerDetails(analysis, index) {
    const name = analysis.names[index];
    const key = normalizeName(name);
    const nameColumn = findNameColumn(analysis.headers);
    const givenByCategory = Object.fromEntries(CATEGORIES.map((category) => [category, { positive: 0, negative: 0 }]));
    const knownNames = new Set(analysis.names.map(normalizeName));
    const questions = analysis.questionMeta.map((meta, questionIndex) => {
      const received = [];
      const given = [];
      analysis.rows.forEach((row) => {
        const picks = [...new Set(splitNames(row[nameColumn + 1 + questionIndex] || "").map(normalizeName))];
        if (picks.includes(key)) received.push(row[nameColumn]);
        if (normalizeName(row[nameColumn] || "") === key) {
          given.push(row[nameColumn + 1 + questionIndex] || "Nessuna scelta");
          if (meta.category) {
            givenByCategory[meta.category][meta.positive ? "positive" : "negative"] += picks.filter((pick) => pick !== key && knownNames.has(pick)).length;
          }
        }
      });
      return { ...meta, given: given.join("; ") || "Nessuna risposta", received };
    });
    const relationships = analysis.names.flatMap((other, otherIndex) => {
      if (otherIndex === index) return [];
      const m = analysis.matrices;
      return [{
        name: other,
        positiveGiven: m.posMatrix[index][otherIndex],
        positiveReceived: m.posMatrix[otherIndex][index],
        negativeGiven: m.negMatrix[index][otherIndex],
        negativeReceived: m.negMatrix[otherIndex][index],
        reciprocalPositive: m.reciprociPos[index][otherIndex],
        reciprocalNegative: m.reciprociNeg[index][otherIndex],
        bond: m.forzaLegami[index][otherIndex],
        antagonism: m.forzaAntagonismo[index][otherIndex],
        unreturnedGiven: m.nonRicambiatePos[index][otherIndex],
        unreturnedReceived: m.nonRicambiatePos[otherIndex][index]
      }];
    });
    return { name, questions, relationships, givenByCategory };
  }

  function renderPlayerTable(headers, rows) {
    const wrap = document.createElement("div");
    wrap.className = "table-wrap player-table-wrap";
    const table = document.createElement("table");
    table.className = "matrix-table";
    const head = table.createTHead().insertRow();
    headers.forEach((label) => {
      const cell = document.createElement("th");
      cell.scope = "col";
      cell.textContent = label;
      head.appendChild(cell);
    });
    const body = table.createTBody();
    rows.forEach((values) => {
      const row = body.insertRow();
      values.forEach((value, index) => {
        const cell = document.createElement(index ? "td" : "th");
        if (!index) cell.scope = "row";
        cell.textContent = value;
        row.appendChild(cell);
      });
    });
    wrap.appendChild(table);
    return wrap;
  }

  function renderPlayerBars(title, rows, series) {
    const chart = document.createElement("section");
    chart.className = "player-chart";
    const heading = document.createElement("h3");
    heading.textContent = title;
    chart.appendChild(heading);
    const maximum = Math.max(1, ...rows.flatMap((row) => row.values));
    rows.forEach((row) => {
      const group = document.createElement("div");
      group.className = "player-bar-group";
      const label = document.createElement("strong");
      label.textContent = row.name;
      group.appendChild(label);
      row.values.forEach((value, index) => {
        const line = document.createElement("div");
        line.className = "player-bar-line";
        const caption = document.createElement("span");
        caption.textContent = series[index].label;
        const track = document.createElement("div");
        track.className = "player-bar-track";
        const bar = document.createElement("div");
        bar.className = `player-bar ${series[index].tone}`;
        bar.style.width = `${value / maximum * 100}%`;
        track.appendChild(bar);
        track.setAttribute("aria-hidden", "true");
        const number = document.createElement("strong");
        number.textContent = formatNumber(value);
        line.append(caption, track, number);
        group.appendChild(line);
      });
      chart.appendChild(group);
    });
    return chart;
  }

  function renderPlayerNetwork(name, relationships, metric, label) {
    const wrapper = document.createElement("div");
    wrapper.className = "player-network";
    const connections = relationships.filter((row) => row[metric] > 0);
    if (!connections.length) {
      const empty = document.createElement("p");
      empty.className = "note";
      empty.textContent = "Nessun legame per questo indicatore.";
      wrapper.appendChild(empty);
      return wrapper;
    }
    const ns = "http://www.w3.org/2000/svg";
    const svgElement = (tag, attrs, text) => {
      const el = document.createElementNS(ns, tag);
      Object.entries(attrs).forEach(([key, value]) => el.setAttribute(key, value));
      if (text != null) el.textContent = text;
      return el;
    };
    const height = Math.max(320, connections.length * 52 + 40);
    const svg = svgElement("svg", { viewBox: `0 0 800 ${height}`, role: "img", "aria-label": `${label}: relazioni di ${name}` });
    svg.appendChild(svgElement("title", {}, `${label} — ${name}`));
    const maximum = Math.max(...connections.map((row) => row[metric]));
    connections.forEach((row, index) => {
      const y = 40 + index * (height - 80) / Math.max(1, connections.length - 1);
      const line = svgElement("line", { x1: 220, y1: height / 2, x2: 500, y2: y, stroke: "#688d83", "stroke-width": 1 + row[metric] / maximum * 5 });
      line.appendChild(svgElement("title", {}, `${row.name}: ${row[metric]}`));
      svg.appendChild(line);
      svg.appendChild(svgElement("circle", { cx: 500, cy: y, r: 7, fill: "#688d83" }));
      svg.appendChild(svgElement("text", { x: 518, y: y + 5, "font-size": 14, fill: "#1c201a" }, `${row.name} (${row[metric]})`));
    });
    svg.appendChild(svgElement("circle", { cx: 220, cy: height / 2, r: 12, fill: "#b65a2a" }));
    svg.appendChild(svgElement("text", { x: 200, y: height / 2 + 5, "text-anchor": "end", "font-size": 14, "font-weight": "bold", fill: "#1c201a" }, name));
    wrapper.appendChild(svg);
    return wrapper;
  }

  function renderPlayerSection(analysis) {
    const section = document.createElement("div");
    section.className = "player-section";
    const heading = document.createElement("h2");
    heading.textContent = "Scheda giocatrice";
    section.appendChild(heading);
    if (!analysis.names.length) return section;
    const control = document.createElement("label");
    control.className = "player-control";
    control.appendChild(document.createTextNode("Giocatrice"));
    const select = document.createElement("select");
    analysis.names.forEach((name, index) => select.add(new Option(name, index)));
    select.value = String(playerSelections.get(analysis) || 0);
    control.appendChild(select);
    section.appendChild(control);
    const content = document.createElement("div");
    section.appendChild(content);
    const metrics = [
      ["positiveReceived", "Scelte positive ricevute"], ["positiveGiven", "Scelte positive date"],
      ["negativeReceived", "Scelte negative ricevute"], ["negativeGiven", "Scelte negative date"],
      ["reciprocalPositive", "Reciproci positivi"], ["reciprocalNegative", "Reciproci negativi"],
      ["bond", "Forza legami"], ["antagonism", "Forza antagonismo"],
      ["unreturnedGiven", "Positive date e non ricambiate"], ["unreturnedReceived", "Positive ricevute e non ricambiate"]
    ];
    let selectedMetric = metrics[0][0];
    const render = () => {
      const index = Number(select.value);
      playerSelections.set(analysis, index);
      const details = getPlayerDetails(analysis, index);
      const summary = analysis.summaryRows[index];
      const classification = analysis.classifications[index];
      content.replaceChildren();
      const title = document.createElement("h3");
      title.textContent = details.name;
      content.appendChild(title);
      const labels = document.createElement("p");
      labels.className = "player-labels";
      labels.textContent = classification.label;
      content.appendChild(labels);
      const cards = document.createElement("div");
      cards.className = "player-stats";
      metrics.forEach(([key, label]) => {
        const card = document.createElement("div");
        const caption = document.createElement("span");
        caption.textContent = label;
        const value = document.createElement("strong");
        value.textContent = formatNumber(details.relationships.reduce((sum, row) => sum + row[key], 0));
        card.append(caption, value);
        cards.appendChild(card);
      });
      content.appendChild(cards);
      const balance = document.createElement("p");
      balance.textContent = `Equilibrio relazionale: ${classification.equilibrio}`;
      content.appendChild(balance);
      content.appendChild(renderPlayerTable(
        ["Ambito", "Scelte ricevute +", "Scelte ricevute −", "Influenza positiva", "Influenza negativa", "Etichetta"],
        [
          ["Totale", summary.positiveReceived.Totali, summary.negativeReceived.Totali, classification.inflPos, classification.inflNeg, classification.label],
          ...CATEGORIES.map((category) => [category, summary.positiveReceived[category], summary.negativeReceived[category], classification.byCategory[category].inflPos, classification.byCategory[category].inflNeg, classification.byCategory[category].label])
        ]
      ));
      const note = document.createElement("p");
      note.className = "note";
      note.textContent = "Le influenze mantengono il confronto con la media dell’intera squadra. I grafici e le relazioni includono solo la giocatrice selezionata. Le etichette descrivono le percezioni espresse nelle risposte.";
      content.appendChild(note);
      content.appendChild(renderPlayerBars("Scelte ricevute per categoria", CATEGORIES.map((category) => ({
        name: category, values: [summary.positiveReceived[category], summary.negativeReceived[category]]
      })), [{ label: "Positive", tone: "positive" }, { label: "Negative", tone: "negative" }]));
      content.appendChild(renderPlayerBars("Scelte date per categoria", CATEGORIES.map((category) => ({
        name: category, values: [details.givenByCategory[category].positive, details.givenByCategory[category].negative]
      })), [{ label: "Positive", tone: "positive" }, { label: "Negative", tone: "negative" }]));
      const relationHeading = document.createElement("h3");
      relationHeading.textContent = "Relazioni con le compagne";
      content.appendChild(relationHeading);
      const metricControl = document.createElement("label");
      metricControl.className = "player-control";
      metricControl.appendChild(document.createTextNode("Indicatore del grafico e della rete"));
      const metricSelect = document.createElement("select");
      metrics.forEach(([key, label]) => metricSelect.add(new Option(label, key)));
      metricSelect.value = selectedMetric;
      metricControl.appendChild(metricSelect);
      content.appendChild(metricControl);
      const graphs = document.createElement("div");
      content.appendChild(graphs);
      const renderGraphs = () => {
        selectedMetric = metricSelect.value;
        const label = metrics.find(([key]) => key === selectedMetric)[1];
        const ordered = [...details.relationships].sort((a, b) => b[selectedMetric] - a[selectedMetric]);
        graphs.replaceChildren(renderPlayerBars(label, ordered.map((row) => ({ name: row.name, values: [row[selectedMetric]] })), [{ label: "Valore", tone: /negative|Negative|antagonism/.test(selectedMetric) ? "negative" : "positive" }]));
        const graphNote = document.createElement("p");
        graphNote.className = "note";
        graphNote.textContent = `${label}: ogni collegamento riguarda ${details.name}; tra parentesi il valore. La rete esclude i legami tra le altre compagne e quelli con valore zero.`;
        graphs.append(graphNote, renderPlayerNetwork(details.name, ordered, selectedMetric, label));
      };
      metricSelect.addEventListener("change", renderGraphs);
      renderGraphs();
      content.appendChild(renderPlayerTable(["Compagna", ...metrics.map(([, label]) => label)], details.relationships.map((row) => [row.name, ...metrics.map(([key]) => row[key])])));
      const questionHeading = document.createElement("h3");
      questionHeading.textContent = "Dettaglio per domanda";
      content.appendChild(questionHeading);
      const questionNote = document.createElement("p");
      questionNote.className = "note";
      questionNote.textContent = "Tutte le risposte date e i nomi di chi ha scelto la giocatrice. Le domande senza categoria sono visibili qui ma, come nelle analisi generali, sono escluse dai conteggi.";
      content.appendChild(questionNote);
      content.appendChild(renderPlayerTable(["Domanda", "Categoria", "Segno", "Scelte date", "Scelte ricevute da", "Numero ricevute"], details.questions.map((question) => [question.question, question.category || "Non classificata", question.positive ? "+" : "−", question.given, question.received.join(", ") || "Nessuna scelta", question.received.length])));
    };
    select.addEventListener("change", render);
    render();
    return section;
  }

  function renderResponsesSection(analysis) {
    const section = document.createElement("div");
    const heading = document.createElement("h2");
    heading.textContent = "Risposte del modulo";
    section.appendChild(heading);

    const note = document.createElement("p");
    note.className = "note";
    note.textContent = "Naviga per atleta e consulta tutte le risposte organizzate per categoria.";
    section.appendChild(note);

    const list = document.createElement("div");
    list.className = "responses-list";
    section.appendChild(list);

    const responseData = buildResponsesData(analysis);

    list.innerHTML = "";
    responseData.forEach((item) => {
      const details = document.createElement("details");
      details.className = "response-card";

      const summary = document.createElement("summary");
      const title = document.createElement("span");
      title.textContent = item.name;
      const meta = document.createElement("span");
      meta.className = "response-meta";
      meta.textContent = `${item.counts.positive} positive · ${item.counts.negative} negative`;
      summary.appendChild(title);
      summary.appendChild(meta);
      details.appendChild(summary);

      const grid = document.createElement("div");
      grid.className = "response-grid";

      item.groups.forEach((group) => {
        const groupEl = document.createElement("div");
        groupEl.className = "response-group";
        const h4 = document.createElement("h4");
        h4.textContent = group.title;
        const pill = document.createElement("span");
        pill.className = "response-pill";
        pill.textContent = group.count;
        h4.appendChild(pill);
        groupEl.appendChild(h4);

        const items = document.createElement("div");
        items.className = "response-items";
        group.items.forEach((entry) => {
          const row = document.createElement("div");
          row.className = `response-item ${entry.sentiment || ""}`.trim();

          const question = document.createElement("div");
          question.className = "response-question";
          question.textContent = entry.question;

          const answer = document.createElement("div");
          answer.className = "response-answer";
          if (entry.answer) {
            answer.textContent = entry.answer;
          } else {
            answer.textContent = "Nessuna risposta";
            answer.classList.add("empty");
          }

          row.appendChild(question);
          row.appendChild(answer);
          items.appendChild(row);
        });
        groupEl.appendChild(items);
        grid.appendChild(groupEl);
      });

      details.appendChild(grid);
      list.appendChild(details);
    });

    return section;
  }

  function renderTopQuestionsSection(analysis) {
    const section = document.createElement("div");
    const heading = document.createElement("h2");
    heading.textContent = "Domande con pi\u00f9 risposte ricevute";
    section.appendChild(heading);

    const note = document.createElement("p");
    note.className = "note";
    note.textContent = "Per ogni atleta mostra le domande con pi\u00f9 risposte positive e negative ricevute.";
    section.appendChild(note);

    const controls = document.createElement("div");
    controls.className = "top-questions-controls";
    const label = document.createElement("label");
    label.textContent = "Numero di domande per atleta";
    const input = document.createElement("input");
    input.type = "number";
    input.min = "1";
    const maxLimit = Math.max(analysis.questionStats.posLabels.length, analysis.questionStats.negLabels.length, 1);
    input.max = String(maxLimit);
    input.value = "3";
    label.appendChild(input);
    controls.appendChild(label);
    section.appendChild(controls);

    const list = document.createElement("div");
    list.className = "top-questions-list";
    section.appendChild(list);

    const renderList = (limit) => {
      list.innerHTML = "";
      analysis.names.forEach((name, index) => {
        const card = document.createElement("div");
        card.className = "response-card top-questions-card";

        const title = document.createElement("h3");
        title.textContent = name;
        card.appendChild(title);

        const grid = document.createElement("div");
        grid.className = "top-questions-grid";

        grid.appendChild(renderTopQuestionColumn(
          "Positive ricevute",
          analysis.questionStats.posLabels,
          analysis.questionStats.posReceivedByQuestion[index],
          limit,
          "positive"
        ));
        grid.appendChild(renderTopQuestionColumn(
          "Negative ricevute",
          analysis.questionStats.negLabels,
          analysis.questionStats.negReceivedByQuestion[index],
          limit,
          "negative"
        ));

        card.appendChild(grid);
        list.appendChild(card);
      });
    };

    const normalizeLimit = () => {
      const parsed = Number.parseInt(input.value, 10);
      const safe = Number.isFinite(parsed) ? parsed : 1;
      return Math.max(1, Math.min(maxLimit, safe));
    };

    input.addEventListener("input", () => {
      const limit = normalizeLimit();
      input.value = String(limit);
      renderList(limit);
    });

    renderList(normalizeLimit());
    return section;
  }

  function renderTopRecipientsSection(analysis) {
    const section = document.createElement("div");
    const heading = document.createElement("h2");
    heading.textContent = "Atlete pi\u00f9 scelte per domanda";
    section.appendChild(heading);

    const note = document.createElement("p");
    note.className = "note";
    note.textContent = "Per ogni domanda mostra le atlete che hanno ricevuto pi\u00f9 risposte.";
    section.appendChild(note);

    const controls = document.createElement("div");
    controls.className = "top-questions-controls";
    const label = document.createElement("label");
    label.textContent = "Numero di atlete per domanda";
    const input = document.createElement("input");
    input.type = "number";
    input.min = "1";
    input.max = String(Math.max(analysis.names.length, 1));
    input.value = "3";
    label.appendChild(input);
    controls.appendChild(label);
    section.appendChild(controls);

    const list = document.createElement("div");
    list.className = "top-questions-list";
    section.appendChild(list);

    const renderList = (limit) => {
      list.innerHTML = "";
      const grid = document.createElement("div");
      grid.className = "top-questions-grid";
      const posGroup = document.createElement("div");
      const posHeading = document.createElement("h3");
      posHeading.textContent = "Domande positive";
      posGroup.appendChild(posHeading);
      const posList = document.createElement("div");
      posList.className = "top-questions-list";
      analysis.questionStats.posLabels.forEach((question, index) => {
        posList.appendChild(renderTopRecipientsQuestionCard(
          question,
          analysis.names,
          analysis.questionStats.posReceivedByQuestion,
          index,
          limit,
          "positive"
        ));
      });
      posGroup.appendChild(posList);
      grid.appendChild(posGroup);

      const negGroup = document.createElement("div");
      const negHeading = document.createElement("h3");
      negHeading.textContent = "Domande negative";
      negGroup.appendChild(negHeading);
      const negList = document.createElement("div");
      negList.className = "top-questions-list";
      analysis.questionStats.negLabels.forEach((question, index) => {
        negList.appendChild(renderTopRecipientsQuestionCard(
          question,
          analysis.names,
          analysis.questionStats.negReceivedByQuestion,
          index,
          limit,
          "negative"
        ));
      });
      negGroup.appendChild(negList);
      grid.appendChild(negGroup);
      list.appendChild(grid);
    };

    const normalizeLimit = () => {
      const parsed = Number.parseInt(input.value, 10);
      const safe = Number.isFinite(parsed) ? parsed : 1;
      const max = Math.max(analysis.names.length, 1);
      return Math.max(1, Math.min(max, safe));
    };

    input.addEventListener("input", () => {
      const limit = normalizeLimit();
      input.value = String(limit);
      renderList(limit);
    });

    renderList(normalizeLimit());
    return section;
  }

  function renderTopRecipientsQuestionCard(question, names, matrixByAthlete, questionIndex, limit, tone) {
    const card = document.createElement("div");
    card.className = "response-card top-questions-card";
    if (tone === "positive") {
      card.classList.add("top-questions-positive");
    } else if (tone === "negative") {
      card.classList.add("top-questions-negative");
    }

    const title = document.createElement("h3");
    title.textContent = question;
    card.appendChild(title);

    const items = buildTopRecipientsList(names, matrixByAthlete, questionIndex, limit);
    if (!items.length) {
      const empty = document.createElement("p");
      empty.className = "note";
      empty.textContent = "Nessuna risposta.";
      card.appendChild(empty);
      return card;
    }

    const list = document.createElement("ol");
    list.className = "top-questions-items";
    items.forEach((item) => {
      const li = document.createElement("li");
      li.className = "top-questions-item";
      const label = document.createElement("span");
      label.textContent = item.label;
      const count = document.createElement("span");
      count.className = "top-questions-count";
      count.textContent = item.count;
      li.appendChild(label);
      li.appendChild(count);
      list.appendChild(li);
    });
    card.appendChild(list);
    return card;
  }

  function renderTopQuestionColumn(title, labels, counts = [], limit, tone) {
    const column = document.createElement("div");
    column.className = "top-questions-column";
    if (tone === "positive") {
      column.classList.add("top-questions-positive");
    } else if (tone === "negative") {
      column.classList.add("top-questions-negative");
    }

    const heading = document.createElement("h4");
    heading.textContent = title;
    column.appendChild(heading);

    const items = buildTopQuestionList(labels, counts, limit);
    if (!items.length) {
      const empty = document.createElement("p");
      empty.className = "note";
      empty.textContent = "Nessuna risposta.";
      column.appendChild(empty);
      return column;
    }

    const list = document.createElement("ol");
    list.className = "top-questions-items";
    items.forEach((item) => {
      const li = document.createElement("li");
      li.className = "top-questions-item";
      const label = document.createElement("span");
      label.textContent = item.label;
      const count = document.createElement("span");
      count.className = "top-questions-count";
      count.textContent = item.count;
      li.appendChild(label);
      li.appendChild(count);
      list.appendChild(li);
    });
    column.appendChild(list);

    return column;
  }

  function buildTopQuestionList(labels, counts, limit) {
    return labels
      .map((label, index) => ({ label, count: counts[index] || 0, index }))
      .filter((item) => item.count > 0)
      .sort((a, b) => (b.count === a.count ? a.index - b.index : b.count - a.count))
      .slice(0, limit);
  }

  function buildTopRecipientsList(names, matrixByAthlete, questionIndex, limit) {
    return names
      .map((name, index) => ({
        label: name,
        count: (matrixByAthlete[index] || [])[questionIndex] || 0,
        index
      }))
      .filter((item) => item.count > 0)
      .sort((a, b) => (b.count === a.count ? a.index - b.index : b.count - a.count))
      .slice(0, limit);
  }

  function parseList(text) {
    return String(text || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
  }

  function titleCaseName(value) {
    return String(value || "")
      .trim()
      .split(/\s+/)
      .map((part) => {
        if (!part) {
          return "";
        }
        const lower = part.toLowerCase();
        return lower.charAt(0).toUpperCase() + lower.slice(1);
      })
      .join(" ");
  }

  function extractInlineCategory(text, fallback) {
    const raw = String(text || "");
    const match = raw.match(/^\s*\[(Tecniche|Attitudinali|Sociali)\]\s*/i);
    if (!match) {
      return { category: fallback, question: raw.trim() };
    }
    const lowered = match[1].toLowerCase();
    const category = lowered.startsWith("tec")
      ? "Tecniche"
      : lowered.startsWith("att")
        ? "Attitudinali"
        : "Sociali";
    const question = raw.slice(match[0].length).trim();
    return { category, question: question || raw.trim() };
  }

  function parseTeamweaveQuestions(text, categories) {
    const lines = parseList(text);
    const pos = [];
    const neg = [];
    lines.forEach((line, index) => {
      const current = line.trim();
      if (!current) {
        return;
      }
      const pairIndex = Math.floor(index / 2);
      const category = categories[pairIndex] || "Tecniche";
      const question = `[${category}] ${current}`;
      if (index % 2 === 0) {
        pos.push(question);
      } else {
        neg.push(question);
      }
    });
    return { pos, neg };
  }

  function buildTeamweaveFormScript(players, questions, options) {
    const { title, limit } = options;
    const formTitle = title || "Test TeamWeave";
    const description = buildTeamweaveDescription(limit);
    const lines = [
      "function buildTeamWeaveForm() {",
      `  var form = FormApp.create("${escapeQuotes(formTitle)}");`,
      `  form.setDescription("${escapeQuotes(description)}");`,
      `  var players = [${players.map((name) => `"${escapeQuotes(name)}"`).join(", ")}];`,
      "  var nameItem = form.addListItem();",
      "  nameItem.setTitle(\"Seleziona il tuo nome\").setChoiceValues(players).setRequired(true);"
    ];

    const max = Math.max(questions.pos.length, questions.neg.length);
    for (let i = 0; i < max; i += 1) {
      if (questions.pos[i]) {
        lines.push("  var item = form.addCheckboxItem();");
        lines.push(`  item.setTitle("${escapeQuotes(questions.pos[i])}").setChoiceValues(players).setRequired(false);`);
        lines.push(`  item.setValidation(FormApp.createCheckboxValidation().requireSelectAtMost(${limit}).build());`);
      }
      if (questions.neg[i]) {
        lines.push("  var item = form.addCheckboxItem();");
        lines.push(`  item.setTitle("${escapeQuotes(questions.neg[i])}").setChoiceValues(players).setRequired(false);`);
        lines.push(`  item.setValidation(FormApp.createCheckboxValidation().requireSelectAtMost(${limit}).build());`);
      }
    }

    lines.push("  Logger.log(\"Form creato: \" + form.getEditUrl());");
    lines.push("}");
    return lines.join("\n");
  }

  function buildTeamweaveDescription(limit) {
    const max = Number.isFinite(limit) ? limit : 3;
    return [
      "Questo questionario serve a capire meglio le relazioni, la comunicazione e la coesione all’interno della squadra. Le tue risposte aiuteranno l’allenatore a conoscere il gruppo e a migliorare il clima di squadra, l’organizzazione e la collaborazione tra compagne.",
      "",
      "• Non ci sono risposte giuste o sbagliate: rispondi con sincerità.",
      "• Nessuno oltre all’allenatore leggerà le risposte, e non verranno mai condivise con la squadra.",
      "• Non si può rispondere con “tutte”: scegli le migliori " + max + " (o le prime " + max + " che ti vengono in mente).",
      "• Non vi preoccupate di escludere qualcuno: mettete le prime " + max + " persone che vi vengono in mente.",
      "• Potete mettere da 0 a " + max + " risposte. Nelle domande positive è importante sforzarsi di mettere " + max + " risposte (a meno che proprio non ci siano).",
      "• Siate oggettive nelle domande puramente tecniche (es. “chi sceglieresti per fare una squadra forte”).",
      "• Nessuno saprà le vostre risposte oltre all’allenatore: siate sincere.",
      "• Non ha ovviamente senso votare dei non ricevitori su una domanda di ricezione, eccetto rari casi, così come non ha ovviamente senso votare i liberi su domande di battuta o attacco",
      "• Non potete auto-votarvi."
    ].join("\n");
  }

  async function initializeDefaultQuestionSets() {
    populateQuestionSetSelect(DEFAULT_QUESTION_SETS);
    await discoverQuestionSets();
    const preferredUrl = getStoredQuestionSetChoice() || DEFAULT_QUESTION_SETS[0]?.url;
    if (preferredUrl) {
      await loadQuestionSetFromUrl(preferredUrl, { silent: true, persistChoice: false });
    }
  }

  async function discoverQuestionSets() {
    try {
      const response = await fetch(QUESTION_SETS_DIR, { cache: "no-store" });
      if (!response.ok) {
        return;
      }
      const html = await response.text();
      const discovered = extractQuestionSetLinks(html);
      if (discovered.length) {
        populateQuestionSetSelect(discovered);
      }
    } catch (error) {
      return;
    }
  }

  function extractQuestionSetLinks(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");
    const baseUrl = new URL(QUESTION_SETS_DIR, location.href);
    const discovered = Array.from(doc.querySelectorAll("a[href]"))
      .map((link) => link.getAttribute("href") || "")
      .filter((href) => href.toLowerCase().endsWith(".txt"))
      .map((href) => {
        const url = new URL(href, baseUrl);
        const fileName = decodeURIComponent(url.pathname.split("/").pop() || href);
        return {
          name: fileName.replace(/\.txt$/i, ""),
          url: `${QUESTION_SETS_DIR}${encodeURIComponent(fileName)}`
        };
      });
    const byUrl = new Map();
    DEFAULT_QUESTION_SETS.concat(discovered).forEach((set) => {
      byUrl.set(normalizeQuestionSetUrl(set.url), set);
    });
    return Array.from(byUrl.values()).sort((a, b) => a.name.localeCompare(b.name, "it"));
  }

  function normalizeQuestionSetUrl(url) {
    try {
      const parsed = new URL(url, location.href);
      return decodeURIComponent(parsed.pathname).toLowerCase();
    } catch (error) {
      return String(url || "").replace(/%20/g, " ").toLowerCase();
    }
  }

  function populateQuestionSetSelect(sets) {
    if (!twQuestionSetSelect) {
      return;
    }
    const selected = twQuestionSetSelect.value || getStoredQuestionSetChoice() || DEFAULT_QUESTION_SETS[0]?.url || "";
    twQuestionSetSelect.textContent = "";
    sets.forEach((set) => {
      const option = document.createElement("option");
      option.value = set.url;
      option.textContent = set.name;
      twQuestionSetSelect.appendChild(option);
    });
    if (sets.some((set) => set.url === selected)) {
      twQuestionSetSelect.value = selected;
    } else if (sets[0]) {
      twQuestionSetSelect.value = sets[0].url;
    }
  }

  function getStoredQuestionSetChoice() {
    try {
      return localStorage.getItem(QUESTION_SET_CHOICE_KEY);
    } catch (error) {
      return null;
    }
  }

  function saveQuestionSetChoice(url) {
    try {
      localStorage.setItem(QUESTION_SET_CHOICE_KEY, url);
    } catch (error) {
      return;
    }
  }

  function getQuestionSetName(url) {
    const option = Array.from(twQuestionSetSelect?.options || []).find((entry) => entry.value === url);
    if (option) {
      return option.textContent || url;
    }
    const fallback = DEFAULT_QUESTION_SETS.find((entry) => entry.url === url);
    return fallback?.name || decodeURIComponent(url.split("/").pop() || "Set domande").replace(/\.txt$/i, "");
  }

  async function loadQuestionSetFromUrl(url, options = {}) {
    const { silent = false, persistChoice = true } = options;
    try {
      const response = await fetch(url, { cache: "no-store" });
      if (!response.ok) {
        throw new Error("Set domande non trovato.");
      }
      const text = await response.text();
      applyQuestionSetText(text, getQuestionSetName(url));
      if (persistChoice) {
        saveQuestionSetChoice(url);
      }
      if (!silent) {
        setGeneratorStatus(`Set domande caricato: ${getQuestionSetName(url)}.`);
      }
      return true;
    } catch (error) {
      if (!silent) {
        setGeneratorStatus(error.message || "Impossibile caricare il set domande.", true);
      }
      return false;
    }
  }

  function applyQuestionSetText(text, name) {
    const lines = parseList(text);
    if (!lines.length) {
      throw new Error("File domande vuoto.");
    }
    twFormQuestions.value = lines.join("\n");
    twFormQuestions.dispatchEvent(new Event("input"));
    setGeneratorStatus(`${name || "Set domande"}: ${lines.length} domande caricate.`);
  }

  function setGeneratorStatus(message, isError) {
    if (!twQuestionSetStatus) {
      return;
    }
    twQuestionSetStatus.textContent = message;
    twQuestionSetStatus.classList.toggle("error", Boolean(isError));
  }

  function initTeamweaveFormGenerator() {
    if (!twFormTitle || !twFormPlayers || !twFormQuestions) {
      return;
    }

    function getCategorySelections() {
      if (!twFormCategories) {
        return [];
      }
      const selections = [];
      const rows = Array.from(twFormCategories.querySelectorAll(".tw-category-row"));
      rows.forEach((row) => {
        const checked = row.querySelector("input[type=\"radio\"]:checked");
        if (checked) {
          selections.push(checked.value);
        }
      });
      return selections;
    }

    function renderCategorySelectors() {
      if (!twFormCategories) {
        return;
      }
      const lines = parseList(twFormQuestions.value);
      const pairCount = Math.ceil(lines.length / 2);
      const prev = getCategorySelections();
      const detected = [];
      for (let i = 0; i < pairCount; i += 1) {
        const first = extractInlineCategory(lines[i * 2] || "", null);
        const second = extractInlineCategory(lines[i * 2 + 1] || "", null);
        detected.push(first.category || second.category || null);
      }
      twFormCategories.innerHTML = "";
      for (let i = 0; i < pairCount; i += 1) {
        const row = document.createElement("div");
        row.className = "tw-category-row";
        const label = document.createElement("span");
        label.className = "tw-category-label";
        const first = lines[i * 2] || "";
        const second = lines[i * 2 + 1] || "";
        if (second) {
          label.innerHTML = `<strong>${escapeHtml(first)}</strong><span class="tw-category-sep"> / </span>${escapeHtml(second)}`;
        } else {
          label.textContent = first || `Domanda ${i * 2 + 1}`;
        }
        row.appendChild(label);

        const options = document.createElement("div");
        options.className = "tw-category-options";
        ["Tecniche", "Attitudinali", "Sociali"].forEach((value) => {
          const radioLabel = document.createElement("label");
          radioLabel.className = "tw-category-option";
          const input = document.createElement("input");
          input.type = "radio";
          input.name = `tw-cat-${i}`;
          input.value = value;
          const preferred = prev[i] || detected[i] || "Tecniche";
          if (preferred === value) {
            input.checked = true;
          }
          radioLabel.appendChild(input);
          radioLabel.appendChild(document.createTextNode(value));
          options.appendChild(radioLabel);
        });
        row.appendChild(options);
        twFormCategories.appendChild(row);
      }
    }

    function renderPreview(questions) {
      const items = [];
      questions.pos.forEach((q) => items.push(`✅ ${q}`));
      questions.neg.forEach((q) => items.push(`❌ ${q}`));
      if (!items.length) {
        twFormPreview.innerHTML = "<p class=\"note\">Nessuna domanda rilevata.</p>";
        return;
      }
      twFormPreview.innerHTML = items.map((q) => `<div class="question">${escapeHtml(q)}</div>`).join("");
    }

    function convert() {
      const players = parseList(twFormPlayers.value).map(titleCaseName).filter(Boolean);
      if (!players.length) {
        twFormOutput.textContent = "Inserisci almeno una giocatrice.";
        twFormPreview.innerHTML = "";
        return;
      }
      const lines = parseList(twFormQuestions.value);
      if (lines.length % 2 !== 0) {
        twFormOutput.textContent = "Le domande devono essere in numero pari (positiva + negativa).";
        twFormPreview.innerHTML = "";
        return;
      }
      const limit = Number(twFormLimit?.value) || 1;
      const categories = getCategorySelections();
      const questions = parseTeamweaveQuestions(
        lines.map((line) => extractInlineCategory(line, "Tecniche").question).join("\n"),
        categories
      );
      latestTeamweaveScript = buildTeamweaveFormScript(players, questions, {
        title: twFormTitle.value,
        limit
      });
      twFormOutput.textContent = latestTeamweaveScript;
      renderPreview(questions);
    }

    function copyScript() {
      if (!latestTeamweaveScript) {
        convert();
      }
      if (!latestTeamweaveScript) {
        return;
      }
      navigator.clipboard?.writeText(latestTeamweaveScript).then(() => {
        twFormOutput.textContent = `${latestTeamweaveScript}\n// Copiato negli appunti.`;
      }).catch(() => {
        twFormOutput.textContent = `${latestTeamweaveScript}\n// Copia non riuscita.`;
      });
    }

    function downloadScript() {
      if (!latestTeamweaveScript) {
        convert();
      }
      if (!latestTeamweaveScript) {
        return;
      }
      const blob = new Blob([latestTeamweaveScript], { type: "text/plain" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "teamweave-form.gs";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    }

    twFormQuestions.addEventListener("input", renderCategorySelectors);
    renderCategorySelectors();
    initializeDefaultQuestionSets();

    twFormConvert?.addEventListener("click", convert);
    twFormCopy?.addEventListener("click", copyScript);
    twFormDownload?.addEventListener("click", downloadScript);
    twQuestionSetLoad?.addEventListener("click", async () => {
      const selectedUrl = twQuestionSetSelect?.value;
      if (!selectedUrl) {
        setGeneratorStatus("Nessun set domande disponibile.", true);
        return;
      }
      await loadQuestionSetFromUrl(selectedUrl, { silent: false, persistChoice: true });
    });
    twQuestionSetSelect?.addEventListener("change", async () => {
      const selectedUrl = twQuestionSetSelect.value;
      if (selectedUrl) {
        await loadQuestionSetFromUrl(selectedUrl, { silent: false, persistChoice: true });
      }
    });
    twQuestionFileInput?.addEventListener("change", async (event) => {
      const [file] = Array.from(event.target.files || []);
      if (!file) {
        return;
      }
      try {
        const text = await readFile(file);
        applyQuestionSetText(text, file.name);
      } catch (error) {
        setGeneratorStatus(error.message || "Impossibile leggere il file domande.", true);
      } finally {
        event.target.value = "";
      }
    });
    twFormDownloadQuestions?.addEventListener("click", () => {
      const lines = parseList(twFormQuestions.value);
      if (!lines.length) {
        twFormOutput.textContent = "Inserisci almeno una domanda.";
        return;
      }
      const categories = getCategorySelections();
      const tagged = lines.map((line, index) => {
        const pairIndex = Math.floor(index / 2);
        const category = categories[pairIndex] || "Tecniche";
        const cleaned = extractInlineCategory(line, category).question;
        return `[${category}] ${cleaned}`;
      }).join("\n");
      const blob = new Blob([tagged], { type: "text/plain" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "teamweave-domande.txt";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    });
  }

  function buildResponsesData(analysis) {
    const nameIndex = findNameColumn(analysis.headers);
    const firstQuestionIndex = nameIndex + 1;
    const data = [];

    analysis.rows.forEach((row) => {
      const name = row[nameIndex] || "Senza nome";
      const groups = [];
      const counts = { positive: 0, negative: 0 };

      const grouped = {
        Tecniche: [],
        Attitudinali: [],
        Sociali: [],
        Altro: []
      };

      analysis.questionMeta.forEach((meta, idx) => {
        const answer = row[firstQuestionIndex + idx] || "";
        const category = meta.category || "Altro";
        const sentiment = meta.positive ? "positive" : "negative";
        grouped[category].push({ question: meta.question, answer, sentiment });
        if (meta.positive) {
          counts.positive += answer ? 1 : 0;
        } else {
          counts.negative += answer ? 1 : 0;
        }
      });

      Object.keys(grouped).forEach((category) => {
        const items = grouped[category];
        if (!items.length) {
          return;
        }
        groups.push({
          title: category,
          count: items.length,
          items
        });
      });

      data.push({
        name,
        groups,
        counts
      });
    });

    return data;
  }

  function renderNetworkSection({ title, note, names, matrix, strengthMatrix, nodePalette, edgePalette, toggleId }) {
    const section = document.createElement("div");
    const heading = document.createElement("h2");
    heading.textContent = title;
    section.appendChild(heading);

    const noteEl = document.createElement("p");
    noteEl.className = "note";
    noteEl.textContent = note;
    section.appendChild(noteEl);

    const controls = document.createElement("div");
    controls.className = "note";
    const label = document.createElement("label");
    label.style.display = "inline-flex";
    label.style.gap = "0.5rem";
    label.style.alignItems = "center";
    const toggle = document.createElement("input");
    toggle.type = "checkbox";
    toggle.checked = true;
    toggle.id = toggleId;
    const text = document.createElement("span");
    text.textContent = "Colora i legami in base alla forza";
    label.appendChild(toggle);
    label.appendChild(text);
    controls.appendChild(label);
    section.appendChild(controls);

    const legendStats = getNetworkLegendStats(names, matrix, strengthMatrix);
    const legend = document.createElement("div");
    legend.className = "network-legend";
    const nodeLegend = document.createElement("div");
    nodeLegend.className = "legend-row";
    nodeLegend.innerHTML = `<span>Nodi</span><div class="legend-bar" style="background: linear-gradient(90deg, ${nodePalette[0]}, ${nodePalette[1]}, ${nodePalette[2]});"></div><span class="legend-range">${legendStats.minDegree} → ${legendStats.maxDegree}</span>`;
    const edgeLegend = document.createElement("div");
    edgeLegend.className = "legend-row";
    edgeLegend.innerHTML = `<span>Archi</span><div class="legend-bar" style="background: linear-gradient(90deg, ${edgePalette[0]}, ${edgePalette[1]});"></div><span class="legend-range">${legendStats.minStrength} → ${legendStats.maxStrength}</span>`;
    legend.appendChild(nodeLegend);
    legend.appendChild(edgeLegend);
    section.appendChild(legend);

    const network = document.createElement("div");
    network.className = "network";

    const canvas = document.createElement("canvas");
    canvas.width = 900;
    canvas.height = 520;
    network.appendChild(canvas);

    section.appendChild(network);
    initNetwork(
      canvas,
      names,
      matrix,
      strengthMatrix,
      toggle,
      {
        nodePalette,
        edgePalette
      }
    );

    return section;
  }

  function initNetwork(canvas, names, matrix, strengthMatrix, toggleEl, config) {
    const ctx = canvas.getContext("2d");
    const width = canvas.width;
    const height = canvas.height;

    const state = createNetworkState(names, matrix, strengthMatrix, width, height, config);
    canvas._networkState = state;
    drawNetwork(ctx, state);
    attachNetworkHandlers(canvas);
    attachNetworkToggle(canvas, toggleEl);
  }

  function getNetworkLegendStats(names, matrix, strengthMatrix) {
    const edges = [];
    const strengths = [];
    for (let i = 0; i < matrix.length; i += 1) {
      for (let j = i + 1; j < matrix.length; j += 1) {
        if (matrix[i][j] > 0) {
          edges.push([i, j]);
          strengths.push(strengthMatrix?.[i]?.[j] || 0);
        }
      }
    }
    const nodeWeights = sumRows(matrix);
    const minDegree = nodeWeights.length ? Math.min(...nodeWeights) : 0;
    const maxDegree = nodeWeights.length ? Math.max(...nodeWeights) : 0;
    const strengthRange = minMaxArray(strengths);
    return {
      minDegree,
      maxDegree,
      minStrength: formatNumber(strengthRange.min),
      maxStrength: formatNumber(strengthRange.max)
    };
  }

  function createNetworkState(names, matrix, strengthMatrix, width, height, config) {
    const nodes = names.map((name) => ({ name, x: 0, y: 0 }));
    const edges = [];
    const strengths = [];
    for (let i = 0; i < matrix.length; i += 1) {
      for (let j = i + 1; j < matrix.length; j += 1) {
        if (matrix[i][j] > 0) {
          edges.push([i, j]);
          strengths.push(strengthMatrix?.[i]?.[j] || 0);
        }
      }
    }

    layoutGraph(nodes, edges, width, height);

    const nodeWeights = sumRows(matrix);
    const maxWeight = Math.max(...nodeWeights, 1);

    const strengthRange = minMaxArray(strengths);
    return {
      nodes,
      edges,
      strengths,
      strengthRange,
      nodeWeights,
      maxWeight,
      width,
      height,
      dragIndex: null,
      showStrength: true,
      nodePalette: config?.nodePalette || PALETTE.network,
      edgePalette: config?.edgePalette || ["#d5d8dc", "#1a9850"]
    };
  }

  function drawNetwork(ctx, state) {
    const { nodes, edges, nodeWeights, maxWeight, width, height, strengths, strengthRange, showStrength, nodePalette, edgePalette } = state;
    ctx.clearRect(0, 0, width, height);
    ctx.lineWidth = 1.2;

    edges.forEach(([i, j], idx) => {
      if (showStrength) {
        const value = strengths[idx] || 0;
        const t = strengthRange.min === strengthRange.max ? 0 : (value - strengthRange.min) / (strengthRange.max - strengthRange.min);
        ctx.strokeStyle = interpolateBi(edgePalette, t);
      } else {
        ctx.strokeStyle = "rgba(0, 0, 0, 0.2)";
      }
      ctx.beginPath();
      ctx.moveTo(nodes[i].x, nodes[i].y);
      ctx.lineTo(nodes[j].x, nodes[j].y);
      ctx.stroke();
    });

    nodes.forEach((node, idx) => {
      const weight = nodeWeights[idx] || 0;
      const radius = 8 + (weight / maxWeight) * 16;
      const t = maxWeight === 0 ? 0 : weight / maxWeight;
      const color = interpolateTri(nodePalette, t);

      ctx.beginPath();
      ctx.fillStyle = color;
      ctx.strokeStyle = "#111";
      ctx.lineWidth = 1.2;
      ctx.arc(node.x, node.y, radius, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();

      ctx.fillStyle = "#1c201a";
      ctx.font = "12px Trebuchet MS";
      ctx.textAlign = "center";
      ctx.fillText(node.name, node.x, node.y - radius - 6);
    });
  }

  function attachNetworkHandlers(canvas) {
    const state = canvas._networkState;
    if (!state || canvas._networkHandlersAttached) {
      return;
    }
    canvas._networkHandlersAttached = true;
    const ctx = canvas.getContext("2d");

    const getPointer = (event) => {
      const rect = canvas.getBoundingClientRect();
      return {
        x: ((event.clientX - rect.left) / rect.width) * canvas.width,
        y: ((event.clientY - rect.top) / rect.height) * canvas.height
      };
    };

    const findNodeAt = (x, y) => {
      for (let i = state.nodes.length - 1; i >= 0; i -= 1) {
        const node = state.nodes[i];
        const degree = state.degrees[i];
        const radius = 8 + (degree / state.maxDegree) * 16;
        const dist = Math.hypot(node.x - x, node.y - y);
        if (dist <= radius + 6) {
          return i;
        }
      }
      return null;
    };

    const onDown = (event) => {
      const { x, y } = getPointer(event);
      const idx = findNodeAt(x, y);
      if (idx !== null) {
        state.dragIndex = idx;
        state.dragOffset = { x: state.nodes[idx].x - x, y: state.nodes[idx].y - y };
      }
    };

    const onMove = (event) => {
      if (state.dragIndex === null) {
        return;
      }
      const { x, y } = getPointer(event);
      const node = state.nodes[state.dragIndex];
      node.x = Math.min(state.width - 20, Math.max(20, x + state.dragOffset.x));
      node.y = Math.min(state.height - 20, Math.max(20, y + state.dragOffset.y));
      drawNetwork(ctx, state);
    };

    const onUp = () => {
      state.dragIndex = null;
    };

    canvas.addEventListener("mousedown", onDown);
    canvas.addEventListener("mousemove", onMove);
    canvas.addEventListener("mouseup", onUp);
    canvas.addEventListener("mouseleave", onUp);
  }

  function attachNetworkToggle(canvas, toggleEl) {
    const state = canvas._networkState;
    if (!state || !toggleEl) {
      return;
    }
    toggleEl.addEventListener("change", () => {
      state.showStrength = toggleEl.checked;
      const ctx = canvas.getContext("2d");
      drawNetwork(ctx, state);
    });
  }

  function layoutGraph(nodes, edges, width, height) {
    const n = nodes.length;
    const centerX = width / 2;
    const centerY = height / 2;
    const radius = Math.min(width, height) / 3;

    nodes.forEach((node, idx) => {
      const angle = (Math.PI * 2 * idx) / n;
      node.x = centerX + Math.cos(angle) * radius;
      node.y = centerY + Math.sin(angle) * radius;
    });

    const area = width * height;
    const k = Math.sqrt(area / n);

    for (let step = 0; step < 250; step += 1) {
      const disp = nodes.map(() => ({ x: 0, y: 0 }));

      for (let i = 0; i < n; i += 1) {
        for (let j = i + 1; j < n; j += 1) {
          const dx = nodes[i].x - nodes[j].x;
          const dy = nodes[i].y - nodes[j].y;
          const dist = Math.hypot(dx, dy) || 0.001;
          const force = (k * k) / dist;
          const offsetX = (dx / dist) * force;
          const offsetY = (dy / dist) * force;
          disp[i].x += offsetX;
          disp[i].y += offsetY;
          disp[j].x -= offsetX;
          disp[j].y -= offsetY;
        }
      }

      edges.forEach(([i, j]) => {
        const dx = nodes[i].x - nodes[j].x;
        const dy = nodes[i].y - nodes[j].y;
        const dist = Math.hypot(dx, dy) || 0.001;
        const force = (dist * dist) / k;
        const offsetX = (dx / dist) * force;
        const offsetY = (dy / dist) * force;
        disp[i].x -= offsetX;
        disp[i].y -= offsetY;
        disp[j].x += offsetX;
        disp[j].y += offsetY;
      });

      nodes.forEach((node, idx) => {
        node.x = Math.min(width - 40, Math.max(40, node.x + disp[idx].x * 0.02));
        node.y = Math.min(height - 40, Math.max(40, node.y + disp[idx].y * 0.02));
      });
    }
  }

  function countReciprociPairs(matrix) {
    let count = 0;
    for (let i = 0; i < matrix.length; i += 1) {
      for (let j = i + 1; j < matrix.length; j += 1) {
        if (matrix[i][j] > 0) {
          count += 1;
        }
      }
    }
    return count;
  }

  function formatNumber(value) {
    if (!Number.isFinite(value)) {
      return "0";
    }
    return value % 1 === 0 ? value.toString() : value.toFixed(2);
  }

  function formatPercent(value) {
    if (!Number.isFinite(value)) {
      return "0%";
    }
    const percent = value * 100;
    const formatted = percent % 1 === 0 ? percent.toString() : percent.toFixed(1);
    return `${formatted}%`;
  }

  function downloadCsv(csv) {
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "teamweave-data.csv";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function downloadHtml(html, name) {
    const safeName = (name || "teamweave-analisi").replace(/[^a-z0-9-_]+/gi, "-").toLowerCase();
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${safeName || "teamweave-analisi"}.html`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function base64UrlEncode(text) {
    const bytes = new TextEncoder().encode(text);
    let binary = "";
    bytes.forEach((b) => {
      binary += String.fromCharCode(b);
    });
    const base64 = btoa(binary);
    return base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }

  function base64UrlEncodeBytes(bytes) {
    let binary = "";
    bytes.forEach((b) => {
      binary += String.fromCharCode(b);
    });
    const base64 = btoa(binary);
    return base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }

  function base64UrlDecode(text) {
    try {
      const base64 = text.replace(/-/g, "+").replace(/_/g, "/");
      const padded = base64.padEnd(base64.length + (4 - (base64.length % 4)) % 4, "=");
      const binary = atob(padded);
      const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
      return new TextDecoder().decode(bytes);
    } catch (error) {
      setStatus("Link dati non valido.", true);
      return null;
    }
  }

  function base64UrlDecodeBytes(text) {
    try {
      const base64 = text.replace(/-/g, "+").replace(/_/g, "/");
      const padded = base64.padEnd(base64.length + (4 - (base64.length % 4)) % 4, "=");
      const binary = atob(padded);
      return Uint8Array.from(binary, (char) => char.charCodeAt(0));
    } catch (error) {
      return null;
    }
  }

  function packCsvForLink(text) {
    const { headers, rows } = parseCsv(text);
    const nameIndex = findNameColumn(headers);
    const names = [];
    const nameMap = new Map();

    const addName = (raw) => {
      const normalized = normalizeName(raw || "");
      if (!normalized) {
        return;
      }
      if (!nameMap.has(normalized)) {
        nameMap.set(normalized, names.length);
        names.push(raw.trim());
      }
    };

    rows.forEach((row) => {
      addName(row[nameIndex] || "");
    });

    rows.forEach((row) => {
      headers.forEach((_, idx) => {
        if (idx === nameIndex) {
          return;
        }
        splitNames(row[idx] || "").forEach(addName);
      });
    });

    const bytes = [];
    bytes.push(1);
    writeVarint(headers.length, bytes);
    headers.forEach((header) => writeString(header, bytes));
    writeVarint(nameIndex, bytes);
    writeVarint(names.length, bytes);
    names.forEach((name) => writeString(name, bytes));
    writeVarint(rows.length, bytes);

    const questionCount = headers.length - 1;
    rows.forEach((row) => {
      const chooserKey = normalizeName(row[nameIndex] || "");
      const chooserIndex = nameMap.has(chooserKey) ? nameMap.get(chooserKey) : names.length;
      writeVarint(chooserIndex, bytes);
      let answersWritten = 0;
      headers.forEach((_, idx) => {
        if (idx === nameIndex) {
          return;
        }
        const picks = splitNames(row[idx] || "");
        const indices = picks
          .map((pick) => nameMap.get(normalizeName(pick)))
          .filter((value) => Number.isInteger(value));
        writeVarint(indices.length, bytes);
        indices.forEach((value) => writeVarint(value, bytes));
        answersWritten += 1;
      });
      if (answersWritten < questionCount) {
        for (let i = answersWritten; i < questionCount; i += 1) {
          writeVarint(0, bytes);
        }
      }
    });

    return new Uint8Array(bytes);
  }

  function decodePackedCsv(payload) {
    if (!payload) {
      return null;
    }
    if (!payload.startsWith("pack:")) {
      return payload;
    }
    try {
      return unpackCsvFromLink(payload.slice(5));
    } catch (error) {
      setStatus("Link dati non valido.", true);
      return null;
    }
  }

  function unpackCsvFromLink(payload) {
    const data = JSON.parse(payload);
    if (!data || data.v !== 1 || !Array.isArray(data.headers)) {
      throw new Error("Formato link non valido.");
    }
    const headers = data.headers;
    const nameIndex = data.nameIndex;
    const names = data.names || [];
    const rows = data.rows || [];

    const lines = [];
    lines.push(headers.map(escapeCsvCell).join(","));
    rows.forEach((row) => {
      const chooserIndex = row[0];
      const answers = Array.isArray(row[1]) ? row[1] : [];
      const cells = Array(headers.length).fill("");
      if (Number.isInteger(chooserIndex) && names[chooserIndex]) {
        cells[nameIndex] = names[chooserIndex];
      }
      let answerIndex = 0;
      for (let i = 0; i < headers.length; i += 1) {
        if (i === nameIndex) {
          continue;
        }
        const indices = answers[answerIndex] || [];
        const picks = indices
          .map((idx) => names[idx])
          .filter(Boolean);
        cells[i] = picks.join(", ");
        answerIndex += 1;
      }
      lines.push(cells.map(escapeCsvCell).join(","));
    });
    return lines.join("\n");
  }

  function escapeCsvCell(value) {
    const safe = value == null ? "" : String(value);
    const escaped = safe.replace(/"/g, '""');
    if (/[",\n\r]/.test(escaped)) {
      return `"${escaped}"`;
    }
    return escaped;
  }

  function unpackCsvFromBinary(bytes) {
    if (!bytes || !bytes.length) {
      return null;
    }
    let offset = 0;
    const version = bytes[offset];
    offset += 1;
    if (version !== 1) {
      throw new Error("Formato link non supportato.");
    }

    const headerResult = readListOfStrings(bytes, offset);
    const headers = headerResult.value;
    offset = headerResult.offset;

    const nameIndexResult = readVarint(bytes, offset);
    const nameIndex = nameIndexResult.value;
    offset = nameIndexResult.offset;

    const namesResult = readListOfStrings(bytes, offset);
    const names = namesResult.value;
    offset = namesResult.offset;

    const rowCountResult = readVarint(bytes, offset);
    const rowCount = rowCountResult.value;
    offset = rowCountResult.offset;

    const questionCount = Math.max(headers.length - 1, 0);
    const lines = [];
    lines.push(headers.map(escapeCsvCell).join(","));

    for (let r = 0; r < rowCount; r += 1) {
      const chooserResult = readVarint(bytes, offset);
      const chooserIndex = chooserResult.value;
      offset = chooserResult.offset;

      const answers = [];
      for (let q = 0; q < questionCount; q += 1) {
        const countResult = readVarint(bytes, offset);
        const count = countResult.value;
        offset = countResult.offset;
        const indices = [];
        for (let k = 0; k < count; k += 1) {
          const idxResult = readVarint(bytes, offset);
          indices.push(idxResult.value);
          offset = idxResult.offset;
        }
        answers.push(indices);
      }

      const cells = Array(headers.length).fill("");
      if (chooserIndex < names.length) {
        cells[nameIndex] = names[chooserIndex];
      }
      let answerIndex = 0;
      for (let i = 0; i < headers.length; i += 1) {
        if (i === nameIndex) {
          continue;
        }
        const indices = answers[answerIndex] || [];
        const picks = indices.map((idx) => names[idx]).filter(Boolean);
        cells[i] = picks.join(", ");
        answerIndex += 1;
      }
      lines.push(cells.map(escapeCsvCell).join(","));
    }

    return lines.join("\n");
  }

  function writeVarint(value, bytes) {
    let current = Math.max(0, Math.floor(value));
    while (current >= 128) {
      bytes.push((current % 128) + 128);
      current = Math.floor(current / 128);
    }
    bytes.push(current);
  }

  function readVarint(bytes, offset) {
    let value = 0;
    let shift = 0;
    let index = offset;
    while (index < bytes.length) {
      const byte = bytes[index];
      index += 1;
      value += (byte & 0x7f) * (2 ** shift);
      if ((byte & 0x80) === 0) {
        return { value, offset: index };
      }
      shift += 7;
    }
    throw new Error("Link dati non valido.");
  }

  function writeString(text, bytes) {
    const encoded = new TextEncoder().encode(text);
    writeVarint(encoded.length, bytes);
    encoded.forEach((byte) => bytes.push(byte));
  }

  function readString(bytes, offset) {
    const lengthResult = readVarint(bytes, offset);
    const length = lengthResult.value;
    const start = lengthResult.offset;
    const end = start + length;
    if (end > bytes.length) {
      throw new Error("Link dati non valido.");
    }
    const chunk = bytes.slice(start, end);
    return { value: new TextDecoder().decode(chunk), offset: end };
  }

  function readListOfStrings(bytes, offset) {
    const countResult = readVarint(bytes, offset);
    const count = countResult.value;
    let currentOffset = countResult.offset;
    const values = [];
    for (let i = 0; i < count; i += 1) {
      const item = readString(bytes, currentOffset);
      values.push(item.value);
      currentOffset = item.offset;
    }
    return { value: values, offset: currentOffset };
  }

  async function encodeLinkPayload(payload) {
    if (payload instanceof Uint8Array) {
      if (typeof CompressionStream === "function") {
        const compressed = await gzipCompressBytes(payload);
        return `gzb.${base64UrlEncodeBytes(compressed)}`;
      }
      return `bin.${base64UrlEncodeBytes(payload)}`;
    }
    if (typeof CompressionStream === "function") {
      const compressed = await gzipCompress(payload);
      return `gz.${base64UrlEncodeBytes(compressed)}`;
    }
    return `raw.${base64UrlEncode(payload)}`;
  }

  async function decodeLinkPayload(encoded) {
    if (encoded.startsWith("gzb.")) {
      const bytes = base64UrlDecodeBytes(encoded.slice(4));
      if (!bytes) {
        setStatus("Link dati non valido.", true);
        return null;
      }
      const decompressed = await gzipDecompressBytes(bytes);
      if (!decompressed) {
        return null;
      }
      return { kind: "bytes", value: decompressed };
    }
    if (encoded.startsWith("bin.")) {
      const bytes = base64UrlDecodeBytes(encoded.slice(4));
      if (!bytes) {
        setStatus("Link dati non valido.", true);
        return null;
      }
      return { kind: "bytes", value: bytes };
    }
    if (encoded.startsWith("gz.")) {
      const bytes = base64UrlDecodeBytes(encoded.slice(3));
      if (!bytes) {
        setStatus("Link dati non valido.", true);
        return null;
      }
      const text = await gzipDecompress(bytes);
      return text ? { kind: "text", value: text } : null;
    }
    if (encoded.startsWith("raw.")) {
      return { kind: "text", value: base64UrlDecode(encoded.slice(4)) };
    }
    const bytes = base64UrlDecodeBytes(encoded);
    if (bytes) {
      const text = await gzipDecompress(bytes, true);
      if (text) {
        return { kind: "text", value: text };
      }
    }
    const fallback = base64UrlDecode(encoded);
    return fallback ? { kind: "text", value: fallback } : null;
  }

  async function gzipCompressBytes(bytes) {
    const stream = new CompressionStream("gzip");
    const writer = stream.writable.getWriter();
    writer.write(bytes);
    writer.close();
    const buffer = await new Response(stream.readable).arrayBuffer();
    return new Uint8Array(buffer);
  }

  async function gzipDecompressBytes(bytes, silentFailure) {
    if (typeof DecompressionStream !== "function") {
      if (!silentFailure) {
        setStatus("Il browser non supporta la decompressione dei link.", true);
      }
      return null;
    }
    try {
      const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"));
      const buffer = await new Response(stream).arrayBuffer();
      return new Uint8Array(buffer);
    } catch (error) {
      if (!silentFailure) {
        setStatus("Link dati non valido.", true);
      }
      return null;
    }
  }

  async function gzipCompress(text) {
    const encoder = new TextEncoder();
    return gzipCompressBytes(encoder.encode(text));
  }

  async function gzipDecompress(bytes, silentFailure) {
    const decompressed = await gzipDecompressBytes(bytes, silentFailure);
    if (!decompressed) {
      return null;
    }
    return new TextDecoder().decode(decompressed);
  }

  async function initializeFromStorage() {
    const restored = await restoreFromLink();
    if (restored) {
      return;
    }
    if (restoreDatasets()) {
      return;
    }
  }

  function minMaxArray(values) {
    if (!values.length) {
      return { min: 0, max: 0 };
    }
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;
    values.forEach((value) => {
      min = Math.min(min, value);
      max = Math.max(max, value);
    });
    if (!Number.isFinite(min) || !Number.isFinite(max)) {
      return { min: 0, max: 0 };
    }
    return { min, max };
  }

  function flattenMatrix(matrix) {
    return matrix.reduce((acc, row) => acc.concat(row), []);
  }

  function getMinMaxConfig(matrix, rowTotals, colTotals, colorConfig) {
    const matrixValues = flattenMatrix(matrix);
    const totalsValues = []
      .concat(rowTotals || [])
      .concat(colTotals || []);
    const combinedValues = matrixValues.concat(totalsValues);

    const matrixRange = minMaxArray(matrixValues);
    const totalsRange = minMaxArray(totalsValues);
    const combinedRange = minMaxArray(combinedValues);

    const pickRange = (scope) => {
      if (scope === "combined") {
        return combinedRange;
      }
      if (scope === "totals") {
        return totalsRange;
      }
      return matrixRange;
    };

    return {
      matrix: pickRange(colorConfig?.matrix?.scope || "matrix"),
      totals: pickRange(colorConfig?.totals?.scope || "totals")
    };
  }

  function applyHeatmap(cell, value, min, max, palette) {
    if (!palette) {
      return;
    }
    const t = min === max ? 0 : (value - min) / (max - min);
    const color = palette.length === 2 ? interpolateBi(palette, t) : interpolateTri(palette, t);
    cell.style.backgroundColor = color;
    cell.style.color = textColorFor(color);
  }

  function applySummaryColors(table, analysis) {
    const dataRows = Array.from(table.querySelectorAll("tbody tr"));
    const values = analysis.summaryRows;
    if (!dataRows.length) {
      return;
    }

    const columns = [
      { index: 1, palette: PALETTE.posMatrix },
      { index: 2, palette: PALETTE.posMatrix },
      { index: 3, palette: PALETTE.posMatrix },
      { index: 4, palette: PALETTE.posMatrix },
      { index: 5, palette: PALETTE.negMatrix },
      { index: 6, palette: PALETTE.negMatrix },
      { index: 7, palette: PALETTE.negMatrix },
      { index: 8, palette: PALETTE.negMatrix },
      { index: 9, palette: PALETTE.posMatrix },
      { index: 10, palette: PALETTE.negMatrix },
      { index: 11, palette: PALETTE.posMatrix },
      { index: 12, palette: PALETTE.negMatrix },
      { index: 13, palette: PALETTE.posMatrix },
      { index: 14, palette: PALETTE.negMatrix },
      { index: 15, palette: PALETTE.nonRic },
      { index: 16, palette: PALETTE.nonRic }
    ];

    const columnValues = new Map();
    columns.forEach(({ index }) => {
      columnValues.set(index, values.map((row) => getSummaryValue(row, index)));
    });

    const columnRanges = new Map();
    columnValues.forEach((vals, index) => {
      columnRanges.set(index, minMaxArray(vals));
    });

    dataRows.forEach((row, rowIndex) => {
      const cells = row.querySelectorAll("th, td");
      columns.forEach(({ index, palette }) => {
        const cell = cells[index];
        if (!cell) {
          return;
        }
        const value = getSummaryValue(values[rowIndex], index);
        const range = columnRanges.get(index);
        applyHeatmap(cell, value, range.min, range.max, palette);
      });
    });
  }

  function getSummaryValue(row, index) {
    switch (index) {
      case 1:
        return row.positiveReceived.Tecniche;
      case 2:
        return row.positiveReceived.Attitudinali;
      case 3:
        return row.positiveReceived.Sociali;
      case 4:
        return row.positiveReceived.Totali;
      case 5:
        return row.negativeReceived.Tecniche;
      case 6:
        return row.negativeReceived.Attitudinali;
      case 7:
        return row.negativeReceived.Sociali;
      case 8:
        return row.negativeReceived.Totali;
      case 9:
        return row.positiveGiven;
      case 10:
        return row.negativeGiven;
      case 11:
        return row.reciprociPos;
      case 12:
        return row.reciprociNeg;
      case 13:
        return row.forzaLegami;
      case 14:
        return row.forzaAntagonismo;
      case 15:
        return row.nonRicambiateDate;
      case 16:
        return row.nonRicambiateRicevute;
      default:
        return 0;
    }
  }

  function applyLabelFill(cell, value, mode) {
    const normalized = value.toLowerCase();
    if (mode === "positive") {
      if (normalized === "alta") {
        setCellFill(cell, "#ccffcc");
      } else if (normalized === "media") {
        setCellFill(cell, "#ffffcc");
      } else if (normalized === "bassa") {
        setCellFill(cell, "#ffcccc");
      }
      return;
    }

    if (normalized === "alta") {
      setCellFill(cell, "#ffcccc");
    } else if (normalized === "media") {
      setCellFill(cell, "#ffffcc");
    } else if (normalized === "bassa") {
      setCellFill(cell, "#ccffcc");
    }
  }

  function applyEquilibrioFill(cell, value) {
    const normalized = value.toLowerCase();
    if (normalized === "selettiva") {
      setCellFill(cell, "#ffffcc");
    } else if (normalized === "bilanciata") {
      setCellFill(cell, "#ccffcc");
    } else if (normalized === "ignorata") {
      setCellFill(cell, "#ffcccc");
    }
  }

  function applyEtichettaFill(cell, value) {
    const normalized = value.toLowerCase();
    const positiveHints = ["leader positiva", "leader silenziosa", "leader", "stimata", "benvoluta", "apprezzata", "figura di riferimento"];
    const neutralHints = ["neutrale", "bilanciata", "presenza discreta"];
    const negativeHints = ["controversa", "esclusa", "invisibile", "marginale", "trascurata", "non considerata", "dispersa"];

    if (positiveHints.some((hint) => normalized.includes(hint))) {
      setCellFill(cell, "#ccffcc");
      return;
    }
    if (negativeHints.some((hint) => normalized.includes(hint))) {
      setCellFill(cell, "#ffcccc");
      return;
    }
    if (neutralHints.some((hint) => normalized.includes(hint))) {
      setCellFill(cell, "#ffffcc");
    }
  }

  function setCellFill(cell, color) {
    cell.style.backgroundColor = color;
    cell.style.color = "#1c201a";
  }

  function interpolateBi(colors, t) {
    const [a, b] = colors;
    return mixHex(a, b, t);
  }

  function interpolateTri(colors, t) {
    const [a, b, c] = colors;
    if (t <= 0.5) {
      return mixHex(a, b, t / 0.5);
    }
    return mixHex(b, c, (t - 0.5) / 0.5);
  }

  function mixHex(a, b, t) {
    const ca = hexToRgb(a);
    const cb = hexToRgb(b);
    const r = Math.round(ca.r + (cb.r - ca.r) * t);
    const g = Math.round(ca.g + (cb.g - ca.g) * t);
    const bVal = Math.round(ca.b + (cb.b - ca.b) * t);
    return `rgb(${r}, ${g}, ${bVal})`;
  }

  function hexToRgb(hex) {
    const clean = hex.replace("#", "");
    const num = parseInt(clean, 16);
    return {
      r: (num >> 16) & 255,
      g: (num >> 8) & 255,
      b: num & 255
    };
  }

  function textColorFor(rgb) {
    const match = rgb.match(/rgb\\((\\d+),\\s*(\\d+),\\s*(\\d+)\\)/);
    if (!match) {
      return "#1c201a";
    }
    const r = Number(match[1]);
    const g = Number(match[2]);
    const b = Number(match[3]);
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance > 0.65 ? "#1c201a" : "#fdfcf8";
  }

  function extractCode(value) {
    const trimmed = value.trim();
    const match = trimmed.match(/#(?:[^=]*=)?(.+)$/);
    if (match) {
      return match[1];
    }
    if (trimmed.includes(`${LINK_PARAM}=`)) {
      const parts = trimmed.split(`${LINK_PARAM}=`);
      return parts[parts.length - 1].trim();
    }
    return trimmed;
  }
})();
