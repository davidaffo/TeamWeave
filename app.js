(() => {
  const fileInput = document.getElementById("file-input");
  const importCodeInput = document.getElementById("import-code");
  const importCodeBtn = document.getElementById("import-code-btn");
  const datasetBox = document.getElementById("dataset-box");
  const datasetSelect = document.getElementById("dataset-select");
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

  const renderers = new Map();
  const datasets = [];
  let currentAnalysis = null;
  let currentDatasetId = null;
  let datasetCounter = 1;
  let isRestoringDatasets = false;
  const DATASETS_KEY = "teamweave:datasets";
  const DATASET_STATE_KEY = "teamweave:datasetState";
  const VIEW_KEY = "teamweave:lastView";
  const LINK_PARAM = "data";
  const PALETTE = {
    posMatrix: ["#ffffff", "#00a933"],
    posTotals: ["#ffffff", "#81d41a"],
    negMatrix: ["#ffffff", "#ff0000"],
    negTotals: ["#ffffff", "#800080"],
    nonRic: ["#ffffff", "#ff8000"],
    network: ["#d73027", "#fee08b", "#1a9850"]
  };

  if ("serviceWorker" in navigator) {
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
  });

  compareSelectB.addEventListener("change", () => {
    ensureCompareSelection();
    const active = getActiveView() || "responses";
    renderView(active);
    saveDatasetState();
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
      }
      datasets.push(dataset);
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
      const payload = datasets.map((dataset) => ({
        id: dataset.id,
        name: dataset.name,
        csv: dataset.csv
      }));
      localStorage.setItem(DATASETS_KEY, JSON.stringify(payload));
    } catch (error) {
      setStatus("Memoria del browser piena: impossibile salvare i dataset.", true);
    }
  }

  function saveDatasetState() {
    try {
      const state = {
        currentDatasetId,
        compareA: compareSelectA?.value || null,
        compareB: compareSelectB?.value || null
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
      entries.forEach((entry) => {
        if (!entry || !entry.csv) {
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
      const category = CATEGORY_MAP_NORMALIZED.get(normalizeQuestion(question)) || null;
      return {
        question,
        category,
        positive: idx % 2 === 0
      };
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
    const receivedPosByCat = initCategoryMap(names, CATEGORIES);
    const receivedNegByCat = initCategoryMap(names, CATEGORIES);
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
            }
            receivedPosByCat.get(pickKey)[meta.category] += 1;
          } else {
            if (chooserIndex !== pickIndex) {
              negMatrix[chooserIndex][pickIndex] += 1;
            }
            receivedNegByCat.get(pickKey)[meta.category] += 1;
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
        label: synthesizeLabel(inflPos, inflNeg, equilibrio)
      };
    });

    return {
      headers,
      rows: rowsNormalized,
      names,
      matrices: {
        posMatrix,
        negMatrix,
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

  function normalizeQuestion(question) {
    return question
      .trim()
      .replace(/[’‘]/g, "'")
      .replace(/\s+/g, " ")
      .toLowerCase();
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

  function setCurrentDataset(id, options = {}) {
    const { persist = true } = options;
    const dataset = datasets.find((entry) => entry.id === id);
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
  }

  function updateDatasetControls() {
    if (!datasets.length) {
      datasetBox.classList.add("hidden");
      return;
    }
    datasetBox.classList.remove("hidden");
    populateDatasetSelect(datasetSelect, currentDatasetId);
    populateDatasetSelect(compareSelectA, compareSelectA.value);
    populateDatasetSelect(compareSelectB, compareSelectB.value);
    ensureCompareSelection();
    const showCompare = datasets.length >= 2;
    compareRow.classList.toggle("hidden", !showCompare);
    compareNote.classList.toggle("hidden", !showCompare);
    compareSelectA.disabled = !showCompare;
    compareSelectB.disabled = !showCompare;
    if (!isRestoringDatasets) {
      saveDatasetState();
    }
  }

  function populateDatasetSelect(select, selectedId) {
    select.textContent = "";
    datasets.forEach((dataset) => {
      const option = document.createElement("option");
      option.value = dataset.id;
      option.textContent = dataset.name;
      select.appendChild(option);
    });
    const targetId = datasets.some((dataset) => dataset.id === selectedId)
      ? selectedId
      : datasets[0]?.id;
    if (targetId) {
      select.value = targetId;
    }
  }

  function ensureCompareSelection() {
    if (datasets.length < 2) {
      return;
    }
    if (compareSelectA.value === compareSelectB.value) {
      const alternative = datasets.find((dataset) => dataset.id !== compareSelectA.value);
      if (alternative) {
        compareSelectB.value = alternative.id;
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
      case "responses":
        return (analysis, container) => {
          container.appendChild(renderResponsesSection(analysis));
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
    if (datasets.length < 2) {
      return null;
    }
    const datasetA = datasets.find((dataset) => dataset.id === compareSelectA.value);
    const datasetB = datasets.find((dataset) => dataset.id === compareSelectB.value);
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
    if (panels.length === 2) {
      highlightDifferences(panels[0], panels[1]);
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
    firstTh.textContent = "Chi sceglie ↓ / Chi viene scelto →";
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
    firstTh.textContent = "Chi sceglie ↓ / Chi viene scelto →";
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

    const classTable = document.createElement("table");
    classTable.className = "matrix-table";
    const classHead = document.createElement("thead");
    const classRow = document.createElement("tr");
    ["Atleta", "Influenza positiva", "Influenza negativa", "Equilibrio relazionale", "Etichetta sintetica"].forEach((label) => {
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
      [row.name, row.inflPos, row.inflNeg, row.equilibrio, row.label].forEach((value, idx) => {
        const cell = idx === 0 ? document.createElement("th") : document.createElement("td");
        cell.textContent = value;
        if (idx === 1) {
          applyLabelFill(cell, value, "positive");
          if (influenceRank[value]) {
            cell.dataset.rank = influenceRank[value];
          }
        } else if (idx === 2) {
          applyLabelFill(cell, value, "negative");
          if (influenceRank[value]) {
            cell.dataset.rank = influenceRank[value];
          }
        } else if (idx === 3) {
          applyEquilibrioFill(cell, value);
          if (equilibrioRank[value]) {
            cell.dataset.rank = equilibrioRank[value];
          }
        } else if (idx === 4) {
          applyEtichettaFill(cell, value);
        }
        tr.appendChild(cell);
      });
      classBody.appendChild(tr);
    });
    classTable.appendChild(classBody);

    const classWrap = document.createElement("div");
    classWrap.className = "table-wrap";
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
