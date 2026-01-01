(() => {
  const fileInput = document.getElementById("file-input");
  const statusEl = document.getElementById("status");
  const summaryEl = document.getElementById("summary");
  const viewsEl = document.getElementById("views");
  const viewContainer = document.getElementById("view-container");
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
  let currentAnalysis = null;

  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("sw.js").catch(() => undefined);
    });
  }

  fileInput.addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (!file) {
      return;
    }
    readFile(file);
  });


  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      if (!currentAnalysis) {
        return;
      }
      tabs.forEach((btn) => btn.classList.remove("active"));
      tab.classList.add("active");
      const view = tab.dataset.view;
      renderView(view);
    });
  });

  function readFile(file) {
    const reader = new FileReader();
    reader.onload = () => {
      loadCsv(reader.result);
    };
    reader.onerror = () => {
      setStatus("Impossibile leggere il file.", true);
    };
    reader.readAsText(file);
  }

  function loadCsv(text) {
    try {
      const { headers, rows } = parseCsv(text);
      const analysis = analyzeData(headers, rows);
      currentAnalysis = analysis;
      updateSummary(analysis);
      summaryEl.classList.remove("hidden");
      viewsEl.classList.remove("hidden");
      renderView("matrix-pos");
      tabs.forEach((btn) => btn.classList.remove("active"));
      tabs[0].classList.add("active");
      setStatus(`Caricato: ${analysis.rows.length} risposte, ${analysis.names.length} atlete.`, false);
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  function setStatus(message, isError) {
    statusEl.textContent = message;
    statusEl.classList.toggle("error", isError);
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
    const reciprociTot = countReciprociPairs(analysis.matrices.reciprociPos);
    document.getElementById("summary-responses").textContent = analysis.rows.length;
    document.getElementById("summary-athletes").textContent = analysis.names.length;
    document.getElementById("summary-reciproci").textContent = reciprociTot;
    document.getElementById("summary-legami").textContent = formatNumber(analysis.averages.forzaLegami);

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

  function renderView(view) {
    viewContainer.innerHTML = "";
    if (!currentAnalysis) {
      return;
    }

    if (!renderers.has(view)) {
      renderers.set(view, createRenderer(view));
    }
    renderers.get(view)(currentAnalysis, viewContainer);
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
            true
          ));
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
            true
          ));
        };
      case "reciproci-pos":
        return (analysis, container) => {
          container.appendChild(renderReciprociSection(
            "Reciproci positivi",
            "1 indica una scelta reciproca positiva.",
            analysis.names,
            analysis.matrices.reciprociPos
          ));
        };
      case "reciproci-neg":
        return (analysis, container) => {
          container.appendChild(renderReciprociSection(
            "Reciproci negativi",
            "1 indica un reciproco negativo.",
            analysis.names,
            analysis.matrices.reciprociNeg
          ));
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
            true
          ));
        };
      case "riassunto":
        return (analysis, container) => {
          container.appendChild(renderSummarySection(analysis));
        };
      case "rete":
        return (analysis, container) => {
          container.appendChild(renderNetworkSection(analysis));
        };
      default:
        return () => undefined;
    }
  }

  function renderMatrixSection(title, description, names, matrix, rowTotals, colTotals, totalLabel, includeTotalsRow) {
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
      th.textContent = name;
      headerRow.appendChild(th);
    });

    const totalTh = document.createElement("th");
    totalTh.textContent = totalLabel;
    headerRow.appendChild(totalTh);

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    matrix.forEach((row, i) => {
      const tr = document.createElement("tr");
      const rowHeader = document.createElement("th");
      rowHeader.textContent = names[i];
      tr.appendChild(rowHeader);

      row.forEach((cell) => {
        const td = document.createElement("td");
        td.textContent = cell;
        tr.appendChild(td);
      });

      const total = document.createElement("td");
      total.textContent = rowTotals[i];
      tr.appendChild(total);

      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    if (includeTotalsRow && colTotals) {
      const tfoot = document.createElement("tfoot");
      const tr = document.createElement("tr");
      const label = document.createElement("th");
      label.textContent = "Totali ricevuti";
      tr.appendChild(label);

      colTotals.forEach((value) => {
        const td = document.createElement("td");
        td.textContent = value;
        tr.appendChild(td);
      });

      const total = document.createElement("td");
      total.textContent = colTotals.reduce((acc, value) => acc + value, 0);
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

  function renderReciprociSection(title, description, names, matrix) {
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
      th.textContent = name;
      headerRow.appendChild(th);
    });

    const totalTh = document.createElement("th");
    totalTh.textContent = "Totale reciprocità";
    headerRow.appendChild(totalTh);

    const idxTh = document.createElement("th");
    idxTh.textContent = "Indice reciprocità";
    headerRow.appendChild(idxTh);

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    const rowTotals = sumRows(matrix);
    const reciprocalIndex = rowTotals.map((total, idx) => {
      const denom = matrix.length - 1;
      return denom > 0 ? total / denom : 0;
    });

    matrix.forEach((row, i) => {
      const tr = document.createElement("tr");
      const rowHeader = document.createElement("th");
      rowHeader.textContent = names[i];
      tr.appendChild(rowHeader);

      row.forEach((cell) => {
        const td = document.createElement("td");
        td.textContent = cell;
        tr.appendChild(td);
      });

      const total = document.createElement("td");
      total.textContent = rowTotals[i];
      tr.appendChild(total);

      const idxCell = document.createElement("td");
      idxCell.textContent = formatNumber(reciprocalIndex[i]);
      tr.appendChild(idxCell);

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);

    const tfoot = document.createElement("tfoot");
    const tr = document.createElement("tr");
    const label = document.createElement("th");
    label.textContent = "Totali";
    tr.appendChild(label);

    const colTotals = sumColumns(matrix);
    colTotals.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    });

    const totalSum = rowTotals.reduce((acc, value) => acc + value, 0);
    const totalCell = document.createElement("td");
    totalCell.textContent = totalSum;
    tr.appendChild(totalCell);

    const avgIdx = reciprocalIndex.reduce((acc, value) => acc + value, 0) / reciprocalIndex.length;
    const avgCell = document.createElement("td");
    avgCell.textContent = formatNumber(avgIdx);
    tr.appendChild(avgCell);

    tfoot.appendChild(tr);
    table.appendChild(tfoot);

    const wrapper = document.createElement("div");
    wrapper.className = "table-wrap";
    wrapper.appendChild(table);
    section.appendChild(wrapper);

    return section;
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
      "Scelte ricevute + (Tec)",
      "Scelte ricevute + (Att)",
      "Scelte ricevute + (Soc)",
      "Scelte ricevute + (Tot)",
      "Scelte ricevute - (Tec)",
      "Scelte ricevute - (Att)",
      "Scelte ricevute - (Soc)",
      "Scelte ricevute - (Tot)",
      "Scelte date +",
      "Scelte date -",
      "Reciproci +",
      "Reciproci -",
      "Forza legami",
      "Forza antagonismo",
      "Positive date non ric.",
      "Positive ricevute non ric."
    ].forEach((label) => {
      const th = document.createElement("th");
      th.textContent = label;
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
      th.textContent = label;
      classRow.appendChild(th);
    });
    classHead.appendChild(classRow);
    classTable.appendChild(classHead);

    const classBody = document.createElement("tbody");
    analysis.classifications.forEach((row) => {
      const tr = document.createElement("tr");
      [row.name, row.inflPos, row.inflNeg, row.equilibrio, row.label].forEach((value, idx) => {
        const cell = idx === 0 ? document.createElement("th") : document.createElement("td");
        cell.textContent = value;
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

  function renderNetworkSection(analysis) {
    const section = document.createElement("div");
    const heading = document.createElement("h2");
    heading.textContent = "Rete dei reciproci positivi";
    section.appendChild(heading);

    const note = document.createElement("p");
    note.className = "note";
    note.textContent = "Il grafo mostra solo i legami reciproci positivi.";
    section.appendChild(note);

    const network = document.createElement("div");
    network.className = "network";

    const canvas = document.createElement("canvas");
    canvas.width = 900;
    canvas.height = 520;
    network.appendChild(canvas);

    const list = document.createElement("div");
    list.className = "network-list";
    list.innerHTML = "<h3>Clique trovate</h3>";
    const cliques = findCliques(analysis.matrices.reciprociPos, analysis.names);
    if (cliques.length === 0) {
      const p = document.createElement("p");
      p.textContent = "Nessuna clique rilevante.";
      list.appendChild(p);
    } else {
      cliques.forEach((clique) => {
        const p = document.createElement("p");
        p.textContent = `• ${clique.join(", ")}`;
        list.appendChild(p);
      });
    }
    network.appendChild(list);

    section.appendChild(network);
    drawNetwork(canvas, analysis.names, analysis.matrices.reciprociPos);

    return section;
  }

  function drawNetwork(canvas, names, matrix) {
    const ctx = canvas.getContext("2d");
    const width = canvas.width;
    const height = canvas.height;

    const nodes = names.map((name) => ({ name }));
    const edges = [];
    for (let i = 0; i < matrix.length; i += 1) {
      for (let j = i + 1; j < matrix.length; j += 1) {
        if (matrix[i][j] > 0) {
          edges.push([i, j]);
        }
      }
    }

    layoutGraph(nodes, edges, width, height);

    const degrees = nodes.map((_, idx) => edges.filter((e) => e[0] === idx || e[1] === idx).length);
    const maxDegree = Math.max(...degrees, 1);

    ctx.clearRect(0, 0, width, height);
    ctx.strokeStyle = "rgba(0, 0, 0, 0.2)";
    ctx.lineWidth = 1;

    edges.forEach(([i, j]) => {
      ctx.beginPath();
      ctx.moveTo(nodes[i].x, nodes[i].y);
      ctx.lineTo(nodes[j].x, nodes[j].y);
      ctx.stroke();
    });

    nodes.forEach((node, idx) => {
      const degree = degrees[idx];
      const radius = 8 + (degree / maxDegree) * 16;
      const t = degree / maxDegree;
      const hue = 120 * t;

      ctx.beginPath();
      ctx.fillStyle = `hsl(${hue}, 65%, 45%)`;
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

  function findCliques(matrix, names) {
    const n = names.length;
    const adj = Array.from({ length: n }, (_, i) =>
      new Set(matrix[i].map((value, idx) => (value > 0 ? idx : null)).filter((v) => v !== null))
    );
    const cliques = [];

    function bronKerbosch(r, p, x) {
      if (p.size === 0 && x.size === 0) {
        if (r.length >= 3) {
          cliques.push(r.map((idx) => names[idx]));
        }
        return;
      }
      const pArray = Array.from(p);
      for (const v of pArray) {
        const newR = r.concat(v);
        const newP = new Set([...p].filter((u) => adj[v].has(u)));
        const newX = new Set([...x].filter((u) => adj[v].has(u)));
        bronKerbosch(newR, newP, newX);
        p.delete(v);
        x.add(v);
      }
    }

    bronKerbosch([], new Set(Array.from({ length: n }, (_, i) => i)), new Set());
    return cliques;
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
})();
