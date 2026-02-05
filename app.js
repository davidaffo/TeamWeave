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
  const IS_EXPORT_MODE = Boolean(window.TEAMWEAVE_EXPORT_MODE);
  const EXPORT_ORDER = Array.isArray(window.TEAMWEAVE_EXPORT_ORDER)
    ? window.TEAMWEAVE_EXPORT_ORDER.slice()
    : null;
  const INLINE_STYLES = ":root {\n  color-scheme: light;\n  --bg: #f5f1e8;\n  --bg-2: #f0e5d2;\n  --ink: #1c201a;\n  --muted: #4b4f44;\n  --accent: #b65a2a;\n  --accent-2: #2d6a4f;\n  --card: #ffffff;\n  --border: rgba(28, 32, 26, 0.12);\n  --shadow: 0 18px 40px rgba(27, 29, 25, 0.14);\n}\n\n* {\n  box-sizing: border-box;\n}\n\nbody {\n  margin: 0;\n  font-family: \"Palatino Linotype\", \"Book Antiqua\", Palatino, \"Georgia\", serif;\n  color: var(--ink);\n  background: radial-gradient(circle at top left, #fffaf0 0%, var(--bg) 60%, var(--bg-2) 100%);\n  min-height: 100vh;\n}\n\n.main-tabs {\n  display: flex;\n  gap: 0.6rem;\n  flex-wrap: wrap;\n  margin: 0 0 1.2rem;\n}\n\n.main-tab {\n  border: 1px solid var(--border);\n  background: rgba(255, 255, 255, 0.8);\n  color: var(--ink);\n  padding: 0.65rem 1.1rem;\n  border-radius: 999px;\n  font-weight: 600;\n  cursor: pointer;\n}\n\n.main-tab.active {\n  background: var(--accent);\n  color: #fff7f0;\n  border-color: transparent;\n}\n\n.bg-shapes {\n  position: fixed;\n  inset: 0;\n  background:\n    radial-gradient(circle at 10% 20%, rgba(182, 90, 42, 0.18), transparent 45%),\n    radial-gradient(circle at 90% 15%, rgba(45, 106, 79, 0.16), transparent 40%),\n    radial-gradient(circle at 20% 90%, rgba(27, 111, 192, 0.12), transparent 50%);\n  pointer-events: none;\n  z-index: 0;\n}\n\n.hero {\n  position: relative;\n  display: flex;\n  justify-content: space-between;\n  align-items: flex-end;\n  gap: 2rem;\n  padding: 3rem clamp(1.5rem, 4vw, 4rem) 2rem;\n  z-index: 1;\n}\n\n.hero-content h1 {\n  margin: 0 0 0.5rem;\n  font-size: clamp(2.4rem, 4vw, 4rem);\n  letter-spacing: -0.02em;\n}\n\n.eyebrow {\n  text-transform: uppercase;\n  letter-spacing: 0.2em;\n  font-size: 0.8rem;\n  margin: 0 0 0.5rem;\n  color: var(--accent-2);\n}\n\n.subtitle {\n  margin: 0;\n  max-width: 32rem;\n  color: var(--muted);\n  font-size: 1.05rem;\n}\n\n.hero-badge {\n  background: linear-gradient(160deg, #183c2f, #2d6a4f);\n  color: #fdfcf8;\n  padding: 1.2rem 1.6rem;\n  border-radius: 1.2rem;\n  box-shadow: var(--shadow);\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n}\n\n.hero-badge span {\n  display: block;\n  text-transform: uppercase;\n  letter-spacing: 0.2em;\n  font-size: 0.7rem;\n  opacity: 0.8;\n}\n\n.hero-badge strong {\n  font-size: 1.1rem;\n}\n\n.main {\n  position: relative;\n  z-index: 1;\n  padding: 0 clamp(1.5rem, 4vw, 4rem) 4rem;\n}\n\n.panel {\n  background: var(--card);\n  padding: 1.8rem;\n  border-radius: 1.2rem;\n  box-shadow: var(--shadow);\n  border: 1px solid var(--border);\n  margin-bottom: 2rem;\n}\n\n.upload-row {\n  display: flex;\n  flex-wrap: wrap;\n  gap: 1rem;\n  margin-top: 1rem;\n}\n\n.export-box {\n  margin-top: 1.2rem;\n  display: grid;\n  gap: 0.6rem;\n}\n\n.export-box label {\n  font-weight: 600;\n}\n\n.export-row {\n  display: flex;\n  gap: 0.6rem;\n  flex-wrap: wrap;\n}\n\n.export-row input {\n  flex: 1 1 260px;\n  padding: 0.6rem 0.8rem;\n  border-radius: 0.6rem;\n  border: 1px solid var(--border);\n  background: #fdfcf8;\n}\n\n.import-box {\n  margin-top: 1.2rem;\n  display: grid;\n  gap: 0.6rem;\n}\n\n.import-box label {\n  font-weight: 600;\n}\n\n#import-code {\n  width: 100%;\n  min-height: 110px;\n  padding: 0.8rem;\n  border-radius: 0.8rem;\n  border: 1px solid var(--border);\n  background: #fdfcf8;\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n  resize: vertical;\n}\n\n#tw-form-output {\n  white-space: pre-wrap;\n}\n\n.tw-form-box {\n  margin-top: 1.2rem;\n  display: grid;\n  gap: 0.6rem;\n  padding: 1rem;\n  border-radius: 1rem;\n  border: 1px solid var(--border);\n  background: #fffefb;\n}\n\n.tw-categories {\n  display: grid;\n  gap: 0.4rem;\n  margin-top: 0.6rem;\n}\n\n.tw-category-list {\n  display: grid;\n  gap: 0.6rem;\n}\n\n.tw-category-row {\n  display: grid;\n  grid-template-columns: minmax(140px, 1fr) 2fr;\n  gap: 0.6rem;\n  align-items: center;\n  padding: 0.6rem 0.8rem;\n  border-radius: 0.8rem;\n  border: 1px solid rgba(28, 32, 26, 0.08);\n  background: #f9f4ea;\n}\n\n.tw-category-label {\n  font-weight: 600;\n}\n\n.tw-category-sep {\n  color: var(--muted);\n  font-weight: 400;\n  margin: 0 0.35rem;\n}\n\n.tw-category-options {\n  display: flex;\n  gap: 0.8rem;\n  flex-wrap: wrap;\n}\n\n.tw-category-option {\n  display: inline-flex;\n  align-items: center;\n  gap: 0.4rem;\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n}\n\n.tw-category-option input {\n  width: 1rem;\n  height: 1rem;\n}\n\n.tw-form-box h3 {\n  margin: 0;\n}\n\n.form-grid {\n  display: grid;\n  grid-template-columns: minmax(160px, 1fr) minmax(220px, 2fr);\n  gap: 0.6rem 1rem;\n  align-items: center;\n}\n\n.form-grid input,\n.form-grid select {\n  padding: 0.6rem 0.8rem;\n  border-radius: 0.6rem;\n  border: 1px solid var(--border);\n  background: #fdfcf8;\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n}\n\n.tall {\n  min-height: 140px;\n  padding: 0.8rem;\n  border-radius: 0.8rem;\n  border: 1px solid var(--border);\n  background: #fdfcf8;\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n  resize: vertical;\n}\n\n.result.pre {\n  border-radius: 0.8rem;\n  border: 1px solid var(--border);\n  background: #f1e7d2;\n  padding: 0.8rem;\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n  white-space: pre-wrap;\n}\n\n.parsed .question {\n  padding: 0.4rem 0.6rem;\n  border-radius: 0.6rem;\n  border: 1px solid rgba(28, 32, 26, 0.08);\n  background: #fffefb;\n  margin-bottom: 0.4rem;\n}\n\n.dataset-box {\n  margin-top: 1.2rem;\n  display: grid;\n  gap: 0.8rem;\n}\n\n.dataset-row {\n  display: grid;\n  gap: 0.4rem;\n}\n\n.dataset-row label {\n  font-weight: 600;\n}\n\n.dataset-row select {\n  padding: 0.6rem 0.8rem;\n  border-radius: 0.6rem;\n  border: 1px solid var(--border);\n  background: #fdfcf8;\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n}\n\n.dataset-manager {\n  display: grid;\n  gap: 0.6rem;\n}\n\n.dataset-manager-grid {\n  display: grid;\n  gap: 0.6rem;\n}\n\n.dataset-manager-item {\n  display: grid;\n  grid-template-columns: minmax(200px, 1fr) auto;\n  gap: 0.6rem 1rem;\n  align-items: center;\n  padding: 0.5rem 0.7rem;\n  border-radius: 0.8rem;\n  border: 1px solid var(--border);\n  background: #fffefb;\n}\n\n.dataset-manager-toggle {\n  display: flex;\n  align-items: center;\n  gap: 0.6rem;\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n}\n\n.dataset-manager-toggle input {\n  width: 1rem;\n  height: 1rem;\n}\n\n.dataset-manager-order {\n  display: flex;\n  align-items: center;\n  gap: 0.8rem;\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n  font-size: 0.9rem;\n}\n\n.dataset-manager-order label {\n  display: inline-flex;\n  align-items: center;\n  gap: 0.35rem;\n}\n\n.dataset-manager-order input {\n  width: 1rem;\n  height: 1rem;\n}\n\n.dataset-picker {\n  display: grid;\n  gap: 0.6rem;\n  margin-top: 0.8rem;\n}\n\n.dataset-option {\n  display: flex;\n  align-items: center;\n  gap: 0.6rem;\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n}\n\n.dataset-option input {\n  width: 1rem;\n  height: 1rem;\n}\n\n.compare-selects {\n  display: grid;\n  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));\n  gap: 0.6rem;\n}\n\ninput[type=\"file\"] {\n  flex: 1 1 250px;\n  padding: 0.6rem;\n  border-radius: 0.6rem;\n  border: 1px dashed var(--border);\n  background: #fdfcf8;\n}\n\nbutton {\n  border: none;\n  padding: 0.75rem 1.2rem;\n  border-radius: 999px;\n  background: var(--accent);\n  color: #fff7f0;\n  cursor: pointer;\n  font-weight: 600;\n  transition: transform 0.2s ease, box-shadow 0.2s ease;\n}\n\nbutton:hover {\n  transform: translateY(-1px);\n  box-shadow: 0 10px 16px rgba(182, 90, 42, 0.25);\n}\n\n.status {\n  margin-top: 1rem;\n  color: var(--muted);\n  font-size: 0.95rem;\n}\n\n.status.error {\n  color: #9b2c2c;\n}\n\n.summary {\n  display: grid;\n  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));\n  gap: 1rem;\n  margin-bottom: 2rem;\n}\n\n.summary-card {\n  background: rgba(255, 255, 255, 0.9);\n  border-radius: 1rem;\n  padding: 1rem 1.4rem;\n  border: 1px solid var(--border);\n  display: flex;\n  flex-direction: column;\n  gap: 0.4rem;\n}\n\n.summary-card-pos {\n  background: #e4f5e6;\n  border-color: rgba(45, 106, 79, 0.3);\n}\n\n.summary-card-neg {\n  background: #fde3e3;\n  border-color: rgba(215, 48, 39, 0.3);\n}\n\n.summary-card strong {\n  font-size: 1.6rem;\n}\n\n.tabs {\n  display: flex;\n  flex-wrap: wrap;\n  gap: 0.6rem;\n  margin-bottom: 1.5rem;\n}\n\n.tab {\n  background: rgba(255, 255, 255, 0.8);\n  color: var(--ink);\n  border: 1px solid var(--border);\n}\n\n.tab.active {\n  background: var(--accent-2);\n  color: #fdfcf8;\n  border-color: transparent;\n}\n\n.view-container {\n  background: var(--card);\n  border-radius: 1.2rem;\n  padding: 1.4rem;\n  border: 1px solid var(--border);\n  box-shadow: var(--shadow);\n}\n\n.dual-view {\n  background: transparent;\n  border: none;\n  box-shadow: none;\n  padding: 0;\n  display: grid;\n  grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));\n  gap: 1rem;\n}\n\n.view-panel {\n  background: var(--card);\n  border-radius: 1.2rem;\n  padding: 1.4rem;\n  border: 1px solid var(--border);\n  box-shadow: var(--shadow);\n}\n\n.table-wrap {\n  overflow: auto;\n  max-height: 70vh;\n  border-radius: 0.8rem;\n  border: 1px solid var(--border);\n}\n\n.matrix-table {\n  width: max-content;\n  border-collapse: collapse;\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n  font-size: 0.85rem;\n}\n\n.matrix-table th,\n.matrix-table td {\n  border: 1px solid rgba(28, 32, 26, 0.08);\n  padding: 0.35rem 0.5rem;\n  text-align: center;\n  min-width: 2.4rem;\n  background: #fffefb;\n}\n\n.matrix-table td {\n  position: relative;\n  padding-right: 1.3rem;\n  line-height: 1.1;\n}\n\n.diff-arrow {\n  position: absolute;\n  right: 0.3rem;\n  top: 50%;\n  transform: translateY(-50%);\n  font-weight: 800;\n  font-size: 1.05rem;\n  line-height: 1;\n}\n\n.diff-up {\n  color: #1a9850;\n}\n\n.diff-down {\n  color: #d73027;\n}\n\n.matrix-table td.delta-good {\n  background-color: rgba(129, 212, 26, 0.35) !important;\n}\n\n.matrix-table td.delta-bad {\n  background-color: rgba(215, 48, 39, 0.35) !important;\n}\n\n.matrix-table thead th {\n  position: sticky;\n  top: 0;\n  background: #f7f1e4;\n  z-index: 2;\n}\n\n.matrix-table tbody th {\n  position: sticky;\n  left: 0;\n  background: #f7f1e4;\n  text-align: left;\n  z-index: 1;\n}\n\n.matrix-table thead .sticky-right {\n  position: sticky;\n  right: 0;\n  background: #f7f1e4;\n  z-index: 3;\n}\n\n.matrix-table tbody .sticky-right {\n  position: sticky;\n  right: 0;\n  background: #fffefb;\n  z-index: 1;\n}\n\n.matrix-table tfoot .sticky-right {\n  position: sticky;\n  right: 0;\n  background: #f1e7d2;\n  z-index: 2;\n}\n\n.matrix-table tfoot th,\n.matrix-table tfoot td {\n  background: #f1e7d2;\n  font-weight: 600;\n}\n\n.note {\n  color: var(--muted);\n  margin-bottom: 1rem;\n}\n\n.responses-list {\n  display: grid;\n  gap: 1rem;\n}\n\n.response-card {\n  border-radius: 1rem;\n  border: 1px solid var(--border);\n  background: #fffefb;\n  padding: 0.8rem 1rem;\n}\n\n.response-card summary {\n  cursor: pointer;\n  list-style: none;\n  font-weight: 600;\n  display: flex;\n  justify-content: space-between;\n  align-items: center;\n  gap: 1rem;\n}\n\n.response-card summary::-webkit-details-marker {\n  display: none;\n}\n\n.response-meta {\n  display: flex;\n  gap: 0.6rem;\n  font-size: 0.85rem;\n  color: var(--muted);\n}\n\n.response-grid {\n  display: grid;\n  gap: 0.8rem;\n  margin-top: 0.8rem;\n}\n\n.response-group {\n  border-left: 3px solid var(--border);\n  padding-left: 0.8rem;\n}\n\n.response-group h4 {\n  margin: 0 0 0.4rem;\n}\n\n.response-group ul {\n  margin: 0;\n  padding-left: 1rem;\n}\n\n.response-items {\n  display: grid;\n  gap: 0.5rem;\n}\n\n.response-item {\n  display: grid;\n  grid-template-columns: minmax(180px, 1.2fr) minmax(200px, 1fr);\n  gap: 0.8rem;\n  align-items: start;\n  padding: 0.5rem 0.6rem;\n  border-radius: 0.6rem;\n  background: #f9f4ea;\n}\n\n.response-item.positive {\n  background: #e4f5e6;\n}\n\n.response-item.negative {\n  background: #fde3e3;\n}\n\n.response-question {\n  font-weight: 600;\n}\n\n.response-answer {\n  color: var(--muted);\n}\n\n.response-answer.empty {\n  color: #9b8f7a;\n  font-style: italic;\n}\n\n.response-pill {\n  display: inline-block;\n  padding: 0.2rem 0.6rem;\n  border-radius: 999px;\n  font-size: 0.75rem;\n  background: #f1e7d2;\n  margin-left: 0.4rem;\n}\n\n.top-questions-controls {\n  display: flex;\n  gap: 1rem;\n  flex-wrap: wrap;\n  align-items: center;\n  margin-bottom: 1rem;\n}\n\n.top-questions-controls label {\n  display: inline-flex;\n  align-items: center;\n  gap: 0.6rem;\n  font-weight: 600;\n}\n\n.top-questions-controls input {\n  width: 4.5rem;\n  padding: 0.35rem 0.5rem;\n  border-radius: 0.5rem;\n  border: 1px solid var(--border);\n  background: #fdfcf8;\n  font-family: \"Trebuchet MS\", \"Gill Sans\", \"Segoe UI\", sans-serif;\n}\n\n.top-questions-list {\n  display: grid;\n  gap: 1rem;\n}\n\n.top-questions-card h3 {\n  margin: 0 0 0.6rem;\n}\n\n.top-questions-grid {\n  display: grid;\n  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));\n  gap: 1rem;\n}\n\n.top-questions-column h4 {\n  margin: 0 0 0.5rem;\n}\n\n.top-questions-items {\n  margin: 0;\n  padding-left: 1.2rem;\n  display: grid;\n  gap: 0.4rem;\n}\n\n.top-questions-item {\n  display: flex;\n  justify-content: space-between;\n  gap: 0.6rem;\n  align-items: baseline;\n  padding: 0.35rem 0.5rem;\n  border-radius: 0.5rem;\n}\n\n.top-questions-positive .top-questions-item {\n  background: #e4f5e6;\n}\n\n.top-questions-negative .top-questions-item {\n  background: #fde3e3;\n}\n\n.top-questions-count {\n  font-weight: 700;\n  background: #f1e7d2;\n  border-radius: 999px;\n  padding: 0.1rem 0.5rem;\n  white-space: nowrap;\n}\n\n.network {\n  display: grid;\n  grid-template-columns: 1fr;\n  gap: 1.6rem;\n}\n\n.network-legend {\n  display: grid;\n  gap: 0.4rem;\n  margin-bottom: 1rem;\n  font-size: 0.85rem;\n  color: var(--muted);\n}\n\n.legend-row {\n  display: grid;\n  grid-template-columns: auto 1fr auto;\n  align-items: center;\n  gap: 0.6rem;\n}\n\n.legend-bar {\n  height: 8px;\n  border-radius: 999px;\n  border: 1px solid var(--border);\n}\n\n.legend-range {\n  font-size: 0.8rem;\n}\n\n.network canvas {\n  width: 100%;\n  height: 520px;\n  display: block;\n  border-radius: 1rem;\n  border: 1px solid var(--border);\n  background: #fffefb;\n}\n\n.network-list {\n  background: #fffefb;\n  border-radius: 1rem;\n  border: 1px solid var(--border);\n  padding: 1rem;\n  height: 520px;\n  overflow: auto;\n}\n\n.network-list h3 {\n  margin-top: 0;\n}\n\n.hidden {\n  display: none;\n}\n\n.footer {\n  padding: 2rem clamp(1.5rem, 4vw, 4rem);\n  text-align: center;\n  color: var(--muted);\n}\n\n@media (max-width: 900px) {\n  .hero {\n    flex-direction: column;\n    align-items: flex-start;\n  }\n\n  .hero-badge {\n    align-self: stretch;\n  }\n\n  .network {\n    grid-template-columns: 1fr;\n  }\n\n  .network canvas,\n  .network-list {\n    height: 420px;\n  }\n}\n\n@media (max-width: 600px) {\n  .form-grid {\n    grid-template-columns: 1fr;\n  }\n\n  .tw-category-row {\n    grid-template-columns: 1fr;\n  }\n\n  .matrix-table {\n    font-size: 0.75rem;\n  }\n\n  button {\n    width: 100%;\n  }\n}\n";
  const INLINE_APP = "(() => {\n  const fileInput = document.getElementById(\"file-input\");\n  const importCodeInput = document.getElementById(\"import-code\");\n  const importCodeBtn = document.getElementById(\"import-code-btn\");\n  const datasetBox = document.getElementById(\"dataset-box\");\n  const mainTabs = Array.from(document.querySelectorAll(\".main-tab\"));\n  const analysisSection = document.getElementById(\"analysis-section\");\n  const generatorSection = document.getElementById(\"generator-section\");\n  const twFormTitle = document.getElementById(\"tw-form-title\");\n  const twFormLimit = document.getElementById(\"tw-form-limit\");\n  const twFormPlayers = document.getElementById(\"tw-form-players\");\n  const twFormQuestions = document.getElementById(\"tw-form-questions\");\n  const twFormCategories = document.getElementById(\"tw-form-categories\");\n  const twFormConvert = document.getElementById(\"tw-form-convert\");\n  const twFormCopy = document.getElementById(\"tw-form-copy\");\n  const twFormDownload = document.getElementById(\"tw-form-download\");\n  const twFormDownloadQuestions = document.getElementById(\"tw-form-download-questions\");\n  const twFormOutput = document.getElementById(\"tw-form-output\");\n  const twFormPreview = document.getElementById(\"tw-form-preview\");\n  const datasetSelect = document.getElementById(\"dataset-select\");\n  const datasetManager = document.getElementById(\"dataset-manager\");\n  const deleteDatasetBtn = document.getElementById(\"delete-dataset-btn\");\n  const compareRow = document.getElementById(\"compare-row\");\n  const compareNote = document.getElementById(\"compare-note\");\n  const compareSelectA = document.getElementById(\"compare-a\");\n  const compareSelectB = document.getElementById(\"compare-b\");\n  const statusEl = document.getElementById(\"status\");\n  const summaryEl = document.getElementById(\"summary\");\n  const viewsEl = document.getElementById(\"views\");\n  const viewContainer = document.getElementById(\"view-container\");\n  const exportLinkBtn = document.getElementById(\"export-link-btn\");\n  const exportCsvBtn = document.getElementById(\"export-csv-btn\");\n  const exportHtmlBtn = document.getElementById(\"export-html-btn\");\n  const exportBox = document.getElementById(\"export-box\");\n  const exportLinkInput = document.getElementById(\"export-link\");\n  const copyLinkBtn = document.getElementById(\"copy-link-btn\");\n  const tabs = Array.from(document.querySelectorAll(\".tab\"));\n\n  const CATEGORY_MAP = new Map([\n    [\"Immagina di andare in trasferta. Chi vorresti in camera con te? (Seleziona 3)\", \"Sociali\"],\n    [\"Immagina di andare in trasferta. Chi non vorresti in camera con te? (Seleziona 3)\", \"Sociali\"],\n    [\"Chi preferiresti che l'allenatore avesse scelto come capitano della squadra? (Seleziona 3)\", \"Attitudinali\"],\n    [\"Chi non sarebbe adatta a fare il capitano? (Seleziona 3)\", \"Attitudinali\"],\n    [\"Immagina di giocare una partita decisiva. Chi vorresti con te in campo? (Seleziona 3)\", \"Tecniche\"],\n    [\"Chi non vorresti con te in campo in una partita decisiva? (Seleziona 3)\", \"Tecniche\"],\n    [\"Sei al tie-break (set dello spareggio). Chi secondo te dovrebbe giocare? (Seleziona 3)\", \"Tecniche\"],\n    [\"Chi non dovrebbe giocare al tie-break? (Seleziona 3)\", \"Tecniche\"],\n    [\"Durante l'allenamento, in un esercizio a coppie, chi vorresti con te? (Seleziona 3)\", \"Attitudinali\"],\n    [\"Durante l'allenamento, chi non vorresti come compagna in un esercizio a coppie? (Seleziona 3)\", \"Attitudinali\"],\n    [\"Durante l'allenamento, chi secondo te d\u00e0 sempre il massimo? (Seleziona 3)\", \"Attitudinali\"],\n    [\"Chi secondo te non d\u00e0 sempre il massimo durante l'allenamento? (Seleziona 3)\", \"Attitudinali\"],\n    [\"In un momento decisivo della partita, chi vorresti andasse a battere? (Seleziona 3)\", \"Tecniche\"],\n    [\"In un momento decisivo, chi preferiresti non mandare a battere? (Seleziona 3)\", \"Tecniche\"],\n    [\"Sei l'alzatore, sul 24-23 per la tua squadra. A chi alzi il pallone? (Seleziona 3)\", \"Tecniche\"],\n    [\"Chi non vorresti che attaccasse sul 24-23? (Seleziona 3)\", \"Tecniche\"],\n    [\"Se una sera sei sola, con chi ti piacerebbe uscire? (Seleziona 3)\", \"Sociali\"],\n    [\"Con chi non ti andrebbe di uscire una sera? (Seleziona 3)\", \"Sociali\"],\n    [\"Devi creare una squadra forte per vincere con altre 3 persone. Chi scegli? (Seleziona 3)\", \"Tecniche\"],\n    [\"Chi non sceglieresti per creare una squadra forte? (Seleziona 3)\", \"Tecniche\"],\n    [\"Chi tra le tue compagne pu\u00f2 dare il maggior contributo in un momento critico della partita? (Seleziona 3)\", \"Attitudinali\"],\n    [\"Chi non pensi possa contribuire in un momento critico della partita? (Seleziona 3)\", \"Attitudinali\"],\n    [\"Hai un problema e vuoi confrontarti con delle compagne di squadra. Con chi ne parli? (Seleziona 3)\", \"Sociali\"],\n    [\"Con chi non ti confronteresti per parlare di un problema? (Seleziona 3)\", \"Sociali\"],\n    [\"La squadra ha dei problemi. Chi potrebbe farsi portavoce presso l\u2019allenatore o i dirigenti? (Seleziona 3)\", \"Attitudinali\"],\n    [\"Chi non sarebbe adatta a farsi portavoce della squadra? (Seleziona 3)\", \"Attitudinali\"],\n    [\"Chi sono secondo te le pi\u00f9 forti della squadra? (Seleziona 3)\", \"Tecniche\"],\n    [\"Chi sono secondo te le pi\u00f9 scarse della squadra? (Seleziona 3)\", \"Tecniche\"],\n    [\"Chi sono le compagne pi\u00f9 simpatiche? (Seleziona 3)\", \"Sociali\"],\n    [\"Chi sono le compagne pi\u00f9 antipatiche? (Seleziona 3)\", \"Sociali\"]\n  ]);\n\n  const CATEGORIES = [\"Tecniche\", \"Attitudinali\", \"Sociali\"];\n  const CATEGORY_MAP_NORMALIZED = new Map(\n    Array.from(CATEGORY_MAP.entries()).map(([question, category]) => [\n      normalizeQuestion(question),\n      category\n    ])\n  );\n\n  const renderers = new Map();\n  const datasets = [];\n  let currentAnalysis = null;\n  let currentDatasetId = null;\n  let datasetCounter = 1;\n  let visibleDatasetIds = new Set();\n  let hasVisibleState = false;\n  let latestTeamweaveScript = \"\";\n  let isRestoringDatasets = false;\n  const DATASETS_KEY = \"teamweave:datasets\";\n  const DATASET_STATE_KEY = \"teamweave:datasetState\";\n  const VIEW_KEY = \"teamweave:lastView\";\n  const LINK_PARAM = \"data\";\n  const PALETTE = {\n    posMatrix: [\"#ffffff\", \"#00a933\"],\n    posTotals: [\"#ffffff\", \"#81d41a\"],\n    negMatrix: [\"#ffffff\", \"#ff0000\"],\n    negTotals: [\"#ffffff\", \"#800080\"],\n    nonRic: [\"#ffffff\", \"#ff8000\"],\n    network: [\"#d73027\", \"#fee08b\", \"#1a9850\"]\n  };\n  const IS_EXPORT_MODE = Boolean(window.TEAMWEAVE_EXPORT_MODE);\n  const EXPORT_ORDER = Array.isArray(window.TEAMWEAVE_EXPORT_ORDER)\n    ? window.TEAMWEAVE_EXPORT_ORDER.slice()\n    : null;\n  const INLINE_STYLES = \"__INLINE_STYLES__\";\n  const INLINE_APP = \"__INLINE_APP__\";\n\n\n  if (\"serviceWorker\" in navigator && !window.TEAMWEAVE_DISABLE_SW) {\n    window.addEventListener(\"load\", () => {\n      navigator.serviceWorker.register(\"sw.js\").then((registration) => {\n        registration.update().catch(() => undefined);\n        registration.addEventListener(\"updatefound\", () => {\n          const newWorker = registration.installing;\n          if (!newWorker) {\n            return;\n          }\n          newWorker.addEventListener(\"statechange\", () => {\n            if (newWorker.state === \"installed\" && navigator.serviceWorker.controller) {\n              newWorker.postMessage({ type: \"SKIP_WAITING\" });\n            }\n          });\n        });\n      }).catch(() => undefined);\n\n      navigator.serviceWorker.addEventListener(\"controllerchange\", () => {\n        window.location.reload();\n      });\n    });\n  }\n\n  fileInput.addEventListener(\"change\", (event) => {\n    const files = Array.from(event.target.files || []);\n    if (!files.length) {\n      return;\n    }\n    readFiles(files);\n    event.target.value = \"\";\n  });\n\n  datasetSelect.addEventListener(\"change\", () => {\n    setCurrentDataset(datasetSelect.value);\n  });\n\n  deleteDatasetBtn.addEventListener(\"click\", () => {\n    const dataset = getCurrentDataset();\n    if (!dataset) {\n      return;\n    }\n    removeDataset(dataset.id);\n  });\n\n  compareSelectA.addEventListener(\"change\", () => {\n    ensureCompareSelection();\n    const active = getActiveView() || \"responses\";\n    renderView(active);\n    saveDatasetState();\n    updateDatasetManager();\n  });\n\n  compareSelectB.addEventListener(\"change\", () => {\n    ensureCompareSelection();\n    const active = getActiveView() || \"responses\";\n    renderView(active);\n    saveDatasetState();\n    updateDatasetManager();\n  });\n\n  importCodeBtn.addEventListener(\"click\", async () => {\n    const value = importCodeInput.value;\n    if (!value || !value.trim()) {\n      setStatus(\"Incolla un codice prima di importare.\", true);\n      return;\n    }\n    const decoded = await decodeLinkPayload(extractCode(value));\n    let csv = null;\n    if (decoded && decoded.kind === \"bytes\") {\n      csv = unpackCsvFromBinary(decoded.value);\n    } else if (decoded && decoded.kind === \"text\") {\n      csv = decodePackedCsv(decoded.value);\n    }\n    if (!csv) {\n      setStatus(\"Codice non valido.\", true);\n      return;\n    }\n    loadCsv(csv, { name: `Codice ${datasetCounter}`, persist: true });\n    datasetCounter += 1;\n  });\n\n  exportLinkBtn.addEventListener(\"click\", async () => {\n    const dataset = getCurrentDataset();\n    if (!dataset) {\n      setStatus(\"Carica un CSV prima di creare il codice.\", true);\n      return;\n    }\n    let payload = dataset.csv;\n    try {\n      payload = packCsvForLink(dataset.csv);\n    } catch (error) {\n      payload = dataset.csv;\n    }\n    const encoded = await encodeLinkPayload(payload);\n    exportLinkInput.value = encoded;\n    exportBox.classList.remove(\"hidden\");\n  });\n\n  copyLinkBtn.addEventListener(\"click\", async () => {\n    const value = exportLinkInput.value;\n    if (!value) {\n      return;\n    }\n    try {\n      await navigator.clipboard.writeText(value);\n      setStatus(\"Codice copiato negli appunti.\", false);\n    } catch (error) {\n      setStatus(\"Impossibile copiare il codice.\", true);\n    }\n  });\n\n  exportCsvBtn.addEventListener(\"click\", () => {\n    const dataset = getCurrentDataset();\n    if (!dataset) {\n      setStatus(\"Nessun CSV salvato in memoria.\", true);\n      return;\n    }\n    downloadCsv(dataset.csv);\n  });\n\n  exportHtmlBtn.addEventListener(\"click\", async () => {\n    const dataset = getCurrentDataset();\n    if (!dataset) {\n      setStatus(\"Carica un CSV prima di esportare l'HTML.\", true);\n      return;\n    }\n    try {\n      const html = await buildStaticHtml(dataset);\n      downloadHtml(html, dataset.name);\n    } catch (error) {\n      setStatus(\"Impossibile esportare l'HTML.\", true);\n    }\n  });\n\n\n  tabs.forEach((tab) => {\n    tab.addEventListener(\"click\", () => {\n      if (!currentAnalysis) {\n        return;\n      }\n      tabs.forEach((btn) => btn.classList.remove(\"active\"));\n      tab.classList.add(\"active\");\n      const view = tab.dataset.view;\n      saveView(view);\n      renderView(view);\n    });\n  });\n\n  mainTabs.forEach((tab) => {\n    tab.addEventListener(\"click\", () => {\n      mainTabs.forEach((btn) => btn.classList.remove(\"active\"));\n      tab.classList.add(\"active\");\n      const target = tab.dataset.main;\n      const isGenerator = target === \"generator\";\n      if (analysisSection) {\n        analysisSection.classList.toggle(\"hidden\", isGenerator);\n      }\n      if (generatorSection) {\n        generatorSection.classList.toggle(\"hidden\", !isGenerator);\n      }\n    });\n  });\n\n  initTeamweaveFormGenerator();\n  initializeFromStorage();\n\n  function readFile(file) {\n    return new Promise((resolve, reject) => {\n      const reader = new FileReader();\n      reader.onload = () => resolve(reader.result);\n      reader.onerror = () => reject(new Error(\"Impossibile leggere il file.\"));\n      reader.readAsText(file);\n    });\n  }\n\n  async function readFiles(files) {\n    for (const file of files) {\n      try {\n        const content = await readFile(file);\n        loadCsv(content, { name: file.name, persist: true });\n      } catch (error) {\n        setStatus(\"Impossibile leggere il file.\", true);\n      }\n    }\n  }\n\n  function loadCsv(text, options = {}) {\n    const { name = `Dataset ${datasets.length + 1}`, persist = true, replaceAll = false } = options;\n    try {\n      const dataset = createDatasetFromCsv(text, name);\n      if (replaceAll) {\n        datasets.length = 0;\n        visibleDatasetIds.clear();\n        hasVisibleState = false;\n      }\n      datasets.push(dataset);\n      visibleDatasetIds.add(dataset.id);\n      datasetCounter = Math.max(datasetCounter, datasets.length + 1);\n      saveDatasets();\n      setCurrentDataset(dataset.id, { persist });\n    } catch (error) {\n      setStatus(error.message, true);\n    }\n  }\n\n  function createDatasetFromCsv(text, name) {\n    const sanitized = text.replace(/^\\uFEFF/, \"\");\n    const { headers, rows } = parseCsv(sanitized);\n    const analysis = analyzeData(headers, rows);\n    return {\n      id: `ds-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,\n      name,\n      csv: sanitized,\n      analysis\n    };\n  }\n\n  function setStatus(message, isError) {\n    statusEl.textContent = message;\n    statusEl.classList.toggle(\"error\", isError);\n  }\n\n  function saveView(view) {\n    try {\n      localStorage.setItem(VIEW_KEY, view);\n    } catch (error) {\n      return;\n    }\n  }\n\n  function restoreView() {\n    try {\n      const cached = localStorage.getItem(VIEW_KEY);\n      if (!cached) {\n        return null;\n      }\n      const valid = tabs.some((tab) => tab.dataset.view === cached);\n      return valid ? cached : null;\n    } catch (error) {\n      return null;\n    }\n  }\n\n  function saveDatasets() {\n    try {\n      const payload = datasets\n        .filter((dataset) => dataset.name !== \"Ultimo CSV\")\n        .map((dataset) => ({\n          id: dataset.id,\n          name: dataset.name,\n          csv: dataset.csv\n        }));\n      localStorage.setItem(DATASETS_KEY, JSON.stringify(payload));\n    } catch (error) {\n      setStatus(\"Memoria del browser piena: impossibile salvare i dataset.\", true);\n    }\n  }\n\n  function pruneLegacyDatasets() {\n    let removed = false;\n    for (let i = datasets.length - 1; i >= 0; i -= 1) {\n      if (datasets[i].name === \"Ultimo CSV\") {\n        datasets.splice(i, 1);\n        removed = true;\n      }\n    }\n    if (removed) {\n      if (!datasets.some((dataset) => dataset.id === currentDatasetId)) {\n        currentDatasetId = datasets[0]?.id || null;\n      }\n      saveDatasets();\n    }\n    return removed;\n  }\n\n  function saveDatasetState() {\n    try {\n      const state = {\n        currentDatasetId,\n        compareA: compareSelectA?.value || null,\n        compareB: compareSelectB?.value || null,\n        visibleIds: Array.from(visibleDatasetIds)\n      };\n      localStorage.setItem(DATASET_STATE_KEY, JSON.stringify(state));\n    } catch (error) {\n      return;\n    }\n  }\n\n  function restoreDatasets() {\n    try {\n      const cached = localStorage.getItem(DATASETS_KEY);\n      if (!cached) {\n        return false;\n      }\n      const entries = JSON.parse(cached);\n      if (!Array.isArray(entries) || entries.length === 0) {\n        return false;\n      }\n      datasets.length = 0;\n      let removedLegacy = false;\n      entries.forEach((entry) => {\n        if (!entry || !entry.csv) {\n          return;\n        }\n        if (entry.name === \"Ultimo CSV\") {\n          removedLegacy = true;\n          return;\n        }\n        try {\n          const dataset = createDatasetFromCsv(entry.csv, entry.name || `Dataset ${datasets.length + 1}`);\n          dataset.id = entry.id || dataset.id;\n          datasets.push(dataset);\n        } catch (error) {\n          return;\n        }\n      });\n      pruneLegacyDatasets();\n      if (removedLegacy) {\n        saveDatasets();\n      }\n      if (!datasets.length) {\n        return false;\n      }\n      isRestoringDatasets = true;\n      datasetCounter = Math.max(datasetCounter, datasets.length + 1);\n      updateDatasetControls();\n      const stateRaw = localStorage.getItem(DATASET_STATE_KEY);\n      let nextId = datasets[0].id;\n      if (stateRaw) {\n        try {\n          const state = JSON.parse(stateRaw);\n          if (Array.isArray(state?.visibleIds)) {\n            hasVisibleState = true;\n            visibleDatasetIds = new Set(state.visibleIds);\n          }\n          if (state?.currentDatasetId && datasets.some((d) => d.id === state.currentDatasetId)) {\n            nextId = state.currentDatasetId;\n          }\n          if (state?.compareA) {\n            compareSelectA.value = state.compareA;\n          }\n          if (state?.compareB) {\n            compareSelectB.value = state.compareB;\n          }\n          ensureCompareSelection();\n        } catch (error) {\n          return false;\n        }\n      }\n      setCurrentDataset(nextId, { persist: false });\n      isRestoringDatasets = false;\n      saveDatasetState();\n      return true;\n    } catch (error) {\n      isRestoringDatasets = false;\n      return false;\n    }\n  }\n\n  async function restoreFromLink() {\n    const hash = location.hash.replace(/^#/, \"\");\n    if (!hash.startsWith(`${LINK_PARAM}=`)) {\n      return false;\n    }\n    const encoded = hash.slice(LINK_PARAM.length + 1);\n    const decoded = await decodeLinkPayload(encoded);\n    let csv = null;\n    if (decoded && decoded.kind === \"bytes\") {\n      csv = unpackCsvFromBinary(decoded.value);\n    } else if (decoded && decoded.kind === \"text\") {\n      csv = decodePackedCsv(decoded.value);\n    }\n    if (csv) {\n      loadCsv(csv, { name: \"Link dati\", persist: true, replaceAll: true });\n      return true;\n    }\n    return false;\n  }\n\n  function parseCsv(text) {\n    const rows = [];\n    let current = [];\n    let value = \"\";\n    let inQuotes = false;\n\n    for (let i = 0; i < text.length; i += 1) {\n      const char = text[i];\n      const next = text[i + 1];\n      if (char === '\"') {\n        if (inQuotes && next === '\"') {\n          value += '\"';\n          i += 1;\n        } else {\n          inQuotes = !inQuotes;\n        }\n      } else if (char === \",\" && !inQuotes) {\n        current.push(value);\n        value = \"\";\n      } else if ((char === \"\\n\" || char === \"\\r\") && !inQuotes) {\n        if (char === \"\\r\" && next === \"\\n\") {\n          i += 1;\n        }\n        current.push(value);\n        value = \"\";\n        if (current.some((cell) => cell.trim() !== \"\")) {\n          rows.push(current);\n        }\n        current = [];\n      } else {\n        value += char;\n      }\n    }\n    if (value.length || current.length) {\n      current.push(value);\n      if (current.some((cell) => cell.trim() !== \"\")) {\n        rows.push(current);\n      }\n    }\n\n    if (rows.length < 2) {\n      throw new Error(\"CSV non valido o vuoto.\");\n    }\n\n    const headers = rows[0].map((cell) => cell.trim());\n    const dataRows = rows.slice(1);\n\n    return { headers, rows: dataRows };\n  }\n\n  function analyzeData(headers, rows) {\n    const nameIndex = findNameColumn(headers);\n    const firstQuestionIndex = nameIndex + 1;\n    if (firstQuestionIndex >= headers.length) {\n      throw new Error(\"CSV senza colonne di domande riconoscibili.\");\n    }\n\n    const questionHeaders = headers.slice(firstQuestionIndex);\n    const questionMeta = questionHeaders.map((question, idx) => {\n      const tagged = extractCategoryTag(question);\n      const normalized = normalizeQuestion(tagged.question);\n      const category = tagged.category || CATEGORY_MAP_NORMALIZED.get(normalized) || null;\n      return {\n        question: tagged.question,\n        category,\n        positive: idx % 2 === 0\n      };\n    });\n    const posQuestionLabels = [];\n    const negQuestionLabels = [];\n    const posIndexByMeta = [];\n    const negIndexByMeta = [];\n    questionMeta.forEach((meta, idx) => {\n      if (meta.positive) {\n        posIndexByMeta[idx] = posQuestionLabels.length;\n        posQuestionLabels.push(meta.question);\n        return;\n      }\n      negIndexByMeta[idx] = negQuestionLabels.length;\n      negQuestionLabels.push(meta.question);\n    });\n\n    const names = [];\n    const nameIndexMap = new Map();\n    const normalizedMap = new Map();\n    const rowsNormalized = rows.map((row) => row.map((cell) => cell.trim()));\n\n    rowsNormalized.forEach((row) => {\n      const rawName = row[nameIndex] || \"\";\n      const normalized = normalizeName(rawName);\n      if (!normalized) {\n        return;\n      }\n      if (!nameIndexMap.has(normalized)) {\n        nameIndexMap.set(normalized, names.length);\n        normalizedMap.set(normalized, rawName.trim());\n        names.push(rawName.trim());\n      }\n    });\n\n    if (names.length === 0) {\n      throw new Error(\"Nessun nome trovato nella colonna 'Seleziona il tuo nome'.\");\n    }\n\n    const size = names.length;\n    const posMatrix = createMatrix(size, 0);\n    const negMatrix = createMatrix(size, 0);\n    const receivedPosByCat = initCategoryMap(names, CATEGORIES);\n    const receivedNegByCat = initCategoryMap(names, CATEGORIES);\n    const posReceivedByQuestion = createRectMatrix(size, posQuestionLabels.length, 0);\n    const negReceivedByQuestion = createRectMatrix(size, negQuestionLabels.length, 0);\n    const unknownSelections = new Set();\n    const missingCategories = new Set();\n\n    rowsNormalized.forEach((row) => {\n      const chooserRaw = row[nameIndex] || \"\";\n      const chooserKey = normalizeName(chooserRaw);\n      if (!chooserKey || !nameIndexMap.has(chooserKey)) {\n        return;\n      }\n      const chooserIndex = nameIndexMap.get(chooserKey);\n\n      questionMeta.forEach((meta, idx) => {\n        const cell = row[firstQuestionIndex + idx] || \"\";\n        const picks = splitNames(cell);\n        const seen = new Set();\n        if (!meta.category) {\n          if (cell.trim() !== \"\") {\n            missingCategories.add(meta.question);\n          }\n          return;\n        }\n\n        picks.forEach((pick) => {\n          const pickKey = normalizeName(pick);\n          if (!pickKey || seen.has(pickKey)) {\n            return;\n          }\n          seen.add(pickKey);\n          if (!nameIndexMap.has(pickKey)) {\n            unknownSelections.add(pick.trim());\n            return;\n          }\n          const pickIndex = nameIndexMap.get(pickKey);\n\n          if (meta.positive) {\n            if (chooserIndex !== pickIndex) {\n              posMatrix[chooserIndex][pickIndex] += 1;\n            }\n            receivedPosByCat.get(pickKey)[meta.category] += 1;\n            const qIndex = posIndexByMeta[idx];\n            if (qIndex != null) {\n              posReceivedByQuestion[pickIndex][qIndex] += 1;\n            }\n          } else {\n            if (chooserIndex !== pickIndex) {\n              negMatrix[chooserIndex][pickIndex] += 1;\n            }\n            receivedNegByCat.get(pickKey)[meta.category] += 1;\n            const qIndex = negIndexByMeta[idx];\n            if (qIndex != null) {\n              negReceivedByQuestion[pickIndex][qIndex] += 1;\n            }\n          }\n        });\n      });\n    });\n\n    const posRowSum = sumRows(posMatrix);\n    const posColSum = sumColumns(posMatrix);\n    const negRowSum = sumRows(negMatrix);\n    const negColSum = sumColumns(negMatrix);\n\n    const reciprociPos = buildReciprociMatrix(posMatrix);\n    const reciprociNeg = buildReciprociMatrix(negMatrix);\n\n    const forzaLegami = buildForzaMatrix(posMatrix);\n    const forzaAntagonismo = buildForzaMatrix(negMatrix);\n\n    const nonRicambiatePos = buildNonRicambiateMatrix(posMatrix, reciprociPos);\n\n    const reciprociPosRowSum = sumRows(reciprociPos);\n    const reciprociNegRowSum = sumRows(reciprociNeg);\n\n    const forzaLegamiRowSum = sumRows(forzaLegami);\n    const forzaAntagonismoRowSum = sumRows(forzaAntagonismo);\n    const nonRicambiateRowSum = sumRows(nonRicambiatePos);\n    const nonRicambiateColSum = sumColumns(nonRicambiatePos);\n\n    const summaryRows = names.map((name, index) => {\n      const key = normalizeName(name);\n      const posCats = receivedPosByCat.get(key);\n      const negCats = receivedNegByCat.get(key);\n      return {\n        name,\n        positiveReceived: {\n          Tecniche: posCats.Tecniche,\n          Attitudinali: posCats.Attitudinali,\n          Sociali: posCats.Sociali,\n          Totali: posColSum[index]\n        },\n        negativeReceived: {\n          Tecniche: negCats.Tecniche,\n          Attitudinali: negCats.Attitudinali,\n          Sociali: negCats.Sociali,\n          Totali: negColSum[index]\n        },\n        positiveGiven: posRowSum[index],\n        negativeGiven: reciprociNegRowSum[index],\n        reciprociPos: reciprociPosRowSum[index],\n        reciprociNeg: reciprociNegRowSum[index],\n        forzaLegami: forzaLegamiRowSum[index],\n        forzaAntagonismo: forzaAntagonismoRowSum[index],\n        nonRicambiateDate: nonRicambiateRowSum[index],\n        nonRicambiateRicevute: nonRicambiateColSum[index]\n      };\n    });\n\n    const averages = computeAverages(summaryRows);\n    const classifications = summaryRows.map((row) => {\n      const inflPos = classifyInfluence(row.positiveReceived.Totali, averages.positiveReceivedTotal);\n      const inflNeg = classifyInfluence(row.negativeReceived.Totali, averages.negativeReceivedTotal);\n      const equilibrio = classifyBalance(\n        row.nonRicambiateDate,\n        row.nonRicambiateRicevute,\n        averages.nonRicambiateDate,\n        averages.nonRicambiateRicevute\n      );\n      return {\n        name: row.name,\n        inflPos,\n        inflNeg,\n        equilibrio,\n        label: synthesizeLabel(inflPos, inflNeg, equilibrio)\n      };\n    });\n\n    return {\n      headers,\n      rows: rowsNormalized,\n      names,\n      matrices: {\n        posMatrix,\n        negMatrix,\n        reciprociPos,\n        reciprociNeg,\n        forzaLegami,\n        forzaAntagonismo,\n        nonRicambiatePos\n      },\n      totals: {\n        posRowSum,\n        posColSum,\n        negRowSum,\n        negColSum,\n        reciprociPosRowSum,\n        reciprociNegRowSum,\n        forzaLegamiRowSum,\n        forzaAntagonismoRowSum,\n        nonRicambiateRowSum,\n        nonRicambiateColSum\n      },\n      questionMeta,\n      summaryRows,\n      averages,\n      classifications,\n      questionStats: {\n        posLabels: posQuestionLabels,\n        negLabels: negQuestionLabels,\n        posReceivedByQuestion,\n        negReceivedByQuestion\n      },\n      warnings: {\n        unknownSelections: Array.from(unknownSelections).sort(),\n        missingCategories: Array.from(missingCategories).sort()\n      }\n    };\n  }\n\n  function findNameColumn(headers) {\n    const idx = headers.findIndex((header) =>\n      header.toLowerCase().includes(\"seleziona il tuo nome\")\n    );\n    if (idx !== -1) {\n      return idx;\n    }\n    return 1;\n  }\n\n  function normalizeName(name) {\n    return name.trim().replace(/\\s+/g, \" \").toLowerCase();\n  }\n\n  function extractCategoryTag(question) {\n    const raw = String(question || \"\");\n    const match = raw.match(/^\\s*\\[(Tecniche|Attitudinali|Sociali)\\]\\s*/i);\n    if (!match) {\n      return { question: raw, category: null };\n    }\n    const cleaned = raw.slice(match[0].length).trim();\n    const lowered = match[1].toLowerCase();\n    const category = lowered.startsWith(\"tec\")\n      ? \"Tecniche\"\n      : lowered.startsWith(\"att\")\n        ? \"Attitudinali\"\n        : \"Sociali\";\n    return { question: cleaned || raw, category };\n  }\n\n  function normalizeQuestion(question) {\n    return question\n      .trim()\n      .replace(/[\u2019\u2018]/g, \"'\")\n      .replace(/\\s+/g, \" \")\n      .toLowerCase();\n  }\n\n  function escapeQuotes(text) {\n    return String(text || \"\")\n      .replace(/\\\\/g, \"\\\\\\\\\")\n      .replace(/\"/g, '\\\\\"')\n      .replace(/\\r?\\n/g, \"\\\\n\");\n  }\n\n  function splitNames(value) {\n    if (!value) {\n      return [];\n    }\n    return value\n      .split(\",\")\n      .map((part) => part.trim())\n      .filter(Boolean);\n  }\n\n  function createMatrix(size, fillValue) {\n    return Array.from({ length: size }, () => Array(size).fill(fillValue));\n  }\n\n  function createRectMatrix(rows, cols, fillValue) {\n    return Array.from({ length: rows }, () => Array(cols).fill(fillValue));\n  }\n\n  function initCategoryMap(names, categories) {\n    const map = new Map();\n    names.forEach((name) => {\n      const entry = {};\n      categories.forEach((cat) => {\n        entry[cat] = 0;\n      });\n      map.set(normalizeName(name), entry);\n    });\n    return map;\n  }\n\n  function sumRows(matrix) {\n    return matrix.map((row) => row.reduce((acc, value) => acc + value, 0));\n  }\n\n  function sumColumns(matrix) {\n    const size = matrix.length;\n    const sums = Array(size).fill(0);\n    matrix.forEach((row) => {\n      row.forEach((value, index) => {\n        sums[index] += value;\n      });\n    });\n    return sums;\n  }\n\n  function buildReciprociMatrix(matrix) {\n    const size = matrix.length;\n    const result = createMatrix(size, 0);\n    for (let i = 0; i < size; i += 1) {\n      for (let j = 0; j < size; j += 1) {\n        if (i === j) {\n          result[i][j] = 0;\n        } else if (matrix[i][j] > 0 && matrix[j][i] > 0) {\n          result[i][j] = 1;\n        }\n      }\n    }\n    return result;\n  }\n\n  function buildForzaMatrix(matrix) {\n    const size = matrix.length;\n    const result = createMatrix(size, 0);\n    for (let i = 0; i < size; i += 1) {\n      for (let j = 0; j < size; j += 1) {\n        if (i === j) {\n          result[i][j] = 0;\n        } else {\n          result[i][j] = Math.min(matrix[i][j], matrix[j][i]);\n        }\n      }\n    }\n    return result;\n  }\n\n  function buildNonRicambiateMatrix(matrix, reciproci) {\n    const size = matrix.length;\n    const result = createMatrix(size, 0);\n    for (let i = 0; i < size; i += 1) {\n      for (let j = 0; j < size; j += 1) {\n        if (matrix[i][j] > 0 && reciproci[i][j] === 0) {\n          result[i][j] = matrix[i][j];\n        }\n      }\n    }\n    return result;\n  }\n\n  function computeAverages(rows) {\n    const count = rows.length || 1;\n    const totals = {\n      positiveReceivedTecniche: 0,\n      positiveReceivedAttitudinali: 0,\n      positiveReceivedSociali: 0,\n      positiveReceivedTotal: 0,\n      negativeReceivedTecniche: 0,\n      negativeReceivedAttitudinali: 0,\n      negativeReceivedSociali: 0,\n      negativeReceivedTotal: 0,\n      positiveGiven: 0,\n      negativeGiven: 0,\n      reciprociPos: 0,\n      reciprociNeg: 0,\n      forzaLegami: 0,\n      forzaAntagonismo: 0,\n      nonRicambiateDate: 0,\n      nonRicambiateRicevute: 0\n    };\n\n    rows.forEach((row) => {\n      totals.positiveReceivedTecniche += row.positiveReceived.Tecniche;\n      totals.positiveReceivedAttitudinali += row.positiveReceived.Attitudinali;\n      totals.positiveReceivedSociali += row.positiveReceived.Sociali;\n      totals.positiveReceivedTotal += row.positiveReceived.Totali;\n      totals.negativeReceivedTecniche += row.negativeReceived.Tecniche;\n      totals.negativeReceivedAttitudinali += row.negativeReceived.Attitudinali;\n      totals.negativeReceivedSociali += row.negativeReceived.Sociali;\n      totals.negativeReceivedTotal += row.negativeReceived.Totali;\n      totals.positiveGiven += row.positiveGiven;\n      totals.negativeGiven += row.negativeGiven;\n      totals.reciprociPos += row.reciprociPos;\n      totals.reciprociNeg += row.reciprociNeg;\n      totals.forzaLegami += row.forzaLegami;\n      totals.forzaAntagonismo += row.forzaAntagonismo;\n      totals.nonRicambiateDate += row.nonRicambiateDate;\n      totals.nonRicambiateRicevute += row.nonRicambiateRicevute;\n    });\n\n    const averages = {};\n    Object.keys(totals).forEach((key) => {\n      averages[key] = totals[key] / count;\n    });\n    return averages;\n  }\n\n  function classifyInfluence(value, average) {\n    if (!Number.isFinite(value)) {\n      return \"\";\n    }\n    if (value >= average) {\n      return \"Alta\";\n    }\n    if (value >= average / 2) {\n      return \"Media\";\n    }\n    return \"Bassa\";\n  }\n\n  function classifyBalance(nonRicDate, nonRicRecv, avgDate, avgRecv) {\n    if (!Number.isFinite(nonRicDate) || !Number.isFinite(nonRicRecv)) {\n      return \"\";\n    }\n    if (nonRicRecv - nonRicDate >= avgDate) {\n      return \"Selettiva\";\n    }\n    if (nonRicDate - nonRicRecv >= avgRecv) {\n      return \"Ignorata\";\n    }\n    return \"Bilanciata\";\n  }\n\n  function synthesizeLabel(inflPos, inflNeg, equilibrio) {\n    if (inflPos === \"Alta\") {\n      if (inflNeg === \"Alta\") {\n        if (equilibrio === \"Selettiva\") {\n          return \"Leader controversa\";\n        }\n        if (equilibrio === \"Bilanciata\") {\n          return \"Figura di riferimento\";\n        }\n        return \"Figura controversa\";\n      }\n      if (inflNeg === \"Media\") {\n        if (equilibrio === \"Selettiva\") {\n          return \"Leader silenziosa\";\n        }\n        if (equilibrio === \"Bilanciata\") {\n          return \"Benvoluta\";\n        }\n        return \"Stimata\";\n      }\n      if (equilibrio === \"Selettiva\") {\n        return \"Leader positiva\";\n      }\n      if (equilibrio === \"Bilanciata\") {\n        return \"Stimata\";\n      }\n      return \"Apprezzata\";\n    }\n\n    if (inflPos === \"Media\") {\n      if (inflNeg === \"Alta\") {\n        if (equilibrio === \"Ignorata\") {\n          return \"Invisibile e controversa\";\n        }\n        if (equilibrio === \"Bilanciata\") {\n          return \"Controversa\";\n        }\n        return \"Marginale\";\n      }\n      if (inflNeg === \"Media\") {\n        if (equilibrio === \"Selettiva\") {\n          return \"Bilanciata\";\n        }\n        if (equilibrio === \"Bilanciata\") {\n          return \"Neutrale\";\n        }\n        return \"Presenza discreta\";\n      }\n      if (equilibrio === \"Selettiva\") {\n        return \"Apprezzata\";\n      }\n      if (equilibrio === \"Bilanciata\") {\n        return \"Stimata\";\n      }\n      return \"Presenza discreta\";\n    }\n\n    if (inflNeg === \"Alta\") {\n      if (equilibrio === \"Ignorata\") {\n        return \"Esclusa\";\n      }\n      if (equilibrio === \"Bilanciata\") {\n        return \"Altruista\";\n      }\n      return \"Dispersa\";\n    }\n    if (inflNeg === \"Media\") {\n      if (equilibrio === \"Ignorata\") {\n        return \"Trascurata\";\n      }\n      return \"Non considerata\";\n    }\n    return \"Invisibile\";\n  }\n\n  function updateSummary(analysis) {\n    document.getElementById(\"summary-athletes\").textContent = analysis.names.length;\n    document.getElementById(\"summary-reciproci-pos\").textContent = formatPercent(\n      computeReciprocityIndex(analysis.matrices.reciprociPos)\n    );\n    document.getElementById(\"summary-reciproci-neg\").textContent = formatPercent(\n      computeReciprocityIndex(analysis.matrices.reciprociNeg)\n    );\n\n    const warnings = [];\n    if (analysis.warnings.unknownSelections.length) {\n      warnings.push(`Nomi fuori lista: ${analysis.warnings.unknownSelections.join(\", \")}`);\n    }\n    if (analysis.warnings.missingCategories.length) {\n      warnings.push(\"Alcune domande non sono mappate alle categorie.\");\n    }\n    if (warnings.length) {\n      setStatus(warnings.join(\" \"), true);\n    }\n  }\n\n  function getCurrentDataset() {\n    return datasets.find((dataset) => dataset.id === currentDatasetId) || null;\n  }\n\n  function removeDataset(id) {\n    const index = datasets.findIndex((dataset) => dataset.id === id);\n    if (index === -1) {\n      return;\n    }\n    datasets.splice(index, 1);\n    visibleDatasetIds.delete(id);\n    saveDatasets();\n    if (!datasets.length) {\n      currentDatasetId = null;\n      currentAnalysis = null;\n      summaryEl.classList.add(\"hidden\");\n      viewsEl.classList.add(\"hidden\");\n      viewContainer.innerHTML = \"\";\n      datasetBox.classList.add(\"hidden\");\n      saveDatasetState();\n      setStatus(\"Dataset rimosso.\", false);\n      return;\n    }\n    const nextId = datasets[0].id;\n    setCurrentDataset(nextId, { persist: false });\n    setStatus(\"Dataset rimosso.\", false);\n  }\n\n  function ensureVisibleDatasetIds() {\n    if (IS_EXPORT_MODE) {\n      return false;\n    }\n    const validIds = new Set(datasets.map((dataset) => dataset.id));\n    let changed = false;\n    visibleDatasetIds.forEach((id) => {\n      if (!validIds.has(id)) {\n        visibleDatasetIds.delete(id);\n        changed = true;\n      }\n    });\n    if (!visibleDatasetIds.size && !hasVisibleState) {\n      datasets.forEach((dataset) => visibleDatasetIds.add(dataset.id));\n      changed = true;\n    }\n    return changed;\n  }\n\n  function getVisibleDatasetIds() {\n    if (IS_EXPORT_MODE) {\n      return new Set(datasets.map((dataset) => dataset.id));\n    }\n    ensureVisibleDatasetIds();\n    return new Set(visibleDatasetIds);\n  }\n\n  function getVisibleDatasets() {\n    const visibleIds = getVisibleDatasetIds();\n    return datasets.filter((dataset) => visibleIds.has(dataset.id));\n  }\n\n  function updateDatasetManager() {\n    if (!datasetManager) {\n      return;\n    }\n    datasetManager.textContent = \"\";\n    if (!datasets.length) {\n      return;\n    }\n    const visibleIds = getVisibleDatasetIds();\n    const list = document.createElement(\"div\");\n    list.className = \"dataset-manager-grid\";\n    datasets.forEach((dataset) => {\n      const row = document.createElement(\"div\");\n      row.className = \"dataset-manager-item\";\n\n      const toggleLabel = document.createElement(\"label\");\n      toggleLabel.className = \"dataset-manager-toggle\";\n      const checkbox = document.createElement(\"input\");\n      checkbox.type = \"checkbox\";\n      checkbox.checked = visibleIds.has(dataset.id);\n      checkbox.addEventListener(\"change\", () => {\n        if (checkbox.checked) {\n          visibleDatasetIds.add(dataset.id);\n        } else {\n          visibleDatasetIds.delete(dataset.id);\n        }\n        hasVisibleState = true;\n        const active = getActiveView() || \"responses\";\n        updateDatasetControls();\n        renderView(active);\n      });\n      const name = document.createElement(\"span\");\n      name.textContent = dataset.name;\n      toggleLabel.appendChild(checkbox);\n      toggleLabel.appendChild(name);\n\n      const orderGroup = document.createElement(\"div\");\n      orderGroup.className = \"dataset-manager-order\";\n      const radioA = document.createElement(\"input\");\n      radioA.type = \"radio\";\n      radioA.name = \"dataset-order-a\";\n      radioA.value = dataset.id;\n      radioA.checked = compareSelectA.value === dataset.id;\n      radioA.addEventListener(\"change\", () => {\n        if (!visibleDatasetIds.has(dataset.id)) {\n          visibleDatasetIds.add(dataset.id);\n          checkbox.checked = true;\n          hasVisibleState = true;\n        }\n        compareSelectA.value = dataset.id;\n        ensureCompareSelection();\n        const active = getActiveView() || \"responses\";\n        renderView(active);\n        saveDatasetState();\n        updateDatasetManager();\n      });\n      const labelA = document.createElement(\"label\");\n      labelA.appendChild(radioA);\n      labelA.appendChild(document.createTextNode(\"A\"));\n\n      const radioB = document.createElement(\"input\");\n      radioB.type = \"radio\";\n      radioB.name = \"dataset-order-b\";\n      radioB.value = dataset.id;\n      radioB.checked = compareSelectB.value === dataset.id;\n      radioB.addEventListener(\"change\", () => {\n        if (!visibleDatasetIds.has(dataset.id)) {\n          visibleDatasetIds.add(dataset.id);\n          checkbox.checked = true;\n          hasVisibleState = true;\n        }\n        compareSelectB.value = dataset.id;\n        ensureCompareSelection();\n        const active = getActiveView() || \"responses\";\n        renderView(active);\n        saveDatasetState();\n        updateDatasetManager();\n      });\n      const labelB = document.createElement(\"label\");\n      labelB.appendChild(radioB);\n      labelB.appendChild(document.createTextNode(\"B\"));\n\n      orderGroup.appendChild(labelA);\n      orderGroup.appendChild(labelB);\n\n      row.appendChild(toggleLabel);\n      row.appendChild(orderGroup);\n      list.appendChild(row);\n    });\n    datasetManager.appendChild(list);\n  }\n\n  function setCurrentDataset(id, options = {}) {\n    const { persist = true } = options;\n    const visibleIds = getVisibleDatasetIds();\n    let targetId = id;\n    if (visibleIds.size && !visibleIds.has(targetId)) {\n      const fallback = datasets.find((entry) => visibleIds.has(entry.id));\n      targetId = fallback ? fallback.id : id;\n    }\n    const dataset = datasets.find((entry) => entry.id === targetId);\n    if (!dataset) {\n      return;\n    }\n    currentDatasetId = dataset.id;\n    currentAnalysis = dataset.analysis;\n    updateDatasetControls();\n    if (!isRestoringDatasets) {\n      saveDatasetState();\n    }\n    updateSummary(dataset.analysis);\n    summaryEl.classList.remove(\"hidden\");\n    viewsEl.classList.remove(\"hidden\");\n    const selectedView = getActiveView() || restoreView() || \"responses\";\n    renderView(selectedView);\n    tabs.forEach((btn) => btn.classList.remove(\"active\"));\n    const activeTab = tabs.find((btn) => btn.dataset.view === selectedView) || tabs[0];\n    activeTab.classList.add(\"active\");\n    setStatus(\n      `Caricato: ${dataset.analysis.rows.length} risposte, ${dataset.analysis.names.length} atlete. (${dataset.name})`,\n      false\n    );\n    applyExportModeUI();\n  }\n\n  function updateDatasetControls() {\n    if (!datasets.length) {\n      datasetBox.classList.add(\"hidden\");\n      return;\n    }\n    datasetBox.classList.remove(\"hidden\");\n    const visibleIds = getVisibleDatasetIds();\n    const visibleDatasets = getVisibleDatasets();\n    updateDatasetManager();\n    if (!visibleDatasets.length) {\n      currentDatasetId = null;\n      currentAnalysis = null;\n      summaryEl.classList.add(\"hidden\");\n      viewsEl.classList.add(\"hidden\");\n      viewContainer.classList.add(\"hidden\");\n      compareRow.classList.add(\"hidden\");\n      compareNote.classList.add(\"hidden\");\n      compareSelectA.value = \"\";\n      compareSelectB.value = \"\";\n      compareSelectA.selectedIndex = -1;\n      compareSelectB.selectedIndex = -1;\n      if (!isRestoringDatasets) {\n        saveDatasetState();\n      }\n      return;\n    }\n    if (currentDatasetId && !visibleIds.has(currentDatasetId)) {\n      setCurrentDataset(visibleDatasets[0].id, { persist: false });\n      return;\n    }\n    viewContainer.classList.remove(\"hidden\");\n    populateDatasetSelect(datasetSelect, currentDatasetId, visibleDatasets);\n    populateDatasetSelect(compareSelectA, compareSelectA.value, visibleDatasets);\n    populateDatasetSelect(compareSelectB, compareSelectB.value, visibleDatasets);\n    ensureCompareSelection();\n    const showCompare = visibleDatasets.length >= 2;\n    compareRow.classList.toggle(\"hidden\", !showCompare);\n    compareNote.classList.toggle(\"hidden\", !showCompare);\n    compareSelectA.disabled = !showCompare;\n    compareSelectB.disabled = !showCompare;\n    if (!isRestoringDatasets) {\n      saveDatasetState();\n    }\n    applyExportModeUI();\n  }\n\n  function applyExportModeUI() {\n    if (!IS_EXPORT_MODE) {\n      return;\n    }\n    const panel = document.querySelector(\".upload-panel\");\n    if (!panel) {\n      return;\n    }\n    const title = panel.querySelector(\"h2\");\n    if (title) {\n      title.textContent = \"Dataset caricati\";\n    }\n    const description = panel.querySelector(\"p\");\n    if (description) {\n      description.textContent = \"Seleziona uno o due dataset da visualizzare.\";\n    }\n    if (analysisSection) {\n      analysisSection.classList.remove(\"hidden\");\n    }\n    if (generatorSection) {\n      generatorSection.classList.add(\"hidden\");\n    }\n    if (mainTabs.length) {\n      mainTabs.forEach((btn) => {\n        btn.classList.toggle(\"active\", btn.dataset.main === \"analysis\");\n        btn.classList.add(\"hidden\");\n      });\n    }\n    panel.querySelectorAll(\n      \"button, .export-box, .import-box, .tw-form-box, .upload-row, .dataset-row, #compare-row, #compare-note, #status, #file-input, #import-code, #export-link\"\n    ).forEach((el) => {\n      el.classList.add(\"hidden\");\n    });\n    let picker = panel.querySelector(\".dataset-picker\");\n    if (!picker) {\n      picker = document.createElement(\"div\");\n      picker.className = \"dataset-picker\";\n      panel.appendChild(picker);\n    }\n    picker.innerHTML = \"\";\n    if (!datasets.length) {\n      const empty = document.createElement(\"p\");\n      empty.className = \"note\";\n      empty.textContent = \"Nessun dataset.\";\n      picker.appendChild(empty);\n      return;\n    }\n    const selectedIds = new Set(getExportSelectedIds());\n    datasets.forEach((dataset) => {\n      const label = document.createElement(\"label\");\n      label.className = \"dataset-option\";\n      const checkbox = document.createElement(\"input\");\n      checkbox.type = \"checkbox\";\n      checkbox.value = dataset.id;\n      checkbox.checked = selectedIds.has(dataset.id);\n      checkbox.addEventListener(\"change\", () => {\n        handleExportCheckboxChange(picker, checkbox);\n      });\n      const name = document.createElement(\"span\");\n      name.textContent = dataset.name;\n      label.appendChild(checkbox);\n      label.appendChild(name);\n      picker.appendChild(label);\n    });\n    if (selectedIds.size) {\n      summaryEl.classList.remove(\"hidden\");\n      viewsEl.classList.remove(\"hidden\");\n      viewContainer.classList.remove(\"hidden\");\n    }\n  }\n\n  function getExportSelectedIds() {\n    const ids = [];\n    if (compareSelectA?.value) {\n      ids.push(compareSelectA.value);\n    }\n    if (compareSelectB?.value && compareSelectB.value !== compareSelectA.value) {\n      ids.push(compareSelectB.value);\n    }\n    if (!ids.length && currentDatasetId) {\n      ids.push(currentDatasetId);\n    }\n    return ids;\n  }\n\n  function handleExportCheckboxChange(container, changed) {\n    const inputs = Array.from(container.querySelectorAll(\"input[type=\\\"checkbox\\\"]\"));\n    const selected = inputs.filter((input) => input.checked).map((input) => input.value);\n    applyExportSelection(selected);\n  }\n\n  function orderByExport(ids) {\n    if (!IS_EXPORT_MODE || !EXPORT_ORDER || !EXPORT_ORDER.length) {\n      return ids.slice();\n    }\n    const orderMap = new Map(EXPORT_ORDER.map((id, index) => [id, index]));\n    return ids\n      .map((id, index) => ({\n        id,\n        index,\n        order: orderMap.has(id) ? orderMap.get(id) : Number.MAX_SAFE_INTEGER\n      }))\n      .sort((a, b) => (a.order === b.order ? a.index - b.index : a.order - b.order))\n      .map((entry) => entry.id);\n  }\n\n  function applyExportSelection(ids) {\n    const unique = Array.from(new Set(ids));\n    if (!unique.length) {\n      currentDatasetId = null;\n      currentAnalysis = null;\n      compareSelectA.value = \"\";\n      compareSelectB.value = \"\";\n      compareSelectA.selectedIndex = -1;\n      compareSelectB.selectedIndex = -1;\n      summaryEl.classList.add(\"hidden\");\n      viewsEl.classList.add(\"hidden\");\n      viewContainer.classList.add(\"hidden\");\n      return;\n    }\n    summaryEl.classList.remove(\"hidden\");\n    viewsEl.classList.remove(\"hidden\");\n    viewContainer.classList.remove(\"hidden\");\n    if (unique.length === 1) {\n      compareSelectA.value = unique[0];\n      compareSelectB.value = unique[0];\n      setCurrentDataset(unique[0], { persist: false });\n      return;\n    }\n    const ordered = orderByExport(unique);\n    compareSelectA.value = ordered[0];\n    compareSelectB.value = ordered[1];\n    setCurrentDataset(ordered[0], { persist: false });\n  }\n\n  function populateDatasetSelect(select, selectedId, source = datasets) {\n    select.textContent = \"\";\n    source.forEach((dataset) => {\n      const option = document.createElement(\"option\");\n      option.value = dataset.id;\n      option.textContent = dataset.name;\n      select.appendChild(option);\n    });\n    const targetId = source.some((dataset) => dataset.id === selectedId)\n      ? selectedId\n      : source[0]?.id;\n    if (targetId) {\n      select.value = targetId;\n    }\n  }\n\n  function ensureCompareSelection() {\n    if (IS_EXPORT_MODE) {\n      return;\n    }\n    const visibleDatasets = getVisibleDatasets();\n    if (visibleDatasets.length < 2) {\n      if (visibleDatasets.length === 1) {\n        compareSelectA.value = visibleDatasets[0].id;\n        compareSelectB.value = visibleDatasets[0].id;\n      }\n      return;\n    }\n    const visibleIds = visibleDatasets.map((dataset) => dataset.id);\n    if (!visibleIds.includes(compareSelectA.value)) {\n      compareSelectA.value = visibleIds[0];\n    }\n    if (!visibleIds.includes(compareSelectB.value)) {\n      compareSelectB.value = visibleIds[1] || visibleIds[0];\n    }\n    if (compareSelectA.value === compareSelectB.value) {\n      const alternative = visibleIds.find((id) => id !== compareSelectA.value);\n      if (alternative) {\n        compareSelectB.value = alternative;\n      }\n    }\n  }\n\n  function getActiveView() {\n    const active = tabs.find((tab) => tab.classList.contains(\"active\"));\n    return active ? active.dataset.view : null;\n  }\n\n  function renderView(view) {\n    viewContainer.innerHTML = \"\";\n    const compareDatasets = getCompareDatasets();\n    if (!currentAnalysis && !compareDatasets) {\n      return;\n    }\n\n    if (view === \"confronto\") {\n      viewContainer.classList.remove(\"dual-view\");\n      viewContainer.appendChild(renderCompareSection());\n      return;\n    }\n\n    if (compareDatasets) {\n      renderDualView(view, compareDatasets);\n      return;\n    }\n\n    viewContainer.classList.remove(\"dual-view\");\n    if (!renderers.has(view)) {\n      renderers.set(view, createRenderer(view));\n    }\n      renderers.get(view)(currentAnalysis, viewContainer);\n  }\n\n  function computeReciprocityIndex(matrix) {\n    const rowTotals = sumRows(matrix);\n    const denom = matrix.length - 1;\n    if (denom <= 0) {\n      return 0;\n    }\n    const total = rowTotals.reduce((acc, value) => acc + value / denom, 0);\n    return total / rowTotals.length;\n  }\n\n  function createRenderer(view) {\n    switch (view) {\n      case \"matrix-pos\":\n        return (analysis, container) => {\n          container.appendChild(renderMatrixSection(\n            \"Matrice positiva\",\n            \"Conta quante volte ogni atleta ha scelto positivamente un'altra.\",\n            analysis.names,\n            analysis.matrices.posMatrix,\n            analysis.totals.posRowSum,\n            analysis.totals.posColSum,\n            \"Totale scelte fatte\",\n            {\n              matrix: { palette: PALETTE.posMatrix, scope: \"matrix\" },\n              totals: { palette: PALETTE.posTotals, scope: \"totals\" }\n            },\n            true,\n            \"Totale scelte ricevute\"\n          ));\n          appendNetworkForView(analysis, container, view);\n        };\n      case \"responses\":\n        return (analysis, container) => {\n          container.appendChild(renderResponsesSection(analysis));\n        };\n      case \"top-domande\":\n        return (analysis, container) => {\n          container.appendChild(renderTopQuestionsSection(analysis));\n        };\n      case \"top-ricevute\":\n        return (analysis, container) => {\n          container.appendChild(renderTopRecipientsSection(analysis));\n        };\n      case \"matrix-neg\":\n        return (analysis, container) => {\n          container.appendChild(renderMatrixSection(\n            \"Matrice negativa\",\n            \"Conta quante volte ogni atleta ha indicato negativamente un'altra.\",\n            analysis.names,\n            analysis.matrices.negMatrix,\n            analysis.totals.negRowSum,\n            analysis.totals.negColSum,\n            \"Totale scelte fatte\",\n            {\n              matrix: { palette: PALETTE.negMatrix, scope: \"matrix\" },\n              totals: { palette: PALETTE.negTotals, scope: \"totals\" }\n            },\n            true,\n            \"Totale scelte ricevute\"\n          ));\n          appendNetworkForView(analysis, container, view);\n        };\n      case \"reciproci-pos\":\n        return (analysis, container) => {\n          container.appendChild(renderReciprociSection(\n            \"Reciproci positivi\",\n            \"1 indica una scelta reciproca positiva.\",\n            analysis.names,\n            analysis.matrices.reciprociPos,\n            {\n              matrix: { palette: PALETTE.posMatrix, scope: \"matrix\" },\n              totals: { palette: PALETTE.posTotals, scope: \"totals\" }\n            }\n          ));\n          container.appendChild(renderNetworkSection({\n            title: \"Rete dei reciproci positivi\",\n            note: \"Il grafo mostra solo i legami reciproci positivi.\",\n            names: analysis.names,\n            matrix: analysis.matrices.reciprociPos,\n            strengthMatrix: analysis.matrices.forzaLegami,\n            nodePalette: PALETTE.network,\n            edgePalette: [\"#d5d8dc\", \"#1a9850\"],\n            toggleId: \"link-strength-toggle-pos\"\n          }));\n        };\n      case \"reciproci-neg\":\n        return (analysis, container) => {\n          container.appendChild(renderReciprociSection(\n            \"Reciproci negativi\",\n            \"1 indica un reciproco negativo.\",\n            analysis.names,\n            analysis.matrices.reciprociNeg,\n            {\n              matrix: { palette: PALETTE.negMatrix, scope: \"matrix\" },\n              totals: { palette: PALETTE.negTotals, scope: \"totals\" }\n            }\n          ));\n          container.appendChild(renderNetworkSection({\n            title: \"Rete dei reciproci negativi\",\n            note: \"Il grafo mostra solo i legami reciproci negativi.\",\n            names: analysis.names,\n            matrix: analysis.matrices.reciprociNeg,\n            strengthMatrix: analysis.matrices.forzaAntagonismo,\n            nodePalette: [...PALETTE.network].reverse(),\n            edgePalette: [\"#d5d8dc\", \"#d73027\"],\n            toggleId: \"link-strength-toggle-neg\"\n          }));\n        };\n      case \"forza-legami\":\n        return (analysis, container) => {\n          container.appendChild(renderMatrixSection(\n            \"Forza legami\",\n            \"Minimo tra scelte reciproche positive.\",\n            analysis.names,\n            analysis.matrices.forzaLegami,\n            analysis.totals.forzaLegamiRowSum,\n            null,\n            \"Totale forza\",\n            {\n              matrix: { palette: PALETTE.posMatrix, scope: \"combined\" },\n              totals: { palette: PALETTE.posMatrix, scope: \"combined\" }\n            },\n            false\n          ));\n        };\n      case \"forza-antagonismo\":\n        return (analysis, container) => {\n          container.appendChild(renderMatrixSection(\n            \"Forza antagonismo\",\n            \"Minimo tra scelte reciproche negative.\",\n            analysis.names,\n            analysis.matrices.forzaAntagonismo,\n            analysis.totals.forzaAntagonismoRowSum,\n            null,\n            \"Totale forza\",\n            {\n              matrix: { palette: PALETTE.negMatrix, scope: \"combined\" },\n              totals: { palette: PALETTE.negMatrix, scope: \"combined\" }\n            },\n            false\n          ));\n        };\n      case \"non-ricambiate\":\n        return (analysis, container) => {\n          container.appendChild(renderMatrixSection(\n            \"Positive non ricambiate\",\n            \"Scelte positive non ricambiate.\",\n            analysis.names,\n            analysis.matrices.nonRicambiatePos,\n            analysis.totals.nonRicambiateRowSum,\n            analysis.totals.nonRicambiateColSum,\n            \"Positive date e non ricambiate\",\n            {\n              matrix: { palette: PALETTE.nonRic, scope: \"combined\" },\n              totals: { palette: PALETTE.nonRic, scope: \"combined\" }\n            },\n            true,\n            \"Positive ricevute non ricambiate\"\n          ));\n        };\n      case \"riassunto\":\n        return (analysis, container) => {\n          container.appendChild(renderSummarySection(analysis));\n        };\n      case \"confronto\":\n        return (_, container) => {\n          container.appendChild(renderCompareSection());\n        };\n      default:\n        return () => undefined;\n    }\n  }\n\n  function appendNetworkForView(analysis, container, view) {\n    const config = getNetworkConfigForView(view, analysis);\n    if (!config) {\n      return;\n    }\n    container.appendChild(renderNetworkSection(config));\n  }\n\n  function getNetworkConfigForView(view, analysis) {\n    const mapping = {\n      \"matrix-pos\": {\n        title: \"Grafo scelte positive\",\n        note: \"Il grafo mostra le scelte positive ricevute.\",\n        matrix: transposeMatrix(analysis.matrices.posMatrix),\n        strengthMatrix: transposeMatrix(analysis.matrices.posMatrix),\n        nodePalette: PALETTE.network,\n        edgePalette: [\"#d5d8dc\", \"#1a9850\"]\n      },\n      \"matrix-neg\": {\n        title: \"Grafo scelte negative\",\n        note: \"Il grafo mostra le scelte negative ricevute.\",\n        matrix: transposeMatrix(analysis.matrices.negMatrix),\n        strengthMatrix: transposeMatrix(analysis.matrices.negMatrix),\n        nodePalette: [...PALETTE.network].reverse(),\n        edgePalette: [\"#d5d8dc\", \"#d73027\"]\n      },\n      \"reciproci-pos\": {\n        title: \"Rete dei reciproci positivi\",\n        note: \"Il grafo mostra solo i legami reciproci positivi.\",\n        matrix: analysis.matrices.reciprociPos,\n        strengthMatrix: analysis.matrices.forzaLegami,\n        nodePalette: PALETTE.network,\n        edgePalette: [\"#d5d8dc\", \"#1a9850\"]\n      },\n      \"reciproci-neg\": {\n        title: \"Rete dei reciproci negativi\",\n        note: \"Il grafo mostra solo i legami reciproci negativi.\",\n        matrix: analysis.matrices.reciprociNeg,\n        strengthMatrix: analysis.matrices.forzaAntagonismo,\n        nodePalette: [...PALETTE.network].reverse(),\n        edgePalette: [\"#d5d8dc\", \"#d73027\"]\n      },\n      \"forza-legami\": {\n        title: \"Grafo forza legami\",\n        note: \"Il grafo usa la forza dei legami positivi.\",\n        matrix: analysis.matrices.forzaLegami,\n        strengthMatrix: analysis.matrices.forzaLegami,\n        nodePalette: PALETTE.network,\n        edgePalette: [\"#d5d8dc\", \"#1a9850\"]\n      },\n      \"forza-antagonismo\": {\n        title: \"Grafo forza antagonismo\",\n        note: \"Il grafo usa la forza delle relazioni negative.\",\n        matrix: analysis.matrices.forzaAntagonismo,\n        strengthMatrix: analysis.matrices.forzaAntagonismo,\n        nodePalette: [...PALETTE.network].reverse(),\n        edgePalette: [\"#d5d8dc\", \"#d73027\"]\n      },\n      \"non-ricambiate\": {\n        title: \"Grafo positive non ricambiate\",\n        note: \"Il grafo mostra le scelte positive non ricambiate.\",\n        matrix: analysis.matrices.nonRicambiatePos,\n        strengthMatrix: analysis.matrices.nonRicambiatePos,\n        nodePalette: PALETTE.network,\n        edgePalette: [\"#d5d8dc\", \"#1a9850\"]\n      },\n      \"responses\": {\n        title: \"Grafo scelte positive\",\n        note: \"Il grafo mostra le scelte positive tra atlete.\",\n        matrix: analysis.matrices.posMatrix,\n        strengthMatrix: analysis.matrices.posMatrix,\n        nodePalette: PALETTE.network,\n        edgePalette: [\"#d5d8dc\", \"#1a9850\"]\n      },\n      \"riassunto\": {\n        title: \"Grafo scelte positive\",\n        note: \"Il grafo mostra le scelte positive tra atlete.\",\n        matrix: analysis.matrices.posMatrix,\n        strengthMatrix: analysis.matrices.posMatrix,\n        nodePalette: PALETTE.network,\n        edgePalette: [\"#d5d8dc\", \"#1a9850\"]\n      }\n    };\n\n    const config = mapping[view];\n    if (!config) {\n      return null;\n    }\n    return {\n      ...config,\n      names: analysis.names,\n      toggleId: `link-strength-toggle-${view}-${Math.random().toString(16).slice(2, 8)}`\n    };\n  }\n\n  function transposeMatrix(matrix) {\n    const size = matrix.length;\n    if (!size) {\n      return [];\n    }\n    const result = createMatrix(size, 0);\n    for (let i = 0; i < size; i += 1) {\n      for (let j = 0; j < size; j += 1) {\n        result[j][i] = matrix[i][j];\n      }\n    }\n    return result;\n  }\n\n  function getCompareDatasets() {\n    const visibleDatasets = getVisibleDatasets();\n    if (visibleDatasets.length < 2) {\n      return null;\n    }\n    const datasetA = visibleDatasets.find((dataset) => dataset.id === compareSelectA.value);\n    const datasetB = visibleDatasets.find((dataset) => dataset.id === compareSelectB.value);\n    if (!datasetA || !datasetB || datasetA.id === datasetB.id) {\n      return null;\n    }\n    return { datasetA, datasetB };\n  }\n\n  function renderDualView(view, { datasetA, datasetB }) {\n    viewContainer.classList.add(\"dual-view\");\n    if (!renderers.has(view)) {\n      renderers.set(view, createRenderer(view));\n    }\n    const renderer = renderers.get(view);\n    const panels = [];\n    [datasetA, datasetB].forEach((dataset) => {\n      const panel = document.createElement(\"div\");\n      panel.className = \"view-panel\";\n      const label = document.createElement(\"p\");\n      label.className = \"note\";\n      label.textContent = dataset.name;\n      panel.appendChild(label);\n      renderer(dataset.analysis, panel);\n      viewContainer.appendChild(panel);\n      panels.push(panel);\n    });\n    if (panels.length === 2) {\n      highlightDifferences(panels[0], panels[1]);\n      syncDualScroll(panels[0], panels[1]);\n    }\n  }\n\n  function highlightDifferences(panelA, panelB) {\n    const tablesA = Array.from(panelA.querySelectorAll(\"table\"));\n    const tablesB = Array.from(panelB.querySelectorAll(\"table\"));\n    const count = Math.min(tablesA.length, tablesB.length);\n    for (let i = 0; i < count; i += 1) {\n      compareTables(tablesA[i], tablesB[i]);\n    }\n  }\n\n  function syncDualScroll(panelA, panelB) {\n    const wrapsA = Array.from(panelA.querySelectorAll(\".table-wrap\"));\n    const wrapsB = Array.from(panelB.querySelectorAll(\".table-wrap\"));\n    const count = Math.min(wrapsA.length, wrapsB.length);\n    for (let i = 0; i < count; i += 1) {\n      attachScrollSync(wrapsA[i], wrapsB[i]);\n    }\n  }\n\n  function attachScrollSync(elA, elB) {\n    const key = \"data-sync-attached\";\n    if (elA.getAttribute(key) && elB.getAttribute(key)) {\n      return;\n    }\n    let isSyncing = false;\n    const sync = (source, target) => {\n      if (isSyncing) {\n        return;\n      }\n      isSyncing = true;\n      target.scrollLeft = source.scrollLeft;\n      target.scrollTop = source.scrollTop;\n      requestAnimationFrame(() => {\n        isSyncing = false;\n      });\n    };\n    elA.addEventListener(\"scroll\", () => sync(elA, elB));\n    elB.addEventListener(\"scroll\", () => sync(elB, elA));\n    elA.setAttribute(key, \"true\");\n    elB.setAttribute(key, \"true\");\n  }\n\n  function compareTables(tableA, tableB) {\n    const rowsA = getTableDataRows(tableA);\n    const rowsB = getTableDataRows(tableB);\n    const rowCount = Math.min(rowsA.length, rowsB.length);\n    for (let r = 0; r < rowCount; r += 1) {\n      const cellsA = rowsA[r];\n      const cellsB = rowsB[r];\n      const cellCount = Math.min(cellsA.length, cellsB.length);\n      for (let c = 0; c < cellCount; c += 1) {\n        const cellA = cellsA[c];\n        const cellB = cellsB[c];\n        const valueA = parseCellValue(cellA);\n        const valueB = parseCellValue(cellB);\n        if (valueA !== null && valueB !== null) {\n          if (valueA !== valueB) {\n            markDiffCell(cellA, 0);\n            markDiffCell(cellB, valueB - valueA);\n          }\n        } else if (cellA.textContent.trim() !== cellB.textContent.trim()) {\n          markDiffCell(cellA, 0);\n          markDiffCell(cellB, 0);\n        }\n      }\n    }\n  }\n\n  function getTableDataRows(table) {\n    const rows = Array.from(table.querySelectorAll(\"tbody tr, tfoot tr\"));\n    return rows.map((row) => Array.from(row.querySelectorAll(\"td\")));\n  }\n\n  function parseCellValue(cell) {\n    if (!cell) {\n      return null;\n    }\n    if (cell.dataset && cell.dataset.rank) {\n      const rank = Number(cell.dataset.rank);\n      return Number.isFinite(rank) ? rank : null;\n    }\n    const text = cell.textContent;\n    if (!text) {\n      return null;\n    }\n    const cleaned = text.trim().replace(\",\", \".\");\n    if (cleaned === \"\u2014\") {\n      return null;\n    }\n    if (cleaned.endsWith(\"%\")) {\n      const num = Number.parseFloat(cleaned.slice(0, -1));\n      return Number.isFinite(num) ? num / 100 : null;\n    }\n    const value = Number.parseFloat(cleaned.replace(/[^0-9.-]/g, \"\"));\n    return Number.isFinite(value) ? value : null;\n  }\n\n  function markDiffCell(cell, delta) {\n    if (!Number.isFinite(delta) || delta === 0) {\n      return;\n    }\n    const arrow = document.createElement(\"span\");\n    arrow.className = delta > 0 ? \"diff-arrow diff-up\" : \"diff-arrow diff-down\";\n    arrow.textContent = delta > 0 ? \"\u25b2\" : \"\u25bc\";\n    cell.appendChild(arrow);\n  }\n\n  function renderMatrixSection(title, description, names, matrix, rowTotals, colTotals, totalLabel, colorConfig, includeTotalsRow, totalsRowLabel = \"Totali\") {\n    const section = document.createElement(\"div\");\n    const heading = document.createElement(\"h2\");\n    heading.textContent = title;\n    const note = document.createElement(\"p\");\n    note.className = \"note\";\n    note.textContent = description;\n    section.appendChild(heading);\n    section.appendChild(note);\n\n    const table = document.createElement(\"table\");\n    table.className = \"matrix-table\";\n    const thead = document.createElement(\"thead\");\n    const headerRow = document.createElement(\"tr\");\n\n    const firstTh = document.createElement(\"th\");\n    firstTh.innerHTML = \"Chi viene scelto \u2192<br>Chi sceglie \u2193\";\n    headerRow.appendChild(firstTh);\n\n    names.forEach((name) => {\n      const th = document.createElement(\"th\");\n      setHeaderText(th, name);\n      headerRow.appendChild(th);\n    });\n\n    const totalTh = document.createElement(\"th\");\n    setHeaderText(totalTh, totalLabel);\n    totalTh.classList.add(\"sticky-right\");\n    headerRow.appendChild(totalTh);\n\n    thead.appendChild(headerRow);\n    table.appendChild(thead);\n\n    const tbody = document.createElement(\"tbody\");\n    const minMax = getMinMaxConfig(matrix, rowTotals, colTotals, colorConfig);\n    matrix.forEach((row, i) => {\n      const tr = document.createElement(\"tr\");\n      const rowHeader = document.createElement(\"th\");\n      rowHeader.textContent = names[i];\n      tr.appendChild(rowHeader);\n\n      row.forEach((cell) => {\n        const td = document.createElement(\"td\");\n        td.textContent = cell;\n        if (colorConfig?.matrix) {\n          const range = minMax.matrix;\n          applyHeatmap(td, cell, range.min, range.max, colorConfig.matrix.palette);\n        }\n        tr.appendChild(td);\n      });\n\n      const total = document.createElement(\"td\");\n      total.textContent = rowTotals[i];\n      total.classList.add(\"sticky-right\");\n      if (colorConfig?.totals) {\n        const range = minMax.totals;\n        applyHeatmap(total, rowTotals[i], range.min, range.max, colorConfig.totals.palette);\n      }\n      tr.appendChild(total);\n\n      tbody.appendChild(tr);\n    });\n    table.appendChild(tbody);\n\n    if (includeTotalsRow && colTotals) {\n      const tfoot = document.createElement(\"tfoot\");\n      const tr = document.createElement(\"tr\");\n      const label = document.createElement(\"th\");\n      label.textContent = totalsRowLabel;\n      tr.appendChild(label);\n\n      colTotals.forEach((value) => {\n        const td = document.createElement(\"td\");\n        td.textContent = value;\n        if (colorConfig?.totals) {\n          const range = minMax.totals;\n          applyHeatmap(td, value, range.min, range.max, colorConfig.totals.palette);\n        }\n        tr.appendChild(td);\n      });\n\n      const total = document.createElement(\"td\");\n      total.textContent = colTotals.reduce((acc, value) => acc + value, 0);\n      total.classList.add(\"sticky-right\");\n      tr.appendChild(total);\n\n      tfoot.appendChild(tr);\n      table.appendChild(tfoot);\n    }\n\n    const wrapper = document.createElement(\"div\");\n    wrapper.className = \"table-wrap\";\n    wrapper.appendChild(table);\n    section.appendChild(wrapper);\n\n    return section;\n  }\n\n  function renderReciprociSection(title, description, names, matrix, colorConfig) {\n    const section = document.createElement(\"div\");\n    const heading = document.createElement(\"h2\");\n    heading.textContent = title;\n    const note = document.createElement(\"p\");\n    note.className = \"note\";\n    note.textContent = description;\n    section.appendChild(heading);\n    section.appendChild(note);\n\n    const table = document.createElement(\"table\");\n    table.className = \"matrix-table\";\n    const thead = document.createElement(\"thead\");\n    const headerRow = document.createElement(\"tr\");\n    const firstTh = document.createElement(\"th\");\n    firstTh.innerHTML = \"Chi viene scelto \u2192<br>Chi sceglie \u2193\";\n    headerRow.appendChild(firstTh);\n\n    names.forEach((name) => {\n      const th = document.createElement(\"th\");\n      setHeaderText(th, name);\n      headerRow.appendChild(th);\n    });\n\n    const totalTh = document.createElement(\"th\");\n    setHeaderText(totalTh, \"Totale reciprocit\u00e0\");\n    totalTh.classList.add(\"sticky-right\");\n    headerRow.appendChild(totalTh);\n\n    const idxTh = document.createElement(\"th\");\n    setHeaderText(idxTh, \"Indice reciprocit\u00e0\");\n    headerRow.appendChild(idxTh);\n\n    thead.appendChild(headerRow);\n    table.appendChild(thead);\n\n    const tbody = document.createElement(\"tbody\");\n    const rowTotals = sumRows(matrix);\n    const colTotals = sumColumns(matrix);\n    const reciprocalIndex = rowTotals.map((total, idx) => {\n      const denom = matrix.length - 1;\n      return denom > 0 ? total / denom : 0;\n    });\n\n    const minMax = getMinMaxConfig(matrix, rowTotals, colTotals, colorConfig);\n    matrix.forEach((row, i) => {\n      const tr = document.createElement(\"tr\");\n      const rowHeader = document.createElement(\"th\");\n      rowHeader.textContent = names[i];\n      tr.appendChild(rowHeader);\n\n      row.forEach((cell) => {\n        const td = document.createElement(\"td\");\n        td.textContent = cell;\n        if (colorConfig?.matrix) {\n          const range = minMax.matrix;\n          applyHeatmap(td, cell, range.min, range.max, colorConfig.matrix.palette);\n        }\n        tr.appendChild(td);\n      });\n\n      const total = document.createElement(\"td\");\n      total.textContent = rowTotals[i];\n      total.classList.add(\"sticky-right\");\n      if (colorConfig?.totals) {\n        const range = minMax.totals;\n        applyHeatmap(total, rowTotals[i], range.min, range.max, colorConfig.totals.palette);\n      }\n      tr.appendChild(total);\n\n      const idxCell = document.createElement(\"td\");\n      idxCell.textContent = formatPercent(reciprocalIndex[i]);\n      tr.appendChild(idxCell);\n\n      tbody.appendChild(tr);\n    });\n\n    table.appendChild(tbody);\n\n    const tfoot = document.createElement(\"tfoot\");\n    const tr = document.createElement(\"tr\");\n    const label = document.createElement(\"th\");\n    label.textContent = \"Totali\";\n    tr.appendChild(label);\n\n    colTotals.forEach((value) => {\n      const td = document.createElement(\"td\");\n      td.textContent = value;\n      if (colorConfig?.totals) {\n        const range = minMax.totals;\n        applyHeatmap(td, value, range.min, range.max, colorConfig.totals.palette);\n      }\n      tr.appendChild(td);\n    });\n\n    const totalSum = rowTotals.reduce((acc, value) => acc + value, 0);\n    const totalCell = document.createElement(\"td\");\n    totalCell.textContent = totalSum;\n    totalCell.classList.add(\"sticky-right\");\n    tr.appendChild(totalCell);\n\n    const avgIdx = reciprocalIndex.reduce((acc, value) => acc + value, 0) / reciprocalIndex.length;\n    const avgCell = document.createElement(\"td\");\n    avgCell.textContent = formatPercent(avgIdx);\n    tr.appendChild(avgCell);\n\n    tfoot.appendChild(tr);\n    table.appendChild(tfoot);\n\n    const wrapper = document.createElement(\"div\");\n    wrapper.className = \"table-wrap\";\n    wrapper.appendChild(table);\n    section.appendChild(wrapper);\n\n    return section;\n  }\n\n  function setHeaderText(cell, text) {\n    const value = text == null ? \"\" : String(text);\n    const words = value.trim().split(/\\s+/).filter(Boolean);\n    cell.textContent = \"\";\n    if (words.length <= 1) {\n      cell.textContent = value;\n      return;\n    }\n    const lines = [];\n    words.forEach((word) => {\n      if (!lines.length) {\n        lines.push(word);\n        return;\n      }\n      if (word.length <= 1) {\n        lines[lines.length - 1] = `${lines[lines.length - 1]} ${word}`;\n        return;\n      }\n      lines.push(word);\n    });\n    lines.forEach((line) => {\n      const span = document.createElement(\"span\");\n      span.textContent = line;\n      span.style.display = \"block\";\n      cell.appendChild(span);\n    });\n  }\n\n  async function buildStaticHtml(dataset) {\n    const baseHtml = document.documentElement.outerHTML;\n    const [styles, appJs] = await Promise.all([\n      getAssetText(\"styles.css\", \".css\", \"Seleziona styles.css\"),\n      getAssetText(\"app.js\", \".js\", \"Seleziona app.js\")\n    ]);\n\n    const exportedDatasets = datasets.length ? datasets : [dataset];\n    const state = {\n      currentDatasetId,\n      compareA: compareSelectA?.value || null,\n      compareB: compareSelectB?.value || null\n    };\n    const activeView = getActiveView() || restoreView() || \"responses\";\n    const baseOrder = [\n      state.compareA,\n      state.compareB,\n      ...exportedDatasets.map((entry) => entry.id)\n    ].filter(Boolean);\n    const exportOrder = Array.from(new Set(baseOrder));\n\n    const datasetsJson = escapeScriptTag(JSON.stringify(exportedDatasets.map((entry) => ({\n      id: entry.id,\n      name: entry.name,\n      csv: entry.csv\n    }))));\n    const exportOrderJson = escapeScriptTag(JSON.stringify(exportOrder));\n    const stateJson = escapeScriptTag(JSON.stringify(state));\n    const viewJson = escapeScriptTag(JSON.stringify(activeView));\n\n    const preload = `\n      window.TEAMWEAVE_DISABLE_SW = true;\n      window.TEAMWEAVE_EXPORT_MODE = true;\n      window.TEAMWEAVE_EXPORT_ORDER = ${exportOrderJson};\n      (function() {\n        const datasets = ${datasetsJson};\n        const state = ${stateJson};\n        const view = ${viewJson};\n        localStorage.setItem(${JSON.stringify(DATASETS_KEY)}, JSON.stringify(datasets));\n        localStorage.setItem(${JSON.stringify(DATASET_STATE_KEY)}, JSON.stringify(state));\n        localStorage.setItem(${JSON.stringify(VIEW_KEY)}, view);\n      })();\n    `;\n\n    const safeAppJs = escapeScriptTag(appJs);\n    const inlineScript = `<script>${preload}\\n${safeAppJs}</script>`;\n    const inlineStyles = `<style>\\n${styles}\\n</style>`;\n\n    const manifestPattern = new RegExp(\"<link[^>]*manifest\\\\.json[^>]*>\", \"i\");\n    const stylesPattern = new RegExp(\"<link[^>]*styles\\\\.css[^>]*>\", \"i\");\n    const scriptPattern = new RegExp(\"<script[^>]*app\\\\.js[^>]*><\" + \"/script>\", \"i\");\n\n    let html = `<!doctype html>\\n${baseHtml}`\n      .replace(manifestPattern, \"\")\n      .replace(stylesPattern, inlineStyles)\n      .replace(scriptPattern, inlineScript);\n\n    if (!html.includes(inlineScript)) {\n      html = html.replace(\"</body>\", `${inlineScript}\\n</body>`);\n    }\n    if (!html.includes(inlineStyles)) {\n      html = html.replace(\"</head>\", `${inlineStyles}\\n</head>`);\n    }\n\n    return html;\n  }\n\n  async function getAssetText(url, accept, label) {\n    try {\n      const response = await fetch(url, { cache: \"no-store\" });\n      if (response.ok) {\n        return await response.text();\n      }\n    } catch (error) {\n      // Fallback to inline content for blocked fetch.\n    }\n    if (url.endsWith(\"styles.css\") && INLINE_STYLES !== \"__INLINE_STYLES__\") {\n      return INLINE_STYLES;\n    }\n    if (url.endsWith(\"app.js\") && INLINE_APP !== \"__INLINE_APP__\") {\n      return INLINE_APP;\n    }\n    throw new Error(`Impossibile recuperare: ${label}`);\n  }\n\n  function escapeScriptTag(value) {\n    const pattern = new RegExp(\"<\" + \"/script\", \"gi\");\n    return String(value || \"\").replace(pattern, \"<\\\\/script\");\n  }\n\n  function escapeHtml(value) {\n    return String(value || \"\")\n      .replace(/&/g, \"&amp;\")\n      .replace(/</g, \"&lt;\")\n      .replace(/>/g, \"&gt;\")\n      .replace(/\"/g, \"&quot;\")\n      .replace(/'/g, \"&#39;\");\n  }\n\n  function renderSummarySection(analysis) {\n    const section = document.createElement(\"div\");\n    const heading = document.createElement(\"h2\");\n    heading.textContent = \"Dati riassuntivi\";\n    section.appendChild(heading);\n\n    const note = document.createElement(\"p\");\n    note.className = \"note\";\n    note.textContent = \"Indicatori principali per atleta + classificazione automatica.\";\n    section.appendChild(note);\n\n    const table = document.createElement(\"table\");\n    table.className = \"matrix-table\";\n\n    const thead = document.createElement(\"thead\");\n    const row1 = document.createElement(\"tr\");\n    [\n      \"Atleta\",\n      \"Scelte ricevute + (Tecniche)\",\n      \"Scelte ricevute + (Attitudinali)\",\n      \"Scelte ricevute + (Sociali)\",\n      \"Scelte ricevute + (Totali)\",\n      \"Scelte ricevute - (Tecniche)\",\n      \"Scelte ricevute - (Attitudinali)\",\n      \"Scelte ricevute - (Sociali)\",\n      \"Scelte ricevute - (Totali)\",\n      \"Scelte date +\",\n      \"Scelte date -\",\n      \"Reciproci +\",\n      \"Reciproci -\",\n      \"Forza legami\",\n      \"Forza antagonismo\",\n      \"Positive date e non ricambiate\",\n      \"Positive ricevute e non ricambiate\"\n    ].forEach((label) => {\n      const th = document.createElement(\"th\");\n      setHeaderText(th, label);\n      row1.appendChild(th);\n    });\n    thead.appendChild(row1);\n    table.appendChild(thead);\n\n    const tbody = document.createElement(\"tbody\");\n    analysis.summaryRows.forEach((row) => {\n      const tr = document.createElement(\"tr\");\n      const cells = [\n        row.name,\n        row.positiveReceived.Tecniche,\n        row.positiveReceived.Attitudinali,\n        row.positiveReceived.Sociali,\n        row.positiveReceived.Totali,\n        row.negativeReceived.Tecniche,\n        row.negativeReceived.Attitudinali,\n        row.negativeReceived.Sociali,\n        row.negativeReceived.Totali,\n        row.positiveGiven,\n        row.negativeGiven,\n        row.reciprociPos,\n        row.reciprociNeg,\n        row.forzaLegami,\n        row.forzaAntagonismo,\n        row.nonRicambiateDate,\n        row.nonRicambiateRicevute\n      ];\n\n      cells.forEach((value, index) => {\n        const cell = index === 0 ? document.createElement(\"th\") : document.createElement(\"td\");\n        cell.textContent = value;\n        tr.appendChild(cell);\n      });\n      tbody.appendChild(tr);\n    });\n    table.appendChild(tbody);\n\n    const tfoot = document.createElement(\"tfoot\");\n    const tr = document.createElement(\"tr\");\n    const averageLabel = document.createElement(\"th\");\n    averageLabel.textContent = \"Media di squadra\";\n    tr.appendChild(averageLabel);\n\n    const avgCells = [\n      analysis.averages.positiveReceivedTecniche,\n      analysis.averages.positiveReceivedAttitudinali,\n      analysis.averages.positiveReceivedSociali,\n      analysis.averages.positiveReceivedTotal,\n      analysis.averages.negativeReceivedTecniche,\n      analysis.averages.negativeReceivedAttitudinali,\n      analysis.averages.negativeReceivedSociali,\n      analysis.averages.negativeReceivedTotal,\n      analysis.averages.positiveGiven,\n      analysis.averages.negativeGiven,\n      analysis.averages.reciprociPos,\n      analysis.averages.reciprociNeg,\n      analysis.averages.forzaLegami,\n      analysis.averages.forzaAntagonismo,\n      analysis.averages.nonRicambiateDate,\n      analysis.averages.nonRicambiateRicevute\n    ];\n\n    avgCells.forEach((value) => {\n      const td = document.createElement(\"td\");\n      td.textContent = formatNumber(value);\n      tr.appendChild(td);\n    });\n\n    tfoot.appendChild(tr);\n    table.appendChild(tfoot);\n\n    const wrapper = document.createElement(\"div\");\n    wrapper.className = \"table-wrap\";\n    wrapper.appendChild(table);\n    section.appendChild(wrapper);\n\n    applySummaryColors(table, analysis);\n\n    const classHeading = document.createElement(\"h3\");\n    classHeading.textContent = \"Classificazione\";\n    classHeading.style.marginTop = \"2rem\";\n    section.appendChild(classHeading);\n\n    const classTable = document.createElement(\"table\");\n    classTable.className = \"matrix-table\";\n    const classHead = document.createElement(\"thead\");\n    const classRow = document.createElement(\"tr\");\n    [\"Atleta\", \"Influenza positiva\", \"Influenza negativa\", \"Equilibrio relazionale\", \"Etichetta sintetica\"].forEach((label) => {\n      const th = document.createElement(\"th\");\n      setHeaderText(th, label);\n      classRow.appendChild(th);\n    });\n    classHead.appendChild(classRow);\n    classTable.appendChild(classHead);\n\n    const classBody = document.createElement(\"tbody\");\n    const influenceRank = { Bassa: 1, Media: 2, Alta: 3 };\n    const equilibrioRank = { Ignorata: 1, Bilanciata: 2, Selettiva: 3 };\n    analysis.classifications.forEach((row) => {\n      const tr = document.createElement(\"tr\");\n      [row.name, row.inflPos, row.inflNeg, row.equilibrio, row.label].forEach((value, idx) => {\n        const cell = idx === 0 ? document.createElement(\"th\") : document.createElement(\"td\");\n        cell.textContent = value;\n        if (idx === 1) {\n          applyLabelFill(cell, value, \"positive\");\n          if (influenceRank[value]) {\n            cell.dataset.rank = influenceRank[value];\n          }\n        } else if (idx === 2) {\n          applyLabelFill(cell, value, \"negative\");\n          if (influenceRank[value]) {\n            cell.dataset.rank = influenceRank[value];\n          }\n        } else if (idx === 3) {\n          applyEquilibrioFill(cell, value);\n          if (equilibrioRank[value]) {\n            cell.dataset.rank = equilibrioRank[value];\n          }\n        } else if (idx === 4) {\n          applyEtichettaFill(cell, value);\n        }\n        tr.appendChild(cell);\n      });\n      classBody.appendChild(tr);\n    });\n    classTable.appendChild(classBody);\n\n    const classWrap = document.createElement(\"div\");\n    classWrap.className = \"table-wrap\";\n    classWrap.appendChild(classTable);\n    section.appendChild(classWrap);\n\n    return section;\n  }\n\n  function renderCompareSection() {\n    const section = document.createElement(\"div\");\n    const heading = document.createElement(\"h2\");\n    heading.textContent = \"Confronto test\";\n    section.appendChild(heading);\n\n    const note = document.createElement(\"p\");\n    note.className = \"note\";\n    note.textContent = \"Le differenze sono calcolate come B - A.\";\n    section.appendChild(note);\n\n    if (datasets.length < 2) {\n      const empty = document.createElement(\"p\");\n      empty.textContent = \"Carica almeno due dataset per vedere il confronto.\";\n      section.appendChild(empty);\n      return section;\n    }\n\n    const datasetA = datasets.find((dataset) => dataset.id === compareSelectA.value);\n    const datasetB = datasets.find((dataset) => dataset.id === compareSelectB.value);\n    if (!datasetA || !datasetB) {\n      const empty = document.createElement(\"p\");\n      empty.textContent = \"Seleziona due dataset validi.\";\n      section.appendChild(empty);\n      return section;\n    }\n\n    const title = document.createElement(\"p\");\n    title.className = \"note\";\n    title.textContent = `A: ${datasetA.name} | B: ${datasetB.name}`;\n    section.appendChild(title);\n\n    const summaryTable = document.createElement(\"table\");\n    summaryTable.className = \"matrix-table\";\n    const summaryHead = document.createElement(\"thead\");\n    const summaryRow = document.createElement(\"tr\");\n    [\"Metrica\", \"A\", \"B\", \"Delta\"].forEach((label) => {\n      const th = document.createElement(\"th\");\n      setHeaderText(th, label);\n      summaryRow.appendChild(th);\n    });\n    summaryHead.appendChild(summaryRow);\n    summaryTable.appendChild(summaryHead);\n\n    const summaryBody = document.createElement(\"tbody\");\n    const summaryMetrics = [\n      {\n        label: \"Risposte\",\n        get: (analysis) => analysis.rows.length,\n        polarity: \"neutral\"\n      },\n      {\n        label: \"Atlete\",\n        get: (analysis) => analysis.names.length,\n        polarity: \"neutral\"\n      },\n      {\n        label: \"Indice reciprocita +\",\n        get: (analysis) => computeReciprocityIndex(analysis.matrices.reciprociPos),\n        polarity: \"higher\"\n      },\n      {\n        label: \"Indice reciprocita -\",\n        get: (analysis) => computeReciprocityIndex(analysis.matrices.reciprociNeg),\n        polarity: \"lower\"\n      }\n    ];\n\n    summaryMetrics.forEach((metric) => {\n      const tr = document.createElement(\"tr\");\n      const label = document.createElement(\"th\");\n      label.textContent = metric.label;\n      tr.appendChild(label);\n      const aValue = metric.get(datasetA.analysis);\n      const bValue = metric.get(datasetB.analysis);\n      const delta = Number(bValue) - Number(aValue);\n      const cells = [formatCompareValue(aValue), formatCompareValue(bValue), formatCompareDelta(delta, metric.label)];\n      cells.forEach((value, idx) => {\n        const td = document.createElement(\"td\");\n        td.textContent = value;\n        if (idx === 2) {\n          applyDeltaClass(td, delta, metric.polarity);\n        }\n        tr.appendChild(td);\n      });\n      summaryBody.appendChild(tr);\n    });\n    summaryTable.appendChild(summaryBody);\n\n    const summaryWrap = document.createElement(\"div\");\n    summaryWrap.className = \"table-wrap\";\n    summaryWrap.appendChild(summaryTable);\n    section.appendChild(summaryWrap);\n\n    const detailHeading = document.createElement(\"h3\");\n    detailHeading.textContent = \"Differenze per atleta\";\n    detailHeading.style.marginTop = \"2rem\";\n    section.appendChild(detailHeading);\n\n    const compareTable = document.createElement(\"table\");\n    compareTable.className = \"matrix-table\";\n    const head = document.createElement(\"thead\");\n    const headerRow = document.createElement(\"tr\");\n    [\n      \"Atleta\",\n      \"Delta scelte ricevute + (Totali)\",\n      \"Delta scelte ricevute - (Totali)\",\n      \"Delta reciproci +\",\n      \"Delta reciproci -\",\n      \"Delta positive ricevute non ricambiate\",\n      \"Delta positive date non ricambiate\"\n    ].forEach((label) => {\n      const th = document.createElement(\"th\");\n      setHeaderText(th, label);\n      headerRow.appendChild(th);\n    });\n    head.appendChild(headerRow);\n    compareTable.appendChild(head);\n\n    const body = document.createElement(\"tbody\");\n    const mapA = buildSummaryMap(datasetA.analysis.summaryRows);\n    const mapB = buildSummaryMap(datasetB.analysis.summaryRows);\n    const allNames = Array.from(\n      new Set([...datasetA.analysis.names, ...datasetB.analysis.names].map((name) => normalizeName(name)))\n    );\n    const nameLookup = new Map();\n    datasetA.analysis.names.forEach((name) => nameLookup.set(normalizeName(name), name));\n    datasetB.analysis.names.forEach((name) => {\n      if (!nameLookup.has(normalizeName(name))) {\n        nameLookup.set(normalizeName(name), name);\n      }\n    });\n    allNames.forEach((key) => {\n      const rowA = mapA.get(key);\n      const rowB = mapB.get(key);\n      const tr = document.createElement(\"tr\");\n      const nameCell = document.createElement(\"th\");\n      nameCell.textContent = nameLookup.get(key) || key;\n      tr.appendChild(nameCell);\n      const metrics = [\n        { value: diffValue(rowA?.positiveReceived?.Totali, rowB?.positiveReceived?.Totali), polarity: \"higher\" },\n        { value: diffValue(rowA?.negativeReceived?.Totali, rowB?.negativeReceived?.Totali), polarity: \"lower\" },\n        { value: diffValue(rowA?.reciprociPos, rowB?.reciprociPos), polarity: \"higher\" },\n        { value: diffValue(rowA?.reciprociNeg, rowB?.reciprociNeg), polarity: \"lower\" },\n        { value: diffValue(rowA?.nonRicambiateRicevute, rowB?.nonRicambiateRicevute), polarity: \"lower\" },\n        { value: diffValue(rowA?.nonRicambiateDate, rowB?.nonRicambiateDate), polarity: \"lower\" }\n      ];\n      metrics.forEach((metric) => {\n        const td = document.createElement(\"td\");\n        td.textContent = formatCompareDelta(metric.value);\n        applyDeltaClass(td, metric.value, metric.polarity);\n        tr.appendChild(td);\n      });\n      body.appendChild(tr);\n    });\n    compareTable.appendChild(body);\n\n    const compareWrap = document.createElement(\"div\");\n    compareWrap.className = \"table-wrap\";\n    compareWrap.appendChild(compareTable);\n    section.appendChild(compareWrap);\n\n    return section;\n  }\n\n  function buildSummaryMap(rows) {\n    const map = new Map();\n    rows.forEach((row) => {\n      map.set(normalizeName(row.name), row);\n    });\n    return map;\n  }\n\n  function diffValue(valueA, valueB) {\n    if (!Number.isFinite(valueA) || !Number.isFinite(valueB)) {\n      return null;\n    }\n    return valueB - valueA;\n  }\n\n  function formatCompareValue(value) {\n    if (!Number.isFinite(value)) {\n      return \"\u2014\";\n    }\n    if (value % 1 !== 0 && value <= 1) {\n      return formatPercent(value);\n    }\n    return formatNumber(value);\n  }\n\n  function formatCompareDelta(value, label) {\n    if (!Number.isFinite(value)) {\n      return \"\u2014\";\n    }\n    const sign = value > 0 ? \"+\" : \"\";\n    if (label && label.includes(\"Indice\")) {\n      return `${sign}${formatPercent(value)}`;\n    }\n    return `${sign}${formatNumber(value)}`;\n  }\n\n  function applyDeltaClass(cell, value, polarity) {\n    if (!Number.isFinite(value) || value === 0 || polarity === \"neutral\") {\n      return;\n    }\n    const isPositive = value > 0;\n    if (polarity === \"higher\") {\n      cell.classList.add(isPositive ? \"delta-good\" : \"delta-bad\");\n      return;\n    }\n    if (polarity === \"lower\") {\n      cell.classList.add(isPositive ? \"delta-bad\" : \"delta-good\");\n    }\n  }\n\n  function renderResponsesSection(analysis) {\n    const section = document.createElement(\"div\");\n    const heading = document.createElement(\"h2\");\n    heading.textContent = \"Risposte del modulo\";\n    section.appendChild(heading);\n\n    const note = document.createElement(\"p\");\n    note.className = \"note\";\n    note.textContent = \"Naviga per atleta e consulta tutte le risposte organizzate per categoria.\";\n    section.appendChild(note);\n\n    const list = document.createElement(\"div\");\n    list.className = \"responses-list\";\n    section.appendChild(list);\n\n    const responseData = buildResponsesData(analysis);\n\n    list.innerHTML = \"\";\n    responseData.forEach((item) => {\n      const details = document.createElement(\"details\");\n      details.className = \"response-card\";\n\n      const summary = document.createElement(\"summary\");\n      const title = document.createElement(\"span\");\n      title.textContent = item.name;\n      const meta = document.createElement(\"span\");\n      meta.className = \"response-meta\";\n      meta.textContent = `${item.counts.positive} positive \u00b7 ${item.counts.negative} negative`;\n      summary.appendChild(title);\n      summary.appendChild(meta);\n      details.appendChild(summary);\n\n      const grid = document.createElement(\"div\");\n      grid.className = \"response-grid\";\n\n      item.groups.forEach((group) => {\n        const groupEl = document.createElement(\"div\");\n        groupEl.className = \"response-group\";\n        const h4 = document.createElement(\"h4\");\n        h4.textContent = group.title;\n        const pill = document.createElement(\"span\");\n        pill.className = \"response-pill\";\n        pill.textContent = group.count;\n        h4.appendChild(pill);\n        groupEl.appendChild(h4);\n\n        const items = document.createElement(\"div\");\n        items.className = \"response-items\";\n        group.items.forEach((entry) => {\n          const row = document.createElement(\"div\");\n          row.className = `response-item ${entry.sentiment || \"\"}`.trim();\n\n          const question = document.createElement(\"div\");\n          question.className = \"response-question\";\n          question.textContent = entry.question;\n\n          const answer = document.createElement(\"div\");\n          answer.className = \"response-answer\";\n          if (entry.answer) {\n            answer.textContent = entry.answer;\n          } else {\n            answer.textContent = \"Nessuna risposta\";\n            answer.classList.add(\"empty\");\n          }\n\n          row.appendChild(question);\n          row.appendChild(answer);\n          items.appendChild(row);\n        });\n        groupEl.appendChild(items);\n        grid.appendChild(groupEl);\n      });\n\n      details.appendChild(grid);\n      list.appendChild(details);\n    });\n\n    return section;\n  }\n\n  function renderTopQuestionsSection(analysis) {\n    const section = document.createElement(\"div\");\n    const heading = document.createElement(\"h2\");\n    heading.textContent = \"Domande con pi\\u00f9 risposte ricevute\";\n    section.appendChild(heading);\n\n    const note = document.createElement(\"p\");\n    note.className = \"note\";\n    note.textContent = \"Per ogni atleta mostra le domande con pi\\u00f9 risposte positive e negative ricevute.\";\n    section.appendChild(note);\n\n    const controls = document.createElement(\"div\");\n    controls.className = \"top-questions-controls\";\n    const label = document.createElement(\"label\");\n    label.textContent = \"Numero di domande per atleta\";\n    const input = document.createElement(\"input\");\n    input.type = \"number\";\n    input.min = \"1\";\n    const maxLimit = Math.max(analysis.questionStats.posLabels.length, analysis.questionStats.negLabels.length, 1);\n    input.max = String(maxLimit);\n    input.value = \"3\";\n    label.appendChild(input);\n    controls.appendChild(label);\n    section.appendChild(controls);\n\n    const list = document.createElement(\"div\");\n    list.className = \"top-questions-list\";\n    section.appendChild(list);\n\n    const renderList = (limit) => {\n      list.innerHTML = \"\";\n      analysis.names.forEach((name, index) => {\n        const card = document.createElement(\"div\");\n        card.className = \"response-card top-questions-card\";\n\n        const title = document.createElement(\"h3\");\n        title.textContent = name;\n        card.appendChild(title);\n\n        const grid = document.createElement(\"div\");\n        grid.className = \"top-questions-grid\";\n\n        grid.appendChild(renderTopQuestionColumn(\n          \"Positive ricevute\",\n          analysis.questionStats.posLabels,\n          analysis.questionStats.posReceivedByQuestion[index],\n          limit,\n          \"positive\"\n        ));\n        grid.appendChild(renderTopQuestionColumn(\n          \"Negative ricevute\",\n          analysis.questionStats.negLabels,\n          analysis.questionStats.negReceivedByQuestion[index],\n          limit,\n          \"negative\"\n        ));\n\n        card.appendChild(grid);\n        list.appendChild(card);\n      });\n    };\n\n    const normalizeLimit = () => {\n      const parsed = Number.parseInt(input.value, 10);\n      const safe = Number.isFinite(parsed) ? parsed : 1;\n      return Math.max(1, Math.min(maxLimit, safe));\n    };\n\n    input.addEventListener(\"input\", () => {\n      const limit = normalizeLimit();\n      input.value = String(limit);\n      renderList(limit);\n    });\n\n    renderList(normalizeLimit());\n    return section;\n  }\n\n  function renderTopRecipientsSection(analysis) {\n    const section = document.createElement(\"div\");\n    const heading = document.createElement(\"h2\");\n    heading.textContent = \"Atlete pi\\u00f9 scelte per domanda\";\n    section.appendChild(heading);\n\n    const note = document.createElement(\"p\");\n    note.className = \"note\";\n    note.textContent = \"Per ogni domanda mostra le atlete che hanno ricevuto pi\\u00f9 risposte.\";\n    section.appendChild(note);\n\n    const controls = document.createElement(\"div\");\n    controls.className = \"top-questions-controls\";\n    const label = document.createElement(\"label\");\n    label.textContent = \"Numero di atlete per domanda\";\n    const input = document.createElement(\"input\");\n    input.type = \"number\";\n    input.min = \"1\";\n    input.max = String(Math.max(analysis.names.length, 1));\n    input.value = \"3\";\n    label.appendChild(input);\n    controls.appendChild(label);\n    section.appendChild(controls);\n\n    const list = document.createElement(\"div\");\n    list.className = \"top-questions-list\";\n    section.appendChild(list);\n\n    const renderList = (limit) => {\n      list.innerHTML = \"\";\n      const grid = document.createElement(\"div\");\n      grid.className = \"top-questions-grid\";\n      const posGroup = document.createElement(\"div\");\n      const posHeading = document.createElement(\"h3\");\n      posHeading.textContent = \"Domande positive\";\n      posGroup.appendChild(posHeading);\n      const posList = document.createElement(\"div\");\n      posList.className = \"top-questions-list\";\n      analysis.questionStats.posLabels.forEach((question, index) => {\n        posList.appendChild(renderTopRecipientsQuestionCard(\n          question,\n          analysis.names,\n          analysis.questionStats.posReceivedByQuestion,\n          index,\n          limit,\n          \"positive\"\n        ));\n      });\n      posGroup.appendChild(posList);\n      grid.appendChild(posGroup);\n\n      const negGroup = document.createElement(\"div\");\n      const negHeading = document.createElement(\"h3\");\n      negHeading.textContent = \"Domande negative\";\n      negGroup.appendChild(negHeading);\n      const negList = document.createElement(\"div\");\n      negList.className = \"top-questions-list\";\n      analysis.questionStats.negLabels.forEach((question, index) => {\n        negList.appendChild(renderTopRecipientsQuestionCard(\n          question,\n          analysis.names,\n          analysis.questionStats.negReceivedByQuestion,\n          index,\n          limit,\n          \"negative\"\n        ));\n      });\n      negGroup.appendChild(negList);\n      grid.appendChild(negGroup);\n      list.appendChild(grid);\n    };\n\n    const normalizeLimit = () => {\n      const parsed = Number.parseInt(input.value, 10);\n      const safe = Number.isFinite(parsed) ? parsed : 1;\n      const max = Math.max(analysis.names.length, 1);\n      return Math.max(1, Math.min(max, safe));\n    };\n\n    input.addEventListener(\"input\", () => {\n      const limit = normalizeLimit();\n      input.value = String(limit);\n      renderList(limit);\n    });\n\n    renderList(normalizeLimit());\n    return section;\n  }\n\n  function renderTopRecipientsQuestionCard(question, names, matrixByAthlete, questionIndex, limit, tone) {\n    const card = document.createElement(\"div\");\n    card.className = \"response-card top-questions-card\";\n    if (tone === \"positive\") {\n      card.classList.add(\"top-questions-positive\");\n    } else if (tone === \"negative\") {\n      card.classList.add(\"top-questions-negative\");\n    }\n\n    const title = document.createElement(\"h3\");\n    title.textContent = question;\n    card.appendChild(title);\n\n    const items = buildTopRecipientsList(names, matrixByAthlete, questionIndex, limit);\n    if (!items.length) {\n      const empty = document.createElement(\"p\");\n      empty.className = \"note\";\n      empty.textContent = \"Nessuna risposta.\";\n      card.appendChild(empty);\n      return card;\n    }\n\n    const list = document.createElement(\"ol\");\n    list.className = \"top-questions-items\";\n    items.forEach((item) => {\n      const li = document.createElement(\"li\");\n      li.className = \"top-questions-item\";\n      const label = document.createElement(\"span\");\n      label.textContent = item.label;\n      const count = document.createElement(\"span\");\n      count.className = \"top-questions-count\";\n      count.textContent = item.count;\n      li.appendChild(label);\n      li.appendChild(count);\n      list.appendChild(li);\n    });\n    card.appendChild(list);\n    return card;\n  }\n\n  function renderTopQuestionColumn(title, labels, counts = [], limit, tone) {\n    const column = document.createElement(\"div\");\n    column.className = \"top-questions-column\";\n    if (tone === \"positive\") {\n      column.classList.add(\"top-questions-positive\");\n    } else if (tone === \"negative\") {\n      column.classList.add(\"top-questions-negative\");\n    }\n\n    const heading = document.createElement(\"h4\");\n    heading.textContent = title;\n    column.appendChild(heading);\n\n    const items = buildTopQuestionList(labels, counts, limit);\n    if (!items.length) {\n      const empty = document.createElement(\"p\");\n      empty.className = \"note\";\n      empty.textContent = \"Nessuna risposta.\";\n      column.appendChild(empty);\n      return column;\n    }\n\n    const list = document.createElement(\"ol\");\n    list.className = \"top-questions-items\";\n    items.forEach((item) => {\n      const li = document.createElement(\"li\");\n      li.className = \"top-questions-item\";\n      const label = document.createElement(\"span\");\n      label.textContent = item.label;\n      const count = document.createElement(\"span\");\n      count.className = \"top-questions-count\";\n      count.textContent = item.count;\n      li.appendChild(label);\n      li.appendChild(count);\n      list.appendChild(li);\n    });\n    column.appendChild(list);\n\n    return column;\n  }\n\n  function buildTopQuestionList(labels, counts, limit) {\n    return labels\n      .map((label, index) => ({ label, count: counts[index] || 0, index }))\n      .filter((item) => item.count > 0)\n      .sort((a, b) => (b.count === a.count ? a.index - b.index : b.count - a.count))\n      .slice(0, limit);\n  }\n\n  function buildTopRecipientsList(names, matrixByAthlete, questionIndex, limit) {\n    return names\n      .map((name, index) => ({\n        label: name,\n        count: (matrixByAthlete[index] || [])[questionIndex] || 0,\n        index\n      }))\n      .filter((item) => item.count > 0)\n      .sort((a, b) => (b.count === a.count ? a.index - b.index : b.count - a.count))\n      .slice(0, limit);\n  }\n\n  function parseList(text) {\n    return String(text || \"\")\n      .split(/\\r?\\n/)\n      .map((line) => line.trim())\n      .filter(Boolean);\n  }\n\n  function titleCaseName(value) {\n    return String(value || \"\")\n      .trim()\n      .split(/\\s+/)\n      .map((part) => {\n        if (!part) {\n          return \"\";\n        }\n        const lower = part.toLowerCase();\n        return lower.charAt(0).toUpperCase() + lower.slice(1);\n      })\n      .join(\" \");\n  }\n\n  function extractInlineCategory(text, fallback) {\n    const raw = String(text || \"\");\n    const match = raw.match(/^\\s*\\[(Tecniche|Attitudinali|Sociali)\\]\\s*/i);\n    if (!match) {\n      return { category: fallback, question: raw.trim() };\n    }\n    const lowered = match[1].toLowerCase();\n    const category = lowered.startsWith(\"tec\")\n      ? \"Tecniche\"\n      : lowered.startsWith(\"att\")\n        ? \"Attitudinali\"\n        : \"Sociali\";\n    const question = raw.slice(match[0].length).trim();\n    return { category, question: question || raw.trim() };\n  }\n\n  function parseTeamweaveQuestions(text, categories) {\n    const lines = parseList(text);\n    const pos = [];\n    const neg = [];\n    lines.forEach((line, index) => {\n      const current = line.trim();\n      if (!current) {\n        return;\n      }\n      const pairIndex = Math.floor(index / 2);\n      const category = categories[pairIndex] || \"Tecniche\";\n      const question = `[${category}] ${current}`;\n      if (index % 2 === 0) {\n        pos.push(question);\n      } else {\n        neg.push(question);\n      }\n    });\n    return { pos, neg };\n  }\n\n  function buildTeamweaveFormScript(players, questions, options) {\n    const { title, limit } = options;\n    const formTitle = title || \"Test TeamWeave\";\n    const description = buildTeamweaveDescription(limit);\n    const lines = [\n      \"function buildTeamWeaveForm() {\",\n      `  var form = FormApp.create(\"${escapeQuotes(formTitle)}\");`,\n      `  form.setDescription(\"${escapeQuotes(description)}\");`,\n      `  var players = [${players.map((name) => `\"${escapeQuotes(name)}\"`).join(\", \")}];`,\n      \"  var nameItem = form.addListItem();\",\n      \"  nameItem.setTitle(\\\"Seleziona il tuo nome\\\").setChoiceValues(players).setRequired(true);\"\n    ];\n\n    const max = Math.max(questions.pos.length, questions.neg.length);\n    for (let i = 0; i < max; i += 1) {\n      if (questions.pos[i]) {\n        lines.push(\"  var item = form.addCheckboxItem();\");\n        lines.push(`  item.setTitle(\"${escapeQuotes(questions.pos[i])}\").setChoiceValues(players).setRequired(false);`);\n        lines.push(`  item.setValidation(FormApp.createCheckboxValidation().requireSelectAtMost(${limit}).build());`);\n      }\n      if (questions.neg[i]) {\n        lines.push(\"  var item = form.addCheckboxItem();\");\n        lines.push(`  item.setTitle(\"${escapeQuotes(questions.neg[i])}\").setChoiceValues(players).setRequired(false);`);\n        lines.push(`  item.setValidation(FormApp.createCheckboxValidation().requireSelectAtMost(${limit}).build());`);\n      }\n    }\n\n    lines.push(\"  Logger.log(\\\"Form creato: \\\" + form.getEditUrl());\");\n    lines.push(\"}\");\n    return lines.join(\"\\n\");\n  }\n\n  function buildTeamweaveDescription(limit) {\n    const max = Number.isFinite(limit) ? limit : 3;\n    return [\n      \"Questo questionario serve a capire meglio le relazioni, la comunicazione e la coesione all\u2019interno della squadra. Le tue risposte aiuteranno l\u2019allenatore a conoscere il gruppo e a migliorare il clima di squadra, l\u2019organizzazione e la collaborazione tra compagne.\",\n      \"\",\n      \"\u2022 Non ci sono risposte giuste o sbagliate: rispondi con sincerit\u00e0.\",\n      \"\u2022 Nessuno oltre all\u2019allenatore legger\u00e0 le risposte, e non verranno mai condivise con la squadra.\",\n      \"\u2022 Non si pu\u00f2 rispondere con \u201ctutte\u201d: scegli le migliori \" + max + \" (o le prime \" + max + \" che ti vengono in mente).\",\n      \"\u2022 Non vi preoccupate di escludere qualcuno: mettete le prime \" + max + \" persone che vi vengono in mente.\",\n      \"\u2022 Potete mettere da 0 a \" + max + \" risposte. Nelle domande positive \u00e8 importante sforzarsi di mettere \" + max + \" risposte (a meno che proprio non ci siano).\",\n      \"\u2022 Siate oggettive nelle domande puramente tecniche (es. \u201cchi sceglieresti per fare una squadra forte\u201d).\",\n      \"\u2022 Nessuno sapr\u00e0 le vostre risposte oltre all\u2019allenatore: siate sincere.\",\n      \"\u2022 Non potete auto-votarvi.\"\n    ].join(\"\\n\");\n  }\n\n  function initTeamweaveFormGenerator() {\n    if (!twFormTitle || !twFormPlayers || !twFormQuestions) {\n      return;\n    }\n\n    function getCategorySelections() {\n      if (!twFormCategories) {\n        return [];\n      }\n      const selections = [];\n      const rows = Array.from(twFormCategories.querySelectorAll(\".tw-category-row\"));\n      rows.forEach((row) => {\n        const checked = row.querySelector(\"input[type=\\\"radio\\\"]:checked\");\n        if (checked) {\n          selections.push(checked.value);\n        }\n      });\n      return selections;\n    }\n\n    function renderCategorySelectors() {\n      if (!twFormCategories) {\n        return;\n      }\n      const lines = parseList(twFormQuestions.value);\n      const pairCount = Math.ceil(lines.length / 2);\n      const prev = getCategorySelections();\n      const detected = [];\n      for (let i = 0; i < pairCount; i += 1) {\n        const first = extractInlineCategory(lines[i * 2] || \"\", null);\n        const second = extractInlineCategory(lines[i * 2 + 1] || \"\", null);\n        detected.push(first.category || second.category || null);\n      }\n      twFormCategories.innerHTML = \"\";\n      for (let i = 0; i < pairCount; i += 1) {\n        const row = document.createElement(\"div\");\n        row.className = \"tw-category-row\";\n        const label = document.createElement(\"span\");\n        label.className = \"tw-category-label\";\n        const first = lines[i * 2] || \"\";\n        const second = lines[i * 2 + 1] || \"\";\n        if (second) {\n          label.innerHTML = `<strong>${escapeHtml(first)}</strong><span class=\"tw-category-sep\"> / </span>${escapeHtml(second)}`;\n        } else {\n          label.textContent = first || `Domanda ${i * 2 + 1}`;\n        }\n        row.appendChild(label);\n\n        const options = document.createElement(\"div\");\n        options.className = \"tw-category-options\";\n        [\"Tecniche\", \"Attitudinali\", \"Sociali\"].forEach((value) => {\n          const radioLabel = document.createElement(\"label\");\n          radioLabel.className = \"tw-category-option\";\n          const input = document.createElement(\"input\");\n          input.type = \"radio\";\n          input.name = `tw-cat-${i}`;\n          input.value = value;\n          const preferred = prev[i] || detected[i] || \"Tecniche\";\n          if (preferred === value) {\n            input.checked = true;\n          }\n          radioLabel.appendChild(input);\n          radioLabel.appendChild(document.createTextNode(value));\n          options.appendChild(radioLabel);\n        });\n        row.appendChild(options);\n        twFormCategories.appendChild(row);\n      }\n    }\n\n    function renderPreview(questions) {\n      const items = [];\n      questions.pos.forEach((q) => items.push(`\u2705 ${q}`));\n      questions.neg.forEach((q) => items.push(`\u274c ${q}`));\n      if (!items.length) {\n        twFormPreview.innerHTML = \"<p class=\\\"note\\\">Nessuna domanda rilevata.</p>\";\n        return;\n      }\n      twFormPreview.innerHTML = items.map((q) => `<div class=\"question\">${escapeHtml(q)}</div>`).join(\"\");\n    }\n\n    function convert() {\n      const players = parseList(twFormPlayers.value).map(titleCaseName).filter(Boolean);\n      if (!players.length) {\n        twFormOutput.textContent = \"Inserisci almeno una giocatrice.\";\n        twFormPreview.innerHTML = \"\";\n        return;\n      }\n      const lines = parseList(twFormQuestions.value);\n      if (lines.length % 2 !== 0) {\n        twFormOutput.textContent = \"Le domande devono essere in numero pari (positiva + negativa).\";\n        twFormPreview.innerHTML = \"\";\n        return;\n      }\n      const limit = Number(twFormLimit?.value) || 1;\n      const categories = getCategorySelections();\n      const questions = parseTeamweaveQuestions(\n        lines.map((line) => extractInlineCategory(line, \"Tecniche\").question).join(\"\\n\"),\n        categories\n      );\n      latestTeamweaveScript = buildTeamweaveFormScript(players, questions, {\n        title: twFormTitle.value,\n        limit\n      });\n      twFormOutput.textContent = latestTeamweaveScript;\n      renderPreview(questions);\n    }\n\n    function copyScript() {\n      if (!latestTeamweaveScript) {\n        convert();\n      }\n      if (!latestTeamweaveScript) {\n        return;\n      }\n      navigator.clipboard?.writeText(latestTeamweaveScript).then(() => {\n        twFormOutput.textContent = `${latestTeamweaveScript}\\n// Copiato negli appunti.`;\n      }).catch(() => {\n        twFormOutput.textContent = `${latestTeamweaveScript}\\n// Copia non riuscita.`;\n      });\n    }\n\n    function downloadScript() {\n      if (!latestTeamweaveScript) {\n        convert();\n      }\n      if (!latestTeamweaveScript) {\n        return;\n      }\n      const blob = new Blob([latestTeamweaveScript], { type: \"text/plain\" });\n      const url = URL.createObjectURL(blob);\n      const a = document.createElement(\"a\");\n      a.href = url;\n      a.download = \"teamweave-form.gs\";\n      document.body.appendChild(a);\n      a.click();\n      a.remove();\n      URL.revokeObjectURL(url);\n    }\n\n    twFormQuestions.addEventListener(\"input\", renderCategorySelectors);\n    renderCategorySelectors();\n\n    twFormConvert?.addEventListener(\"click\", convert);\n    twFormCopy?.addEventListener(\"click\", copyScript);\n    twFormDownload?.addEventListener(\"click\", downloadScript);\n    twFormDownloadQuestions?.addEventListener(\"click\", () => {\n      const lines = parseList(twFormQuestions.value);\n      if (!lines.length) {\n        twFormOutput.textContent = \"Inserisci almeno una domanda.\";\n        return;\n      }\n      const categories = getCategorySelections();\n      const tagged = lines.map((line, index) => {\n        const pairIndex = Math.floor(index / 2);\n        const category = categories[pairIndex] || \"Tecniche\";\n        const cleaned = extractInlineCategory(line, category).question;\n        return `[${category}] ${cleaned}`;\n      }).join(\"\\n\");\n      const blob = new Blob([tagged], { type: \"text/plain\" });\n      const url = URL.createObjectURL(blob);\n      const a = document.createElement(\"a\");\n      a.href = url;\n      a.download = \"teamweave-domande.txt\";\n      document.body.appendChild(a);\n      a.click();\n      a.remove();\n      URL.revokeObjectURL(url);\n    });\n  }\n\n  function buildResponsesData(analysis) {\n    const nameIndex = findNameColumn(analysis.headers);\n    const firstQuestionIndex = nameIndex + 1;\n    const data = [];\n\n    analysis.rows.forEach((row) => {\n      const name = row[nameIndex] || \"Senza nome\";\n      const groups = [];\n      const counts = { positive: 0, negative: 0 };\n\n      const grouped = {\n        Tecniche: [],\n        Attitudinali: [],\n        Sociali: [],\n        Altro: []\n      };\n\n      analysis.questionMeta.forEach((meta, idx) => {\n        const answer = row[firstQuestionIndex + idx] || \"\";\n        const category = meta.category || \"Altro\";\n        const sentiment = meta.positive ? \"positive\" : \"negative\";\n        grouped[category].push({ question: meta.question, answer, sentiment });\n        if (meta.positive) {\n          counts.positive += answer ? 1 : 0;\n        } else {\n          counts.negative += answer ? 1 : 0;\n        }\n      });\n\n      Object.keys(grouped).forEach((category) => {\n        const items = grouped[category];\n        if (!items.length) {\n          return;\n        }\n        groups.push({\n          title: category,\n          count: items.length,\n          items\n        });\n      });\n\n      data.push({\n        name,\n        groups,\n        counts\n      });\n    });\n\n    return data;\n  }\n\n  function renderNetworkSection({ title, note, names, matrix, strengthMatrix, nodePalette, edgePalette, toggleId }) {\n    const section = document.createElement(\"div\");\n    const heading = document.createElement(\"h2\");\n    heading.textContent = title;\n    section.appendChild(heading);\n\n    const noteEl = document.createElement(\"p\");\n    noteEl.className = \"note\";\n    noteEl.textContent = note;\n    section.appendChild(noteEl);\n\n    const controls = document.createElement(\"div\");\n    controls.className = \"note\";\n    const label = document.createElement(\"label\");\n    label.style.display = \"inline-flex\";\n    label.style.gap = \"0.5rem\";\n    label.style.alignItems = \"center\";\n    const toggle = document.createElement(\"input\");\n    toggle.type = \"checkbox\";\n    toggle.checked = true;\n    toggle.id = toggleId;\n    const text = document.createElement(\"span\");\n    text.textContent = \"Colora i legami in base alla forza\";\n    label.appendChild(toggle);\n    label.appendChild(text);\n    controls.appendChild(label);\n    section.appendChild(controls);\n\n    const legendStats = getNetworkLegendStats(names, matrix, strengthMatrix);\n    const legend = document.createElement(\"div\");\n    legend.className = \"network-legend\";\n    const nodeLegend = document.createElement(\"div\");\n    nodeLegend.className = \"legend-row\";\n    nodeLegend.innerHTML = `<span>Nodi</span><div class=\"legend-bar\" style=\"background: linear-gradient(90deg, ${nodePalette[0]}, ${nodePalette[1]}, ${nodePalette[2]});\"></div><span class=\"legend-range\">${legendStats.minDegree} \u2192 ${legendStats.maxDegree}</span>`;\n    const edgeLegend = document.createElement(\"div\");\n    edgeLegend.className = \"legend-row\";\n    edgeLegend.innerHTML = `<span>Archi</span><div class=\"legend-bar\" style=\"background: linear-gradient(90deg, ${edgePalette[0]}, ${edgePalette[1]});\"></div><span class=\"legend-range\">${legendStats.minStrength} \u2192 ${legendStats.maxStrength}</span>`;\n    legend.appendChild(nodeLegend);\n    legend.appendChild(edgeLegend);\n    section.appendChild(legend);\n\n    const network = document.createElement(\"div\");\n    network.className = \"network\";\n\n    const canvas = document.createElement(\"canvas\");\n    canvas.width = 900;\n    canvas.height = 520;\n    network.appendChild(canvas);\n\n    section.appendChild(network);\n    initNetwork(\n      canvas,\n      names,\n      matrix,\n      strengthMatrix,\n      toggle,\n      {\n        nodePalette,\n        edgePalette\n      }\n    );\n\n    return section;\n  }\n\n  function initNetwork(canvas, names, matrix, strengthMatrix, toggleEl, config) {\n    const ctx = canvas.getContext(\"2d\");\n    const width = canvas.width;\n    const height = canvas.height;\n\n    const state = createNetworkState(names, matrix, strengthMatrix, width, height, config);\n    canvas._networkState = state;\n    drawNetwork(ctx, state);\n    attachNetworkHandlers(canvas);\n    attachNetworkToggle(canvas, toggleEl);\n  }\n\n  function getNetworkLegendStats(names, matrix, strengthMatrix) {\n    const edges = [];\n    const strengths = [];\n    for (let i = 0; i < matrix.length; i += 1) {\n      for (let j = i + 1; j < matrix.length; j += 1) {\n        if (matrix[i][j] > 0) {\n          edges.push([i, j]);\n          strengths.push(strengthMatrix?.[i]?.[j] || 0);\n        }\n      }\n    }\n    const nodeWeights = sumRows(matrix);\n    const minDegree = nodeWeights.length ? Math.min(...nodeWeights) : 0;\n    const maxDegree = nodeWeights.length ? Math.max(...nodeWeights) : 0;\n    const strengthRange = minMaxArray(strengths);\n    return {\n      minDegree,\n      maxDegree,\n      minStrength: formatNumber(strengthRange.min),\n      maxStrength: formatNumber(strengthRange.max)\n    };\n  }\n\n  function createNetworkState(names, matrix, strengthMatrix, width, height, config) {\n    const nodes = names.map((name) => ({ name, x: 0, y: 0 }));\n    const edges = [];\n    const strengths = [];\n    for (let i = 0; i < matrix.length; i += 1) {\n      for (let j = i + 1; j < matrix.length; j += 1) {\n        if (matrix[i][j] > 0) {\n          edges.push([i, j]);\n          strengths.push(strengthMatrix?.[i]?.[j] || 0);\n        }\n      }\n    }\n\n    layoutGraph(nodes, edges, width, height);\n\n    const nodeWeights = sumRows(matrix);\n    const maxWeight = Math.max(...nodeWeights, 1);\n\n    const strengthRange = minMaxArray(strengths);\n    return {\n      nodes,\n      edges,\n      strengths,\n      strengthRange,\n      nodeWeights,\n      maxWeight,\n      width,\n      height,\n      dragIndex: null,\n      showStrength: true,\n      nodePalette: config?.nodePalette || PALETTE.network,\n      edgePalette: config?.edgePalette || [\"#d5d8dc\", \"#1a9850\"]\n    };\n  }\n\n  function drawNetwork(ctx, state) {\n    const { nodes, edges, nodeWeights, maxWeight, width, height, strengths, strengthRange, showStrength, nodePalette, edgePalette } = state;\n    ctx.clearRect(0, 0, width, height);\n    ctx.lineWidth = 1.2;\n\n    edges.forEach(([i, j], idx) => {\n      if (showStrength) {\n        const value = strengths[idx] || 0;\n        const t = strengthRange.min === strengthRange.max ? 0 : (value - strengthRange.min) / (strengthRange.max - strengthRange.min);\n        ctx.strokeStyle = interpolateBi(edgePalette, t);\n      } else {\n        ctx.strokeStyle = \"rgba(0, 0, 0, 0.2)\";\n      }\n      ctx.beginPath();\n      ctx.moveTo(nodes[i].x, nodes[i].y);\n      ctx.lineTo(nodes[j].x, nodes[j].y);\n      ctx.stroke();\n    });\n\n    nodes.forEach((node, idx) => {\n      const weight = nodeWeights[idx] || 0;\n      const radius = 8 + (weight / maxWeight) * 16;\n      const t = maxWeight === 0 ? 0 : weight / maxWeight;\n      const color = interpolateTri(nodePalette, t);\n\n      ctx.beginPath();\n      ctx.fillStyle = color;\n      ctx.strokeStyle = \"#111\";\n      ctx.lineWidth = 1.2;\n      ctx.arc(node.x, node.y, radius, 0, Math.PI * 2);\n      ctx.fill();\n      ctx.stroke();\n\n      ctx.fillStyle = \"#1c201a\";\n      ctx.font = \"12px Trebuchet MS\";\n      ctx.textAlign = \"center\";\n      ctx.fillText(node.name, node.x, node.y - radius - 6);\n    });\n  }\n\n  function attachNetworkHandlers(canvas) {\n    const state = canvas._networkState;\n    if (!state || canvas._networkHandlersAttached) {\n      return;\n    }\n    canvas._networkHandlersAttached = true;\n    const ctx = canvas.getContext(\"2d\");\n\n    const getPointer = (event) => {\n      const rect = canvas.getBoundingClientRect();\n      return {\n        x: ((event.clientX - rect.left) / rect.width) * canvas.width,\n        y: ((event.clientY - rect.top) / rect.height) * canvas.height\n      };\n    };\n\n    const findNodeAt = (x, y) => {\n      for (let i = state.nodes.length - 1; i >= 0; i -= 1) {\n        const node = state.nodes[i];\n        const degree = state.degrees[i];\n        const radius = 8 + (degree / state.maxDegree) * 16;\n        const dist = Math.hypot(node.x - x, node.y - y);\n        if (dist <= radius + 6) {\n          return i;\n        }\n      }\n      return null;\n    };\n\n    const onDown = (event) => {\n      const { x, y } = getPointer(event);\n      const idx = findNodeAt(x, y);\n      if (idx !== null) {\n        state.dragIndex = idx;\n        state.dragOffset = { x: state.nodes[idx].x - x, y: state.nodes[idx].y - y };\n      }\n    };\n\n    const onMove = (event) => {\n      if (state.dragIndex === null) {\n        return;\n      }\n      const { x, y } = getPointer(event);\n      const node = state.nodes[state.dragIndex];\n      node.x = Math.min(state.width - 20, Math.max(20, x + state.dragOffset.x));\n      node.y = Math.min(state.height - 20, Math.max(20, y + state.dragOffset.y));\n      drawNetwork(ctx, state);\n    };\n\n    const onUp = () => {\n      state.dragIndex = null;\n    };\n\n    canvas.addEventListener(\"mousedown\", onDown);\n    canvas.addEventListener(\"mousemove\", onMove);\n    canvas.addEventListener(\"mouseup\", onUp);\n    canvas.addEventListener(\"mouseleave\", onUp);\n  }\n\n  function attachNetworkToggle(canvas, toggleEl) {\n    const state = canvas._networkState;\n    if (!state || !toggleEl) {\n      return;\n    }\n    toggleEl.addEventListener(\"change\", () => {\n      state.showStrength = toggleEl.checked;\n      const ctx = canvas.getContext(\"2d\");\n      drawNetwork(ctx, state);\n    });\n  }\n\n  function layoutGraph(nodes, edges, width, height) {\n    const n = nodes.length;\n    const centerX = width / 2;\n    const centerY = height / 2;\n    const radius = Math.min(width, height) / 3;\n\n    nodes.forEach((node, idx) => {\n      const angle = (Math.PI * 2 * idx) / n;\n      node.x = centerX + Math.cos(angle) * radius;\n      node.y = centerY + Math.sin(angle) * radius;\n    });\n\n    const area = width * height;\n    const k = Math.sqrt(area / n);\n\n    for (let step = 0; step < 250; step += 1) {\n      const disp = nodes.map(() => ({ x: 0, y: 0 }));\n\n      for (let i = 0; i < n; i += 1) {\n        for (let j = i + 1; j < n; j += 1) {\n          const dx = nodes[i].x - nodes[j].x;\n          const dy = nodes[i].y - nodes[j].y;\n          const dist = Math.hypot(dx, dy) || 0.001;\n          const force = (k * k) / dist;\n          const offsetX = (dx / dist) * force;\n          const offsetY = (dy / dist) * force;\n          disp[i].x += offsetX;\n          disp[i].y += offsetY;\n          disp[j].x -= offsetX;\n          disp[j].y -= offsetY;\n        }\n      }\n\n      edges.forEach(([i, j]) => {\n        const dx = nodes[i].x - nodes[j].x;\n        const dy = nodes[i].y - nodes[j].y;\n        const dist = Math.hypot(dx, dy) || 0.001;\n        const force = (dist * dist) / k;\n        const offsetX = (dx / dist) * force;\n        const offsetY = (dy / dist) * force;\n        disp[i].x -= offsetX;\n        disp[i].y -= offsetY;\n        disp[j].x += offsetX;\n        disp[j].y += offsetY;\n      });\n\n      nodes.forEach((node, idx) => {\n        node.x = Math.min(width - 40, Math.max(40, node.x + disp[idx].x * 0.02));\n        node.y = Math.min(height - 40, Math.max(40, node.y + disp[idx].y * 0.02));\n      });\n    }\n  }\n\n  function countReciprociPairs(matrix) {\n    let count = 0;\n    for (let i = 0; i < matrix.length; i += 1) {\n      for (let j = i + 1; j < matrix.length; j += 1) {\n        if (matrix[i][j] > 0) {\n          count += 1;\n        }\n      }\n    }\n    return count;\n  }\n\n  function formatNumber(value) {\n    if (!Number.isFinite(value)) {\n      return \"0\";\n    }\n    return value % 1 === 0 ? value.toString() : value.toFixed(2);\n  }\n\n  function formatPercent(value) {\n    if (!Number.isFinite(value)) {\n      return \"0%\";\n    }\n    const percent = value * 100;\n    const formatted = percent % 1 === 0 ? percent.toString() : percent.toFixed(1);\n    return `${formatted}%`;\n  }\n\n  function downloadCsv(csv) {\n    const blob = new Blob([csv], { type: \"text/csv;charset=utf-8\" });\n    const url = URL.createObjectURL(blob);\n    const link = document.createElement(\"a\");\n    link.href = url;\n    link.download = \"teamweave-data.csv\";\n    document.body.appendChild(link);\n    link.click();\n    link.remove();\n    URL.revokeObjectURL(url);\n  }\n\n  function downloadHtml(html, name) {\n    const safeName = (name || \"teamweave-analisi\").replace(/[^a-z0-9-_]+/gi, \"-\").toLowerCase();\n    const blob = new Blob([html], { type: \"text/html;charset=utf-8\" });\n    const url = URL.createObjectURL(blob);\n    const link = document.createElement(\"a\");\n    link.href = url;\n    link.download = `${safeName || \"teamweave-analisi\"}.html`;\n    document.body.appendChild(link);\n    link.click();\n    link.remove();\n    URL.revokeObjectURL(url);\n  }\n\n  function base64UrlEncode(text) {\n    const bytes = new TextEncoder().encode(text);\n    let binary = \"\";\n    bytes.forEach((b) => {\n      binary += String.fromCharCode(b);\n    });\n    const base64 = btoa(binary);\n    return base64.replace(/\\+/g, \"-\").replace(/\\//g, \"_\").replace(/=+$/, \"\");\n  }\n\n  function base64UrlEncodeBytes(bytes) {\n    let binary = \"\";\n    bytes.forEach((b) => {\n      binary += String.fromCharCode(b);\n    });\n    const base64 = btoa(binary);\n    return base64.replace(/\\+/g, \"-\").replace(/\\//g, \"_\").replace(/=+$/, \"\");\n  }\n\n  function base64UrlDecode(text) {\n    try {\n      const base64 = text.replace(/-/g, \"+\").replace(/_/g, \"/\");\n      const padded = base64.padEnd(base64.length + (4 - (base64.length % 4)) % 4, \"=\");\n      const binary = atob(padded);\n      const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));\n      return new TextDecoder().decode(bytes);\n    } catch (error) {\n      setStatus(\"Link dati non valido.\", true);\n      return null;\n    }\n  }\n\n  function base64UrlDecodeBytes(text) {\n    try {\n      const base64 = text.replace(/-/g, \"+\").replace(/_/g, \"/\");\n      const padded = base64.padEnd(base64.length + (4 - (base64.length % 4)) % 4, \"=\");\n      const binary = atob(padded);\n      return Uint8Array.from(binary, (char) => char.charCodeAt(0));\n    } catch (error) {\n      return null;\n    }\n  }\n\n  function packCsvForLink(text) {\n    const { headers, rows } = parseCsv(text);\n    const nameIndex = findNameColumn(headers);\n    const names = [];\n    const nameMap = new Map();\n\n    const addName = (raw) => {\n      const normalized = normalizeName(raw || \"\");\n      if (!normalized) {\n        return;\n      }\n      if (!nameMap.has(normalized)) {\n        nameMap.set(normalized, names.length);\n        names.push(raw.trim());\n      }\n    };\n\n    rows.forEach((row) => {\n      addName(row[nameIndex] || \"\");\n    });\n\n    rows.forEach((row) => {\n      headers.forEach((_, idx) => {\n        if (idx === nameIndex) {\n          return;\n        }\n        splitNames(row[idx] || \"\").forEach(addName);\n      });\n    });\n\n    const bytes = [];\n    bytes.push(1);\n    writeVarint(headers.length, bytes);\n    headers.forEach((header) => writeString(header, bytes));\n    writeVarint(nameIndex, bytes);\n    writeVarint(names.length, bytes);\n    names.forEach((name) => writeString(name, bytes));\n    writeVarint(rows.length, bytes);\n\n    const questionCount = headers.length - 1;\n    rows.forEach((row) => {\n      const chooserKey = normalizeName(row[nameIndex] || \"\");\n      const chooserIndex = nameMap.has(chooserKey) ? nameMap.get(chooserKey) : names.length;\n      writeVarint(chooserIndex, bytes);\n      let answersWritten = 0;\n      headers.forEach((_, idx) => {\n        if (idx === nameIndex) {\n          return;\n        }\n        const picks = splitNames(row[idx] || \"\");\n        const indices = picks\n          .map((pick) => nameMap.get(normalizeName(pick)))\n          .filter((value) => Number.isInteger(value));\n        writeVarint(indices.length, bytes);\n        indices.forEach((value) => writeVarint(value, bytes));\n        answersWritten += 1;\n      });\n      if (answersWritten < questionCount) {\n        for (let i = answersWritten; i < questionCount; i += 1) {\n          writeVarint(0, bytes);\n        }\n      }\n    });\n\n    return new Uint8Array(bytes);\n  }\n\n  function decodePackedCsv(payload) {\n    if (!payload) {\n      return null;\n    }\n    if (!payload.startsWith(\"pack:\")) {\n      return payload;\n    }\n    try {\n      return unpackCsvFromLink(payload.slice(5));\n    } catch (error) {\n      setStatus(\"Link dati non valido.\", true);\n      return null;\n    }\n  }\n\n  function unpackCsvFromLink(payload) {\n    const data = JSON.parse(payload);\n    if (!data || data.v !== 1 || !Array.isArray(data.headers)) {\n      throw new Error(\"Formato link non valido.\");\n    }\n    const headers = data.headers;\n    const nameIndex = data.nameIndex;\n    const names = data.names || [];\n    const rows = data.rows || [];\n\n    const lines = [];\n    lines.push(headers.map(escapeCsvCell).join(\",\"));\n    rows.forEach((row) => {\n      const chooserIndex = row[0];\n      const answers = Array.isArray(row[1]) ? row[1] : [];\n      const cells = Array(headers.length).fill(\"\");\n      if (Number.isInteger(chooserIndex) && names[chooserIndex]) {\n        cells[nameIndex] = names[chooserIndex];\n      }\n      let answerIndex = 0;\n      for (let i = 0; i < headers.length; i += 1) {\n        if (i === nameIndex) {\n          continue;\n        }\n        const indices = answers[answerIndex] || [];\n        const picks = indices\n          .map((idx) => names[idx])\n          .filter(Boolean);\n        cells[i] = picks.join(\", \");\n        answerIndex += 1;\n      }\n      lines.push(cells.map(escapeCsvCell).join(\",\"));\n    });\n    return lines.join(\"\\n\");\n  }\n\n  function escapeCsvCell(value) {\n    const safe = value == null ? \"\" : String(value);\n    const escaped = safe.replace(/\"/g, '\"\"');\n    if (/[\",\\n\\r]/.test(escaped)) {\n      return `\"${escaped}\"`;\n    }\n    return escaped;\n  }\n\n  function unpackCsvFromBinary(bytes) {\n    if (!bytes || !bytes.length) {\n      return null;\n    }\n    let offset = 0;\n    const version = bytes[offset];\n    offset += 1;\n    if (version !== 1) {\n      throw new Error(\"Formato link non supportato.\");\n    }\n\n    const headerResult = readListOfStrings(bytes, offset);\n    const headers = headerResult.value;\n    offset = headerResult.offset;\n\n    const nameIndexResult = readVarint(bytes, offset);\n    const nameIndex = nameIndexResult.value;\n    offset = nameIndexResult.offset;\n\n    const namesResult = readListOfStrings(bytes, offset);\n    const names = namesResult.value;\n    offset = namesResult.offset;\n\n    const rowCountResult = readVarint(bytes, offset);\n    const rowCount = rowCountResult.value;\n    offset = rowCountResult.offset;\n\n    const questionCount = Math.max(headers.length - 1, 0);\n    const lines = [];\n    lines.push(headers.map(escapeCsvCell).join(\",\"));\n\n    for (let r = 0; r < rowCount; r += 1) {\n      const chooserResult = readVarint(bytes, offset);\n      const chooserIndex = chooserResult.value;\n      offset = chooserResult.offset;\n\n      const answers = [];\n      for (let q = 0; q < questionCount; q += 1) {\n        const countResult = readVarint(bytes, offset);\n        const count = countResult.value;\n        offset = countResult.offset;\n        const indices = [];\n        for (let k = 0; k < count; k += 1) {\n          const idxResult = readVarint(bytes, offset);\n          indices.push(idxResult.value);\n          offset = idxResult.offset;\n        }\n        answers.push(indices);\n      }\n\n      const cells = Array(headers.length).fill(\"\");\n      if (chooserIndex < names.length) {\n        cells[nameIndex] = names[chooserIndex];\n      }\n      let answerIndex = 0;\n      for (let i = 0; i < headers.length; i += 1) {\n        if (i === nameIndex) {\n          continue;\n        }\n        const indices = answers[answerIndex] || [];\n        const picks = indices.map((idx) => names[idx]).filter(Boolean);\n        cells[i] = picks.join(\", \");\n        answerIndex += 1;\n      }\n      lines.push(cells.map(escapeCsvCell).join(\",\"));\n    }\n\n    return lines.join(\"\\n\");\n  }\n\n  function writeVarint(value, bytes) {\n    let current = Math.max(0, Math.floor(value));\n    while (current >= 128) {\n      bytes.push((current % 128) + 128);\n      current = Math.floor(current / 128);\n    }\n    bytes.push(current);\n  }\n\n  function readVarint(bytes, offset) {\n    let value = 0;\n    let shift = 0;\n    let index = offset;\n    while (index < bytes.length) {\n      const byte = bytes[index];\n      index += 1;\n      value += (byte & 0x7f) * (2 ** shift);\n      if ((byte & 0x80) === 0) {\n        return { value, offset: index };\n      }\n      shift += 7;\n    }\n    throw new Error(\"Link dati non valido.\");\n  }\n\n  function writeString(text, bytes) {\n    const encoded = new TextEncoder().encode(text);\n    writeVarint(encoded.length, bytes);\n    encoded.forEach((byte) => bytes.push(byte));\n  }\n\n  function readString(bytes, offset) {\n    const lengthResult = readVarint(bytes, offset);\n    const length = lengthResult.value;\n    const start = lengthResult.offset;\n    const end = start + length;\n    if (end > bytes.length) {\n      throw new Error(\"Link dati non valido.\");\n    }\n    const chunk = bytes.slice(start, end);\n    return { value: new TextDecoder().decode(chunk), offset: end };\n  }\n\n  function readListOfStrings(bytes, offset) {\n    const countResult = readVarint(bytes, offset);\n    const count = countResult.value;\n    let currentOffset = countResult.offset;\n    const values = [];\n    for (let i = 0; i < count; i += 1) {\n      const item = readString(bytes, currentOffset);\n      values.push(item.value);\n      currentOffset = item.offset;\n    }\n    return { value: values, offset: currentOffset };\n  }\n\n  async function encodeLinkPayload(payload) {\n    if (payload instanceof Uint8Array) {\n      if (typeof CompressionStream === \"function\") {\n        const compressed = await gzipCompressBytes(payload);\n        return `gzb.${base64UrlEncodeBytes(compressed)}`;\n      }\n      return `bin.${base64UrlEncodeBytes(payload)}`;\n    }\n    if (typeof CompressionStream === \"function\") {\n      const compressed = await gzipCompress(payload);\n      return `gz.${base64UrlEncodeBytes(compressed)}`;\n    }\n    return `raw.${base64UrlEncode(payload)}`;\n  }\n\n  async function decodeLinkPayload(encoded) {\n    if (encoded.startsWith(\"gzb.\")) {\n      const bytes = base64UrlDecodeBytes(encoded.slice(4));\n      if (!bytes) {\n        setStatus(\"Link dati non valido.\", true);\n        return null;\n      }\n      const decompressed = await gzipDecompressBytes(bytes);\n      if (!decompressed) {\n        return null;\n      }\n      return { kind: \"bytes\", value: decompressed };\n    }\n    if (encoded.startsWith(\"bin.\")) {\n      const bytes = base64UrlDecodeBytes(encoded.slice(4));\n      if (!bytes) {\n        setStatus(\"Link dati non valido.\", true);\n        return null;\n      }\n      return { kind: \"bytes\", value: bytes };\n    }\n    if (encoded.startsWith(\"gz.\")) {\n      const bytes = base64UrlDecodeBytes(encoded.slice(3));\n      if (!bytes) {\n        setStatus(\"Link dati non valido.\", true);\n        return null;\n      }\n      const text = await gzipDecompress(bytes);\n      return text ? { kind: \"text\", value: text } : null;\n    }\n    if (encoded.startsWith(\"raw.\")) {\n      return { kind: \"text\", value: base64UrlDecode(encoded.slice(4)) };\n    }\n    const bytes = base64UrlDecodeBytes(encoded);\n    if (bytes) {\n      const text = await gzipDecompress(bytes, true);\n      if (text) {\n        return { kind: \"text\", value: text };\n      }\n    }\n    const fallback = base64UrlDecode(encoded);\n    return fallback ? { kind: \"text\", value: fallback } : null;\n  }\n\n  async function gzipCompressBytes(bytes) {\n    const stream = new CompressionStream(\"gzip\");\n    const writer = stream.writable.getWriter();\n    writer.write(bytes);\n    writer.close();\n    const buffer = await new Response(stream.readable).arrayBuffer();\n    return new Uint8Array(buffer);\n  }\n\n  async function gzipDecompressBytes(bytes, silentFailure) {\n    if (typeof DecompressionStream !== \"function\") {\n      if (!silentFailure) {\n        setStatus(\"Il browser non supporta la decompressione dei link.\", true);\n      }\n      return null;\n    }\n    try {\n      const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream(\"gzip\"));\n      const buffer = await new Response(stream).arrayBuffer();\n      return new Uint8Array(buffer);\n    } catch (error) {\n      if (!silentFailure) {\n        setStatus(\"Link dati non valido.\", true);\n      }\n      return null;\n    }\n  }\n\n  async function gzipCompress(text) {\n    const encoder = new TextEncoder();\n    return gzipCompressBytes(encoder.encode(text));\n  }\n\n  async function gzipDecompress(bytes, silentFailure) {\n    const decompressed = await gzipDecompressBytes(bytes, silentFailure);\n    if (!decompressed) {\n      return null;\n    }\n    return new TextDecoder().decode(decompressed);\n  }\n\n  async function initializeFromStorage() {\n    const restored = await restoreFromLink();\n    if (restored) {\n      return;\n    }\n    if (restoreDatasets()) {\n      return;\n    }\n  }\n\n  function minMaxArray(values) {\n    if (!values.length) {\n      return { min: 0, max: 0 };\n    }\n    let min = Number.POSITIVE_INFINITY;\n    let max = Number.NEGATIVE_INFINITY;\n    values.forEach((value) => {\n      min = Math.min(min, value);\n      max = Math.max(max, value);\n    });\n    if (!Number.isFinite(min) || !Number.isFinite(max)) {\n      return { min: 0, max: 0 };\n    }\n    return { min, max };\n  }\n\n  function flattenMatrix(matrix) {\n    return matrix.reduce((acc, row) => acc.concat(row), []);\n  }\n\n  function getMinMaxConfig(matrix, rowTotals, colTotals, colorConfig) {\n    const matrixValues = flattenMatrix(matrix);\n    const totalsValues = []\n      .concat(rowTotals || [])\n      .concat(colTotals || []);\n    const combinedValues = matrixValues.concat(totalsValues);\n\n    const matrixRange = minMaxArray(matrixValues);\n    const totalsRange = minMaxArray(totalsValues);\n    const combinedRange = minMaxArray(combinedValues);\n\n    const pickRange = (scope) => {\n      if (scope === \"combined\") {\n        return combinedRange;\n      }\n      if (scope === \"totals\") {\n        return totalsRange;\n      }\n      return matrixRange;\n    };\n\n    return {\n      matrix: pickRange(colorConfig?.matrix?.scope || \"matrix\"),\n      totals: pickRange(colorConfig?.totals?.scope || \"totals\")\n    };\n  }\n\n  function applyHeatmap(cell, value, min, max, palette) {\n    if (!palette) {\n      return;\n    }\n    const t = min === max ? 0 : (value - min) / (max - min);\n    const color = palette.length === 2 ? interpolateBi(palette, t) : interpolateTri(palette, t);\n    cell.style.backgroundColor = color;\n    cell.style.color = textColorFor(color);\n  }\n\n  function applySummaryColors(table, analysis) {\n    const dataRows = Array.from(table.querySelectorAll(\"tbody tr\"));\n    const values = analysis.summaryRows;\n    if (!dataRows.length) {\n      return;\n    }\n\n    const columns = [\n      { index: 1, palette: PALETTE.posMatrix },\n      { index: 2, palette: PALETTE.posMatrix },\n      { index: 3, palette: PALETTE.posMatrix },\n      { index: 4, palette: PALETTE.posMatrix },\n      { index: 5, palette: PALETTE.negMatrix },\n      { index: 6, palette: PALETTE.negMatrix },\n      { index: 7, palette: PALETTE.negMatrix },\n      { index: 8, palette: PALETTE.negMatrix },\n      { index: 9, palette: PALETTE.posMatrix },\n      { index: 10, palette: PALETTE.negMatrix },\n      { index: 11, palette: PALETTE.posMatrix },\n      { index: 12, palette: PALETTE.negMatrix },\n      { index: 13, palette: PALETTE.posMatrix },\n      { index: 14, palette: PALETTE.negMatrix },\n      { index: 15, palette: PALETTE.nonRic },\n      { index: 16, palette: PALETTE.nonRic }\n    ];\n\n    const columnValues = new Map();\n    columns.forEach(({ index }) => {\n      columnValues.set(index, values.map((row) => getSummaryValue(row, index)));\n    });\n\n    const columnRanges = new Map();\n    columnValues.forEach((vals, index) => {\n      columnRanges.set(index, minMaxArray(vals));\n    });\n\n    dataRows.forEach((row, rowIndex) => {\n      const cells = row.querySelectorAll(\"th, td\");\n      columns.forEach(({ index, palette }) => {\n        const cell = cells[index];\n        if (!cell) {\n          return;\n        }\n        const value = getSummaryValue(values[rowIndex], index);\n        const range = columnRanges.get(index);\n        applyHeatmap(cell, value, range.min, range.max, palette);\n      });\n    });\n  }\n\n  function getSummaryValue(row, index) {\n    switch (index) {\n      case 1:\n        return row.positiveReceived.Tecniche;\n      case 2:\n        return row.positiveReceived.Attitudinali;\n      case 3:\n        return row.positiveReceived.Sociali;\n      case 4:\n        return row.positiveReceived.Totali;\n      case 5:\n        return row.negativeReceived.Tecniche;\n      case 6:\n        return row.negativeReceived.Attitudinali;\n      case 7:\n        return row.negativeReceived.Sociali;\n      case 8:\n        return row.negativeReceived.Totali;\n      case 9:\n        return row.positiveGiven;\n      case 10:\n        return row.negativeGiven;\n      case 11:\n        return row.reciprociPos;\n      case 12:\n        return row.reciprociNeg;\n      case 13:\n        return row.forzaLegami;\n      case 14:\n        return row.forzaAntagonismo;\n      case 15:\n        return row.nonRicambiateDate;\n      case 16:\n        return row.nonRicambiateRicevute;\n      default:\n        return 0;\n    }\n  }\n\n  function applyLabelFill(cell, value, mode) {\n    const normalized = value.toLowerCase();\n    if (mode === \"positive\") {\n      if (normalized === \"alta\") {\n        setCellFill(cell, \"#ccffcc\");\n      } else if (normalized === \"media\") {\n        setCellFill(cell, \"#ffffcc\");\n      } else if (normalized === \"bassa\") {\n        setCellFill(cell, \"#ffcccc\");\n      }\n      return;\n    }\n\n    if (normalized === \"alta\") {\n      setCellFill(cell, \"#ffcccc\");\n    } else if (normalized === \"media\") {\n      setCellFill(cell, \"#ffffcc\");\n    } else if (normalized === \"bassa\") {\n      setCellFill(cell, \"#ccffcc\");\n    }\n  }\n\n  function applyEquilibrioFill(cell, value) {\n    const normalized = value.toLowerCase();\n    if (normalized === \"selettiva\") {\n      setCellFill(cell, \"#ffffcc\");\n    } else if (normalized === \"bilanciata\") {\n      setCellFill(cell, \"#ccffcc\");\n    } else if (normalized === \"ignorata\") {\n      setCellFill(cell, \"#ffcccc\");\n    }\n  }\n\n  function applyEtichettaFill(cell, value) {\n    const normalized = value.toLowerCase();\n    const positiveHints = [\"leader positiva\", \"leader silenziosa\", \"leader\", \"stimata\", \"benvoluta\", \"apprezzata\", \"figura di riferimento\"];\n    const neutralHints = [\"neutrale\", \"bilanciata\", \"presenza discreta\"];\n    const negativeHints = [\"controversa\", \"esclusa\", \"invisibile\", \"marginale\", \"trascurata\", \"non considerata\", \"dispersa\"];\n\n    if (positiveHints.some((hint) => normalized.includes(hint))) {\n      setCellFill(cell, \"#ccffcc\");\n      return;\n    }\n    if (negativeHints.some((hint) => normalized.includes(hint))) {\n      setCellFill(cell, \"#ffcccc\");\n      return;\n    }\n    if (neutralHints.some((hint) => normalized.includes(hint))) {\n      setCellFill(cell, \"#ffffcc\");\n    }\n  }\n\n  function setCellFill(cell, color) {\n    cell.style.backgroundColor = color;\n    cell.style.color = \"#1c201a\";\n  }\n\n  function interpolateBi(colors, t) {\n    const [a, b] = colors;\n    return mixHex(a, b, t);\n  }\n\n  function interpolateTri(colors, t) {\n    const [a, b, c] = colors;\n    if (t <= 0.5) {\n      return mixHex(a, b, t / 0.5);\n    }\n    return mixHex(b, c, (t - 0.5) / 0.5);\n  }\n\n  function mixHex(a, b, t) {\n    const ca = hexToRgb(a);\n    const cb = hexToRgb(b);\n    const r = Math.round(ca.r + (cb.r - ca.r) * t);\n    const g = Math.round(ca.g + (cb.g - ca.g) * t);\n    const bVal = Math.round(ca.b + (cb.b - ca.b) * t);\n    return `rgb(${r}, ${g}, ${bVal})`;\n  }\n\n  function hexToRgb(hex) {\n    const clean = hex.replace(\"#\", \"\");\n    const num = parseInt(clean, 16);\n    return {\n      r: (num >> 16) & 255,\n      g: (num >> 8) & 255,\n      b: num & 255\n    };\n  }\n\n  function textColorFor(rgb) {\n    const match = rgb.match(/rgb\\\\((\\\\d+),\\\\s*(\\\\d+),\\\\s*(\\\\d+)\\\\)/);\n    if (!match) {\n      return \"#1c201a\";\n    }\n    const r = Number(match[1]);\n    const g = Number(match[2]);\n    const b = Number(match[3]);\n    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;\n    return luminance > 0.65 ? \"#1c201a\" : \"#fdfcf8\";\n  }\n\n  function extractCode(value) {\n    const trimmed = value.trim();\n    const match = trimmed.match(/#(?:[^=]*=)?(.+)$/);\n    if (match) {\n      return match[1];\n    }\n    if (trimmed.includes(`${LINK_PARAM}=`)) {\n      const parts = trimmed.split(`${LINK_PARAM}=`);\n      return parts[parts.length - 1].trim();\n    }\n    return trimmed;\n  }\n})();";


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
            }
            receivedPosByCat.get(pickKey)[meta.category] += 1;
            const qIndex = posIndexByMeta[idx];
            if (qIndex != null) {
              posReceivedByQuestion[pickIndex][qIndex] += 1;
            }
          } else {
            if (chooserIndex !== pickIndex) {
              negMatrix[chooserIndex][pickIndex] += 1;
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
    if (panels.length === 2) {
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
      "• Non potete auto-votarvi."
    ].join("\n");
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

    twFormConvert?.addEventListener("click", convert);
    twFormCopy?.addEventListener("click", copyScript);
    twFormDownload?.addEventListener("click", downloadScript);
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
