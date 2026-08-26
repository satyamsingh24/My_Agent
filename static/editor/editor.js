/* global pdfjsLib, fabric, jspdf */
(function () {
  "use strict";

  const PAGE_W = 794;
  const PAGE_H = 1123;
  const MARGIN = 50;
  const INNER = PAGE_W - 2 * MARGIN;
  const MIN_TEXT_CHARS = 80;
  const EXPORT_SCALE = 2;
  const MAX_UNDO = 40;
  // Fabric draws the first baseline this far below a text object's top edge
  // when lineHeight is 1, so imported PDF baselines can be placed exactly.
  const BASELINE_RATIO = 0.879;
  const PX_TO_PT = 595 / PAGE_W;
  const DOCX_CDN = "https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js";
  const THEME_KEY = "my-agent-editor-theme";

  const ROLE_WORDS = [
    "developer", "engineer", "analyst", "administrator", "architect",
    "consultant", "manager", "specialist", "designer", "tester", "intern",
    "trainee", "lead", "scientist", "programmer", "devops", "support", "sre",
  ];

  let templateMeta = null;
  let pages = [{ json: null }];
  let currentPage = 0;
  let canvas = null;
  let zoom = 1;
  let dirty = false;
  let cvData = null;
  let importedPlainText = "";
  let activeTemplateId = null;
  let undoStacks = [[]];
  let redoStacks = [[]];
  let suppressHistory = false;

  const $ = (sel) => document.querySelector(sel);
  const statusEl = $("#status");
  const pageLabel = $("#page-label");
  const zoomLabel = $("#zoom-label");
  const templateList = $("#template-list");
  const dialogOverlay = $("#dialog-overlay");
  const dialogText = $("#dialog-text");
  const exportOverlay = $("#export-overlay");
  const previewOverlay = $("#preview-overlay");
  let pendingTemplateId = null;

  function setStatus(msg, kind) {
    statusEl.textContent = msg || "";
    statusEl.className = "status" + (kind ? " " + kind : "");
  }

  function hexColor(h) {
    return "#" + String(h || "000000").replace(/^#/, "");
  }

  function storedTheme() {
    try {
      return window.localStorage.getItem(THEME_KEY);
    } catch (err) {
      return null;
    }
  }

  function rememberTheme(theme) {
    try {
      window.localStorage.setItem(THEME_KEY, theme);
    } catch (err) {
      /* private mode: the choice simply lasts for this page view */
    }
  }

  function setTheme(theme) {
    const dark = theme === "dark";
    document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
    const button = $("#btn-theme");
    button.setAttribute("aria-pressed", dark ? "true" : "false");
    button.title = dark ? "Switch to light mode" : "Switch to dark mode";
    $("#theme-icon").textContent = dark ? "\u2600" : "\u263D";
    $("#theme-label").textContent = dark ? "Light" : "Dark";
  }

  function initTheme() {
    const params = new URLSearchParams(window.location.search);
    const fromApp = (params.get("theme") || "").toLowerCase();
    const prefersDark =
      window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
    // A choice made inside the editor wins, then the theme the app was using,
    // then whatever the browser prefers.
    const theme =
      storedTheme() ||
      (fromApp === "dark" || fromApp === "light" ? fromApp : null) ||
      (prefersDark ? "dark" : "light");
    setTheme(theme);
  }

  function toggleTheme() {
    const next =
      document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark";
    setTheme(next);
    rememberTheme(next);
  }

  function configureBackLink() {
    const params = new URLSearchParams(window.location.search);
    const ret = params.get("return") || "../../";
    const link = $("#back-link");
    link.href = ret;
    link.setAttribute("aria-label", "Back to previous page");
  }

  async function loadTemplates() {
    const res = await fetch("templates.json");
    if (!res.ok) throw new Error("Could not load templates.json");
    templateMeta = await res.json();
    renderTemplateList();
  }

  function templatePreviewSrc(id) {
    return "previews/" + id + ".png";
  }

  function renderTemplateList() {
    templateList.innerHTML = "";
    const entries = Object.entries(templateMeta.templates);
    entries.forEach(([id, tpl]) => {
      const li = document.createElement("li");
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "template-card";
      btn.setAttribute("role", "option");
      btn.setAttribute("aria-selected", id === activeTemplateId ? "true" : "false");
      btn.dataset.templateId = id;
      btn.innerHTML =
        '<img class="thumb" loading="lazy" alt="' + escapeHtml(tpl.label) + ' template preview" ' +
        'src="' + templatePreviewSrc(id) + '" />' +
        '<div class="label">' + escapeHtml(tpl.label) + "</div>" +
        '<div class="desc">' + escapeHtml(tpl.description) + "</div>";
      const thumb = btn.querySelector("img");
      thumb.addEventListener("error", () => {
        thumb.replaceWith(buildSwatch(tpl));
      });
      btn.addEventListener("click", () => openTemplate(id));
      li.appendChild(btn);
      templateList.appendChild(li);
    });
  }

  function buildSwatch(tpl) {
    const wrap = document.createElement("div");
    wrap.className = "thumb";
    wrap.style.background =
      "linear-gradient(180deg, " + hexColor(tpl.primary) + " 0 18%, " +
      hexColor(tpl.heading_fill) + " 18% 30%, #ffffff 30%)";
    return wrap;
  }

  function openTemplate(id) {
    if (cvData || importedPlainText) {
      applyTemplate(id);
      return;
    }
    const tpl = templateMeta.templates[id];
    pendingTemplateId = id;
    $("#preview-image").src = templatePreviewSrc(id);
    $("#preview-image").alt = tpl.label + " template preview";
    $("#preview-title").textContent = tpl.label;
    $("#preview-desc").textContent =
      tpl.description + " Import a selectable-text PDF first, then this design " +
      "is filled with your own details.";
    $("#preview-apply").textContent = "Import PDF";
    previewOverlay.classList.remove("hidden");
    $("#preview-close").focus();
  }

  function closePreview() {
    previewOverlay.classList.add("hidden");
    pendingTemplateId = null;
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function initCanvas() {
    canvas = new fabric.Canvas("cv-canvas", {
      width: PAGE_W,
      height: PAGE_H,
      backgroundColor: "#ffffff",
      preserveObjectStacking: true,
      selection: true,
    });

    canvas.on("object:modified", onCanvasChange);
    canvas.on("object:added", onCanvasChange);
    canvas.on("object:removed", onCanvasChange);
    canvas.on("selection:created", syncToolbarFromSelection);
    canvas.on("selection:updated", syncToolbarFromSelection);
    canvas.on("selection:cleared", syncToolbarFromSelection);

    enableTouchGestures();
    saveHistory();
  }

  function enableTouchGestures() {
    canvas.on("touch:gesture", function (opt) {
      if (opt.e.touches && opt.e.touches.length === 2) {
        if (opt.self.state === "start") {
          opt.self.lastScale = zoom;
        }
        zoom = Math.min(2.5, Math.max(0.35, opt.self.lastScale * opt.self.scale));
        applyZoom();
      }
    });
  }

  function getCurrentCanvasJson() {
    return canvas.toJSON([
      "selectable", "evented", "lockMovementX", "lockMovementY",
      "lockScalingX", "lockScalingY", "lockRotation", "opacity",
      "lineHeight", "dataRole",
    ]);
  }

  function persistCurrentPage() {
    pages[currentPage].json = getCurrentCanvasJson();
  }

  function loadPage(index) {
    if (index < 0 || index >= pages.length) return;
    persistCurrentPage();
    currentPage = index;
    const data = pages[currentPage].json;
    suppressHistory = true;
    canvas.clear();
    canvas.setBackgroundColor("#ffffff", canvas.renderAll.bind(canvas));
    if (data) {
      canvas.loadFromJSON(data, () => {
        canvas.renderAll();
        suppressHistory = false;
        saveHistory();
        updatePageLabel();
      });
    } else {
      suppressHistory = false;
      saveHistory();
      updatePageLabel();
    }
    dirty = pages.some((p) => p.json !== null);
  }

  function updatePageLabel() {
    pageLabel.textContent = "Page " + (currentPage + 1) + " / " + pages.length;
  }

  function applyZoom() {
    // Fabric re-renders vector text at the new scale, so the page stays sharp.
    // A CSS transform would only stretch the existing bitmap and look blurry.
    zoomLabel.textContent = Math.round(zoom * 100) + "%";
    canvas.setZoom(zoom);
    canvas.setDimensions({
      width: Math.round(PAGE_W * zoom),
      height: Math.round(PAGE_H * zoom),
    });
    canvas.calcOffset();
    canvas.requestRenderAll();
  }

  function fitToView() {
    const wrap = $("#canvas-wrap");
    const pad = 44;
    const availW = wrap.clientWidth - pad;
    const top = wrap.getBoundingClientRect().top;
    const availH = Math.max(320, window.innerHeight - top) - pad;
    zoom = Math.min(availW / PAGE_W, availH / PAGE_H, 1.5);
    zoom = Math.max(0.35, zoom);
    applyZoom();
  }

  function onCanvasChange() {
    if (suppressHistory) return;
    dirty = true;
    saveHistory();
  }

  function saveHistory() {
    if (suppressHistory) return;
    const stack = undoStacks[currentPage] || (undoStacks[currentPage] = []);
    stack.push(JSON.stringify(getCurrentCanvasJson()));
    if (stack.length > MAX_UNDO) stack.shift();
    redoStacks[currentPage] = [];
  }

  function undo() {
    const stack = undoStacks[currentPage] || [];
    if (stack.length <= 1) return;
    const current = stack.pop();
    (redoStacks[currentPage] || (redoStacks[currentPage] = [])).push(current);
    const prev = stack[stack.length - 1];
    suppressHistory = true;
    canvas.loadFromJSON(JSON.parse(prev), () => {
      canvas.renderAll();
      suppressHistory = false;
      dirty = true;
    });
  }

  function redo() {
    const stack = redoStacks[currentPage] || [];
    if (!stack.length) return;
    const next = stack.pop();
    (undoStacks[currentPage] || (undoStacks[currentPage] = [])).push(next);
    suppressHistory = true;
    canvas.loadFromJSON(JSON.parse(next), () => {
      canvas.renderAll();
      suppressHistory = false;
      dirty = true;
    });
  }

  function cleanText(text) {
    const bullets = "\u2022\u2023\u2043\u2219\u25aa\u25cf\u25e6\u00b7";
    for (const ch of bullets) text = text.split(ch).join("-");
    text = text.replace(/\u2013/g, "-").replace(/\u2014/g, "-").replace(/\xa0/g, " ");
    const lines = [];
    let blank = false;
    for (const raw of text.split(/\r?\n/)) {
      const line = raw.replace(/[ \t]+/g, " ").trim();
      if (!line) {
        if (lines.length && !blank) lines.push("");
        blank = true;
        continue;
      }
      blank = false;
      lines.push(line);
    }
    return lines.join("\n").trim();
  }

  function mergeWrappedLines(lines) {
    const merged = [];
    lines.forEach((line) => {
      const previous = merged[merged.length - 1];
      const startsBullet = /^[-*\u2022]/.test(line);
      const isHeadingLike = line === line.toUpperCase() && /[A-Z]/.test(line);
      const continues =
        previous &&
        /^[-*\u2022]/.test(previous) &&
        !startsBullet &&
        !isHeadingLike &&
        (/^[a-z(]/.test(line) || !/[.;:]$/.test(previous));
      if (continues) merged[merged.length - 1] = previous + " " + line;
      else merged.push(line);
    });
    return merged;
  }

  function parseCvSections(text) {
    const meta = templateMeta;
    const lines = mergeWrappedLines(cleanText(text).split("\n"));
    const header = [];
    const sections = {};
    const order = [];
    let current = null;
    let dropping = false;

    for (const raw of lines) {
      const line = raw.trim();
      if (!line) continue;
      const key = line.toUpperCase().replace(/:$/, "");
      if (meta.dropHeadings.includes(key)) {
        dropping = true;
        current = null;
        continue;
      }
      const canonical = meta.headingAliases[key];
      if (canonical) {
        dropping = false;
        current = canonical;
        if (!sections[current]) {
          sections[current] = [];
          order.push(current);
        }
        continue;
      }
      if (dropping) continue;
      if (current === null) header.push(line);
      else sections[current].push(line);
    }
    return { header, sections, order };
  }

  function isContactLine(line) {
    const low = line.toLowerCase();
    return (
      line.includes("@") ||
      /\d{7,}/.test(line.replace(/\D/g, "")) ||
      low.includes("linkedin") ||
      low.includes("github") ||
      line.includes("|")
    );
  }

  function extractCvData(text) {
    const { header, sections } = parseCvSections(text);
    const allLines = cleanText(text).split("\n").map((l) => l.trim()).filter(Boolean);
    const candidates = header.length ? header : allLines.slice(0, 8);
    const sectionHeadings = new Set(templateMeta.sectionHeadings);

    const emailMatch = text.match(/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i);
    const phoneMatch = text.match(/(?<!\d)(?:\+?\d[\d ()-]{7,}\d)(?!\d)/);

    let name = candidates.find(
      (line) =>
        !isContactLine(line) &&
        !sectionHeadings.has(line.toUpperCase().replace(/:$/, "")) &&
        !["RESUME", "CURRICULUM VITAE", "CV"].includes(line.toUpperCase()) &&
        line.split(/\s+/).length >= 1 &&
        line.split(/\s+/).length <= 6
    ) || "";

    let role = candidates.find(
      (line) =>
        line !== name &&
        !isContactLine(line) &&
        ROLE_WORDS.some((w) => line.toLowerCase().includes(w)) &&
        line.split(/\s+/).length <= 10
    ) || "";

    const contactParts = candidates.filter(
      (line) => line !== name && line !== role && (isContactLine(line) || line.includes(","))
    );
    const contact = contactParts.join(" | ");

    const joinSection = (key) => (sections[key] || []).join("\n").trim();

    return {
      name,
      role,
      contact,
      email: emailMatch ? emailMatch[0] : "",
      phone: phoneMatch ? phoneMatch[0].replace(/\s+/g, " ").trim() : "",
      summary: joinSection("SUMMARY"),
      skills: joinSection("SKILLS"),
      experience: joinSection("PROFESSIONAL EXPERIENCE"),
      projects: joinSection("PROJECTS"),
      education: joinSection("EDUCATION"),
      qualification: joinSection("QUALIFICATION"),
      certifications: joinSection("CERTIFICATIONS"),
      internships: joinSection("INTERNSHIPS"),
      competencies: joinSection("CORE COMPETENCIES"),
      achievements: joinSection("ACHIEVEMENTS"),
      languages: joinSection("LANGUAGES"),
      additional: joinSection("ADDITIONAL DETAILS"),
      jdMatched: joinSection("JD-MATCHED SKILLS"),
      rawSections: sections,
      rawHeader: header,
    };
  }

  function makeText(text, opts) {
    const o = Object.assign(
      {
        left: MARGIN,
        top: MARGIN,
        fontFamily: "Arial",
        fontSize: 14,
        fill: "#000000",
        width: INNER,
        textAlign: "left",
        lineHeight: 1.2,
        editable: true,
      },
      opts || {}
    );
    const obj = new fabric.Textbox(String(text || ""), o);
    obj.set("dataRole", o.dataRole || "text");
    return obj;
  }

  function makeHeadingBand(text, y, design, ruled) {
    const objects = [];
    const headingText = String(text).toUpperCase();
    const compact = design.compact;
    const fontSize = 13;
    const bandH = fontSize + (ruled ? 6 : 14);

    if (ruled) {
      objects.push(
        makeText(headingText, {
          top: y,
          fontSize,
          fontWeight: "bold",
          fill: hexColor(design.heading_text),
          dataRole: "heading",
        })
      );
      const lineY = y + fontSize + 4;
      objects.push(
        new fabric.Line([MARGIN, lineY, PAGE_W - MARGIN, lineY], {
          stroke: "#000000",
          strokeWidth: 1,
          selectable: true,
          dataRole: "rule",
        })
      );
      return { objects, height: bandH + (compact ? 4 : 8) };
    }

    objects.push(
      new fabric.Rect({
        left: MARGIN,
        top: y,
        width: INNER,
        height: bandH,
        fill: hexColor(design.heading_fill),
        selectable: true,
        dataRole: "heading-bg",
      })
    );
    objects.push(
      makeText(headingText, {
        top: y + 5,
        left: MARGIN + 6,
        width: INNER - 12,
        fontSize,
        fontWeight: "bold",
        fill: hexColor(design.heading_text),
        dataRole: "heading",
      })
    );
    return { objects, height: bandH + (compact ? 6 : 10) };
  }

  function appendSectionLines(objects, lines, y, design, sectionName) {
    const compact = design.compact;
    const prose = templateMeta.proseSections.includes(sectionName);
    const bullets = templateMeta.bulletSections.includes(sectionName);
    const entry = templateMeta.entrySections.includes(sectionName);
    const bodySize = compact ? 11 : 12;
    const entrySize = compact ? 11.5 : 12.5;
    const gap = compact ? 3 : 6;
    let cursor = y;

    for (const raw of lines) {
      if (cursor > PAGE_H - MARGIN - 20) break;
      const line = raw.trim();
      if (!line) {
        cursor += gap;
        continue;
      }
      let content = line;
      let fontSize = bodySize;
      let fontWeight = "normal";
      let fill = design.ruled_headings ? "#000000" : "#2A2E3A";
      let left = MARGIN;
      let width = INNER;

      if (line.startsWith("-") || line.startsWith("*") || bullets) {
        content = "\u2022 " + line.replace(/^[-*\u2022]\s*/, "");
        left = MARGIN + 8;
        width = INNER - 8;
      } else if (entry && !prose) {
        fontSize = entrySize;
        fontWeight = "bold";
        fill = design.ruled_headings ? "#000000" : "#1F2437";
      }

      const tb = makeText(content, {
        top: cursor,
        left,
        width,
        fontSize,
        fontWeight,
        fill,
        lineHeight: compact ? 1.15 : 1.25,
        dataRole: "body",
      });
      objects.push(tb);
      cursor += (tb.height || fontSize * 1.3) + gap;
    }
    return cursor;
  }

  function buildTemplatePages(templateId, data) {
    const design = templateMeta.templates[templateId];
    if (!design) return [];
    const ruled = !!design.ruled_headings;
    const compact = !!design.compact;
    const pageObjects = [[]];
    let objects = pageObjects[0];
    let y = MARGIN;

    function nextPage() {
      objects = [];
      pageObjects.push(objects);
      y = MARGIN;
    }

    if (data.name) {
      const nameObj = makeText(data.name, {
        top: y,
        fontSize: compact ? 22 : 24,
        fontWeight: "bold",
        fill: hexColor(design.primary),
        dataRole: "name",
      });
      objects.push(nameObj);
      y += (nameObj.height || 28) + (compact ? 4 : 8);
    }

    if (data.role) {
      const roleObj = makeText(data.role, {
        top: y,
        fontSize: compact ? 14 : 15,
        fontWeight: "bold",
        fill: hexColor(design.accent),
        dataRole: "role",
      });
      objects.push(roleObj);
      y += (roleObj.height || 18) + (compact ? 3 : 6);
    }

    if (data.contact) {
      const contactObj = makeText(data.contact, {
        top: y,
        fontSize: compact ? 11 : 12,
        fill: "#2A2E3A",
        dataRole: "contact",
      });
      objects.push(contactObj);
      y += (contactObj.height || 14) + (compact ? 8 : 12);
    }

    const sectionMap = {
      SUMMARY: data.summary,
      SKILLS: data.skills,
      "JD-MATCHED SKILLS": data.jdMatched,
      "CORE COMPETENCIES": data.competencies,
      "PROFESSIONAL EXPERIENCE": data.experience,
      INTERNSHIPS: data.internships,
      PROJECTS: data.projects,
      EDUCATION: data.education,
      QUALIFICATION: data.qualification,
      CERTIFICATIONS: data.certifications,
      ACHIEVEMENTS: data.achievements,
      LANGUAGES: data.languages,
      "ADDITIONAL DETAILS": data.additional,
    };

    for (const heading of templateMeta.sectionOrder) {
      const body = sectionMap[heading];
      if (!body) continue;
      const lines = body.split("\n").filter((l) => l.trim());
      if (!lines.length) continue;
      if (y > PAGE_H - MARGIN - 60) nextPage();

      let band = makeHeadingBand(heading, y, design, ruled);
      objects.push.apply(objects, band.objects);
      y += band.height;

      const prose = templateMeta.proseSections.includes(heading);
      const bullets = templateMeta.bulletSections.includes(heading);
      const entry = templateMeta.entrySections.includes(heading);
      const bodySize = compact ? 11 : 12;
      const gap = compact ? 3 : 6;

      for (const raw of lines) {
        const line = raw.trim();
        if (!line) {
          y += gap;
          continue;
        }
        let content = line;
        let left = MARGIN;
        let width = INNER;
        let fontSize = bodySize;
        let fontWeight = "normal";
        let fill = ruled ? "#000000" : "#2A2E3A";
        if (line.startsWith("-") || line.startsWith("*") || bullets) {
          content = "\u2022 " + line.replace(/^[-*\u2022]\s*/, "");
          left += 8;
          width -= 8;
        } else if (entry && !prose) {
          fontSize = compact ? 11.5 : 12.5;
          fontWeight = "bold";
          fill = ruled ? "#000000" : "#1F2437";
        }
        let textObject = makeText(content, {
          top: y,
          left,
          width,
          fontSize,
          fontWeight,
          fill,
          lineHeight: compact ? 1.15 : 1.25,
          dataRole: "body",
        });
        const needed = (textObject.height || fontSize * 1.3) + gap;
        if (y + needed > PAGE_H - MARGIN) {
          nextPage();
          band = makeHeadingBand(heading, y, design, ruled);
          objects.push.apply(objects, band.objects);
          y += band.height;
          textObject = makeText(content, {
            top: y,
            left,
            width,
            fontSize,
            fontWeight,
            fill,
            lineHeight: compact ? 1.15 : 1.25,
            dataRole: "body",
          });
        }
        objects.push(textObject);
        y += (textObject.height || fontSize * 1.3) + gap;
      }
      y += compact ? 4 : 8;
    }

    return pageObjects.filter((page) => page.length);
  }

  function clearCanvasAndAdd(objects) {
    suppressHistory = true;
    canvas.clear();
    canvas.setBackgroundColor("#ffffff", () => {
      objects.forEach((obj) => canvas.add(obj));
      canvas.renderAll();
      suppressHistory = false;
      saveHistory();
      dirty = true;
      persistCurrentPage();
    });
  }

  function confirmReplace(message) {
    if (!dirty) return true;
    return window.confirm(message);
  }

  function applyTemplate(templateId) {
    if (!cvData && !importedPlainText) {
      setStatus("Import a PDF first so template content can be parsed.", "error");
      return;
    }
    if (!confirmReplace(
      "Applying a template rebuilds the canvas layout. Your text facts are kept, " +
      "but positioning and manual edits on this page will be replaced. Continue?"
    )) {
      return;
    }

    if (!cvData && importedPlainText) {
      cvData = extractCvData(importedPlainText);
    }

    activeTemplateId = templateId;
    renderTemplateList();
    const templatePages = buildTemplatePages(templateId, cvData);
    if (!templatePages.length) {
      setStatus("No parsed content to place in the template.", "error");
      return;
    }
    pages = templatePages.map((objects) => ({
      json: {
        version: canvas.version || "5.3.0",
        objects: objects.map((obj) => obj.toObject([
          "selectable", "evented", "dataRole", "lineHeight",
        ])),
        background: "#ffffff",
      },
    }));
    undoStacks = pages.map(() => []);
    redoStacks = pages.map(() => []);
    currentPage = 0;
    suppressHistory = true;
    canvas.loadFromJSON(pages[0].json, () => {
      canvas.renderAll();
      suppressHistory = false;
      saveHistory();
      dirty = true;
      updatePageLabel();
    });
    setStatus("Applied template: " + templateMeta.templates[templateId].label + ".", "success");
  }

  async function importPdf(file) {
    setStatus("Reading PDF...");
    try {
      const buf = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
      const allTextItems = [];

      for (let p = 1; p <= pdf.numPages; p++) {
        const page = await pdf.getPage(p);
        const viewport = page.getViewport({ scale: 1 });
        const content = await page.getTextContent();
        content.items.forEach((item) => {
          if (!("str" in item) || !item.str.trim()) return;
          allTextItems.push({
            page: p - 1,
            str: item.str,
            transform: item.transform,
            itemWidth: item.width,
            viewportW: viewport.width,
            viewportH: viewport.height,
          });
        });
      }

      const pageObjects = groupTextItems(allTextItems);
      importedPlainText = cleanText(
        Object.keys(pageObjects)
          .sort((a, b) => Number(a) - Number(b))
          .map((pageKey) => pageObjects[pageKey].map((item) => item.text).join("\n"))
          .join("\n\n")
      );
      if (importedPlainText.length < MIN_TEXT_CHARS) {
        throw new Error(
          "No readable selectable text was found in this PDF. " +
          "Please upload a selectable-text CV instead of a scanned image PDF."
        );
      }

      cvData = extractCvData(importedPlainText);
      pages = Array.from({ length: pdf.numPages }, () => ({ json: null }));
      undoStacks = pages.map(() => []);
      redoStacks = pages.map(() => []);
      currentPage = 0;

      loadPage(0);
      suppressHistory = true;
      canvas.clear();
      canvas.setBackgroundColor("#ffffff", canvas.renderAll.bind(canvas));

      const newPages = [];

      for (let i = 0; i < pdf.numPages; i++) {
        const objs = (pageObjects[i] || []).map(makeImportedText);
        newPages.push({ json: null, preload: objs });
      }

      pages = newPages.map((p) => ({ json: null }));
      undoStacks = pages.map(() => []);
      redoStacks = pages.map(() => []);

      suppressHistory = true;
      canvas.clear();
      (newPages[0].preload || []).forEach((o) => canvas.add(o));
      canvas.renderAll();
      persistCurrentPage();
      for (let i = 1; i < newPages.length; i++) {
        pages[i].json = {
          version: canvas.version || "5.3.0",
          objects: newPages[i].preload.map((o) => o.toObject([
            "selectable", "evented", "dataRole", "lineHeight",
          ])),
          background: "#ffffff",
        };
      }
      suppressHistory = false;
      saveHistory();
      dirty = false;
      activeTemplateId = null;
      renderTemplateList();
      updatePageLabel();
      setStatus(
        "Imported " + pdf.numPages + " page(s) with selectable text. Choose a template or edit freely.",
        "success"
      );
    } catch (err) {
      console.error(err);
      setStatus(err.message || "PDF import failed.", "error");
    }
  }

  function groupTextItems(items) {
    const byPage = {};
    items.forEach((item) => {
      const scaleX = PAGE_W / item.viewportW;
      const scaleY = PAGE_H / item.viewportH;
      const tx = item.transform[4] * scaleX;
      const baseline = PAGE_H - item.transform[5] * scaleY;
      const fontSize = Math.abs(item.transform[3] || item.transform[0]) * scaleY || 12;
      const measured = (item.itemWidth || 0) * scaleX;
      const page = item.page;
      if (!byPage[page]) byPage[page] = [];
      byPage[page].push({
        x: tx,
        y: baseline,
        fontSize,
        text: item.str,
        width: measured || Math.max(20, item.str.length * fontSize * 0.5),
      });
    });

    const result = {};
    Object.keys(byPage).forEach((pageKey) => {
      const rows = byPage[pageKey].filter((row) => row.text.trim());
      rows.sort((a, b) => a.y - b.y || a.x - b.x);
      const merged = [];
      rows.forEach((row) => {
        const last = merged[merged.length - 1];
        const sameLine = last && Math.abs(last.y - row.y) <= Math.max(2, row.fontSize * 0.35);
        const gap = last ? row.x - (last.x + last.width) : Infinity;
        // Wide gaps usually mean a second column (dates, locations); keeping
        // those as separate objects preserves the original visual position.
        if (sameLine && gap <= row.fontSize * 1.6) {
          const joiner = gap > row.fontSize * 0.18 && !/\s$/.test(last.text) ? " " : "";
          last.text += joiner + row.text;
          last.width = row.x + row.width - last.x;
          last.fontSize = Math.max(last.fontSize, row.fontSize);
        } else {
          merged.push(Object.assign({}, row));
        }
      });
      result[pageKey] = merged.map((row) => ({
        x: row.x,
        y: row.y,
        fontSize: row.fontSize,
        width: row.width,
        text: row.text.replace(/\s+/g, " ").trim(),
      }));
    });
    return result;
  }

  function makeImportedText(item) {
    const fontSize = Math.max(6, Math.min(40, item.fontSize));
    const obj = new fabric.IText(item.text, {
      left: item.x,
      top: item.y - fontSize * BASELINE_RATIO,
      fontFamily: "Arial",
      fontSize,
      lineHeight: 1,
      fill: "#111827",
      objectCaching: false,
    });
    obj.set("dataRole", "imported");
    // Nudge the width back to the source metrics so column alignment survives
    // the substitution of Arial for the original embedded font.
    const natural = obj.width || 0;
    if (natural > 0 && item.width > 0) {
      const ratio = item.width / natural;
      if (ratio > 0.7 && ratio < 1.35) obj.set("scaleX", ratio);
    }
    return obj;
  }

  function getActiveObject() {
    return canvas.getActiveObject();
  }

  function addTextObject() {
    const obj = makeText("New text", {
      left: MARGIN + 20,
      top: MARGIN + 20,
      fontSize: Number($("#font-size").value) || 14,
      fill: $("#text-color").value,
      fontFamily: $("#font-family").value,
    });
    canvas.add(obj);
    canvas.setActiveObject(obj);
    canvas.requestRenderAll();
  }

  function deleteSelected() {
    const active = canvas.getActiveObjects();
    if (!active.length) return;
    active.forEach((obj) => canvas.remove(obj));
    canvas.discardActiveObject();
    canvas.requestRenderAll();
  }

  function duplicateSelected() {
    const active = getActiveObject();
    if (!active) return;
    active.clone((cloned) => {
      cloned.set({ left: (active.left || 0) + 16, top: (active.top || 0) + 16 });
      canvas.add(cloned);
      canvas.setActiveObject(cloned);
      canvas.requestRenderAll();
    });
  }

  function editSelectedText() {
    const active = getActiveObject();
    if (!active || active.type !== "textbox" && active.type !== "i-text" && active.type !== "text") {
      setStatus("Select a text object to edit.", "error");
      return;
    }
    dialogText.value = active.text || "";
    dialogOverlay.classList.remove("hidden");
    dialogText.focus();
    dialogOverlay.dataset.targetId = active.__uid || String(canvas.getObjects().indexOf(active));
  }

  function closeDialog(apply) {
    if (apply) {
      const active = getActiveObject();
      if (active && (active.type === "textbox" || active.type === "i-text" || active.type === "text")) {
        active.set("text", dialogText.value);
        canvas.requestRenderAll();
      }
    }
    dialogOverlay.classList.add("hidden");
  }

  function syncToolbarFromSelection() {
    const active = getActiveObject();
    if (!active) return;
    if (active.type === "textbox" || active.type === "i-text" || active.type === "text") {
      $("#font-family").value = active.fontFamily || "Arial";
      $("#font-size").value = Math.round(active.fontSize || 14);
      $("#text-color").value = rgbToHex(active.fill || "#000000");
      $("#text-align").value = active.textAlign || "left";
      $("#line-height").value = active.lineHeight || 1.2;
      $("#opacity").value = active.opacity != null ? active.opacity : 1;
      $("#btn-bold").setAttribute("aria-pressed", active.fontWeight === "bold" ? "true" : "false");
      $("#btn-italic").setAttribute("aria-pressed", active.fontStyle === "italic" ? "true" : "false");
    }
    if (active.lockMovementX && active.lockMovementY) {
      $("#btn-lock").setAttribute("aria-pressed", "true");
    } else {
      $("#btn-lock").setAttribute("aria-pressed", "false");
    }
  }

  function rgbToHex(color) {
    if (typeof color === "string" && color.startsWith("#")) {
      return color.length === 7 ? color : "#000000";
    }
    return "#000000";
  }

  function applyStyleToSelection(props) {
    const active = getActiveObject();
    if (!active) return;
    if (active.type === "activeSelection") {
      active.getObjects().forEach((obj) => obj.set(props));
    } else {
      active.set(props);
    }
    canvas.requestRenderAll();
  }

  function addRectangle() {
    const rect = new fabric.Rect({
      left: MARGIN + 30,
      top: MARGIN + 30,
      width: 160,
      height: 80,
      fill: "transparent",
      stroke: "#475569",
      strokeWidth: 1,
      dataRole: "shape",
    });
    canvas.add(rect);
    canvas.setActiveObject(rect);
    canvas.requestRenderAll();
  }

  function addLineShape() {
    const line = new fabric.Line([MARGIN, MARGIN + 60, MARGIN + 200, MARGIN + 60], {
      stroke: "#000000",
      strokeWidth: 1,
      dataRole: "shape",
    });
    canvas.add(line);
    canvas.setActiveObject(line);
    canvas.requestRenderAll();
  }

  function uploadImage(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
      fabric.Image.fromURL(e.target.result, (img) => {
        const maxW = INNER * 0.45;
        if (img.width > maxW) img.scaleToWidth(maxW);
        img.set({
          left: PAGE_W - MARGIN - (img.getScaledWidth() || 120),
          top: MARGIN,
          dataRole: "image",
        });
        canvas.add(img);
        canvas.setActiveObject(img);
        canvas.requestRenderAll();
      }, { crossOrigin: "anonymous" });
    };
    reader.readAsDataURL(file);
  }

  function layerStep(direction) {
    const active = getActiveObject();
    if (!active) return;
    if (direction > 0) active.bringForward();
    else active.sendBackwards();
    canvas.requestRenderAll();
  }

  function toggleLock() {
    const active = getActiveObject();
    if (!active) return;
    const locked = !(active.lockMovementX && active.lockMovementY);
    active.set({
      lockMovementX: locked,
      lockMovementY: locked,
      lockScalingX: locked,
      lockScalingY: locked,
      lockRotation: locked,
      selectable: true,
      evented: true,
    });
    $("#btn-lock").setAttribute("aria-pressed", locked ? "true" : "false");
    canvas.requestRenderAll();
  }

  function clearPage() {
    if (!confirmReplace("Clear all objects on this page?")) return;
    suppressHistory = true;
    canvas.clear();
    canvas.setBackgroundColor("#ffffff", () => {
      suppressHistory = false;
      saveHistory();
      dirty = true;
      persistCurrentPage();
      setStatus("Page cleared.", "success");
    });
  }

  function resetAll() {
    if (!confirmReplace("Reset all pages and remove imported content?")) return;
    pages = [{ json: null }];
    undoStacks = [[]];
    redoStacks = [[]];
    currentPage = 0;
    cvData = null;
    importedPlainText = "";
    activeTemplateId = null;
    dirty = false;
    renderTemplateList();
    suppressHistory = true;
    canvas.clear();
    canvas.setBackgroundColor("#ffffff", () => {
      suppressHistory = false;
      saveHistory();
      updatePageLabel();
      setStatus("Editor reset.", "success");
    });
  }

  function addPage() {
    persistCurrentPage();
    pages.push({ json: null });
    undoStacks.push([]);
    redoStacks.push([]);
    loadPage(pages.length - 1);
    suppressHistory = true;
    canvas.clear();
    canvas.setBackgroundColor("#ffffff", () => {
      suppressHistory = false;
      saveHistory();
      setStatus("Blank page added.", "success");
    });
  }

  function removePage() {
    if (pages.length <= 1) {
      setStatus("At least one page is required.", "error");
      return;
    }
    if (!window.confirm("Remove the current page?")) return;
    pages.splice(currentPage, 1);
    undoStacks.splice(currentPage, 1);
    redoStacks.splice(currentPage, 1);
    loadPage(Math.min(currentPage, pages.length - 1));
    setStatus("Page removed.", "success");
  }

  function loadForExport(json) {
    return new Promise((resolve) => {
      suppressHistory = true;
      canvas.loadFromJSON(json || { objects: [], background: "#ffffff" }, () => {
        canvas.renderAll();
        resolve();
      });
    });
  }

  async function exportPdf() {
    setStatus("Generating high-resolution PDF...");
    persistCurrentPage();
    const savedPage = currentPage;
    const savedJson = pages[savedPage].json;
    canvas.setZoom(1);
    canvas.setDimensions({ width: PAGE_W, height: PAGE_H });

    try {
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4" });
      const pdfW = pdf.internal.pageSize.getWidth();
      const pdfH = pdf.internal.pageSize.getHeight();

      for (let i = 0; i < pages.length; i++) {
        if (i > 0) pdf.addPage();
        await loadForExport(pages[i].json);
        const dataUrl = canvas.toDataURL({
          format: "png",
          multiplier: EXPORT_SCALE,
          enableRetinaScaling: false,
        });
        pdf.addImage(dataUrl, "PNG", 0, 0, pdfW, pdfH, undefined, "FAST");
      }

      pdf.save("cv-edited.pdf");
      setStatus("Exported cv-edited.pdf (" + pages.length + " page(s)).", "success");
    } catch (err) {
      console.error(err);
      setStatus("PDF export failed: " + (err.message || "unknown error"), "error");
    } finally {
      await loadForExport(savedJson);
      currentPage = savedPage;
      suppressHistory = false;
      applyZoom();
      updatePageLabel();
    }
  }

  function loadScriptOnce(src) {
    return new Promise((resolve, reject) => {
      const existing = document.querySelector('script[data-src="' + src + '"]');
      if (existing) {
        if (existing.dataset.loaded === "1") resolve();
        else existing.addEventListener("load", () => resolve());
        existing.addEventListener("error", () => reject(new Error("Script failed")));
        return;
      }
      const el = document.createElement("script");
      el.src = src;
      el.crossOrigin = "anonymous";
      el.dataset.src = src;
      el.addEventListener("load", () => {
        el.dataset.loaded = "1";
        resolve();
      });
      el.addEventListener("error", () =>
        reject(new Error("Could not load the Word export library. Check your internet connection."))
      );
      document.head.appendChild(el);
    });
  }

  function textBlocksForPage(json) {
    const objects = (json && json.objects) || [];
    const blocks = objects
      .filter(
        (obj) =>
          (obj.type === "textbox" || obj.type === "i-text" || obj.type === "text") &&
          String(obj.text || "").trim()
      )
      .map((obj) => ({
        text: String(obj.text).replace(/\s+/g, " ").trim(),
        top: obj.top || 0,
        left: obj.left || 0,
        size: (obj.fontSize || 12) * (obj.scaleY || 1),
        bold: obj.fontWeight === "bold" || Number(obj.fontWeight) >= 600,
      }))
      .sort((a, b) => a.top - b.top || a.left - b.left);

    const lines = [];
    blocks.forEach((block) => {
      const last = lines[lines.length - 1];
      if (last && Math.abs(last.top - block.top) <= Math.max(3, block.size * 0.4)) {
        last.text += " " + block.text;
        last.size = Math.max(last.size, block.size);
        last.bold = last.bold || block.bold;
      } else {
        lines.push(Object.assign({}, block));
      }
    });
    return lines;
  }

  function buildDocxParagraphs(docx, firstPage) {
    return function (line, index) {
      const { Paragraph, TextRun, HeadingLevel } = docx;
      const bullet = /^[\u2022*-]\s+/.test(line.text);
      const text = bullet ? line.text.replace(/^[\u2022*-]\s+/, "") : line.text;
      const isHeading =
        !bullet && line.text.length <= 45 && line.text === line.text.toUpperCase() &&
        /[A-Z]/.test(line.text);
      const halfPoints = Math.max(16, Math.min(56, Math.round(line.size * PX_TO_PT * 2)));
      const options = {
        children: [new TextRun({ text, bold: line.bold || isHeading, size: halfPoints })],
        spacing: { after: bullet ? 40 : 100 },
      };
      if (bullet) options.bullet = { level: 0 };
      if (line.size >= 18) options.heading = HeadingLevel.HEADING_1;
      else if (isHeading) options.heading = HeadingLevel.HEADING_2;
      if (!firstPage && index === 0) options.pageBreakBefore = true;
      return new Paragraph(options);
    };
  }

  async function exportDocx() {
    setStatus("Preparing Word document...");
    persistCurrentPage();
    try {
      await loadScriptOnce(DOCX_CDN);
      const docx = window.docx;
      if (!docx) throw new Error("Word export library did not load.");

      const children = [];
      pages.forEach((page, pageIndex) => {
        const lines = textBlocksForPage(page.json);
        lines.map(buildDocxParagraphs(docx, pageIndex === 0)).forEach((p) => children.push(p));
      });

      if (!children.length) {
        setStatus("There is no text on the canvas to place in a Word file.", "error");
        return;
      }

      const doc = new docx.Document({
        styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
        sections: [{ properties: {}, children }],
      });
      const blob = await docx.Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "cv-edited.docx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      setStatus(
        "Exported cv-edited.docx as single-column text. Review spacing in Word before sending.",
        "success"
      );
    } catch (err) {
      console.error(err);
      setStatus("Word export failed: " + (err.message || "unknown error"), "error");
    }
  }

  function openExportDialog() {
    exportOverlay.classList.remove("hidden");
    const checked = exportOverlay.querySelector('input[name="export-format"]:checked');
    if (checked) checked.focus();
  }

  function closeExportDialog() {
    exportOverlay.classList.add("hidden");
  }

  async function runExport() {
    const choice = exportOverlay.querySelector('input[name="export-format"]:checked');
    const format = choice ? choice.value : "pdf";
    closeExportDialog();
    if (format === "docx") await exportDocx();
    else await exportPdf();
  }

  function bindEvents() {
    $("#pdf-import").addEventListener("change", (e) => {
      const file = e.target.files && e.target.files[0];
      if (file) importPdf(file);
      e.target.value = "";
    });

    $("#btn-theme").addEventListener("click", toggleTheme);
    $("#btn-export").addEventListener("click", openExportDialog);
    $("#export-cancel").addEventListener("click", closeExportDialog);
    $("#export-confirm").addEventListener("click", runExport);
    exportOverlay.addEventListener("click", (e) => {
      if (e.target === exportOverlay) closeExportDialog();
    });
    $("#preview-close").addEventListener("click", closePreview);
    $("#preview-apply").addEventListener("click", () => {
      const id = pendingTemplateId;
      closePreview();
      if (id && (cvData || importedPlainText)) applyTemplate(id);
      else $("#pdf-import").click();
    });
    previewOverlay.addEventListener("click", (e) => {
      if (e.target === previewOverlay) closePreview();
    });
    $("#btn-add-text").addEventListener("click", addTextObject);
    $("#btn-edit-text").addEventListener("click", editSelectedText);
    $("#btn-delete").addEventListener("click", deleteSelected);
    $("#btn-duplicate").addEventListener("click", duplicateSelected);
    $("#btn-rect").addEventListener("click", addRectangle);
    $("#btn-line").addEventListener("click", addLineShape);
    $("#btn-forward").addEventListener("click", () => layerStep(1));
    $("#btn-backward").addEventListener("click", () => layerStep(-1));
    $("#btn-lock").addEventListener("click", toggleLock);
    $("#btn-undo").addEventListener("click", undo);
    $("#btn-redo").addEventListener("click", redo);
    $("#btn-clear").addEventListener("click", clearPage);
    $("#btn-reset").addEventListener("click", resetAll);
    $("#btn-zoom-in").addEventListener("click", () => { zoom = Math.min(2.5, zoom + 0.1); applyZoom(); });
    $("#btn-zoom-out").addEventListener("click", () => { zoom = Math.max(0.35, zoom - 0.1); applyZoom(); });
    $("#btn-fit").addEventListener("click", fitToView);
    $("#btn-prev-page").addEventListener("click", () => loadPage(currentPage - 1));
    $("#btn-next-page").addEventListener("click", () => loadPage(currentPage + 1));
    $("#btn-add-page").addEventListener("click", addPage);
    $("#btn-remove-page").addEventListener("click", removePage);

    $("#font-family").addEventListener("change", (e) => {
      applyStyleToSelection({ fontFamily: e.target.value });
    });
    $("#font-size").addEventListener("change", (e) => {
      applyStyleToSelection({ fontSize: Number(e.target.value) || 14 });
    });
    $("#text-color").addEventListener("input", (e) => {
      applyStyleToSelection({ fill: e.target.value });
    });
    $("#text-align").addEventListener("change", (e) => {
      applyStyleToSelection({ textAlign: e.target.value });
    });
    $("#line-height").addEventListener("change", (e) => {
      applyStyleToSelection({ lineHeight: Number(e.target.value) || 1.2 });
    });
    $("#opacity").addEventListener("input", (e) => {
      applyStyleToSelection({ opacity: Number(e.target.value) });
    });
    $("#btn-bold").addEventListener("click", () => {
      const active = getActiveObject();
      const next = active && active.fontWeight === "bold" ? "normal" : "bold";
      applyStyleToSelection({ fontWeight: next });
      $("#btn-bold").setAttribute("aria-pressed", next === "bold" ? "true" : "false");
    });
    $("#btn-italic").addEventListener("click", () => {
      const active = getActiveObject();
      const next = active && active.fontStyle === "italic" ? "normal" : "italic";
      applyStyleToSelection({ fontStyle: next });
      $("#btn-italic").setAttribute("aria-pressed", next === "italic" ? "true" : "false");
    });

    $("#image-upload").addEventListener("change", (e) => {
      const file = e.target.files && e.target.files[0];
      uploadImage(file);
      e.target.value = "";
    });

    $("#dialog-cancel").addEventListener("click", () => closeDialog(false));
    $("#dialog-ok").addEventListener("click", () => closeDialog(true));
    dialogOverlay.addEventListener("click", (e) => {
      if (e.target === dialogOverlay) closeDialog(false);
    });

    document.addEventListener("keydown", (e) => {
      const mod = e.ctrlKey || e.metaKey;
      if (mod && e.key === "z") { e.preventDefault(); undo(); }
      if (mod && e.key === "y") { e.preventDefault(); redo(); }
      if (mod && e.key === "d") { e.preventDefault(); duplicateSelected(); }
      if (mod && e.key === "b") { e.preventDefault(); $("#btn-bold").click(); }
      if (mod && e.key === "i") { e.preventDefault(); $("#btn-italic").click(); }
      if (e.key === "Delete" || e.key === "Backspace") {
        const tag = document.activeElement && document.activeElement.tagName;
        if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
        if (getActiveObject()) {
          e.preventDefault();
          deleteSelected();
        }
      }
      if (e.key === "t" && !mod && document.activeElement === document.body) {
        addTextObject();
      }
      if (e.key === "Escape") {
        closeDialog(false);
        closeExportDialog();
        closePreview();
      }
    });

    window.addEventListener("resize", fitToView);
    window.addEventListener("beforeunload", (event) => {
      if (!dirty) return;
      event.preventDefault();
      event.returnValue = "";
    });
  }

  async function boot() {
    initTheme();
    configureBackLink();
    try {
      await loadTemplates();
    } catch (err) {
      setStatus(err.message, "error");
      return;
    }
    initCanvas();
    bindEvents();
    fitToView();
    updatePageLabel();
    setStatus("Import a selectable-text PDF to begin, or add objects manually.");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
