"use strict";

function formatDate(d) {
  if (!d) return "—";
  var parts = d.split("-");
  if (parts.length !== 3) return d;
  return parts[2] + "/" + parts[1] + "/" + parts[0];
}

(function () {
  if (window._SOP_SYS_LOADED_) return;
  window._SOP_SYS_LOADED_ = true;

  // ---------- CONFIG ----------
  var DEPARTMENTS = ["All", "Finance", "Procurement", "BTU", "Marketing", "Human Resources", "IT", "Sales", "Warehousing"];
  var CATEGORIES = ["All", "Process", "Guideline", "Form", "Procedure", "Checklist", "Records", "Doc. Approval"];
  var STATUSES = ["All", "Planning", "In Progress", "Active", "Completed", "Done", "Pending", "Archived", "Overdue"];

  // =========================================================
  // FIREBASE SETUP
  // PUT YOUR REAL FIREBASE CONFIG HERE
  // =========================================================
  var firebaseConfig = {
    apiKey: "YOUR_API_KEY",
    authDomain: "YOUR_PROJECT.firebaseapp.com",
    projectId: "YOUR_PROJECT_ID",
    storageBucket: "YOUR_PROJECT.firebasestorage.app",
    messagingSenderId: "YOUR_SENDER_ID",
    appId: "YOUR_APP_ID"
  };

  if (!firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);
  }

  var fdb = firebase.firestore();
  var fstorage = firebase.storage();
  var sopsCollection = fdb.collection("sops");

  // ---------- DB ----------
  var db = { sops: [] };

  async function loadDB() {
    try {
      var snap = await sopsCollection.get();
      var rows = [];

      snap.forEach(function (doc) {
        var data = doc.data() || {};
        if (!data.id) data.id = doc.id;
        rows.push(data);
      });

      db = { sops: rows };
      normalizeDB();
      return db;
    } catch (e) {
      console.error("Firestore load failed:", e);
      db = { sops: [] };
      return db;
    }
  }

  async function saveSopToCloud(sop) {
    if (!sop || !sop.id) return;
    await sopsCollection.doc(sop.id).set(sop);
  }

  async function deleteSopFromCloud(sopId) {
    if (!sopId) return;
    await sopsCollection.doc(sopId).delete();
  }

  function saveDB() {
    // no localStorage anymore
    // keep function so the rest of your app still works without breaking
  }

  // ---------- FILE STORAGE : FIREBASE STORAGE ----------
  async function saveFileBlob(file) {
    try {
      var cleanName = (file.name || "file").replace(/[^\w.\-]+/g, "_");
      var path = "sop_files/" + Date.now() + "_" + Math.random().toString(36).slice(2) + "_" + cleanName;

      var storageRef = fstorage.ref().child(path);
      var uploadTask = await storageRef.put(file);
      var downloadURL = await uploadTask.ref.getDownloadURL();

      return {
        id: path,
        name: file.name,
        type: file.type || "",
        url: downloadURL
      };
    } catch (err) {
      console.error("Firebase file save failed:", err);
      throw err;
    }
  }

  async function getFileBlob(fileId) {
    try {
      if (!fileId) return null;
      var url = await fstorage.ref().child(fileId).getDownloadURL();
      return {
        id: fileId,
        url: url
      };
    } catch (err) {
      console.error("Firebase file read failed:", err);
      return null;
    }
  }

  function openStoredFile(fileMeta) {
    if (!fileMeta) {
      alert("No file found.");
      return;
    }

    if (fileMeta.url) {
      window.open(fileMeta.url, "_blank");
      return;
    }

    if (!fileMeta.id) {
      alert("No file found.");
      return;
    }

    getFileBlob(fileMeta.id).then(function (record) {
      if (!record || !record.url) {
        alert("Stored file not found.");
        return;
      }
      window.open(record.url, "_blank");
    }).catch(function () {
      alert("Could not open stored file.");
    });
  }

  // ---------- HELPERS ----------
  function uid() {
    return Math.random().toString(36).slice(2) + Date.now().toString(36);
  }

  function today() {
    var d = new Date();
    var mm = String(d.getMonth() + 1).padStart(2, "0");
    var dd = String(d.getDate()).padStart(2, "0");
    return d.getFullYear() + "-" + mm + "-" + dd;
  }

  function escapeHtml(s) {
    var v = String(s == null ? "" : s);
    return v.replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
  }

  function isUrl(s) {
    return /^https?:\/\//i.test((s || "").trim());
  }

  function badgeClass(status) {
    var s = (status || "").toLowerCase();
    if (s.indexOf("done") >= 0 || s.indexOf("completed") >= 0) return "done";
    if (s.indexOf("active") >= 0) return "active";
    if (s.indexOf("progress") >= 0) return "progress";
    if (s.indexOf("planning") >= 0) return "planning";
    if (s.indexOf("overdue") >= 0) return "pending";
    if (s.indexOf("archived") >= 0) return "done";
    return "pending";
  }

  function nextVersion(latest) {
    if (!latest || !/^v\d+(\.\d+)?$/i.test(latest)) return "V1.0";
    var n = latest.slice(1);
    var parts = n.split(".");
    var major = parseInt(parts[0] || "1", 10);
    var minor = parseInt(parts[1] || "0", 10);
    return "V" + major + "." + (minor + 1);
  }

  function byId(id) {
    return document.getElementById(id);
  }

  function fillSelect(sel, arr) {
    if (!sel) return;
    sel.innerHTML = "";
    for (var i = 0; i < arr.length; i++) {
      var o = document.createElement("option");
      o.value = arr[i];
      o.textContent = arr[i];
      sel.appendChild(o);
    }
  }

  function ensureFilesArray(obj) {
    if (!obj) return;
    if (!Array.isArray(obj.files)) obj.files = [];
    if (obj.file && obj.files.length === 0) obj.files = [obj.file];
    if (!obj.file && obj.files.length) obj.file = obj.files[0];
  }

  function openAnyFiles(obj) {
    if (!obj) return false;

    ensureFilesArray(obj);

    if (obj.files && obj.files.length) {
      obj.files.forEach(function (file) {
        openStoredFile(file);
      });
      return true;
    }

    if (obj.fileLink && isUrl(obj.fileLink)) {
      window.open(obj.fileLink, "_blank");
      return true;
    }

    return false;
  }

  function ensureProcessObj(sop) {
    if (!sop) return;

    if (!sop.process || typeof sop.process !== "object") {
      sop.process = { notes: "", fileLink: "", file: null, files: [], updatedAt: "" };
    } else {
      if (typeof sop.process.notes !== "string") sop.process.notes = "";
      if (typeof sop.process.fileLink !== "string") sop.process.fileLink = "";
      if (!("file" in sop.process)) sop.process.file = null;
      if (!Array.isArray(sop.process.files)) sop.process.files = [];
      if (typeof sop.process.updatedAt !== "string") sop.process.updatedAt = "";
    }

    ensureFilesArray(sop.process);
  }

  function normalizeDB() {
    if (!db || !Array.isArray(db.sops)) db = { sops: [] };

    db.sops.forEach(function (sop) {
      ensureFilesArray(sop);
      ensureProcessObj(sop);

      if (!Array.isArray(sop.versions)) sop.versions = [];
      sop.versions.forEach(function (v) {
        ensureFilesArray(v);
      });
    });
  }

  function hasProcessFile(sop) {
    ensureProcessObj(sop);
    return !!((sop.process.files && sop.process.files.length) || (sop.process.fileLink && isUrl(sop.process.fileLink)));
  }

  // ---------- ELEMENTS ----------
  var navItems = document.querySelectorAll(".navItem");
  var viewTemplates = byId("viewTemplates");
  var viewProcess = byId("viewProcess");
  var viewDash = byId("viewDash");

  var pageTitle = byId("pageTitle");
  var pageDesc = byId("pageDesc");

  var kpiStack = byId("kpiStack");
  var kpiSection = byId("kpiSection");
  var deptChips = byId("deptChips");

  var q = byId("q");
  var fDept = byId("fDept");
  var fCat = byId("fCat");
  var fStatus = byId("fStatus");
  var btnClear = byId("btnClear");

  var btnNew = byId("btnNew");
  var btnExportXlsx = byId("btnExportXlsx");
  var btnImportXlsx = byId("btnImportXlsx");
  var importXlsxFile = byId("importXlsxFile");
  var btnLogout = byId("btnLogout");

  var tbody = byId("tbody");
  var tbodyProcess = byId("tbodyProcess");

  var processSubtitle = byId("processSubtitle");
  var pTitle = byId("pTitle");
  var pMeta = byId("pMeta");
  var pDept = byId("pDept");
  var pCat = byId("pCat");
  var pStatus = byId("pStatus");
  var pAct = byId("pAct");
  var pOwner = byId("pOwner");
  var pLatest = byId("pLatest");
  var pNotes = byId("pNotes");
  var pFileTag = byId("pFileTag");
  var pFileWrap = byId("pFileWrap");
  var pTimeline = byId("pTimeline");
  var btnOpenDetails = byId("btnOpenDetails");
  var btnAddVersionQuick = byId("btnAddVersionQuick");
  var btnNewProcess = byId("btnNewProcess");

  var pillTotal = byId("pillTotal");
  var pillDept = byId("pillDept");
  var overdueList = byId("overdueList");
  var recentList = byId("recentList");
  var insightsEl = byId("insights");
  var chartStatusCanvas = byId("chartStatus");
  var chartDeptCanvas = byId("chartDept");
  var chartStatus = null;
  var chartDept = null;

  var dlgSop = byId("dlgSop");
  var dlgTitle = byId("dlgTitle");
  var sopForm = byId("sopForm");
  var btnCancel = byId("btnCancel");
  var btnDeleteSop = byId("btnDeleteSop");
  var xSop = byId("xSop");

  var sTitle = byId("sTitle");
  var sDept = byId("sDept");
  var sCat = byId("sCat");
  var sStatus = byId("sStatus");
  var sAct = byId("sAct");
  var sOwner = byId("sOwner");
  var sNotes = byId("sNotes");
  var sFile = byId("sFile");
  var sFilePick = byId("sFilePick");

  var dlgDetails = byId("dlgDetails");
  var btnEditSop = byId("btnEditSop");
  var xDetails = byId("xDetails");
  var dName = byId("dName");
  var dDept = byId("dDept");
  var dCat = byId("dCat");
  var dStatus = byId("dStatus");
  var dAct = byId("dAct");
  var dLatest = byId("dLatest");
  var dCount = byId("dCount");
  var dOwner = byId("dOwner");
  var dFile = byId("dFile");
  var dNotes = byId("dNotes");
  var vbody = byId("vbody");
  var btnAddVersion = byId("btnAddVersion");

  var dlgVersion = byId("dlgVersion");
  var verForm = byId("verForm");
  var vVer = byId("vVer");
  var vDate = byId("vDate");
  var vStatus = byId("vStatus");
  var vSum = byId("vSum");
  var vFile = byId("vFile");
  var vFilePick = byId("vFilePick");
  var btnVerCancel = byId("btnVerCancel");
  var xVer = byId("xVer");

  var navDash = byId("navDash");

  // ---------- STATE ----------
  var activeView = "templates";
  var activeDept = "All";
  var editingId = null;
  var detailsId = null;
  var selectedProcessId = null;
  var currentRole = (sessionStorage.getItem("role") || "").toLowerCase();

  // ---------- PERMISSIONS ----------
  function getPermissions(role) {
    role = (role || "").toLowerCase();

    if (role === "admin") {
      return {
        canViewDashboard: true,
        canViewKpi: true,
        canCreateSop: true,
        canCreateProcess: true,
        canEditSop: true,
        canDeleteSop: true,
        canImportExcel: true,
        canAddVersion: true
      };
    }

    if (role === "manager") {
      return {
        canViewDashboard: true,
        canViewKpi: true,
        canCreateSop: false,
        canCreateProcess: false,
        canEditSop: false,
        canDeleteSop: false,
        canImportExcel: false,
        canAddVersion: false
      };
    }

    return {
      canViewDashboard: false,
      canViewKpi: false,
      canCreateSop: false,
      canCreateProcess: false,
      canEditSop: false,
      canDeleteSop: false,
      canImportExcel: false,
      canAddVersion: false
    };
  }

  function can(roleKey) {
    return getPermissions(currentRole)[roleKey];
  }

  function applyPermissions() {
    if (btnNew) btnNew.style.display = can("canCreateSop") ? "" : "none";
    if (btnNewProcess) btnNewProcess.style.display = can("canCreateProcess") ? "" : "none";
    if (btnImportXlsx) btnImportXlsx.style.display = can("canImportExcel") ? "" : "none";
    if (kpiSection) kpiSection.style.display = can("canViewKpi") ? "" : "none";
    if (navDash) navDash.style.display = can("canViewDashboard") ? "" : "none";
    if (btnEditSop) btnEditSop.style.display = can("canEditSop") ? "" : "none";
    if (btnDeleteSop) btnDeleteSop.style.display = can("canDeleteSop") ? "" : "none";
    if (btnAddVersion) btnAddVersion.style.display = can("canAddVersion") ? "" : "none";
    if (btnAddVersionQuick) btnAddVersionQuick.style.display = can("canAddVersion") ? "" : "none";

    if (!can("canViewDashboard") && activeView === "dash") {
      activeView = "templates";
    }
  }

  // ---------- INIT ----------
  fillSelect(fDept, DEPARTMENTS);
  fillSelect(fCat, CATEGORIES);
  fillSelect(fStatus, STATUSES);

  fillSelect(sDept, DEPARTMENTS.filter(function (x) { return x !== "All"; }));
  fillSelect(sCat, CATEGORIES.filter(function (x) { return x !== "All"; }));
  fillSelect(sStatus, STATUSES.filter(function (x) { return x !== "All"; }));
  fillSelect(vStatus, STATUSES.filter(function (x) { return x !== "All"; }));

  if (sAct) sAct.value = today();

  // ---------- VIEW ----------
  function setView(view) {
    if (view === "dash" && !can("canViewDashboard")) {
      view = "templates";
    }

    activeView = view;

    var chipsRow = document.querySelector(".chipsRow");
    var filtersSection = document.querySelector(".filters");

    if (chipsRow) chipsRow.style.display = (view === "dash") ? "none" : "";
    if (filtersSection) filtersSection.style.display = (view === "dash") ? "none" : "";

    for (var i = 0; i < navItems.length; i++) {
      navItems[i].classList.toggle("active", navItems[i].getAttribute("data-view") === view);
    }

    if (viewTemplates) viewTemplates.classList.toggle("active", view === "templates");
    if (viewProcess) viewProcess.classList.toggle("active", view === "process");
    if (viewDash) viewDash.classList.toggle("active", view === "dash");

    if (pageTitle && pageDesc) {
      if (view === "templates") {
        pageTitle.textContent = "Company Templates";
        pageDesc.textContent = "Standardize processes, improve efficiency, track versions.";
      } else if (view === "process") {
        pageTitle.textContent = "Process & Guidelines";
        pageDesc.textContent = "Open linked process files and keep them connected to SOP templates.";
      } else {
        pageTitle.textContent = "Dashboard";
        pageDesc.textContent = "Live analytics, health widgets, and performance insights.";
      }
    }

    if (view === "dash") renderDashboard();
  }

  // ---------- FILTERING ----------
  function getFiltered() {
    var rows = db.sops.slice();

    if (activeDept !== "All") {
      rows = rows.filter(function (s) {
        return s.department === activeDept;
      });
    }

    var qq = (q && q.value ? q.value : "").trim().toLowerCase();
    if (qq) {
      rows = rows.filter(function (s) {
        return (s.title || "").toLowerCase().indexOf(qq) >= 0;
      });
    }

    if (fDept && fDept.value !== "All") {
      rows = rows.filter(function (s) { return s.department === fDept.value; });
    }

    if (fCat && fCat.value !== "All") {
      rows = rows.filter(function (s) { return s.category === fCat.value; });
    }

    if (fStatus && fStatus.value !== "All") {
      rows = rows.filter(function (s) { return s.status === fStatus.value; });
    }

    return rows;
  }

  function openFile(sopOrVersion) {
    if (!openAnyFiles(sopOrVersion)) {
      alert("No file found.");
    }
  }

  // ---------- KPI ----------
  function countStatuses() {
    var map = {};
    for (var i = 0; i < db.sops.length; i++) {
      var k = db.sops[i].status || "Unknown";
      map[k] = (map[k] || 0) + 1;
    }
    return map;
  }

  function renderKPI() {
    if (!kpiStack) return;

    var total = db.sops.length;
    var c = countStatuses();

    var kpis = [
      { label: "Total SOPs", value: total },
      { label: "Active", value: c["Active"] || 0 },
      { label: "In Progress", value: c["In Progress"] || 0 },
      { label: "Overdue", value: c["Overdue"] || 0 },
      { label: "Archived", value: c["Archived"] || 0 }
    ];

    kpiStack.innerHTML = "";
    for (var i = 0; i < kpis.length; i++) {
      var div = document.createElement("div");
      div.className = "kpiCard";
      div.innerHTML =
        '<div class="kpiTop">' +
          '<div class="kpiLabel">' + escapeHtml(kpis[i].label) + '</div>' +
          '<div class="kpiValue">' + kpis[i].value + '</div>' +
        '</div>';
      kpiStack.appendChild(div);
    }
  }

  // ---------- CHIPS ----------
  function renderDeptChips() {
    if (!deptChips) return;

    deptChips.innerHTML = "";
    for (var i = 0; i < DEPARTMENTS.length; i++) {
      (function (dep) {
        var b = document.createElement("button");
        b.className = "chip" + (activeDept === dep ? " active" : "");
        b.textContent = dep;
        b.onclick = function () {
          activeDept = dep;
          renderAll();
        };
        deptChips.appendChild(b);
      })(DEPARTMENTS[i]);
    }
  }

  // ---------- TABLE: TEMPLATES ----------
  function renderTemplatesTable() {
    if (!tbody) return;
    tbody.innerHTML = "";

    var rows = getFiltered();

    for (var i = 0; i < rows.length; i++) {
      (function (sop) {
        ensureProcessObj(sop);
        ensureFilesArray(sop);

        var versionsCount = (sop.versions || []).length;
        var fileCell = "—";

        if (sop.files && sop.files.length) {
          fileCell = sop.files.map(function (file, index) {
            return '<a href="#" class="linkBtn fileOpen" data-file-index="' + index + '">Open ' + (index + 1) + '</a>';
          }).join(" ");
        } else if (sop.fileLink && isUrl(sop.fileLink)) {
          fileCell = '<a href="#" class="linkBtn fileOpen">Open 1</a>';
        }

        var ownerNotes = [];
        if (sop.owner) ownerNotes.push(sop.owner);
        if (sop.notes) ownerNotes.push(sop.notes);

        var tr = document.createElement("tr");
        tr.innerHTML =
          '<td><a href="#" class="tplLink"><b>' + escapeHtml(sop.title) + '</b></a></td>' +
          '<td>' + escapeHtml(sop.department) + '</td>' +
          '<td>' + escapeHtml(sop.category) + '</td>' +
          '<td>' + escapeHtml(sop.latestVersion || "—") + '</td>' +
          '<td>' + versionsCount + '</td>' +
          '<td>' + formatDate(sop.activationDate) + '</td>' +
          '<td><span class="badge ' + badgeClass(sop.status) + '">' + escapeHtml(sop.status) + '</span></td>' +
          '<td>' + fileCell + '</td>' +
          '<td>' + escapeHtml(ownerNotes.join(" • ") || "—") + '</td>';

        tr.addEventListener("click", function (e) {
          if (e.target && e.target.classList && e.target.classList.contains("fileOpen")) {
            e.preventDefault();

            if (sop.files && sop.files.length) {
              var idx = parseInt(e.target.getAttribute("data-file-index"), 10);
              if (!isNaN(idx) && sop.files[idx]) {
                openStoredFile(sop.files[idx]);
                return;
              }
            }

            openFile(sop);
            return;
          }

          if (e.target && e.target.classList && e.target.classList.contains("tplLink")) {
            e.preventDefault();
            selectedProcessId = sop.id;
            setView("process");
            renderAll();
            renderProcessPanel(sop.id);
            return;
          }

          openDetails(sop.id);
        });

        var tplLink = tr.querySelector(".tplLink");
        if (tplLink) {
          tplLink.addEventListener("click", function (e) {
            e.preventDefault();
            e.stopPropagation();
            selectedProcessId = sop.id;
            setView("process");
            renderAll();
            renderProcessPanel(sop.id);
          });
        }

        tbody.appendChild(tr);
      })(rows[i]);
    }
  }

  // ---------- TABLE: PROCESS ----------
  function renderProcessTable() {
    if (!tbodyProcess) return;
    tbodyProcess.innerHTML = "";

    var rows = getFiltered();

    for (var i = 0; i < rows.length; i++) {
      (function (sop) {
        ensureProcessObj(sop);

        var versionsCount = (sop.versions || []).length;
        var fileCell = "—";

        if (sop.process.files && sop.process.files.length) {
          fileCell = sop.process.files.map(function (file, index) {
            return '<a href="#" class="linkBtn procOpen" data-file-index="' + index + '">Open ' + (index + 1) + '</a>';
          }).join(" ");
        } else if (sop.process.fileLink && isUrl(sop.process.fileLink)) {
          fileCell = '<a href="#" class="linkBtn procOpen" data-file-index="0">Open 1</a>';
        }

        var tr = document.createElement("tr");
        tr.innerHTML =
          '<td><b>' + escapeHtml(sop.title) + '</b></td>' +
          '<td>' + escapeHtml(sop.department) + '</td>' +
          '<td>' + escapeHtml(sop.category) + '</td>' +
          '<td>' + escapeHtml(sop.latestVersion || "—") + '</td>' +
          '<td>' + versionsCount + '</td>' +
          '<td>' + fileCell + '</td>' +
          '<td><span class="badge ' + badgeClass(sop.status) + '">' + escapeHtml(sop.status) + '</span></td>' +
          '<td>' + formatDate(sop.activationDate) + '</td>';

        if (selectedProcessId === sop.id) {
          tr.style.background = "rgba(37,99,235,0.06)";
        }

        tr.addEventListener("click", function (e) {
          if (e.target && e.target.classList && e.target.classList.contains("procOpen")) {
            e.preventDefault();
            e.stopPropagation();

            var idx = parseInt(e.target.getAttribute("data-file-index"), 10);

            if (sop.process.files && sop.process.files.length) {
              if (!isNaN(idx) && sop.process.files[idx]) {
                openStoredFile(sop.process.files[idx]);
                return;
              }
            }

            if (sop.process.fileLink && isUrl(sop.process.fileLink)) {
              window.open(sop.process.fileLink, "_blank");
              return;
            }

            alert("No process file found.");
            return;
          }

          selectedProcessId = sop.id;
          renderProcessTable();
          renderProcessPanel(sop.id);
        });

        tbodyProcess.appendChild(tr);
      })(rows[i]);
    }

    if (processSubtitle) {
      processSubtitle.textContent = selectedProcessId ? "Focused SOP selected." : "Select an SOP to view details.";
    }
  }

  function renderProcessPanel(id) {
    var sop = db.sops.find(function (x) { return x.id === id; });
    if (!sop) return;

    ensureProcessObj(sop);

    if (pTitle) pTitle.textContent = sop.title;
    if (pMeta) pMeta.textContent = (sop.department || "—") + " • " + (sop.category || "—");
    if (pDept) pDept.textContent = sop.department || "—";
    if (pCat) pCat.textContent = sop.category || "—";
    if (pStatus) pStatus.textContent = sop.status || "—";
    if (pAct) pAct.textContent = formatDate(sop.activationDate);
    if (pOwner) pOwner.textContent = sop.owner || "—";
    if (pLatest) pLatest.textContent = sop.latestVersion || "—";

    var procNotesValue = (sop.process && sop.process.notes ? sop.process.notes : "");
    if (pNotes) pNotes.textContent = procNotesValue || sop.notes || "—";

    if (pFileWrap) pFileWrap.innerHTML = "";

    var hasProc = hasProcessFile(sop);
    if (pFileTag) pFileTag.textContent = hasProc ? "Process File Ready" : "No Process File";

    if (pFileWrap && hasProc) {
      if (sop.process.files && sop.process.files.length) {
        sop.process.files.forEach(function (file, index) {
          var a = document.createElement("a");
          a.href = "#";
          a.className = "linkBtn";
          a.textContent = "Open: " + (file.name || ("File " + (index + 1)));
          a.onclick = function (e) {
            e.preventDefault();
            openStoredFile(file);
          };
          pFileWrap.appendChild(a);
        });
      } else if (sop.process.fileLink && isUrl(sop.process.fileLink)) {
        var linkA = document.createElement("a");
        linkA.href = "#";
        linkA.className = "linkBtn";
        linkA.textContent = "Open linked file";
        linkA.onclick = function (e) {
          e.preventDefault();
          window.open(sop.process.fileLink, "_blank");
        };
        pFileWrap.appendChild(linkA);
      }
    }

    if (pTimeline) {
      pTimeline.innerHTML = "";
      var vers = (sop.versions || []).slice().reverse();

      for (var j = 0; j < vers.length; j++) {
        var v = vers[j];
        var div = document.createElement("div");
        div.className = "timeItem";
        div.innerHTML =
          '<div class="timeDot"></div>' +
          '<div class="timeBody">' +
          '<b>' + escapeHtml(v.version) + ' • ' + escapeHtml(v.status) + '</b>' +
          '<span>' + formatDate(v.date) + ' — ' + escapeHtml(v.summary || "") + '</span>' +
          '</div>';
        pTimeline.appendChild(div);
      }
    }

    if (btnOpenDetails) {
      btnOpenDetails.disabled = false;
      btnOpenDetails.onclick = function () {
        openDetails(sop.id);
      };
    }

    if (btnAddVersionQuick) {
      btnAddVersionQuick.disabled = !can("canAddVersion");
      btnAddVersionQuick.onclick = function () {
        if (!can("canAddVersion")) {
          alert("You do not have permission to add a version.");
          return;
        }
        detailsId = sop.id;
        openAddVersion();
      };
    }
  }

  // ---------- SOP CRUD ----------
  function openNew() {
    if (!can("canCreateSop")) {
      alert("You do not have permission to create a new SOP.");
      return;
    }

    editingId = null;
    if (dlgTitle) dlgTitle.textContent = "New SOP";
    if (btnDeleteSop) btnDeleteSop.style.display = "none";

    sTitle.value = "";
    sDept.value = sDept.options[0] ? sDept.options[0].value : "";
    sCat.value = sCat.options[0] ? sCat.options[0].value : "Process";
    sStatus.value = "In Progress";
    sAct.value = today();
    sOwner.value = "";
    sNotes.value = "";
    sFile.value = "";
    if (sFilePick) sFilePick.value = "";

    dlgSop.showModal();
  }

  function openEdit(id) {
    if (!can("canEditSop")) {
      alert("You do not have permission to edit SOPs.");
      return;
    }

    var sop = db.sops.find(function (x) { return x.id === id; });
    if (!sop) return;

    editingId = id;
    if (dlgTitle) dlgTitle.textContent = "Edit SOP";
    if (btnDeleteSop) btnDeleteSop.style.display = can("canDeleteSop") ? "inline-flex" : "none";

    sTitle.value = sop.title || "";
    sDept.value = sop.department || sDept.options[0].value;
    sCat.value = sop.category || "Process";
    sStatus.value = sop.status || "In Progress";
    sAct.value = sop.activationDate || today();
    sOwner.value = sop.owner || "";
    sNotes.value = sop.notes || "";
    sFile.value = sop.fileLink || "";
    if (sFilePick) sFilePick.value = "";

    dlgSop.showModal();
  }

  async function saveSop() {
    if (!editingId && !can("canCreateSop")) {
      alert("You do not have permission to create a new SOP.");
      return;
    }

    if (editingId && !can("canEditSop")) {
      alert("You do not have permission to edit SOPs.");
      return;
    }

    var title = (sTitle.value || "").trim();
    if (!title) return;

    var payload = {
      title: title,
      department: sDept.value,
      category: sCat.value,
      status: sStatus.value,
      activationDate: sAct.value || "",
      owner: (sOwner.value || "").trim(),
      notes: (sNotes.value || "").trim(),
      fileLink: (sFile.value || "").trim(),
      updatedAt: new Date().toISOString()
    };

    var filesArr = [];
    if (sFilePick && sFilePick.files && sFilePick.files.length) {
      for (var i = 0; i < sFilePick.files.length; i++) {
        var savedMeta = await saveFileBlob(sFilePick.files[i]);
        filesArr.push(savedMeta);
      }
    }
    var id = uid();

      var firstVer = {
        id: uid(),
        version: "V1.0",
        date: payload.activationDate || today(),
        status: payload.status,
        summary: "Initial creation",
        fileLink: payload.fileLink || "",
        files: filesArr.slice(),
        file: filesArr.length ? filesArr[0] : null
      };

      var newSop = {
        id: id,
        title: payload.title,
        department: payload.department,
        category: payload.category,
        status: payload.status,
        activationDate: payload.activationDate,
        owner: payload.owner,
        notes: payload.notes,
        fileLink: payload.fileLink,
        files: filesArr,
        file: filesArr.length ? filesArr[0] : null,
        latestVersion: "V1.0",
        updatedAt: payload.updatedAt,
        versions: [firstVer],
        process: {
          notes: "",
          fileLink: "",
          file: null,
          files: [],
          updatedAt: ""
        }
      };

      db.sops.push(newSop);
      await saveSopToCloud(newSop);

    } else {
      var sop = db.sops.find(function (x) { return x.id === editingId; });
      if (!sop) return;

      sop.title = payload.title;
      sop.department = payload.department;
      sop.category = payload.category;
      sop.status = payload.status;
      sop.activationDate = payload.activationDate;
      sop.owner = payload.owner;
      sop.notes = payload.notes;
      sop.fileLink = payload.fileLink;
      sop.updatedAt = payload.updatedAt;

      if (filesArr.length) {
        sop.files = filesArr;
        sop.file = filesArr[0];
      }

      ensureProcessObj(sop);
      await saveSopToCloud(sop);
    }

    dlgSop.close();
    renderAll();
  }

  async function deleteSop() {
    if (!can("canDeleteSop")) {
      alert("You do not have permission to delete SOPs.");
      return;
    }

    if (!editingId) return;
    if (!confirm("Delete this SOP?")) return;

    db.sops = db.sops.filter(function (x) { return x.id !== editingId; });

    await deleteSopFromCloud(editingId);

    dlgSop.close();

    if (detailsId === editingId) detailsId = null;
    if (selectedProcessId === editingId) selectedProcessId = null;

    renderAll();
  }

  // ---------- DETAILS / VERSIONS ----------
  function openDetails(id) {
    var sop = db.sops.find(function (x) { return x.id === id; });
    if (!sop) return;

    ensureFilesArray(sop);
    detailsId = id;

    dName.textContent = sop.title;
    dDept.textContent = sop.department || "—";
    dCat.textContent = sop.category || "—";
    dStatus.textContent = sop.status || "—";
    dAct.textContent = sop.activationDate || "—";
    dLatest.textContent = sop.latestVersion || "—";
    dCount.textContent = String((sop.versions || []).length);
    dOwner.textContent = sop.owner || "—";
    dNotes.textContent = sop.notes || "—";

    if (btnEditSop) btnEditSop.style.display = can("canEditSop") ? "" : "none";
    if (btnAddVersion) btnAddVersion.style.display = can("canAddVersion") ? "" : "none";

    dFile.innerHTML = "";

    if (sop.files && sop.files.length) {
      sop.files.forEach(function (file, index) {
        var a = document.createElement("a");
        a.href = "#";
        a.className = "linkBtn";
        a.textContent = "Open: " + (file.name || ("File " + (index + 1)));
        a.onclick = function (e) {
          e.preventDefault();
          openStoredFile(file);
        };
        dFile.appendChild(a);
      });
    } else if (sop.fileLink && isUrl(sop.fileLink)) {
      var sopLink = document.createElement("a");
      sopLink.href = "#";
      sopLink.className = "linkBtn";
      sopLink.textContent = "Open Link";
      sopLink.onclick = function (e) {
        e.preventDefault();
        window.open(sop.fileLink, "_blank");
      };
      dFile.appendChild(sopLink);
    } else {
      dFile.textContent = "—";
    }

    vbody.innerHTML = "";

    (sop.versions || []).forEach(function (v) {
      ensureFilesArray(v);

      var fileCell = "—";
      if (v.files && v.files.length) {
        fileCell = v.files.map(function (file, index) {
          return '<a href="#" class="linkBtn verFileOpen" data-file-index="' + index + '">Open ' + (index + 1) + '</a>';
        }).join(" ");
      } else if (v.fileLink && isUrl(v.fileLink)) {
        fileCell = '<a href="#" class="linkBtn verFileOpen">Open</a>';
      }

      var tr = document.createElement("tr");
      tr.innerHTML =
        '<td><b>' + escapeHtml(v.version) + '</b></td>' +
        '<td>' + formatDate(v.date) + '</td>' +
        '<td><span class="badge ' + badgeClass(v.status) + '">' + escapeHtml(v.status) + '</span></td>' +
        '<td>' + escapeHtml(v.summary || "") + '</td>' +
        '<td>' + fileCell + '</td>';

      tr.addEventListener("click", function (e) {
        if (e.target && e.target.classList && e.target.classList.contains("verFileOpen")) {
          e.preventDefault();

          if (v.files && v.files.length) {
            var idx = parseInt(e.target.getAttribute("data-file-index"), 10);
            if (!isNaN(idx) && v.files[idx]) {
              openStoredFile(v.files[idx]);
              return;
            }
          }

          if (v.fileLink && isUrl(v.fileLink)) {
            window.open(v.fileLink, "_blank");
            return;
          }

          alert("No file found.");
        }
      });

      vbody.appendChild(tr);
    });

    dlgDetails.showModal();
  }

  function openAddVersion() {
    if (!can("canAddVersion")) {
      alert("You do not have permission to add a version.");
      return;
    }

    var sop = db.sops.find(function (x) { return x.id === detailsId; });
    if (!sop) return;

    vVer.value = nextVersion(sop.latestVersion || "V1.0");
    vDate.value = today();
    vStatus.value = sop.status || "In Progress";
    vSum.value = "";
    vFile.value = "";
    if (vFilePick) vFilePick.value = "";

    dlgVersion.showModal();
  }

  async function saveVersion() {
    if (!can("canAddVersion")) {
      alert("You do not have permission to add a version.");
      return;
    }

    var sop = db.sops.find(function (x) { return x.id === detailsId; });
    if (!sop) return;

    var version = (vVer.value || "").trim();
    var date = (vDate.value || today()).trim();
    var status = (vStatus.value || "").trim();
    var summary = (vSum.value || "").trim();
    var fileLink = (vFile.value || "").trim();

    if (!version || !summary) return;

    var filesArr = [];
    if (vFilePick && vFilePick.files && vFilePick.files.length) {
      for (var i = 0; i < vFilePick.files.length; i++) {
        var savedMeta = await saveFileBlob(vFilePick.files[i]);
        filesArr.push(savedMeta);
      }
    }

    sop.versions = sop.versions || [];
    sop.versions.push({
      id: uid(),
      version: version,
      date: date,
      status: status,
      summary: summary,
      fileLink: fileLink,
      files: filesArr,
      file: filesArr.length ? filesArr[0] : null
    });

    var lastVersion = sop.versions[sop.versions.length - 1];
    if (lastVersion) {
      if (lastVersion.files && lastVersion.files.length) {
        sop.files = lastVersion.files;
        sop.file = lastVersion.files[0];
      } else if (lastVersion.fileLink) {
        sop.fileLink = lastVersion.fileLink;
      }

      sop.latestVersion = lastVersion.version || sop.latestVersion;
    }

    sop.status = status;
    sop.updatedAt = new Date().toISOString();

    if (!sop.files) sop.files = [];
    if (!sop.file && filesArr.length) sop.file = filesArr[0];
    if (!sop.fileLink && fileLink) sop.fileLink = fileLink;

    await saveSopToCloud(sop);

    dlgVersion.close();
    renderAll();
    openDetails(sop.id);
  }

  // ---------- PROCESS DIALOG ----------
  var dlgProc = null;
  var procForm = null;
  var procSelect = null;
  var procNotes = null;
  var procFileLink = null;
  var procFilePick = null;
  var procCancel = null;
  var procX = null;

  function ensureProcDialog() {
    if (dlgProc) return;

    var html =
      '<dialog id="dlgProc" class="dlg">' +
        '<form method="dialog" class="dlgBody" id="procForm">' +
          '<div class="dlgHeader">' +
            '<h2 id="procTitle">Add Process</h2>' +
            '<button class="iconBtn" value="cancel" aria-label="Close" type="button" id="xProc">✕</button>' +
          '</div>' +

          '<div class="grid">' +
            '<label class="span2">Linked SOP Template' +
              '<select id="procSop"></select>' +
            '</label>' +

            '<label class="span2">Process Notes / Guidelines' +
              '<input id="procNotes" placeholder="Short process description..." />' +
            '</label>' +

            '<label class="span2">Process File (PDF / DOCX) from Desktop' +
              '<input id="procFilePick" type="file" multiple accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg" />' +
              '<small class="small">Stored in Firebase Storage.</small>' +
            '</label>' +

            '<label class="span2">Or Process File Link (SharePoint / Drive URL)' +
              '<input id="procFileLink" placeholder="https://..." />' +
            '</label>' +
          '</div>' +

          '<div class="dlgFooter">' +
            '<div></div>' +
            '<div class="rightBtns">' +
              '<button class="btn secondary" type="button" id="procCancel">Cancel</button>' +
              '<button class="btn primary" type="submit" id="procSave">Save Process</button>' +
            '</div>' +
          '</div>' +
        '</form>' +
      '</dialog>';

    var wrap = document.createElement("div");
    wrap.innerHTML = html;
    document.body.appendChild(wrap.firstChild);

    dlgProc = byId("dlgProc");
    procForm = byId("procForm");
    procSelect = byId("procSop");
    procNotes = byId("procNotes");
    procFileLink = byId("procFileLink");
    procFilePick = byId("procFilePick");
    procCancel = byId("procCancel");
    procX = byId("xProc");

    if (procCancel) {
      procCancel.onclick = function () {
        dlgProc.close();
      };
    }

    if (procX) {
      procX.onclick = function () {
        dlgProc.close();
      };
    }

    if (procForm) {
      procForm.addEventListener("submit", function (e) {
        e.preventDefault();
        saveProcessFromDialog();
      });
    }
  }

  function openAddProcessDialog(defaultSopId) {
    if (!can("canCreateProcess")) {
      alert("You do not have permission to add a new process.");
      return;
    }

    ensureProcDialog();

    procSelect.innerHTML = "";
    db.sops.forEach(function (s) {
      var opt = document.createElement("option");
      opt.value = s.id;
      opt.textContent = s.title + " (" + (s.department || "—") + ")";
      procSelect.appendChild(opt);
    });

    if (defaultSopId) procSelect.value = defaultSopId;

    var sop = db.sops.find(function (x) { return x.id === procSelect.value; });
    if (sop) {
      ensureProcessObj(sop);
      procNotes.value = sop.process.notes || "";
      procFileLink.value = sop.process.fileLink || "";
      if (procFilePick) procFilePick.value = "";
    } else {
      procNotes.value = "";
      procFileLink.value = "";
      if (procFilePick) procFilePick.value = "";
    }

    procSelect.onchange = function () {
      var sop2 = db.sops.find(function (x) { return x.id === procSelect.value; });
      if (!sop2) return;

      ensureProcessObj(sop2);
      procNotes.value = sop2.process.notes || "";
      procFileLink.value = sop2.process.fileLink || "";
      if (procFilePick) procFilePick.value = "";
    };

    dlgProc.showModal();
  }

  async function saveProcessFromDialog() {
    if (!can("canCreateProcess")) {
      alert("You do not have permission to add a new process.");
      return;
    }

    var sopId = (procSelect && procSelect.value) ? procSelect.value : "";
    var sop = db.sops.find(function (x) { return x.id === sopId; });
    if (!sop) {
      dlgProc.close();
      return;
    }

    ensureProcessObj(sop);

    var notes = (procNotes.value || "").trim();
    var link = (procFileLink.value || "").trim();

    var filesArr = [];
    if (procFilePick && procFilePick.files && procFilePick.files.length) {
      for (var i = 0; i < procFilePick.files.length; i++) {
        var savedMeta = await saveFileBlob(procFilePick.files[i]);
        filesArr.push(savedMeta);
      }
    }

    sop.process.notes = notes;
    sop.process.fileLink = link;

    if (filesArr.length) {
      sop.process.files = filesArr;
      sop.process.file = filesArr[0];
    } else if (!link) {
      sop.process.files = [];
      sop.process.file = null;
    }

    sop.process.updatedAt = new Date().toISOString();
    sop.updatedAt = new Date().toISOString();

    await saveSopToCloud(sop);

    dlgProc.close();

    selectedProcessId = sop.id;
    renderAll();
    setView("process");
    renderProcessPanel(sop.id);
  }

  // ---------- EXCEL ----------
  function exportExcel() {
    if (typeof XLSX === "undefined") {
      alert("SheetJS not loaded (XLSX). Check your internet or CDN.");
      return;
    }

    var wb = XLSX.utils.book_new();

    var sopsRows = db.sops.map(function (s) {
      ensureProcessObj(s);
      ensureFilesArray(s);

      return {
        SOP_ID: s.id,
        TemplateName: s.title,
        Department: s.department,
        Category: s.category,
        Status: s.status,
        Date: s.activationDate,
        Owner: s.owner,
        Notes: s.notes,
        FileLink: s.fileLink,
        StoredFileName: s.file && s.file.name ? s.file.name : "",
        LatestVersion: s.latestVersion,
        UpdatedAt: s.updatedAt,
        ProcessNotes: s.process && s.process.notes ? s.process.notes : "",
        ProcessFileLink: s.process && s.process.fileLink ? s.process.fileLink : "",
        ProcessStoredFileName: s.process && s.process.file && s.process.file.name ? s.process.file.name : "",
        ProcessUpdatedAt: s.process && s.process.updatedAt ? s.process.updatedAt : ""
      };
    });

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sopsRows), "SOPs");

    var verRows = [];
    db.sops.forEach(function (sop) {
      (sop.versions || []).forEach(function (v) {
        verRows.push({
          SOP_ID: sop.id,
          TemplateName: sop.title,
          Version: v.version,
          Date: v.date,
          Status: v.status,
          Summary: v.summary,
          FileLink: v.fileLink,
          StoredFileName: v.file && v.file.name ? v.file.name : ""
        });
      });
    });

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(verRows), "Versions");
    XLSX.writeFile(wb, "SOP_System_" + today() + ".xlsx");
  }

  function importExcel(file) {
    if (!can("canImportExcel")) {
      alert("You do not have permission to import Excel.");
      return;
    }

    if (typeof XLSX === "undefined") {
      alert("SheetJS not loaded (XLSX). Check your internet or CDN.");
      return;
    }

    var reader = new FileReader();

    reader.onload = async function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var wb = XLSX.read(data, { type: "array" });

        var wsSops = wb.Sheets["SOPs"];
        if (!wsSops) throw new Error("No SOPs sheet");

        var sopsJson = XLSX.utils.sheet_to_json(wsSops, { defval: "" });
        var newSops = [];

        sopsJson.forEach(function (r) {
          var title = String(r.TemplateName || "").trim();
          if (!title) return;

          var id = String(r.SOP_ID || uid()).trim();

          newSops.push({
            id: id,
            title: title,
            department: String(r.Department || "Finance").trim(),
            category: String(r.Category || "Process").trim(),
            status: String(r.Status || "In Progress").trim(),
            activationDate: String(r.Date || "").trim(),
            owner: String(r.Owner || "").trim(),
            notes: String(r.Notes || "").trim(),
            fileLink: String(r.FileLink || "").trim(),
            file: null,
            files: [],
            latestVersion: String(r.LatestVersion || "V1.0").trim(),
            updatedAt: String(r.UpdatedAt || new Date().toISOString()).trim(),
            versions: [],
            process: {
              notes: String(r.ProcessNotes || "").trim(),
              fileLink: String(r.ProcessFileLink || "").trim(),
              file: null,
              files: [],
              updatedAt: String(r.ProcessUpdatedAt || "").trim()
            }
          });
        });

        var map = {};
        newSops.forEach(function (s) {
          map[s.id] = s;
        });

        var wsVers = wb.Sheets["Versions"];
        if (wsVers) {
          var versJson = XLSX.utils.sheet_to_json(wsVers, { defval: "" });

          versJson.forEach(function (r) {
            var sopId = String(r.SOP_ID || "").trim();
            if (!map[sopId]) return;

            map[sopId].versions.push({
              id: uid(),
              version: String(r.Version || "V1.0").trim(),
              date: String(r.Date || today()).trim(),
              status: String(r.Status || map[sopId].status).trim(),
              summary: String(r.Summary || "Imported").trim(),
              fileLink: String(r.FileLink || "").trim(),
              file: null,
              files: []
            });
          });
        }

        newSops.forEach(function (s) {
          ensureProcessObj(s);
          ensureFilesArray(s);

          if (!s.versions.length) {
            s.versions = [{
              id: uid(),
              version: s.latestVersion || "V1.0",
              date: s.activationDate || today(),
              status: s.status,
              summary: "Imported",
              fileLink: s.fileLink || "",
              file: null,
              files: []
            }];
          }

          s.versions.forEach(function (v) {
            ensureFilesArray(v);
          });
        });

        db.sops = newSops;
        selectedProcessId = null;
        detailsId = null;

        var oldSnap = await sopsCollection.get();
        var deletePromises = [];
        oldSnap.forEach(function (doc) {
          deletePromises.push(sopsCollection.doc(doc.id).delete());
        });
        await Promise.all(deletePromises);

        var savePromises = [];
        newSops.forEach(function (s) {
          savePromises.push(saveSopToCloud(s));
        });
        await Promise.all(savePromises);

        renderAll();
        alert("Import successful.");
      } catch (err) {
        console.error(err);
        alert("Import failed. Ensure sheets are named SOPs and Versions.");
      }
    };

    reader.readAsArrayBuffer(file);
  }

  // ---------- DASHBOARD ----------
  function getStatusCounts() {
    var map = {};
    db.sops.forEach(function (s) {
      var k = s.status || "Unknown";
      map[k] = (map[k] || 0) + 1;
    });
    return map;
  }

  function getDeptCounts() {
    var map = {};
    db.sops.forEach(function (s) {
      var k = s.department || "Unknown";
      map[k] = (map[k] || 0) + 1;
    });
    return map;
  }

  function renderDashboardLists() {
    if (pillTotal) pillTotal.textContent = db.sops.length + " SOP(s)";

    var deptSet = {};
    db.sops.forEach(function (s) {
      deptSet[s.department || ""] = true;
    });

    var deptCount = Object.keys(deptSet).filter(function (x) { return x; }).length;
    if (pillDept) pillDept.textContent = deptCount + " dept(s)";

    if (overdueList) {
      overdueList.innerHTML = "";
      var overdue = db.sops.filter(function (s) {
        return (s.status || "").toLowerCase().indexOf("overdue") >= 0;
      });

      if (!overdue.length) {
        overdueList.innerHTML = '<div class="listItem"><b>All good</b><span>No overdue SOPs</span></div>';
      } else {
        overdue.slice(0, 12).forEach(function (s) {
          var div = document.createElement("div");
          div.className = "listItem";
          div.innerHTML =
            "<b>" + escapeHtml(s.title) + "</b>" +
            "<span>" + escapeHtml(s.department) + " • " + escapeHtml(s.activationDate || "—") + "</span>";
          div.onclick = function () {
            openDetails(s.id);
          };
          overdueList.appendChild(div);
        });
      }
    }

    if (recentList) {
      recentList.innerHTML = "";
      var sorted = db.sops.slice().sort(function (a, b) {
        return String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
      });

      if (!sorted.length) {
        recentList.innerHTML = '<div class="listItem"><b>No data</b><span>Add your first SOP</span></div>';
      } else {
        sorted.slice(0, 12).forEach(function (s) {
          var div = document.createElement("div");
          div.className = "listItem";
          div.innerHTML =
            "<b>" + escapeHtml(s.title) + "</b>" +
            "<span>Updated: " + escapeHtml((s.updatedAt || "").slice(0, 10) || "—") + " • " + escapeHtml(s.status || "") + "</span>";
          div.onclick = function () {
            openDetails(s.id);
          };
          recentList.appendChild(div);
        });
      }
    }
  }

  function renderInsights() {
    if (!insightsEl) return;

    var total = db.sops.length;
    var sMap = getStatusCounts();
    var dMap = getDeptCounts();

    var topDept = "—";
    var topVal = 0;

    Object.keys(dMap).forEach(function (k) {
      if (k !== "All" && dMap[k] > topVal) {
        topVal = dMap[k];
        topDept = k;
      }
    });

    var overdue = sMap["Overdue"] || 0;
    var active = sMap["Active"] || 0;
    var prog = sMap["In Progress"] || 0;

    var cards = [
      {
        title: "Operational Risk",
        body: overdue
          ? ("You have " + overdue + " overdue SOP(s). Assign an owner and update versions.")
          : "No overdue SOPs. Great compliance."
      },
      {
        title: "Execution Health",
        body: total
          ? (active + " Active • " + prog + " In Progress out of " + total + " SOPs.")
          : "Start by creating your first SOP."
      },
      {
        title: "Department Focus",
        body: topDept !== "—"
          ? (topDept + " holds the most SOPs. Consider a quarterly review there.")
          : "Add department ownership to unlock insights."
      }
    ];

    insightsEl.innerHTML = "";
    cards.forEach(function (c) {
      var div = document.createElement("div");
      div.className = "insight";
      div.innerHTML =
        '<div class="insightTitle">' + escapeHtml(c.title) + '</div>' +
        '<div class="insightBody">' + escapeHtml(c.body) + '</div>';
      insightsEl.appendChild(div);
    });
  }

  function renderDashboardCharts() {
    if (typeof Chart === "undefined") return;

    var sMap = getStatusCounts();
    var sLabels = Object.keys(sMap);
    var sData = sLabels.map(function (k) { return sMap[k]; });

    var dMap = getDeptCounts();
    var dLabels = DEPARTMENTS.filter(function (d) {
      return d !== "All" && dMap[d];
    });
    var dData = dLabels.map(function (k) { return dMap[k] || 0; });

    if (chartStatus) chartStatus.destroy();
    if (chartDept) chartDept.destroy();

    var tinyLegend = {
      position: "bottom",
      labels: {
        boxWidth: 10,
        boxHeight: 10,
        padding: 10,
        font: { size: 10, weight: "700" }
      }
    };

    if (chartStatusCanvas) {
      chartStatus = new Chart(chartStatusCanvas, {
        type: "doughnut",
        data: {
          labels: sLabels,
          datasets: [{ data: sData, borderWidth: 0 }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: "70%",
          plugins: {
            legend: tinyLegend,
            tooltip: { bodyFont: { size: 11 } }
          }
        }
      });
    }

    if (chartDeptCanvas) {
      chartDept = new Chart(chartDeptCanvas, {
        type: "bar",
        data: {
          labels: dLabels,
          datasets: [{
            data: dData,
            borderWidth: 0,
            barThickness: 18,
            maxBarThickness: 22
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: { bodyFont: { size: 11 } }
          },
          scales: {
            x: {
              ticks: {
                font: { size: 10, weight: "700" },
                maxRotation: 0,
                minRotation: 0
              }
            },
            y: {
              ticks: {
                font: { size: 10, weight: "700" },
                precision: 0
              },
              grid: { drawBorder: false }
            }
          }
        }
      });
    }
  }

  function renderDashboard() {
    renderDashboardLists();
    renderDashboardCharts();
    renderInsights();
  }

  // ---------- RENDER ----------
  function renderAll() {
    applyPermissions();
    renderDeptChips();
    renderKPI();
    renderTemplatesTable();
    renderProcessTable();

    if (selectedProcessId) {
      renderProcessPanel(selectedProcessId);
    } else {
      if (pTitle) pTitle.textContent = "No SOP Selected";
      if (pMeta) pMeta.textContent = "Click an SOP from the list.";
      if (pDept) pDept.textContent = "—";
      if (pCat) pCat.textContent = "—";
      if (pStatus) pStatus.textContent = "—";
      if (pAct) pAct.textContent = "—";
      if (pOwner) pOwner.textContent = "—";
      if (pLatest) pLatest.textContent = "—";
      if (pNotes) pNotes.textContent = "—";
      if (pFileTag) pFileTag.textContent = "No File";
      if (pFileWrap) pFileWrap.innerHTML = "";
      if (pTimeline) pTimeline.innerHTML = "";
      if (btnOpenDetails) btnOpenDetails.disabled = true;
      if (btnAddVersionQuick) btnAddVersionQuick.disabled = true;
    }

    if (activeView === "dash") renderDashboard();
  }

  // ---------- LOGOUT ----------
  function logoutUser() {
    sessionStorage.removeItem("role");

    try {
      if (dlgSop && dlgSop.open) dlgSop.close();
      if (dlgDetails && dlgDetails.open) dlgDetails.close();
      if (dlgVersion && dlgVersion.open) dlgVersion.close();
      if (typeof dlgProc !== "undefined" && dlgProc && dlgProc.open) dlgProc.close();
    } catch (e) {}

    currentRole = "";
    activeView = "templates";
    editingId = null;
    detailsId = null;
    selectedProcessId = null;
    activeDept = "All";

    if (q) q.value = "";
    if (fDept) fDept.value = "All";
    if (fCat) fCat.value = "All";
    if (fStatus) fStatus.value = "All";

    var loginOverlay = byId("loginOverlay");
    var appShell = document.querySelector(".appShell");
    var loginUser = byId("loginUser");
    var loginPass = byId("loginPass");
    var loginMsg = byId("loginMsg");

    if (appShell) appShell.style.display = "none";
    if (loginOverlay) loginOverlay.style.display = "flex";

    if (loginUser) loginUser.value = "";
    if (loginPass) loginPass.value = "";
    if (loginMsg) {
      loginMsg.textContent = "";
      loginMsg.style.color = "";
    }

    applyPermissions();
    setView("templates");
  }

  // ---------- EVENTS ----------
  navItems.forEach(function (btn) {
    btn.addEventListener("click", function () {
      var targetView = btn.getAttribute("data-view");
      setView(targetView);
    });
  });

  if (btnNew) btnNew.onclick = openNew;

  if (btnCancel) {
    btnCancel.onclick = function () {
      dlgSop.close();
    };
  }

  if (xSop) {
    xSop.onclick = function () {
      dlgSop.close();
    };
  }

  if (btnDeleteSop) btnDeleteSop.onclick = deleteSop;

  if (sopForm) {
    sopForm.addEventListener("submit", function (e) {
      e.preventDefault();
      saveSop();
    });
  }

  if (btnEditSop) {
    btnEditSop.onclick = function () {
      if (detailsId) openEdit(detailsId);
    };
  }

  if (xDetails) {
    xDetails.onclick = function () {
      dlgDetails.close();
    };
  }

  if (btnAddVersion) {
    btnAddVersion.onclick = function () {
      openAddVersion();
    };
  }

  if (btnVerCancel) {
    btnVerCancel.onclick = function () {
      dlgVersion.close();
    };
  }

  if (xVer) {
    xVer.onclick = function () {
      dlgVersion.close();
    };
  }

  if (verForm) {
    verForm.addEventListener("submit", function (e) {
      e.preventDefault();
      saveVersion();
    });
  }

  [q, fDept, fCat, fStatus].forEach(function (el) {
    if (!el) return;
    el.addEventListener("input", function () {
      renderAll();
    });
    el.addEventListener("change", function () {
      renderAll();
    });
  });

  if (btnClear) {
    btnClear.onclick = function () {
      if (q) q.value = "";
      if (fDept) fDept.value = "All";
      if (fCat) fCat.value = "All";
      if (fStatus) fStatus.value = "All";
      activeDept = "All";
      renderAll();
    };
  }

  if (btnExportXlsx) {
    btnExportXlsx.onclick = function () {
      exportExcel();
    };
  }

  if (btnImportXlsx) {
    btnImportXlsx.onclick = function () {
      if (!can("canImportExcel")) {
        alert("You do not have permission to import Excel.");
        return;
      }
      if (importXlsxFile) importXlsxFile.click();
    };
  }

  if (importXlsxFile) {
    importXlsxFile.onchange = function () {
      if (importXlsxFile.files && importXlsxFile.files[0]) {
        importExcel(importXlsxFile.files[0]);
      }
      importXlsxFile.value = "";
    };
  }

  if (btnNewProcess) {
    btnNewProcess.onclick = function () {
      if (!can("canCreateProcess")) {
        alert("You do not have permission to add a new process.");
        return;
      }

      var defaultId = selectedProcessId || (db.sops[0] ? db.sops[0].id : "");
      if (!defaultId) {
        alert("Add an SOP first, then add a process linked to it.");
        return;
      }

      openAddProcessDialog(defaultId);
    };
  }

  if (btnLogout) {
    btnLogout.onclick = function () {
      logoutUser();
    };
  }

  // ---------- START ----------
  (async function startApp() {
    await loadDB();
    renderAll();
    applyPermissions();
    setView("templates");
  })();

})();

// ---------- SERVICE WORKER ----------
if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("service-worker.js").catch(function (err) {
    console.error("Service Worker registration failed:", err);
  });
}

// ---------- LOGIN ----------
const users = {
  manager: "1234",
  employee: "1111",
  admin: "0000"
};

const loginBtn = document.getElementById("loginBtn");
const loginOverlay = document.getElementById("loginOverlay");
const appShell = document.querySelector(".appShell");
const loginUserInput = document.getElementById("loginUser");
const loginPassInput = document.getElementById("loginPass");
const loginMsg = document.getElementById("loginMsg");

function showAppForRole(role) {
  if (loginOverlay) loginOverlay.style.display = "none";
  if (appShell) appShell.style.display = "flex";
  sessionStorage.setItem("role", role);

  if (typeof window !== "undefined") {
    setTimeout(function () {
      window.location.reload();
    }, 50);
  }
}

function doLogin() {
  var user = loginUserInput ? loginUserInput.value.trim().toLowerCase() : "";
  var pass = loginPassInput ? loginPassInput.value.trim() : "";

  if (users[user] && users[user] === pass) {
    if (loginMsg) {
      loginMsg.style.color = "green";
      loginMsg.textContent = "Login successful! Welcome " + user + ".";
    }

    setTimeout(function () {
      showAppForRole(user);
    }, 500);
  } else {
    if (loginMsg) {
      loginMsg.style.color = "red";
      loginMsg.textContent = "Incorrect username or password";
    }
  }
}

if (appShell) {
  appShell.style.display = "none";
}

var savedRole = (sessionStorage.getItem("role") || "").toLowerCase();
if (savedRole && users[savedRole]) {
  if (loginOverlay) loginOverlay.style.display = "none";
  if (appShell) appShell.style.display = "flex";
} else {
  if (loginOverlay) loginOverlay.style.display = "flex";
  if (appShell) appShell.style.display = "none";
}

if (loginBtn) {
  loginBtn.addEventListener("click", doLogin);
}

function enterLogin(e) {
  if (e.key === "Enter") {
    doLogin();
  }
}

if (loginUserInput) loginUserInput.addEventListener("keypress", enterLogin);
if (loginPassInput) loginPassInput.addEventListener("keypress", enterLogin);
