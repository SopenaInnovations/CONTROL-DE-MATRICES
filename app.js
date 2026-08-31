/* ============================================================
   CONTROL DE MATRICES · Grupo Sopena
   Firebase Firestore (colecciones: retoques, retiradas, extrusion, imports)
   ============================================================ */

const {
  db, auth, collection, doc, getDocs, setDoc, updateDoc, deleteDoc,
  query, orderBy, writeBatch, onSnapshot, signInAnonymously, onAuthStateChanged
} = window.__fb;

// ---------- ESTADO EN MEMORIA ----------
let RETOQUES = [];     // { id, fecha, operario, matriz, punto, source }
let RETIRADAS = [];    // { id, fecha, referencia, punto, motivo, eliminado }
let EXTRUSION = [];    // { id, fecha, extruidas, retiradas }
let IMPORTS = [];      // { id, fecha, nuevosRetoques, nuevasRetiradas, archivo }
let unsubs = [];

// ---------- UTILIDADES ----------
function toast(msg, type=""){
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.className = "toast show " + type;
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(()=> t.classList.remove("show"), 3200);
}

function normMatriz(v){
  if (v === null || v === undefined) return "";
  return String(v).trim().toUpperCase();
}

function excelDateToISO(v){
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0,10);
  if (typeof v === "number"){
    // fecha serial de Excel
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
  }
  if (typeof v === "string"){
    const parts = v.split("/");
    if (parts.length === 3){
      // maneja formatos rotos tipo "30/04/04/2026" tomando los 2 últimos como mes/año validos
      const day = parts[0].padStart(2,"0");
      const month = parts[parts.length-2].padStart(2,"0");
      const year = parts[parts.length-1];
      if (year.length === 4) return `${year}-${month}-${day}`;
    }
    const asDate = new Date(v);
    if (!isNaN(asDate)) return asDate.toISOString().slice(0,10);
  }
  return null;
}

function fmtDate(iso){
  if (!iso) return "—";
  const [y,m,d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

function idFromParts(...parts){
  return parts.map(p => String(p ?? "").trim().toUpperCase().replace(/[^A-Z0-9]/g,"_")).join("__").slice(0,400);
}

// ---------- LOGIN ----------
document.getElementById("btn-login").addEventListener("click", doLogin);
document.getElementById("login-pass").addEventListener("keydown", e => { if (e.key === "Enter") doLogin(); });

function doLogin(){
  const email = document.getElementById("login-email").value.trim();
  const pass = document.getElementById("login-pass").value;
  const errEl = document.getElementById("login-error");
  errEl.textContent = "";
  if (!email || !pass){ errEl.textContent = "Introduce correo y contraseña."; return; }

  import("https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js").then(({ signInWithEmailAndPassword }) => {
    signInWithEmailAndPassword(auth, email, pass)
      .catch(err => {
        errEl.textContent = "No se pudo iniciar sesión. Revisa el correo/contraseña.";
        console.error(err);
      });
  });
}

document.getElementById("btn-logout").addEventListener("click", () => {
  auth.signOut();
});

onAuthStateChanged(auth, user => {
  if (user){
    document.getElementById("login-screen").style.display = "none";
    document.getElementById("app").style.display = "block";
    startListeners();
  } else {
    document.getElementById("login-screen").style.display = "flex";
    document.getElementById("app").style.display = "none";
    unsubs.forEach(u => u());
    unsubs = [];
  }
});

// ---------- TABS ----------
document.querySelectorAll(".tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById(btn.dataset.view).classList.add("active");
  });
});

// ---------- LISTENERS FIRESTORE ----------
function startListeners(){
  const setStatus = txt => document.getElementById("sync-status").textContent = txt;

  unsubs.push(onSnapshot(collection(db, "retoques"), snap => {
    RETOQUES = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    renderAll();
  }, err => { console.error(err); setStatus("Error de sincronización"); }));

  unsubs.push(onSnapshot(collection(db, "retiradas"), snap => {
    RETIRADAS = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    renderAll();
  }, err => console.error(err)));

  unsubs.push(onSnapshot(collection(db, "extrusion"), snap => {
    EXTRUSION = snap.docs.map(d => ({ id: d.id, ...d.data() })).sort((a,b)=> (a.fecha||"").localeCompare(b.fecha||""));
    renderExtrusion();
    renderDashboard();
  }, err => console.error(err)));

  unsubs.push(onSnapshot(collection(db, "imports"), snap => {
    IMPORTS = snap.docs.map(d => ({ id: d.id, ...d.data() })).sort((a,b)=> (b.fecha||"").localeCompare(a.fecha||""));
    renderImports();
  }, err => console.error(err)));
}

function renderAll(){
  renderRetoques();
  renderRetiradas();
  renderPapelera();
  renderDashboard();
  renderHistorico();
}

// ============================================================
// IMPORTACIÓN DE EXCEL
// ============================================================
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("dragover", e => { e.preventDefault(); dropzone.classList.add("drag"); });
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("drag"));
dropzone.addEventListener("drop", e => {
  e.preventDefault();
  dropzone.classList.remove("drag");
  if (e.dataTransfer.files.length) handleExcelFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener("change", () => {
  if (fileInput.files.length) handleExcelFile(fileInput.files[0]);
});

function handleExcelFile(file){
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array", cellDates: true });
      const sheetName = wb.SheetNames.includes("Datos") ? "Datos" : wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      await processImportRows(rows, file.name);
    } catch (err){
      console.error(err);
      toast("No se pudo leer el archivo. Comprueba que es el Excel correcto.", "err");
    }
  };
  reader.readAsArrayBuffer(file);
}

async function processImportRows(rows, filename){
  // Salta la fila de cabecera
  const dataRows = rows.slice(1);

  let nuevosRetoques = 0, nuevasRetiradas = 0, ignorados = 0;
  const batch = writeBatch(db);
  let opsInBatch = 0;
  const existingRetoqueIds = new Set(RETOQUES.map(r => r.id));
  const existingRetiradaIds = new Set(RETIRADAS.map(r => r.id));

  for (const row of dataRows){
    // ---- Columnas A-D: retoques ----
    const [fechaRaw, operario, matriz, punto, , retiradaLabel, fechaRetRaw, referencia, puntoRet, motivo] = row;

    if (operario || matriz !== null && matriz !== undefined){
      const fechaISO = excelDateToISO(fechaRaw);
      if (fechaISO && operario && (matriz !== null && matriz !== undefined && matriz !== "")){
        const opNorm = String(operario).trim().toUpperCase();
        const matrizNorm = normMatriz(matriz);
        const puntoNorm = (punto === null || punto === undefined) ? "" : String(punto).trim();
        const id = idFromParts(fechaISO, opNorm, matrizNorm, puntoNorm);
        if (!existingRetoqueIds.has(id)){
          batch.set(doc(db, "retoques", id), {
            fecha: fechaISO,
            operario: opNorm,
            matriz: matrizNorm,
            punto: puntoNorm,
            source: "excel",
            importado: new Date().toISOString()
          });
          existingRetoqueIds.add(id);
          nuevosRetoques++;
          opsInBatch++;
        }
      }
    }

    // ---- Columnas F-J: retiradas ----
    if (referencia !== null && referencia !== undefined && referencia !== "" && fechaRetRaw){
      const fechaRetISO = excelDateToISO(fechaRetRaw);
      const refNorm = normMatriz(referencia);
      const puntoRetNorm = (puntoRet === null || puntoRet === undefined) ? "" : String(puntoRet).trim();
      const motivoNorm = motivo ? String(motivo).trim() : "";
      if (fechaRetISO){
        const id = idFromParts(fechaRetISO, refNorm, puntoRetNorm, motivoNorm);
        if (!existingRetiradaIds.has(id)){
          batch.set(doc(db, "retiradas", id), {
            fecha: fechaRetISO,
            referencia: refNorm,
            punto: puntoRetNorm,
            motivo: motivoNorm,
            eliminado: false,
            source: "excel",
            importado: new Date().toISOString()
          });
          existingRetiradaIds.add(id);
          nuevasRetiradas++;
          opsInBatch++;
        }
      }
    }

    if (opsInBatch >= 400){
      await batch.commit();
      opsInBatch = 0;
    }
  }

  if (opsInBatch > 0) await batch.commit();

  // Registrar la importación en el histórico
  await setDoc(doc(collection(db, "imports")), {
    fecha: new Date().toISOString(),
    nuevosRetoques,
    nuevasRetiradas,
    archivo: filename
  });

  document.getElementById("import-summary").innerHTML = `
    <div class="line"><span>Retoques nuevos añadidos</span><b>${nuevosRetoques}</b></div>
    <div class="line"><span>Matrices retiradas nuevas</span><b>${nuevasRetiradas}</b></div>
    <div class="line"><span>Filas ya existentes (ignoradas)</span><b>${dataRows.length - nuevosRetoques - nuevasRetiradas}</b></div>
  `;
  toast(`Importación completada: ${nuevosRetoques} retoques y ${nuevasRetiradas} retiradas nuevas.`, "ok");
}

function renderImports(){
  const tbody = document.getElementById("tbl-imports");
  if (!IMPORTS.length){ tbody.innerHTML = `<tr><td colspan="4" class="empty">Aún no se ha importado ningún archivo.</td></tr>`; return; }
  tbody.innerHTML = IMPORTS.map(imp => `
    <tr>
      <td>${new Date(imp.fecha).toLocaleString("es-ES")}</td>
      <td>${imp.nuevosRetoques}</td>
      <td>${imp.nuevasRetiradas}</td>
      <td>${imp.archivo || "—"}</td>
    </tr>
  `).join("");
}

// ============================================================
// CÁLCULO DE EFECTIVIDAD POR MATRICERO
// Un retoque se considera "fallido" si esa misma matriz aparece
// en la lista de retiradas en una fecha posterior (o igual) al retoque.
// Si una matriz fue retocada varias veces antes de retirarse, el
// fallo se atribuye al retoque más reciente anterior a la retirada.
// ============================================================
function computeEffectiveness(){
  // Agrupa retiradas por matriz normalizada
  const retiradasPorMatriz = {};
  RETIRADAS.filter(r => !r.eliminado).forEach(r => {
    if (!retiradasPorMatriz[r.referencia]) retiradasPorMatriz[r.referencia] = [];
    retiradasPorMatriz[r.referencia].push(r.fecha);
  });

  // Agrupa retoques por matriz normalizada, ordenados por fecha
  const retoquesPorMatriz = {};
  RETOQUES.forEach(r => {
    if (!retoquesPorMatriz[r.matriz]) retoquesPorMatriz[r.matriz] = [];
    retoquesPorMatriz[r.matriz].push(r);
  });

  const fallidoIds = new Set();

  Object.entries(retiradasPorMatriz).forEach(([matriz, fechasRetirada]) => {
    const toques = (retoquesPorMatriz[matriz] || []).slice().sort((a,b) => (a.fecha||"").localeCompare(b.fecha||""));
    fechasRetirada.forEach(fechaRet => {
      // Busca el retoque más reciente con fecha <= fecha de retirada, que no esté ya marcado
      let candidato = null;
      for (const t of toques){
        if (fallidoIds.has(t.id)) continue;
        if (!t.fecha || t.fecha <= fechaRet){
          candidato = t;
        }
      }
      if (candidato) fallidoIds.add(candidato.id);
    });
  });

  // Agrega por matricero
  const stats = {}; // operario -> {total, fallidos}
  RETOQUES.forEach(r => {
    if (!stats[r.operario]) stats[r.operario] = { total: 0, fallidos: 0 };
    stats[r.operario].total++;
    if (fallidoIds.has(r.id)) stats[r.operario].fallidos++;
  });

  const result = Object.entries(stats).map(([operario, s]) => ({
    operario,
    total: s.total,
    fallidos: s.fallidos,
    efectivos: s.total - s.fallidos,
    pct: s.total ? Math.round(((s.total - s.fallidos) / s.total) * 1000) / 10 : 0
  })).sort((a,b) => b.total - a.total);

  return { result, fallidoIds };
}

// ============================================================
// RENDER: RETOQUES
// ============================================================
function renderRetoques(){
  const filtro = document.getElementById("filtro-operario").value;
  const { fallidoIds } = computeEffectiveness();

  // refresca opciones del selector
  const select = document.getElementById("filtro-operario");
  const current = select.value;
  const operarios = [...new Set(RETOQUES.map(r => r.operario))].sort();
  select.innerHTML = `<option value="">Todos los matriceros</option>` +
    operarios.map(o => `<option value="${o}" ${o===current?"selected":""}>${o}</option>`).join("");

  let list = RETOQUES.slice().sort((a,b) => (b.fecha||"").localeCompare(a.fecha||""));
  if (filtro) list = list.filter(r => r.operario === filtro);
  list = list.slice(0, 400); // limita render por rendimiento

  const tbody = document.getElementById("tbl-retoques");
  if (!list.length){ tbody.innerHTML = `<tr><td colspan="6" class="empty">No hay retoques registrados todavía.</td></tr>`; return; }

  tbody.innerHTML = list.map(r => {
    const fallo = fallidoIds.has(r.id);
    return `
      <tr>
        <td>${fmtDate(r.fecha)}</td>
        <td>${r.operario}</td>
        <td>${r.matriz}</td>
        <td>${r.punto || "—"}</td>
        <td>${fallo ? `<span class="pill no">Retirada después</span>` : `<span class="pill ok">Efectivo</span>`}</td>
        <td><button class="btn ghost btn-sm" onclick="deleteRetoque('${r.id}')">Eliminar</button></td>
      </tr>
    `;
  }).join("");
}

document.getElementById("filtro-operario").addEventListener("change", renderRetoques);

document.getElementById("btn-add-retoque").addEventListener("click", async () => {
  const fecha = prompt("Fecha (AAAA-MM-DD):", new Date().toISOString().slice(0,10));
  if (!fecha) return;
  const operario = prompt("Matricero:");
  if (!operario) return;
  const matriz = prompt("Nº de matriz:");
  if (!matriz) return;
  const punto = prompt("Punto de matriz (opcional):", "");

  const opNorm = operario.trim().toUpperCase();
  const matrizNorm = normMatriz(matriz);
  const id = idFromParts(fecha, opNorm, matrizNorm, punto || "");
  await setDoc(doc(db, "retoques", id), {
    fecha, operario: opNorm, matriz: matrizNorm, punto: punto || "",
    source: "manual", importado: new Date().toISOString()
  });
  toast("Retoque añadido.", "ok");
});

window.deleteRetoque = async function(id){
  if (!confirm("¿Eliminar este retoque? No se puede deshacer.")) return;
  await deleteDoc(doc(db, "retoques", id));
  toast("Retoque eliminado.");
};

// ============================================================
// RENDER: RETIRADAS + PAPELERA
// ============================================================
document.getElementById("btn-save-retirada").addEventListener("click", async () => {
  const fecha = document.getElementById("ret-fecha").value;
  const referencia = document.getElementById("ret-referencia").value.trim();
  const punto = document.getElementById("ret-punto").value.trim();
  const motivo = document.getElementById("ret-motivo").value.trim();
  const editId = document.getElementById("ret-edit-id").value;
  const msg = document.getElementById("ret-form-msg");

  if (!fecha || !referencia){
    msg.textContent = "Fecha y referencia son obligatorias.";
    msg.style.color = "var(--bad)";
    return;
  }

  const refNorm = normMatriz(referencia);
  const id = editId || idFromParts(fecha, refNorm, punto, motivo, Date.now());

  await setDoc(doc(db, "retiradas", id), {
    fecha, referencia: refNorm, punto, motivo,
    eliminado: false, source: "manual",
    importado: new Date().toISOString()
  }, { merge: true });

  msg.textContent = editId ? "Retirada actualizada." : "Retirada añadida.";
  msg.style.color = "var(--good)";
  document.getElementById("ret-edit-id").value = "";
  document.getElementById("ret-fecha").value = "";
  document.getElementById("ret-referencia").value = "";
  document.getElementById("ret-punto").value = "";
  document.getElementById("ret-motivo").value = "";
  toast("Guardado correctamente.", "ok");
});

function renderRetiradas(){
  const activas = RETIRADAS.filter(r => !r.eliminado).sort((a,b) => (b.fecha||"").localeCompare(a.fecha||""));
  document.getElementById("ret-count").textContent = `${activas.length} matrices retiradas`;
  const tbody = document.getElementById("tbl-retiradas");
  if (!activas.length){ tbody.innerHTML = `<tr><td colspan="5" class="empty">No hay matrices retiradas registradas.</td></tr>`; return; }

  tbody.innerHTML = activas.slice(0,400).map(r => `
    <tr>
      <td>${fmtDate(r.fecha)}</td>
      <td>${r.referencia}</td>
      <td>${r.punto || "—"}</td>
      <td>${r.motivo || "—"}</td>
      <td>
        <button class="btn ghost btn-sm" onclick="editRetirada('${r.id}')">Editar</button>
        <button class="btn danger btn-sm" onclick="trashRetirada('${r.id}')">Eliminar</button>
      </td>
    </tr>
  `).join("");
}

window.editRetirada = function(id){
  const r = RETIRADAS.find(x => x.id === id);
  if (!r) return;
  document.getElementById("ret-edit-id").value = id;
  document.getElementById("ret-fecha").value = r.fecha || "";
  document.getElementById("ret-referencia").value = r.referencia || "";
  document.getElementById("ret-punto").value = r.punto || "";
  document.getElementById("ret-motivo").value = r.motivo || "";
  document.getElementById("v-retiradas").scrollIntoView({ behavior: "smooth" });
};

window.trashRetirada = async function(id){
  await updateDoc(doc(db, "retiradas", id), { eliminado: true, eliminadoEn: new Date().toISOString() });
  toast("Movida a la papelera.");
};

function renderPapelera(){
  const eliminadas = RETIRADAS.filter(r => r.eliminado).sort((a,b) => (b.eliminadoEn||"").localeCompare(a.eliminadoEn||""));
  const tbody = document.getElementById("tbl-papelera");
  if (!eliminadas.length){ tbody.innerHTML = `<tr><td colspan="5" class="empty">La papelera está vacía.</td></tr>`; return; }

  tbody.innerHTML = eliminadas.map(r => `
    <tr>
      <td>${fmtDate(r.fecha)}</td>
      <td>${r.referencia}</td>
      <td>${r.punto || "—"}</td>
      <td>${r.motivo || "—"}</td>
      <td>
        <button class="btn secondary btn-sm" onclick="restoreRetirada('${r.id}')">Restaurar</button>
        <button class="btn danger btn-sm" onclick="deleteForever('${r.id}')">Eliminar definitivo</button>
      </td>
    </tr>
  `).join("");
}

window.restoreRetirada = async function(id){
  await updateDoc(doc(db, "retiradas", id), { eliminado: false });
  toast("Restaurada.", "ok");
};

window.deleteForever = async function(id){
  if (!confirm("Esto elimina el registro definitivamente. ¿Continuar?")) return;
  await deleteDoc(doc(db, "retiradas", id));
  toast("Eliminado definitivamente.");
};

// ============================================================
// EXTRUSIÓN DIARIA
// ============================================================
document.getElementById("btn-save-extrusion").addEventListener("click", async () => {
  const fecha = document.getElementById("ext-fecha").value;
  const extruidas = Number(document.getElementById("ext-extruidas").value || 0);
  const retiradas = Number(document.getElementById("ext-retiradas").value || 0);
  const editId = document.getElementById("ext-edit-id").value;
  const msg = document.getElementById("ext-form-msg");

  if (!fecha){ msg.textContent = "Indica una fecha."; msg.style.color = "var(--bad)"; return; }

  const id = editId || fecha;
  await setDoc(doc(db, "extrusion", id), { fecha, extruidas, retiradas }, { merge: true });

  msg.textContent = "Guardado.";
  msg.style.color = "var(--good)";
  document.getElementById("ext-edit-id").value = "";
  document.getElementById("ext-extruidas").value = "";
  document.getElementById("ext-retiradas").value = "";
  toast("Totales del día guardados.", "ok");
});

function renderExtrusion(){
  const tbody = document.getElementById("tbl-extrusion");
  if (!EXTRUSION.length){ tbody.innerHTML = `<tr><td colspan="5" class="empty">Sin datos todavía.</td></tr>`; return; }
  const list = EXTRUSION.slice().reverse();
  tbody.innerHTML = list.map(e => {
    const pct = e.extruidas ? Math.round((e.retiradas / e.extruidas) * 1000)/10 : 0;
    return `
      <tr>
        <td>${fmtDate(e.fecha)}</td>
        <td>${e.extruidas}</td>
        <td>${e.retiradas}</td>
        <td>${pct}%</td>
        <td>
          <button class="btn ghost btn-sm" onclick="editExtrusion('${e.id}')">Editar</button>
          <button class="btn danger btn-sm" onclick="deleteExtrusion('${e.id}')">Eliminar</button>
        </td>
      </tr>
    `;
  }).join("");
}

window.editExtrusion = function(id){
  const e = EXTRUSION.find(x => x.id === id);
  if (!e) return;
  document.getElementById("ext-edit-id").value = id;
  document.getElementById("ext-fecha").value = e.fecha;
  document.getElementById("ext-extruidas").value = e.extruidas;
  document.getElementById("ext-retiradas").value = e.retiradas;
  document.getElementById("v-extrusion").scrollIntoView({ behavior: "smooth" });
};

window.deleteExtrusion = async function(id){
  if (!confirm("¿Eliminar este registro de extrusión?")) return;
  await deleteDoc(doc(db, "extrusion", id));
  toast("Eliminado.");
};

// ============================================================
// HISTÓRICO
// ============================================================
function renderHistorico(){
  const { result } = computeEffectiveness();
  const tbody = document.getElementById("tbl-historico");
  if (!result.length){ tbody.innerHTML = `<tr><td colspan="5" class="empty">Todavía no hay datos suficientes.</td></tr>`; return; }
  tbody.innerHTML = result.map(r => `
    <tr>
      <td><b>${r.operario}</b></td>
      <td>${r.total}</td>
      <td>${r.fallidos}</td>
      <td>${r.efectivos}</td>
      <td>${pctPill(r.pct)}</td>
    </tr>
  `).join("");
}

function pctPill(pct){
  const cls = pct >= 90 ? "ok" : pct >= 70 ? "warn" : "no";
  return `<span class="pill ${cls}">${pct}%</span>`;
}

// ============================================================
// DASHBOARD + GRÁFICAS
// ============================================================
let chartEfectividad, chartExtrusion, chartMotivos;

function renderDashboard(){
  const { result } = computeEffectiveness();
  const totalRetoques = RETOQUES.length;
  const totalRetiradas = RETIRADAS.filter(r => !r.eliminado).length;
  const pctGlobal = totalRetoques ? Math.round(((totalRetoques - result.reduce((s,r)=>s+r.fallidos,0)) / totalRetoques) * 1000)/10 : 0;
  const ultimoDia = EXTRUSION[EXTRUSION.length - 1];

  document.getElementById("dash-kpis").innerHTML = `
    <div class="kpi"><div class="num">${totalRetoques}</div><div class="label">Retoques totales registrados</div></div>
    <div class="kpi bad"><div class="num">${totalRetiradas}</div><div class="label">Matrices retiradas activas</div></div>
    <div class="kpi good"><div class="num">${pctGlobal}%</div><div class="label">Efectividad global de retoque</div></div>
    <div class="kpi"><div class="num">${ultimoDia ? ultimoDia.extruidas : "—"}</div><div class="label">Extruidas (último día)</div></div>
  `;

  // Gráfica efectividad por matricero
  const ctx1 = document.getElementById("chart-efectividad");
  const labels1 = result.map(r => r.operario);
  const data1 = result.map(r => r.pct);
  if (chartEfectividad) chartEfectividad.destroy();
  chartEfectividad = new Chart(ctx1, {
    type: "bar",
    data: { labels: labels1, datasets: [{ label: "% Efectividad", data: data1, backgroundColor: "#003DA5", borderRadius: 5 }] },
    options: {
      responsive:true, maintainAspectRatio:false,
      scales: { y: { beginAtZero:true, max:100, ticks:{ callback:v=>v+"%" } } },
      plugins: { legend: { display:false } }
    }
  });

  // Gráfica extrusión vs retiradas
  const ctx2 = document.getElementById("chart-extrusion");
  const dataExt = EXTRUSION.slice(-30);
  if (chartExtrusion) chartExtrusion.destroy();
  chartExtrusion = new Chart(ctx2, {
    type: "line",
    data: {
      labels: dataExt.map(e => fmtDate(e.fecha)),
      datasets: [
        { label: "Extruidas", data: dataExt.map(e=>e.extruidas), borderColor:"#003DA5", backgroundColor:"rgba(0,61,165,.12)", tension:.25, fill:true },
        { label: "Retiradas", data: dataExt.map(e=>e.retiradas), borderColor:"#d64545", backgroundColor:"rgba(214,69,69,.12)", tension:.25, fill:true }
      ]
    },
    options: { responsive:true, maintainAspectRatio:false, scales: { y: { beginAtZero:true } } }
  });

  // Gráfica retiradas por motivo
  const motivos = {};
  RETIRADAS.filter(r => !r.eliminado && r.motivo).forEach(r => {
    motivos[r.motivo] = (motivos[r.motivo] || 0) + 1;
  });
  const motivoEntries = Object.entries(motivos).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const ctx3 = document.getElementById("chart-motivos");
  if (chartMotivos) chartMotivos.destroy();
  chartMotivos = new Chart(ctx3, {
    type: "bar",
    data: { labels: motivoEntries.map(m=>m[0]), datasets:[{ label:"Retiradas", data: motivoEntries.map(m=>m[1]), backgroundColor:"#e0a319", borderRadius:5 }]},
    options: { indexAxis:"y", responsive:true, maintainAspectRatio:false, plugins:{ legend:{ display:false } }, scales:{ x:{ beginAtZero:true } } }
  });
}
