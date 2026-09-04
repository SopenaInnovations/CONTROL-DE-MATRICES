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
let histFiltro = { desde: null, hasta: null }; // filtro de fechas del histórico

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
  let batch = writeBatch(db);
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
      batch = writeBatch(db);
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
function claveMatrizPunto(matriz, punto){
  const m = normMatriz(matriz);
  const p = (punto === null || punto === undefined) ? "" : String(punto).trim().toUpperCase();
  return `${m}::${p}`;
}

function computeEffectiveness(retoquesSrc, retiradasSrc){
  const retoquesList = retoquesSrc || RETOQUES;
  const retiradasList = (retiradasSrc || RETIRADAS).filter(r => !r.eliminado);

  // Agrupa retiradas por matriz+punto exactos (dos puntos distintos de la misma
  // referencia son piezas físicas distintas y no deben cruzarse entre sí)
  const retiradasPorClave = {};
  retiradasList.forEach(r => {
    const clave = claveMatrizPunto(r.referencia, r.punto);
    if (!retiradasPorClave[clave]) retiradasPorClave[clave] = [];
    retiradasPorClave[clave].push(r);
  });

  // Agrupa retoques por matriz+punto exactos, ordenados por fecha
  const retoquesPorClave = {};
  retoquesList.forEach(r => {
    const clave = claveMatrizPunto(r.matriz, r.punto);
    if (!retoquesPorClave[clave]) retoquesPorClave[clave] = [];
    retoquesPorClave[clave].push(r);
  });

  // Construye un mapa por id para consultas rápidas
  const retoquesById = {};
  retoquesList.forEach(r => { retoquesById[r.id] = r; });

  const atribuidoA = {};        // retoque.id -> retirada a la que quedó asociado
  const retiradaToRetoque = {}; // retirada.id -> retoque object (o null), para trazabilidad

  Object.entries(retiradasPorClave).forEach(([clave, retiradasDeClave]) => {
    const toques = (retoquesPorClave[clave] || []).slice().sort((a,b) => (a.fecha||"").localeCompare(b.fecha||""));
    const retiradasOrdenadas = retiradasDeClave.slice().sort((a,b) => (a.fecha||"").localeCompare(b.fecha||""));
    const usados = new Set();
    retiradasOrdenadas.forEach(retirada => {
      let candidato = null;
      for (const t of toques){
        if (usados.has(t.id)) continue;
        if (!t.fecha || t.fecha <= retirada.fecha){
          candidato = t;
        }
      }
      if (candidato){ usados.add(candidato.id); atribuidoA[candidato.id] = retirada; }
      retiradaToRetoque[retirada.id] = candidato || null;
    });
  });

  // Un retoque atribuido a una retirada cuenta EN CONTRA (sin resolver) mientras
  // nadie haya vuelto a retocar esa misma matriz+punto desde entonces. En cuanto
  // existe un retoque posterior a esa retirada, el retoque antiguo pasa a
  // Efectivo (ya cumplió su función) y la responsabilidad se traslada al nuevo.
  const fallidoIds = new Set();
  Object.entries(atribuidoA).forEach(([retoqueId, retirada]) => {
    const retoque = retoquesById[retoqueId];
    const clave = claveMatrizPunto(retoque.matriz, retoque.punto);
    const hayRetoquePosterior = (retoquesPorClave[clave] || []).some(t => t.fecha && t.fecha > retirada.fecha);
    if (!hayRetoquePosterior) fallidoIds.add(retoqueId);
  });

  // El retoque más reciente de cada matriz+punto que no esté ya contando como
  // "sin resolver" y que no tenga ninguna retirada posterior queda "Pendiente":
  // todavía no hay confirmación de que esa pieza haya extruido bien.
  const pendienteIds = new Set();
  Object.entries(retoquesPorClave).forEach(([clave, toques]) => {
    const ordenados = toques.slice().sort((a,b) => (a.fecha||"").localeCompare(b.fecha||""));
    const ultimo = ordenados[ordenados.length - 1];
    if (!ultimo || fallidoIds.has(ultimo.id)) return;
    const retiradasClave = retiradasPorClave[clave] || [];
    const hayRetiradaPosterior = retiradasClave.some(r => r.fecha && r.fecha >= ultimo.fecha);
    if (!hayRetiradaPosterior) pendienteIds.add(ultimo.id);
  });

  // Agrega por matricero (los retoques pendientes no cuentan aún en el total)
  const stats = {}; // operario -> {total, fallidos}
  retoquesList.forEach(r => {
    if (pendienteIds.has(r.id)) return;
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

  // Nº de veces que cada matriz+punto exacto ha sido retirada (dentro del rango filtrado)
  const vecesPorMatriz = {};
  retiradasList.forEach(r => {
    const clave = claveMatrizPunto(r.referencia, r.punto);
    vecesPorMatriz[clave] = (vecesPorMatriz[clave] || 0) + 1;
  });

  return { result, fallidoIds, pendienteIds, retiradaToRetoque, vecesPorMatriz };
}

function inRange(fecha, desde, hasta){
  if (!fecha) return false;
  if (desde && fecha < desde) return false;
  if (hasta && fecha > hasta) return false;
  return true;
}

function getFilteredForHistorico(){
  const { desde, hasta } = histFiltro;
  if (!desde && !hasta) return { retoques: RETOQUES, retiradas: RETIRADAS };
  return {
    retoques: RETOQUES.filter(r => inRange(r.fecha, desde, hasta)),
    retiradas: RETIRADAS.filter(r => inRange(r.fecha, desde, hasta))
  };
}

// ============================================================
// RENDER: RETOQUES
// ============================================================
function renderRetoques(){
  const filtro = document.getElementById("filtro-operario").value;
  const { fallidoIds, pendienteIds } = computeEffectiveness();

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
    const pendiente = pendienteIds.has(r.id);
    let estado;
    if (fallo) estado = `<span class="pill no">Retirada · sin resolver</span>`;
    else if (pendiente) estado = `<span class="pill warn">Pendiente · sin confirmar</span>`;
    else estado = `<span class="pill ok">Efectivo</span>`;
    return `
      <tr>
        <td>${fmtDate(r.fecha)}</td>
        <td>${r.operario}</td>
        <td>${r.matriz}</td>
        <td>${r.punto || "—"}</td>
        <td>${estado}</td>
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
  const { retoques, retiradas } = getFilteredForHistorico();
  const { result, retiradaToRetoque, vecesPorMatriz } = computeEffectiveness(retoques, retiradas);

  const tbody = document.getElementById("tbl-historico");
  if (!result.length){ tbody.innerHTML = `<tr><td colspan="5" class="empty">No hay datos en el rango seleccionado.</td></tr>`; }
  else {
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

  // Info del filtro activo
  const info = document.getElementById("hist-filtro-info");
  if (histFiltro.desde || histFiltro.hasta){
    info.textContent = `Mostrando datos entre ${histFiltro.desde ? fmtDate(histFiltro.desde) : "el inicio"} y ${histFiltro.hasta ? fmtDate(histFiltro.hasta) : "hoy"}.`;
  } else {
    info.textContent = "Mostrando todo el histórico (sin filtro de fechas).";
  }

  // Detalle de retiradas: fecha, referencia, motivo, veces retirada, matricero
  const activas = retiradas.filter(r => !r.eliminado).sort((a,b) => (b.fecha||"").localeCompare(a.fecha||""));
  const tbodyDet = document.getElementById("tbl-hist-detalle");
  if (!activas.length){ tbodyDet.innerHTML = `<tr><td colspan="5" class="empty">No hay matrices retiradas en el rango seleccionado.</td></tr>`; return; }

  tbodyDet.innerHTML = activas.map(r => {
    const toque = retiradaToRetoque[r.id];
    return `
      <tr>
        <td>${fmtDate(r.fecha)}</td>
        <td>${r.referencia}</td>
        <td>${r.motivo || "—"}</td>
        <td>${vecesPorMatriz[claveMatrizPunto(r.referencia, r.punto)] || 1}</td>
        <td>${toque ? toque.operario : "—"}</td>
      </tr>
    `;
  }).join("");
}

// ---------- FILTRO DE FECHAS DEL HISTÓRICO ----------
document.getElementById("btn-hist-filtrar").addEventListener("click", () => {
  histFiltro.desde = document.getElementById("hist-desde").value || null;
  histFiltro.hasta = document.getElementById("hist-hasta").value || null;
  renderHistorico();
});

document.getElementById("btn-hist-limpiar").addEventListener("click", () => {
  histFiltro = { desde: null, hasta: null };
  document.getElementById("hist-desde").value = "";
  document.getElementById("hist-hasta").value = "";
  renderHistorico();
});

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

// ============================================================
// EXPORTACIÓN A EXCEL
// ============================================================
function downloadWorkbook(sheets, filename){
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => {
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0,31));
  });
  XLSX.writeFile(wb, filename);
}

document.getElementById("btn-export-retiradas").addEventListener("click", () => {
  const activas = RETIRADAS.filter(r => !r.eliminado).sort((a,b) => (b.fecha||"").localeCompare(a.fecha||""));
  if (!activas.length){ toast("No hay matrices retiradas para exportar.", "err"); return; }
  const rows = activas.map(r => ({
    "Fecha": fmtDate(r.fecha),
    "Referencia": r.referencia,
    "Punto": r.punto || "",
    "Motivo": r.motivo || ""
  }));
  downloadWorkbook([{ name: "Matrices retiradas", rows }], `matrices_retiradas_${new Date().toISOString().slice(0,10)}.xlsx`);
  toast("Excel descargado.", "ok");
});

document.getElementById("btn-export-historico").addEventListener("click", () => {
  const { retoques, retiradas } = getFilteredForHistorico();
  const { result, retiradaToRetoque, vecesPorMatriz } = computeEffectiveness(retoques, retiradas);

  const resumenRows = result.map(r => ({
    "Matricero": r.operario,
    "Retoques totales": r.total,
    "Retirados tras retoque": r.fallidos,
    "Efectivos": r.efectivos,
    "% Efectividad": r.pct
  }));

  const activas = retiradas.filter(r => !r.eliminado).sort((a,b) => (b.fecha||"").localeCompare(a.fecha||""));
  const detalleRows = activas.map(r => {
    const toque = retiradaToRetoque[r.id];
    return {
      "Fecha": fmtDate(r.fecha),
      "Referencia": r.referencia,
      "Motivo": r.motivo || "",
      "Nº veces retirada": vecesPorMatriz[claveMatrizPunto(r.referencia, r.punto)] || 1,
      "Matricero (último retoque)": toque ? toque.operario : ""
    };
  });

  if (!resumenRows.length && !detalleRows.length){ toast("No hay datos en el rango seleccionado para exportar.", "err"); return; }

  const rango = (histFiltro.desde || histFiltro.hasta)
    ? `${histFiltro.desde || "inicio"}_a_${histFiltro.hasta || "hoy"}`
    : "completo";

  downloadWorkbook([
    { name: "Efectividad por matricero", rows: resumenRows },
    { name: "Detalle retiradas", rows: detalleRows }
  ], `historico_control_matrices_${rango}.xlsx`);
  toast("Excel del histórico descargado.", "ok");
});

// ============================================================
// PRIORIZACIÓN DE RETOQUES · ORDEN DE TRABAJO DIARIO
// Cruza el Excel de pedidos (fecha de entrega + referencia de matriz)
// con el Excel de matrices actualmente en circuito de retoque.
// ============================================================
let ordenPedidosData = null;   // { refToFechas: { REF: [Date,...] } }
let ordenMatricesData = null;  // [{ matrizCompleto, referencia, punto, observacion }]
let ordenActivasData = null;   // { porReferencia: { REF: Set(['REF-PUNTO', ...]) } }
let ULTIMO_ORDEN = null;       // último listado generado, para exportar a Excel

function normHeader(s){
  return String(s || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function findHeaderRow(rows, keywords){
  for (let i = 0; i < Math.min(rows.length, 20); i++){
    const row = rows[i] || [];
    const rowNorm = row.map(normHeader);
    const found = keywords.every(kw => rowNorm.some(cell => cell.includes(kw)));
    if (found) return i;
  }
  return -1;
}

function colIndexByKeyword(headerRow, keyword){
  return headerRow.findIndex(cell => normHeader(cell).includes(keyword));
}

function colIndexExact(headerRow, keyword){
  return headerRow.findIndex(cell => normHeader(cell) === keyword);
}

function findHeaderRowExact(rows, keyword){
  for (let i = 0; i < Math.min(rows.length, 20); i++){
    const row = rows[i] || [];
    if (row.some(cell => normHeader(cell) === keyword)) return i;
  }
  return -1;
}

function excelDateToJSDate(v){
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "number"){
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  if (typeof v === "string"){
    const parts = v.split("/");
    if (parts.length === 3){
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[parts.length - 2], 10);
      const year = parseInt(parts[parts.length - 1], 10);
      if (year > 1900) return new Date(year, month - 1, day);
    }
    const asDate = new Date(v);
    if (!isNaN(asDate)) return asDate;
  }
  return null;
}

function loadWorkbookRows(file){
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const sheetName = wb.SheetNames.includes("Datos") ? "Datos" : wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
        resolve(rows);
      } catch (err){ reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function setupDropzoneFile(dropzoneId, inputId, onFile){
  const dz = document.getElementById(dropzoneId);
  const input = document.getElementById(inputId);
  dz.addEventListener("click", () => input.click());
  dz.addEventListener("dragover", e => { e.preventDefault(); dz.classList.add("drag"); });
  dz.addEventListener("dragleave", () => dz.classList.remove("drag"));
  dz.addEventListener("drop", e => {
    e.preventDefault();
    dz.classList.remove("drag");
    if (e.dataTransfer.files.length) onFile(e.dataTransfer.files[0]);
  });
  input.addEventListener("change", () => { if (input.files.length) onFile(input.files[0]); });
}

function checkOrdenReady(){
  document.getElementById("btn-generar-orden").disabled = !(ordenPedidosData && ordenMatricesData);
}

setupDropzoneFile("dropzone-pedidos", "file-pedidos", async (file) => {
  const statusEl = document.getElementById("pedidos-status");
  try {
    const rows = await loadWorkbookRows(file);
    const headerIdx = findHeaderRow(rows, ["FECHA_ENTREGA"]) >= 0
      ? findHeaderRow(rows, ["FECHA_ENTREGA"])
      : findHeaderRow(rows, ["FECHA ENTREGA"]);
    if (headerIdx < 0) throw new Error("No se encontró la columna FECHA_ENTREGA");
    const header = rows[headerIdx];
    const idxFecha = colIndexByKeyword(header, "FECHA_ENTREGA") >= 0
      ? colIndexByKeyword(header, "FECHA_ENTREGA")
      : colIndexByKeyword(header, "FECHA ENTREGA");
    // La referencia de matriz vive en "Perf.M" (código de perfil sin punto ni sufijos),
    // NO en "Útil o Matriz" (que solo está rellena en una minoría de filas) ni en "Perfil".
    const idxRef = colIndexByKeyword(header, "PERF.M") >= 0
      ? colIndexByKeyword(header, "PERF.M")
      : colIndexByKeyword(header, "PERFM");

    if (idxFecha < 0 || idxRef < 0) throw new Error("No se encontraron las columnas necesarias");

    const refToFechas = {};
    for (let i = headerIdx + 1; i < rows.length; i++){
      const r = rows[i];
      if (!r) continue;
      const refRaw = r[idxRef];
      const fechaRaw = r[idxFecha];
      if (refRaw && fechaRaw){
        const fecha = excelDateToJSDate(fechaRaw);
        if (fecha){
          const ref = normMatriz(refRaw);
          if (!refToFechas[ref]) refToFechas[ref] = [];
          refToFechas[ref].push(fecha);
        }
      }
    }
    ordenPedidosData = { refToFechas };
    statusEl.textContent = `✓ Cargado: ${Object.keys(refToFechas).length} referencias con pedido pendiente`;
    statusEl.style.color = "var(--good)";
    checkOrdenReady();
  } catch (err){
    console.error(err);
    statusEl.textContent = "Error al leer el archivo. ¿Es el Excel de pedidos correcto?";
    statusEl.style.color = "var(--bad)";
  }
});

setupDropzoneFile("dropzone-matrices", "file-matrices", async (file) => {
  const statusEl = document.getElementById("matrices-status");
  try {
    const rows = await loadWorkbookRows(file);

    let headerIdx = -1;
    let idxMatriz = -1;

    // Formato "maestro" (como Matrices Activas): la cabecera SIEMPRE está en la fila 0
    // y tiene una columna exacta "Código". Se comprueba primero y solo en la fila 0,
    // porque en este formato la columna "Tipo de Útil" contiene el texto "MATRIZ" en
    // cada fila de datos, lo que confundiría una búsqueda más amplia de esa palabra.
    if (rows[0] && colIndexExact(rows[0], "CODIGO") >= 0){
      headerIdx = 0;
      idxMatriz = colIndexExact(rows[0], "CODIGO");
    }

    // Formato "informe" (el habitual hasta ahora): columna exacta "Matriz",
    // en alguna fila dentro de las primeras 20 (puede haber filas de título antes).
    if (idxMatriz < 0){
      headerIdx = findHeaderRowExact(rows, "MATRIZ");
      if (headerIdx >= 0) idxMatriz = colIndexExact(rows[headerIdx], "MATRIZ");
    }

    if (headerIdx < 0 || idxMatriz < 0) throw new Error("No se encontró ninguna columna de código de matriz (ni 'Matriz' ni 'Código')");

    const header = rows[headerIdx];
    // La columna de observaciones/incidencia solo existe en el formato "informe";
    // en el formato "maestro" simplemente no habrá, y se deja en blanco.
    const idxObs = colIndexByKeyword(header, "OBSERVA");

    const list = [];
    const vistos = new Set();
    for (let i = headerIdx + 1; i < rows.length; i++){
      const r = rows[i];
      if (!r) continue;
      const matrizCompleto = r[idxMatriz];
      if (!matrizCompleto) continue;
      const key = String(matrizCompleto).trim().toUpperCase();
      if (vistos.has(key)) continue;
      vistos.add(key);
      const partes = key.split("-");
      const punto = partes.length > 1 ? partes[partes.length - 1].trim() : "";
      const referencia = normMatriz(partes[0]);
      list.push({
        matrizCompleto: key,
        referencia,
        punto,
        observacion: idxObs >= 0 ? (r[idxObs] || "") : ""
      });
    }
    ordenMatricesData = list;
    statusEl.textContent = `✓ Cargado: ${list.length} matrices para retocar`;
    statusEl.style.color = "var(--good)";
    checkOrdenReady();
  } catch (err){
    console.error(err);
    statusEl.textContent = "Error al leer el archivo. ¿Es el Excel de matrices correcto?";
    statusEl.style.color = "var(--bad)";
  }
});

setupDropzoneFile("dropzone-activas", "file-activas", async (file) => {
  const statusEl = document.getElementById("activas-status");
  try {
    const rows = await loadWorkbookRows(file);
    const headerIdx = findHeaderRow(rows, ["PERFIL", "CODIGO"]);
    if (headerIdx < 0) throw new Error("No se encontraron las columnas Perfil/Código");
    const header = rows[headerIdx];
    const idxPerfil = colIndexByKeyword(header, "PERFIL");
    const idxCodigo = colIndexByKeyword(header, "CODIGO");
    if (idxPerfil < 0 || idxCodigo < 0) throw new Error("No se encontraron las columnas necesarias");

    const porReferencia = {};
    let total = 0;
    for (let i = headerIdx + 1; i < rows.length; i++){
      const r = rows[i];
      if (!r) continue;
      const perfil = r[idxPerfil];
      const codigo = r[idxCodigo];
      if (!perfil || !codigo) continue;
      const ref = normMatriz(perfil);
      const code = String(codigo).trim().toUpperCase();
      if (!porReferencia[ref]) porReferencia[ref] = new Set();
      porReferencia[ref].add(code);
      total++;
    }
    ordenActivasData = { porReferencia };
    statusEl.textContent = `✓ Cargado: ${total} puntos de matriz activos (${Object.keys(porReferencia).length} referencias)`;
    statusEl.style.color = "var(--good)";
  } catch (err){
    console.error(err);
    statusEl.textContent = "Error al leer el archivo. ¿Es el Excel de matrices activas correcto?";
    statusEl.style.color = "var(--bad)";
  }
});

function fmtDateJS(d){
  if (!d) return "—";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yy = d.getFullYear();
  return `${dd}/${mm}/${yy}`;
}

document.getElementById("btn-generar-orden").addEventListener("click", () => {
  if (!ordenPedidosData || !ordenMatricesData) return;

  const hoy = new Date(); hoy.setHours(0,0,0,0);

  // Excluye matrices cuya referencia tiene OTRO punto de matriz ya activo/disponible
  let excluidasPorOtroPunto = 0;
  const candidatas = ordenMatricesData.filter(m => {
    if (!ordenActivasData) return true;
    const activos = ordenActivasData.porReferencia[m.referencia];
    if (!activos) return true; // no consta ningún punto activo para esa referencia -> se mantiene
    const tieneOtroPuntoActivo = [...activos].some(code => code !== m.matrizCompleto);
    if (tieneOtroPuntoActivo) excluidasPorOtroPunto++;
    return !tieneOtroPuntoActivo;
  });

  const filas = candidatas.map(m => {
    const fechas = ordenPedidosData.refToFechas[m.referencia];
    const fechaMin = fechas && fechas.length ? fechas.reduce((a,b) => a < b ? a : b) : null;
    const vencida = fechaMin ? fechaMin < hoy : false;
    return { ...m, fechaEntrega: fechaMin, vencida };
  });

  const conFecha = filas.filter(f => f.fechaEntrega).sort((a,b) => a.fechaEntrega - b.fechaEntrega);
  const sinFecha = filas.filter(f => !f.fechaEntrega);
  const ordenFinal = [...conFecha, ...sinFecha].map((f, i) => ({ ...f, orden: i + 1 }));

  ULTIMO_ORDEN = ordenFinal;
  renderOrdenTrabajo(ordenFinal, excluidasPorOtroPunto);
});

function renderOrdenTrabajo(ordenFinal, excluidasPorOtroPunto){
  document.getElementById("card-orden-resultado").style.display = "block";
  document.getElementById("orden-logo").src = "data:image/png;base64," + window.LOGO_B64;
  document.getElementById("orden-fecha-title").textContent =
    `Grupo Sopena · Generado el ${fmtDateJS(new Date())}`;
  document.getElementById("btn-export-orden").style.display = "inline-block";

  const vencidas = ordenFinal.filter(f => f.vencida).length;
  const conPedido = ordenFinal.filter(f => f.fechaEntrega).length;
  const sinPedido = ordenFinal.length - conPedido;

  document.getElementById("orden-kpis").innerHTML = `
    <div class="kpi"><div class="num">${ordenFinal.length}</div><div class="label">Matrices en el orden</div></div>
    <div class="kpi bad"><div class="num">${vencidas}</div><div class="label">Vencidas · muy urgentes</div></div>
    <div class="kpi"><div class="num">${conPedido}</div><div class="label">Con pedido pendiente</div></div>
    <div class="kpi"><div class="num">${sinPedido}</div><div class="label">Sin pedido (al final)</div></div>
    ${excluidasPorOtroPunto > 0 ? `<div class="kpi good"><div class="num">${excluidasPorOtroPunto}</div><div class="label">Excluidas · otro punto disponible</div></div>` : ""}
  `;

  const tbody = document.getElementById("tbl-orden");
  tbody.innerHTML = ordenFinal.map(f => {
    let claseFila = "";
    let estado = "";
    if (f.vencida){
      claseFila = "row-urgente";
      estado = `<span class="pill no">VENCIDA · MUY URGENTE</span>`;
    } else if (f.fechaEntrega){
      estado = `<span class="pill warn">Con pedido</span>`;
    } else {
      claseFila = "row-sin-pedido";
      estado = `<span class="muted">Sin pedido</span>`;
    }
    return `
      <tr class="${claseFila}">
        <td>${f.orden}</td>
        <td>${f.referencia}</td>
        <td>${f.punto || "—"}</td>
        <td>${fmtDateJS(f.fechaEntrega)}</td>
        <td>${estado}</td>
        <td>${f.observacion || ""}</td>
      </tr>
    `;
  }).join("");
}

document.getElementById("btn-export-orden").addEventListener("click", () => {
  if (!ULTIMO_ORDEN){ toast("Genera primero el orden de trabajo.", "err"); return; }

  const headers = ["Orden", "Matriz", "Punto", "Fecha entrega", "Estado", "Observación"];
  const dataRows = ULTIMO_ORDEN.map(f => [
    f.orden,
    f.referencia,
    f.punto || "",
    f.fechaEntrega ? fmtDateJS(f.fechaEntrega) : "",
    f.vencida ? "VENCIDA - MUY URGENTE" : (f.fechaEntrega ? "Con pedido" : "Sin pedido"),
    f.observacion || ""
  ]);

  const aoa = [
    ["GRUPO SOPENA — ORDEN DE TRABAJO · MATRICERÍA"],
    [`Generado el ${fmtDateJS(new Date())}`],
    [],
    headers,
    ...dataRows
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!merges"] = [{ s: { r:0, c:0 }, e: { r:0, c:5 } }, { s: { r:1, c:0 }, e: { r:1, c:5 } }];
  ws["!cols"] = [{ wch:8 }, { wch:12 }, { wch:8 }, { wch:14 }, { wch:22 }, { wch:50 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Orden de trabajo");
  XLSX.writeFile(wb, `orden_trabajo_matriceria_${new Date().toISOString().slice(0,10)}.xlsx`);
  toast("Excel del orden de trabajo descargado.", "ok");
});
