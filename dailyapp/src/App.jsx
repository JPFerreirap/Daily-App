import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient.js";
import {
  Wallet,
  Dumbbell,
  ArrowLeft,
  Plus,
  Minus,
  Trash2,
  Upload,
  Download,
  Database,
  List,
  BarChart3,
  Archive,
  RotateCcw,
  RefreshCw,
  Pencil,
  Check,
  X,
  PieChart as PieChartIcon,
  CreditCard,
  Banknote,
  FileText,
  AlertTriangle,
} from "lucide-react";
import {
  PieChart,
  Pie,
  Cell,
  Tooltip,
  ResponsiveContainer,
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
} from "recharts";

/* ---------- Paleta y helpers de estilo ---------- */
const COLORS = {
  bg: "#EFEDE6",
  card: "#FBFAF7",
  ink: "#1F2A24",
  sub: "#5A6660",
  finanzas: "#2F5D8C",
  deportes: "#2B5D63",
  line: "#DAD6C9",
};

const TIPO_STYLES = {
  activo: { bg: "#E4EDE3", text: "#3B6E4F", label: "Activo" },
  pasivo: { bg: "#F3E3DE", text: "#9C4A2E", label: "Pasivo" },
  ingreso: { bg: "#E3EAF0", text: "#33587A", label: "Ingreso" },
  gasto: { bg: "#F2E8D8", text: "#8C6E2F", label: "Gasto" },
};
const TIPOS_VALIDOS = ["activo", "pasivo", "ingreso", "gasto"];
const PIE_COLORS = ["#2F5D8C", "#4C7FAE", "#6497B8", "#1F4368", "#7FA8C9", "#345E7A", "#5A83A3", "#9CBAD6", "#274A63", "#88AEC9"];

const inputStyle = {
  border: "1px solid #DAD6C9",
  borderRadius: "0.6rem",
  padding: "0.5rem 0.7rem",
  fontFamily: "'Work Sans', sans-serif",
  fontSize: "0.9rem",
  background: "#fff",
  color: COLORS.ink,
  width: "100%",
};

const thStyle = { padding: "0.4rem 0.5rem", fontSize: "0.75rem", color: COLORS.sub, fontWeight: 600, whiteSpace: "nowrap" };
const thFiltroStyle = { padding: "0 0.5rem 0.5rem" };
const filtroInputStyle = { ...inputStyle, padding: "0.3rem 0.5rem", fontSize: "0.78rem", minWidth: "6rem" };
const tdStyle = { padding: "0.35rem 0.5rem", verticalAlign: "top" };
const cellInputStyle = { padding: "0.35rem 0.5rem", fontSize: "0.82rem", minWidth: "7rem" };

const btnStyle = (accent) => ({
  background: accent,
  color: "#fff",
  border: "none",
  borderRadius: "0.6rem",
  padding: "0.55rem 1rem",
  fontFamily: "'Work Sans', sans-serif",
  fontSize: "0.9rem",
  fontWeight: 500,
  cursor: "pointer",
  display: "flex",
  alignItems: "center",
  gap: "0.4rem",
});

const ghostBtnStyle = (accent) => ({
  background: "#fff",
  color: accent,
  border: `1px solid ${accent}`,
  borderRadius: "0.6rem",
  padding: "0.5rem 0.9rem",
  fontFamily: "'Work Sans', sans-serif",
  fontSize: "0.85rem",
  fontWeight: 500,
  cursor: "pointer",
  display: "flex",
  alignItems: "center",
  gap: "0.4rem",
});

/* ---------- Cuentas iniciales (cargadas desde Cuentas.xlsx) ----------
   Nota: "Sueldo" es un Activo (no Ingreso) porque se deja como saldo al
   inicio del mes y se va vaciando hacia "Cuenta" a medida que se gasta.
   "MUST" es un Pasivo porque se deja provisionado al partir el mes como
   obligación pendiente, ya que se cobrará durante ese mes. */
const DEFAULT_CUENTAS = [
  { nombre: "Auto", tipo: "gasto", descripcion: "Pejaes, bencina, seguro, etc" },
  { nombre: "Comida", tipo: "gasto", descripcion: "Comida de la rutina" },
  { nombre: "Compras", tipo: "gasto", descripcion: "Ropa, cosas, etc" },
  { nombre: "Depto", tipo: "gasto", descripcion: "Arriendo, ggcc, cuentas, etc" },
  { nombre: "Gastos Bancarios", tipo: "gasto", descripcion: "Seguro Banco, comisiones, intereses" },
  { nombre: "Gastos Fijos", tipo: "gasto", descripcion: "Celular, suscripciones, Must" },
  { nombre: "Ingresos", tipo: "ingreso", descripcion: "Sueldo u otros que no sean reembolsos" },
  { nombre: "Movilidad", tipo: "gasto", descripcion: "Transporte que no sea el auto" },
  { nombre: "Regalos / Eventos", tipo: "gasto", descripcion: "Regalos de cumpleaños, eventos de la oficina (Navidad)" },
  { nombre: "Salidas", tipo: "gasto", descripcion: "Antojitos, carrete, cine, salir a comer, etc" },
  { nombre: "Bienestar", tipo: "gasto", descripcion: "Farmacia, doctores, gimnasio, deporte, barbería" },
  { nombre: "Viajes / Fechas", tipo: "gasto", descripcion: "Viajes, 18 sept, etc" },
  { nombre: "BIP", tipo: "activo", descripcion: "Tarjeta BIP" },
  { nombre: "BIP QR", tipo: "activo", descripcion: "BIP QR Mercado Pago" },
  { nombre: "Credito", tipo: "pasivo", descripcion: "Credito de consumo" },
  { nombre: "Credito Auto", tipo: "pasivo", descripcion: "Credito para comprar el auto" },
  { nombre: "Cuenta", tipo: "activo", descripcion: "Cuenta Corriente" },
  { nombre: "Dasmi", tipo: "pasivo", descripcion: "Credito con Dasmi" },
  { nombre: "Linea de Credito", tipo: "pasivo", descripcion: "Linea de Credito de la cuenta" },
  { nombre: "MUST", tipo: "pasivo", descripcion: "Aporte de socio" },
  { nombre: "Por Cobrar", tipo: "activo", descripcion: "Reembolsos, tricount amigos, etc" },
  { nombre: "Por Pagar", tipo: "pasivo", descripcion: "Reembolsos, tricount amigos, etc" },
  { nombre: "S25", tipo: "pasivo", descripcion: "Cuotas del celular" },
  { nombre: "Sueldo", tipo: "activo", descripcion: "Sueldo" },
  { nombre: "Tarjeta", tipo: "pasivo", descripcion: "Tarjeta de crédito" },
  { nombre: "Tarjeta Lider", tipo: "pasivo", descripcion: "Tarjeta de crédito Lider" },
  { nombre: "Tarjeta USD", tipo: "pasivo", descripcion: "Tarjeta de crédito en USD" },
  { nombre: "Tempo", tipo: "activo", descripcion: "Cuenta Tempo" },
  { nombre: "Tricount", tipo: "activo", descripcion: "Tricount del departamento" },
].map((c, i) => ({ id: `seed-${i}`, estado: "activa", esMedio: false, ...c }));

/* ---------- Persistencia ---------- */
async function loadFinanzas(userId) {
  let data = { cuentas: [], movimientos: [], montosFijos: [], indicadores: null, fechasFacturacion: {} };
  try {
    const { data: row, error } = await supabase.from("app_data").select("data").eq("user_id", userId).maybeSingle();
    if (error) throw error;
    if (row?.data) data = { ...data, ...row.data };
  } catch (e) {
    console.error("No se pudo cargar", e);
  }
  if (!data.cuentas || data.cuentas.length === 0) {
    data.cuentas = DEFAULT_CUENTAS;
  }
  return data;
}
async function saveFinanzas(userId, data) {
  try {
    const { error } = await supabase.from("app_data").upsert({ user_id: userId, data, updated_at: new Date().toISOString() }, { onConflict: "user_id" });
    if (error) throw error;
  } catch (e) {
    console.error("No se pudo guardar", e);
  }
}

/* ---------- Indicadores económicos (USD / UF) ---------- */
async function fetchIndicadores() {
  // API pública oficial de indicadores económicos de Chile — no requiere
  // llave ni IA, funciona igual dentro de Claude y en la app ya desplegada.
  const response = await fetch("https://mindicador.cl/api");
  const data = await response.json();
  if (!data?.dolar?.valor || !data?.uf?.valor) throw new Error("No se pudo obtener el valor del día.");
  return {
    usd: Number(data.dolar.valor),
    uf: Number(data.uf.valor),
    fecha: (data.dolar.fecha || new Date().toISOString()).slice(0, 10),
    actualizado: new Date().toISOString(),
  };
}
function convertirACLP(monto, moneda, indicadores) {
  // Conversión automática usando el valor del día consultado en mindicador.cl
  // (guardado en "indicadores" en Datos). Si nunca se ha actualizado, el
  // equivalente en CLP puede ser 0 hasta que el usuario presione "Actualizar".
  if (moneda === "CLP") return monto;
  if (moneda === "UF") return monto * (indicadores?.uf || 0);
  if (moneda === "USD") return monto * (indicadores?.usd || 0);
  return monto;
}

/* ---------- Utilidades de fecha/periodo ---------- */
function getPeriodRange(tipo, anio, sub) {
  let startMonth, endMonthExclusive;
  if (tipo === "mes") {
    startMonth = sub - 1;
    endMonthExclusive = sub;
  } else if (tipo === "trimestre") {
    startMonth = (sub - 1) * 3;
    endMonthExclusive = startMonth + 3;
  } else if (tipo === "semestre") {
    startMonth = (sub - 1) * 6;
    endMonthExclusive = startMonth + 6;
  } else {
    startMonth = 0;
    endMonthExclusive = 12;
  }
  const start = new Date(anio, startMonth, 1);
  const end = new Date(anio, endMonthExclusive, 0);
  return { start, end };
}

function fmtMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("es-CL", { style: "currency", currency: "CLP", maximumFractionDigits: 0 });
}

const MESES_ABREV = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
// Formatea un Mes Impacto "YYYY-MM" como "ene/26" (abreviación estándar del
// mes + últimos 2 dígitos del año), para mostrarlo en toda la app.
function fmtMesImpacto(mesStr) {
  if (!mesStr) return "";
  const [y, m] = mesStr.split("-").map(Number);
  if (!y || !m || m < 1 || m > 12) return mesStr;
  return `${MESES_ABREV[m - 1]}/${String(y).slice(-2)}`;
}

/* ---------- Tabla de multiplicadores (hoja "Relación Cuentas") ----------
   El monto ingresado se aplica RESPETANDO SU SIGNO a la cuenta primaria
   (Medio, o Cuenta Balance si no hay Medio). La cuenta secundaria (Cuenta
   EERR, o Cuenta Balance en una transferencia) recibe el monto multiplicado
   según el tipo de ambas cuentas, para que el efecto contable sea correcto
   en ambos lados del movimiento. */
const MULT_TABLA = {
  "activo|activo": -1,
  "activo|pasivo": 1,
  "activo|ingreso": 1,
  "activo|gasto": -1,
  "pasivo|activo": 1,
  "pasivo|pasivo": -1,
  "pasivo|ingreso": -1,
  "pasivo|gasto": 1,
};
function multiplicador(tipoPrimaria, tipoSecundaria) {
  return MULT_TABLA[`${tipoPrimaria}|${tipoSecundaria}`] ?? 1;
}
// Determina qué campo del movimiento actúa como cuenta primaria/secundaria
// según su tipo de movimiento (ver TIPOS_MOVIMIENTO más abajo).
function primariaDe(m) {
  if (m.tipoMov === "ajuste") return m.cuentaBalance;
  if (m.tipoMov === "saldo_inicial") return m.medio || m.cuentaBalance;
  return m.medio;
}
function secundariaDe(m) {
  if (m.tipoMov === "transferencia") return m.cuentaBalance;
  if (m.tipoMov === "saldo_inicial") return null;
  return m.cuentaEERR;
}
// Dado un movimiento (tipoMov + medio/cuentaBalance/cuentaEERR) y el monto
// ya en CLP, calcula el monto que le corresponde a la cuenta primaria (tal
// cual, respetando el signo) y a la secundaria (multiplicado según su tipo).
function calcularMontos(tipoMov, medio, cuentaBalance, cuentaEERR, montoCLP, cuentas) {
  const tipoLookup = {};
  cuentas.forEach((c) => (tipoLookup[c.nombre] = c.tipo));
  const m = { tipoMov, medio, cuentaBalance, cuentaEERR };
  const p = primariaDe(m);
  const s = secundariaDe(m);
  const tp = p ? tipoLookup[p] : null;
  const ts = s ? tipoLookup[s] : null;
  const montoSecundario = s && tp && ts ? montoCLP * multiplicador(tp, ts) : null;
  return { monto: montoCLP, montoSecundario };
}

// Cuentas de medio que admiten "pago en cuotas" al agregar un movimiento.
const CUENTAS_CUOTAS = ["Tarjeta", "Tarjeta Lider", "Tarjeta USD"];

// Suma n meses a un string "YYYY-MM".
function sumarMeses(mesStr, n) {
  const [y, m] = mesStr.split("-").map(Number);
  const total = y * 12 + (m - 1) + n;
  const nuevoY = Math.floor(total / 12);
  const nuevoM = (total % 12) + 1;
  return `${nuevoY}-${String(nuevoM).padStart(2, "0")}`;
}

// Determina el Mes Impacto de un movimiento de una tarjeta según su fecha
// real y la Fecha de Facturación de esa tarjeta para el mes calendario de la
// fecha: si la fecha es posterior a la fecha de facturación de ese mes, el
// movimiento cae en el mes siguiente; si coincide o es anterior, cae en ese
// mismo mes (el límite está incluido). Si no hay fecha de facturación
// definida para ese mes, se usa el mes calendario de la fecha tal cual.
function mesImpactoPorFacturacion(tarjeta, fecha, fechasFacturacion) {
  const n = fecha.slice(0, 7);
  const cutoff = fechasFacturacion?.[`${tarjeta}|${n}`];
  if (!cutoff) return n;
  return fecha <= cutoff ? n : sumarMeses(n, 1);
}

// Si un movimiento de tarjeta quedó con Mes Impacto = N pero su fecha real
// es posterior a la Fecha de Facturación de ese mes N, para efectos de EERR
// y Facturación (NO para Balance, que sigue usando el Mes Impacto tal cual
// quedó guardado) se considera que corresponde al mes N+1.
function mesImpactoEfectivo(m, fechasFacturacion) {
  if (!fechasFacturacion || !m.mesImpacto || !m.fecha) return m.mesImpacto;
  const tarjeta = m.tipoMov === "transferencia" ? m.cuentaBalance : m.medio;
  if (!tarjeta || !CUENTAS_CUOTAS.includes(tarjeta)) return m.mesImpacto;
  const cutoff = fechasFacturacion[`${tarjeta}|${m.mesImpacto}`];
  if (cutoff && m.fecha > cutoff) return sumarMeses(m.mesImpacto, 1);
  return m.mesImpacto;
}

// Devuelve una copia de "movimientos" con el Mes Impacto reemplazado por su
// valor efectivo (ver mesImpactoEfectivo) — para usar en EERR/Facturación.
function conMesImpactoEfectivo(movimientos, fechasFacturacion) {
  if (!fechasFacturacion) return movimientos;
  return movimientos.map((m) => {
    const efectivo = mesImpactoEfectivo(m, fechasFacturacion);
    return efectivo === m.mesImpacto ? m : { ...m, mesImpacto: efectivo };
  });
}

// Diferencia en meses entre dos strings "YYYY-MM" (b - a).
function diffMeses(a, b) {
  const [ya, ma] = a.split("-").map(Number);
  const [yb, mb] = b.split("-").map(Number);
  return yb * 12 + mb - (ya * 12 + ma);
}

// Cuánto de una compra en cuotas (con su calendario en "cronograma") se
// factura en un Mes Impacto dado — 0 si ese mes no está en su calendario.
function montoCuotaEnMes(m, mes) {
  if (!m.esCuota) return 0;
  if (m.cronograma) {
    const entry = m.cronograma.find((c) => c.mes === mes);
    return entry ? entry.monto : 0;
  }
  // Sin cronograma (por ejemplo, una fila reclasificada a mano a "Cuotas"
  // desde el desplegable de Transacción): se cuenta como un cargo de un
  // solo mes, en su propio Mes Impacto — igual que se ve en su fila.
  return m.mesImpacto === mes ? m.monto : 0;
}

// Cuánto del ajuste de un refinanciamiento (cronogramaRefinanciamiento) cae
// en un Mes Impacto dado — 0 si ese mes no está en su cronograma.
function montoAjusteRefinanciamientoEnMes(m, mes) {
  if (!m.cronogramaRefinanciamiento) return 0;
  const entry = m.cronogramaRefinanciamiento.find((c) => c.mes === mes);
  return entry ? entry.monto : 0;
}

// Clasifica un movimiento dentro de la facturación de una tarjeta (mismo
// criterio que usa FacturacionTab).
function clasificarFacturacion(m, tarjeta) {
  if (m.esAjusteRefinanciamiento) return null;
  if (m.tipoMov === "transferencia" && m.cuentaBalance === tarjeta) return "Pago";
  if (m.medio !== tarjeta) return null;
  if (m.esPago) return "Pago";
  if (m.esComision) return "Comisiones";
  if (m.esRefinanciamiento) return "Refinanciamiento";
  if (m.esCuota) return "Cuotas";
  if (m.tipoMov === "gasto_ingreso") return "Gasto";
  return null;
}

// Calcula "Saldo Cuotas" de una tarjeta en un Mes Impacto: desde ese mes
// hasta el último mes con saldo (cuotas nuevas de este mes + saldos de
// cuotas futuras cargados a mano, incluidas las anulaciones por
// refinanciamiento, que son saldos negativos del mismo tipo).
function calcularProyeccionCuotas(tarjeta, mes, movimientos) {
  const mapaMeses = {};
  // Saldo Cuotas(mes) = Total(mes-1): acumula todo lo que ya se sabía desde
  // ANTES de este mes y que aún no llega — saldos manuales cargados a mano,
  // más el resto del cronograma de compras en cuotas ingresadas en meses
  // anteriores (sus filas ya se mostraron en su propio mes de entrada), más
  // ajustes de refinanciamiento de meses anteriores. Lo que se ingresa EN
  // ESTE MES no entra aquí — eso se ve en su propia fila, no en Saldo Cuotas.
  const saldosManuales = movimientos.filter(
    (m) => m.tipoMov === "saldo_inicial" && m.facturacionTipo === "cuotasFuturas" && m.mesImpacto <= mes && m.facturacionMesDestino >= mes && primariaDe(m) === tarjeta
  );
  saldosManuales.forEach((m) => {
    mapaMeses[m.facturacionMesDestino] = (mapaMeses[m.facturacionMesDestino] || 0) + m.monto;
  });

  const cuotasPendientes = movimientos.filter((m) => m.esCuota && m.medio === tarjeta && m.mesImpacto < mes && m.cronograma);
  cuotasPendientes.forEach((m) => {
    m.cronograma.forEach((c) => {
      if (c.mes >= mes) mapaMeses[c.mes] = (mapaMeses[c.mes] || 0) + c.monto;
    });
  });

  const refinanciamientosPendientes = movimientos.filter((m) => m.medio === tarjeta && m.mesImpacto < mes && m.cronogramaRefinanciamiento);
  refinanciamientosPendientes.forEach((m) => {
    m.cronogramaRefinanciamiento.forEach((c) => {
      if (c.mes >= mes) mapaMeses[c.mes] = (mapaMeses[c.mes] || 0) + c.monto;
    });
  });

  if (!Object.keys(mapaMeses).length) return [];
  const maxMes = Object.keys(mapaMeses).reduce((max, mm) => (mm > max ? mm : max), mes);
  const totalMeses = diffMeses(mes, maxMes) + 1;
  return Array.from({ length: totalMeses }, (_, k) => {
    const mesCol = sumarMeses(mes, k);
    return { mes: mesCol, monto: mapaMeses[mesCol] || 0 };
  });
}

// Calcula el Total facturado de una tarjeta en un Mes Impacto. "Facturado Mes
// Pasado" se resuelve así: si hay un saldo inicial manual marcado como "mes
// pasado" para ESTE mes, se usa ese valor (típicamente el mes en que se
// empezó a trackear la tarjeta). Si no, y el mes anterior sí tiene datos
// trackeados, el Facturado Mes Pasado de este mes es automáticamente el
// Total del mes anterior (encadenado recursivamente hacia atrás).
function calcularFacturacionMes(tarjeta, mes, movimientos, profundidad) {
  profundidad = profundidad || 0;
  if (profundidad > 240) return { facturadoMesPasado: 0, totalGasto: 0, totalComisiones: 0, totalCuotas: 0, totalPago: 0, totalRefinanciamiento: 0, total: 0 };

  const huboManual = movimientos.some((m) => m.tipoMov === "saldo_inicial" && m.facturacionTipo === "mesPasado" && m.mesImpacto === mes && primariaDe(m) === tarjeta);
  let facturadoMesPasado;
  if (huboManual) {
    facturadoMesPasado = movimientos
      .filter((m) => m.tipoMov === "saldo_inicial" && m.facturacionTipo === "mesPasado" && m.mesImpacto === mes && primariaDe(m) === tarjeta)
      .reduce((a, m) => a + m.monto, 0);
  } else {
    // Se encadena con el mes anterior aunque ESE mes puntual no tenga
    // movimientos propios (su Total puede venir, a su vez, encadenado de
    // más atrás) — pero para no recursar sin necesidad, se corta la cadena
    // apenas se pasa del primer mes con CUALQUIER dato real de esta
    // tarjeta (antes de eso, todo sería $0 de todas formas).
    const mesAnterior = sumarMeses(mes, -1);
    const primerMesConDatos = movimientos.reduce((min, m) => {
      const involucra = m.medio === tarjeta || (m.tipoMov === "transferencia" && m.cuentaBalance === tarjeta) || (m.tipoMov === "saldo_inicial" && primariaDe(m) === tarjeta);
      if (!involucra || !m.mesImpacto) return min;
      return !min || m.mesImpacto < min ? m.mesImpacto : min;
    }, null);
    facturadoMesPasado = primerMesConDatos && mesAnterior >= primerMesConDatos ? calcularFacturacionMes(tarjeta, mesAnterior, movimientos, profundidad + 1).total : 0;
  }

  // El resumen de Totales solo debe reflejar lo que realmente corresponde
  // cobrar EN ESTE MES (según el Mes Impacto real de cada movimiento) — una
  // compra en cuotas cuya primera cuota cae en un mes futuro no debe sumar
  // nada acá todavía, aunque su fila se muestre agrupada en el mes en que
  // se ingresó (eso es solo para la tabla, no para este resumen).
  const movsDelMes = movimientos.filter((m) => m.mesImpacto === mes && clasificarFacturacion(m, tarjeta));
  const totalGasto = movsDelMes.filter((m) => clasificarFacturacion(m, tarjeta) === "Gasto").reduce((a, m) => a + m.monto, 0);
  const totalComisiones = movsDelMes.filter((m) => clasificarFacturacion(m, tarjeta) === "Comisiones").reduce((a, m) => a + m.monto, 0);
  const totalRefinanciamiento = movsDelMes.filter((m) => clasificarFacturacion(m, tarjeta) === "Refinanciamiento").reduce((a, m) => a + m.monto, 0);
  const saldoCuotasEsteMes = movimientos
    .filter((m) => m.tipoMov === "saldo_inicial" && m.facturacionTipo === "cuotasFuturas" && m.facturacionMesDestino === mes && primariaDe(m) === tarjeta)
    .reduce((a, m) => a + m.monto, 0);
  // Cuotas se calcula sobre TODOS los movimientos en cuotas de la tarjeta
  // (no solo los con Mes Impacto = mes), usando lo que su cronograma indica
  // para este mes puntual — así una compra ingresada en enero con primera
  // cuota en marzo no suma nada en enero ni en febrero, solo desde marzo.
  // También se restan acá los ajustes de refinanciamientos anteriores que
  // apunten a este mes (su cronogramaRefinanciamiento), ya que cancelan
  // cuotas que de otra forma se seguirían cobrando.
  const ajustesRefinanciamientoEsteMes = movimientos
    .filter((m) => m.medio === tarjeta)
    .reduce((a, m) => a + montoAjusteRefinanciamientoEnMes(m, mes), 0);
  const totalCuotas =
    movimientos.filter((m) => clasificarFacturacion(m, tarjeta) === "Cuotas").reduce((a, m) => a + montoCuotaEnMes(m, mes), 0) + saldoCuotasEsteMes + ajustesRefinanciamientoEsteMes;
  const totalPago = movsDelMes.filter((m) => clasificarFacturacion(m, tarjeta) === "Pago").reduce((a, m) => a + (m.tipoMov === "transferencia" ? (m.montoSecundario ?? m.monto) : m.monto), 0);
  const total = facturadoMesPasado + totalGasto + totalComisiones + totalCuotas + totalPago + totalRefinanciamiento;

  return { facturadoMesPasado, totalGasto, totalComisiones, totalCuotas, totalPago, totalRefinanciamiento, total };
}

// Reparte "total" en n partes enteras que suman exactamente "total": todas
// iguales (redondeadas) salvo la última, que absorbe el resto.
function splitEnCuotas(total, n) {
  const base = Math.round(total / n);
  const partes = Array(n).fill(base);
  partes[n - 1] = total - base * (n - 1);
  return partes;
}

// Liquida un pago de una tarjeta en USD: marca como pagados (de más antiguo
// a más nuevo) los movimientos en USD sin pagar hasta sumar EXACTAMENTE
// usdAPagar, y los re-precifica al tipo de cambio real implícito del pago
// (montoClpPagado / usdAPagar). Si el último movimiento tocado no calza
// exacto, se parte en dos: la parte que completa el pago (pagada, con el
// tipo de cambio real) y el resto (sin pagar, con su tipo de cambio original).
// Devuelve { movimientos: nuevoArreglo } o { error: "mensaje" } si no hay
// suficiente USD sin pagar para cubrir el monto indicado.
function liquidarPagoUSD(movimientos, cuentas, tarjeta, usdAPagar, montoClpPagado, fecha, liquidacionId) {
  const disponible = movimientos.filter((m) => m.medio === tarjeta && m.montoUSD != null && !m.pagado).reduce((a, m) => a + m.montoUSD, 0);
  if (usdAPagar > disponible + 0.005) {
    return { error: `Solo hay US$${disponible.toLocaleString("es-CL")} sin pagar en ${tarjeta}, no se puede pagar US$${usdAPagar.toLocaleString("es-CL")}.` };
  }
  const impliedRate = montoClpPagado / usdAPagar;
  const candidatos = movimientos.filter((m) => m.medio === tarjeta && m.montoUSD != null && !m.pagado).sort((a, b) => (a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0));

  function recomputarSecundario(m, nuevoMonto) {
    const tipoPrimaria = cuentas.find((c) => c.nombre === m.medio)?.tipo;
    const tipoSecundaria = m.cuentaEERR ? cuentas.find((c) => c.nombre === m.cuentaEERR)?.tipo : null;
    return tipoPrimaria && tipoSecundaria ? nuevoMonto * multiplicador(tipoPrimaria, tipoSecundaria) : null;
  }

  let restante = usdAPagar;
  const actualizados = {};
  const nuevosSplits = [];
  for (const m of candidatos) {
    if (restante <= 0.005) break;
    // Se guarda el estado previo a esta liquidación (preLiquidacion) para
    // poder revertirlo exacto si más adelante se borra el Pago que la generó.
    const preLiquidacion = { monto: m.monto, montoSecundario: m.montoSecundario, montoUSD: m.montoUSD };
    if (m.montoUSD <= restante + 0.005) {
      const nuevoMonto = Math.round(m.montoUSD * impliedRate);
      actualizados[m.id] = { ...m, pagado: true, monto: nuevoMonto, montoSecundario: recomputarSecundario(m, nuevoMonto), liquidacionId, preLiquidacion };
      restante -= m.montoUSD;
    } else {
      const usdPagadoAqui = restante;
      const usdRestante = m.montoUSD - restante;
      const tasaOriginal = m.monto / m.montoUSD;
      const montoParteFinal = Math.round(usdRestante * tasaOriginal);
      // El id original se queda con la parte PAGADA (para que preLiquidacion
      // represente el movimiento completo antes de partirse); el "resto" es
      // un movimiento nuevo, sin pagar, que apunta de vuelta con splitDeId.
      nuevosSplits.push({ ...m, id: m.id + "-resto-" + Date.now(), montoUSD: usdRestante, monto: montoParteFinal, montoSecundario: recomputarSecundario(m, montoParteFinal), pagado: false, splitDeId: m.id });
      const montoPagado = Math.round(usdPagadoAqui * impliedRate);
      actualizados[m.id] = { ...m, montoUSD: usdPagadoAqui, monto: montoPagado, montoSecundario: recomputarSecundario(m, montoPagado), pagado: true, liquidacionId, preLiquidacion };
      restante = 0;
    }
  }
  const nuevosMovimientos = movimientos.map((m) => actualizados[m.id] || m).concat(nuevosSplits);
  return { movimientos: nuevosMovimientos };
}

// Revierte todo lo que una liquidación de pago en USD había marcado como
// pagado/re-precificado (ver liquidarPagoUSD), para cuando se borra el Pago
// que la generó. Cada movimiento afectado vuelve a su estado previo
// (preLiquidacion); si fue partido en dos, se fusiona de vuelta en uno solo.
function revertirLiquidacionUSD(movimientos, liquidacionId) {
  const afectados = movimientos.filter((m) => m.liquidacionId === liquidacionId);
  if (!afectados.length) return movimientos;
  const idsAEliminar = new Set();
  const restaurados = {};
  afectados.forEach((m) => {
    const resto = movimientos.find((x) => x.splitDeId === m.id);
    if (resto) idsAEliminar.add(resto.id);
    restaurados[m.id] = { ...m, ...m.preLiquidacion, pagado: false, liquidacionId: null, preLiquidacion: null };
  });
  return movimientos.filter((m) => !idsAEliminar.has(m.id)).map((m) => restaurados[m.id] || m);
}

/* ================= HOME ================= */
const MODULES = [
  { id: "finanzas", name: "Finanzas", icon: Wallet, accent: COLORS.finanzas },
  { id: "deportes", name: "Deportes", icon: Dumbbell, accent: COLORS.deportes },
];

function Home({ onOpen }) {
  return (
    <div className="w-full max-w-md">
      <p style={{ fontFamily: "'Fraunces', serif", fontSize: "1.5rem", fontWeight: 600, marginBottom: "2rem" }}>
        Mi día a día
      </p>
      <div className="grid grid-cols-2 gap-4">
        {MODULES.map((mod) => {
          const Icon = mod.icon;
          return (
            <button
              key={mod.id}
              onClick={() => onOpen(mod.id)}
              className="aspect-square rounded-2xl flex flex-col items-center justify-center gap-3"
              style={{ background: COLORS.card, border: "2px solid transparent", boxShadow: "0 1px 3px rgba(31,42,36,0.08)" }}
            >
              <Icon size={40} color={mod.accent} strokeWidth={1.75} />
              <span style={{ fontFamily: "'Fraunces', serif", fontSize: "1.15rem", fontWeight: 500, color: mod.accent }}>
                {mod.name}
              </span>
            </button>
          );
        })}
      </div>
    </div>
  );
}

/* ================= FINANZAS ================= */
const FIN_TABS_TOP = [
  { id: "resultados", name: "Resultados", icon: BarChart3 },
  { id: "analisis", name: "Análisis de Gastos", icon: PieChartIcon },
  { id: "facturacion", name: "Facturación", icon: FileText },
];
const FIN_TABS_BOTTOM = [
  { id: "detalle", name: "Detalle", icon: List },
  { id: "datos", name: "Datos", icon: Database },
  { id: "sueldo", name: "Sueldo", icon: Banknote },
];

function Finanzas({ onBack, userId }) {
  const [tab, setTab] = useState("resultados");
  const [cuentas, setCuentas] = useState([]);
  const [movimientos, setMovimientos] = useState([]);
  const [montosFijos, setMontosFijos] = useState([]);
  const [indicadores, setIndicadores] = useState(null);
  const [fechasFacturacion, setFechasFacturacion] = useState({});
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    loadFinanzas(userId).then((d) => {
      setCuentas(d.cuentas || []);
      setMovimientos(d.movimientos || []);
      setMontosFijos(d.montosFijos || []);
      setIndicadores(d.indicadores || null);
      setFechasFacturacion(d.fechasFacturacion || {});
      setLoaded(true);
    });
  }, [userId]);

  useEffect(() => {
    if (loaded) saveFinanzas(userId, { cuentas, movimientos, montosFijos, indicadores, fechasFacturacion });
  }, [cuentas, movimientos, montosFijos, indicadores, fechasFacturacion, loaded, userId]);

  return (
    <div className="w-full max-w-2xl">
      <button onClick={onBack} className="flex items-center gap-1 mb-4" style={{ color: COLORS.sub, background: "none", border: "none", cursor: "pointer", fontFamily: "'Work Sans', sans-serif" }}>
        <ArrowLeft size={18} /> Inicio
      </button>

      <p style={{ fontFamily: "'Fraunces', serif", fontSize: "1.4rem", fontWeight: 600, color: COLORS.finanzas, marginBottom: "1.2rem" }}>
        Finanzas
      </p>

      <div className="flex gap-2 mb-2 flex-wrap">
        {FIN_TABS_TOP.map((t) => {
          const Icon = t.icon;
          const isActive = tab === t.id;
          return (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              className="flex items-center gap-2 rounded-full"
              style={{
                padding: "0.45rem 0.9rem",
                border: `1px solid ${isActive ? COLORS.finanzas : "#DAD6C9"}`,
                background: isActive ? COLORS.finanzas : "#fff",
                color: isActive ? "#fff" : COLORS.ink,
                fontFamily: "'Work Sans', sans-serif",
                fontSize: "0.85rem",
                cursor: "pointer",
              }}
            >
              <Icon size={15} /> {t.name}
            </button>
          );
        })}
      </div>
      <div className="flex gap-2 mb-6 flex-wrap">
        {FIN_TABS_BOTTOM.map((t) => {
          const Icon = t.icon;
          const isActive = tab === t.id;
          return (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              className="flex items-center gap-2 rounded-full"
              style={{
                padding: "0.45rem 0.9rem",
                border: `1px solid ${isActive ? COLORS.finanzas : "#DAD6C9"}`,
                background: isActive ? COLORS.finanzas : "#fff",
                color: isActive ? "#fff" : COLORS.ink,
                fontFamily: "'Work Sans', sans-serif",
                fontSize: "0.85rem",
                cursor: "pointer",
              }}
            >
              <Icon size={15} /> {t.name}
            </button>
          );
        })}
      </div>

      {tab === "datos" && (
        <DatosTab
          cuentas={cuentas}
          setCuentas={setCuentas}
          montosFijos={montosFijos}
          setMontosFijos={setMontosFijos}
          indicadores={indicadores}
          setIndicadores={setIndicadores}
          movimientos={movimientos}
          setMovimientos={setMovimientos}
          fechasFacturacion={fechasFacturacion}
          setFechasFacturacion={setFechasFacturacion}
        />
      )}
      {tab === "detalle" && <DetalleTab cuentas={cuentas} setCuentas={setCuentas} movimientos={movimientos} setMovimientos={setMovimientos} indicadores={indicadores} fechasFacturacion={fechasFacturacion} />}
      {tab === "resultados" && <ResultadosTab cuentas={cuentas} movimientos={movimientos} fechasFacturacion={fechasFacturacion} />}
      {tab === "analisis" && <AnalisisTab cuentas={cuentas} movimientos={movimientos} />}
      {tab === "sueldo" && <SueldoTab />}
      {tab === "facturacion" && <FacturacionTab cuentas={cuentas} movimientos={movimientos} setMovimientos={setMovimientos} fechasFacturacion={fechasFacturacion} setFechasFacturacion={setFechasFacturacion} />}
    </div>
  );
}

/* ---- 1. Datos ---- */
function DatosTab({ cuentas, setCuentas, montosFijos, setMontosFijos, indicadores, setIndicadores, movimientos, setMovimientos, fechasFacturacion, setFechasFacturacion }) {
  const [nombre, setNombre] = useState("");
  const [tipo, setTipo] = useState("activo");
  const [descripcion, setDescripcion] = useState("");
  const [esMedio, setEsMedio] = useState(false);
  const [subview, setSubview] = useState("activas");
  const [actualizando, setActualizando] = useState(false);
  const [errorIndicadores, setErrorIndicadores] = useState("");
  const [editandoNombreId, setEditandoNombreId] = useState(null);
  const [nombreTemp, setNombreTemp] = useState("");
  const [descripcionTemp, setDescripcionTemp] = useState("");
  const [errorEliminarId, setErrorEliminarId] = useState(null);

  function addCuenta() {
    if (!nombre.trim()) return;
    setCuentas([...cuentas, { id: Date.now().toString(), nombre: nombre.trim(), tipo, descripcion: descripcion.trim(), estado: "activa", esMedio }]);
    setNombre("");
    setDescripcion("");
    setEsMedio(false);
  }
  function toggleMedio(id) {
    setCuentas(cuentas.map((c) => (c.id === id ? { ...c, esMedio: !c.esMedio } : c)));
  }
  // Cuántos movimientos referencian esta cuenta (por nombre, en Medio,
  // Cuenta Balance o Cuenta EERR) — si hay alguno, no se puede eliminar.
  function movimientosDeCuenta(nombreCuenta) {
    return movimientos.filter((m) => m.medio === nombreCuenta || m.cuentaBalance === nombreCuenta || m.cuentaEERR === nombreCuenta).length;
  }
  function removeCuenta(id) {
    const cuenta = cuentas.find((c) => c.id === id);
    const cantidad = cuenta ? movimientosDeCuenta(cuenta.nombre) : 0;
    if (cantidad > 0) {
      setErrorEliminarId({ id, cantidad });
      return;
    }
    setErrorEliminarId(null);
    setCuentas(cuentas.filter((c) => c.id !== id));
  }
  function archivarCuenta(id) {
    const hoy = new Date().toISOString().slice(0, 10);
    setCuentas(cuentas.map((c) => (c.id === id ? { ...c, estado: "inactiva", fechaInactivacion: hoy } : c)));
  }
  function reactivarCuenta(id) {
    setCuentas(cuentas.map((c) => (c.id === id ? { ...c, estado: "activa", fechaInactivacion: null } : c)));
  }

  // Renombrar (y/o editar la descripción de) una cuenta debe reflejarse en
  // todas las hojas: los movimientos la referencian por nombre (no por id),
  // así que hay que actualizar Medio, Cuenta Balance y Cuenta EERR en
  // Detalle, y las llaves de Fecha de Facturación en Facturación (que están
  // guardadas como "nombre|mes").
  function guardarEdicionCuenta(id, nuevoNombre, nuevaDescripcion) {
    const nombreNuevo = nuevoNombre.trim();
    if (!nombreNuevo) return;
    const cuenta = cuentas.find((c) => c.id === id);
    if (!cuenta) return;
    const nombreAnterior = cuenta.nombre;

    setCuentas(cuentas.map((c) => (c.id === id ? { ...c, nombre: nombreNuevo, descripcion: nuevaDescripcion.trim() } : c)));

    if (nombreAnterior === nombreNuevo) return;

    setMovimientos(
      movimientos.map((m) => ({
        ...m,
        medio: m.medio === nombreAnterior ? nombreNuevo : m.medio,
        cuentaBalance: m.cuentaBalance === nombreAnterior ? nombreNuevo : m.cuentaBalance,
        cuentaEERR: m.cuentaEERR === nombreAnterior ? nombreNuevo : m.cuentaEERR,
      }))
    );

    if (fechasFacturacion && Object.keys(fechasFacturacion).some((k) => k.startsWith(`${nombreAnterior}|`))) {
      const nuevasFechas = {};
      Object.entries(fechasFacturacion).forEach(([key, val]) => {
        const [tarjetaKey, mesKey] = key.split("|");
        const nuevaKey = tarjetaKey === nombreAnterior ? `${nombreNuevo}|${mesKey}` : key;
        nuevasFechas[nuevaKey] = val;
      });
      setFechasFacturacion(nuevasFechas);
    }
  }

  async function actualizarIndicadores() {
    setActualizando(true);
    setErrorIndicadores("");
    try {
      const nuevos = await fetchIndicadores();
      setIndicadores(nuevos);
    } catch (e) {
      setErrorIndicadores("No se pudo obtener el valor del día. Revisa tu conexión e intenta de nuevo.");
    }
    setActualizando(false);
  }

  const cuentasActivas = cuentas.filter((c) => c.estado !== "inactiva");
  const cuentasPasadas = cuentas.filter((c) => c.estado === "inactiva");

  return (
    <div>
      {/* Valores del día: USD y UF */}
      <div className="rounded-2xl mb-5" style={{ background: COLORS.card, padding: "1rem 1.2rem" }}>
        <div className="flex items-center justify-between flex-wrap gap-2">
          <div className="flex items-center gap-5">
            <div>
              <div style={{ fontSize: "0.75rem", color: COLORS.sub }}>Dólar (USD)</div>
              <div style={{ fontSize: "1.05rem", fontWeight: 600 }}>{indicadores?.usd ? fmtMoney(indicadores.usd) : "—"}</div>
            </div>
            <div>
              <div style={{ fontSize: "0.75rem", color: COLORS.sub }}>UF</div>
              <div style={{ fontSize: "1.05rem", fontWeight: 600 }}>{indicadores?.uf ? fmtMoney(indicadores.uf) : "—"}</div>
            </div>
            {indicadores?.fecha && <div style={{ fontSize: "0.72rem", color: COLORS.sub }}>Valores del {indicadores.fecha}</div>}
          </div>
          <button style={ghostBtnStyle(COLORS.finanzas)} onClick={actualizarIndicadores} disabled={actualizando}>
            <RefreshCw size={15} /> {actualizando ? "Actualizando..." : "Actualizar"}
          </button>
        </div>
        {errorIndicadores && <p style={{ color: "#9C4A2E", fontSize: "0.8rem", marginTop: "0.5rem" }}>{errorIndicadores}</p>}
      </div>

      <div className="flex gap-2 mb-4 flex-wrap">
        <button
          onClick={() => setSubview("activas")}
          className="rounded-full"
          style={{
            padding: "0.4rem 0.8rem",
            border: `1px solid ${subview === "activas" ? COLORS.finanzas : "#DAD6C9"}`,
            background: subview === "activas" ? COLORS.finanzas : "#fff",
            color: subview === "activas" ? "#fff" : COLORS.ink,
            fontSize: "0.82rem",
            cursor: "pointer",
          }}
        >
          Cuentas
        </button>
        <button
          onClick={() => setSubview("fijos")}
          className="rounded-full"
          style={{
            padding: "0.4rem 0.8rem",
            border: `1px solid ${subview === "fijos" ? COLORS.finanzas : "#DAD6C9"}`,
            background: subview === "fijos" ? COLORS.finanzas : "#fff",
            color: subview === "fijos" ? "#fff" : COLORS.ink,
            fontSize: "0.82rem",
            cursor: "pointer",
          }}
        >
          Montos Fijos {montosFijos.length > 0 && `(${montosFijos.length})`}
        </button>
        <button
          onClick={() => setSubview("pasadas")}
          className="rounded-full"
          style={{
            padding: "0.4rem 0.8rem",
            border: `1px solid ${subview === "pasadas" ? COLORS.finanzas : "#DAD6C9"}`,
            background: subview === "pasadas" ? COLORS.finanzas : "#fff",
            color: subview === "pasadas" ? "#fff" : COLORS.ink,
            fontSize: "0.82rem",
            cursor: "pointer",
          }}
        >
          Cuentas Pasadas {cuentasPasadas.length > 0 && `(${cuentasPasadas.length})`}
        </button>
      </div>

      {subview === "activas" && (
        <>
          <p style={{ color: COLORS.sub, fontSize: "0.85rem", marginBottom: "1rem" }}>
            Clasifica cada cuenta, tarjeta o categoría como activo, pasivo, gasto o ingreso.
          </p>
          <div className="flex gap-2 mb-4 flex-wrap items-center">
            <input style={{ ...inputStyle, flex: 2 }} placeholder="Nombre de la cuenta (ej. Cuenta Corriente BCI)" value={nombre} onChange={(e) => setNombre(e.target.value)} />
            <input style={{ ...inputStyle, flex: 2 }} placeholder="Descripción (opcional)" value={descripcion} onChange={(e) => setDescripcion(e.target.value)} />
            <select style={{ ...inputStyle, flex: 1 }} value={tipo} onChange={(e) => setTipo(e.target.value)}>
              <option value="activo">Activo</option>
              <option value="pasivo">Pasivo</option>
              <option value="ingreso">Ingreso</option>
              <option value="gasto">Gasto</option>
            </select>
            <label className="flex items-center gap-1" style={{ fontSize: "0.8rem", color: COLORS.sub, whiteSpace: "nowrap" }}>
              <input type="checkbox" checked={esMedio} onChange={(e) => setEsMedio(e.target.checked)} /> Es medio de pago
            </label>
            <button style={btnStyle(COLORS.finanzas)} onClick={addCuenta}>
              <Plus size={16} /> Agregar
            </button>
          </div>

          {cuentasActivas.length === 0 ? (
            <p style={{ color: COLORS.sub, fontSize: "0.85rem" }}>Aún no defines cuentas.</p>
          ) : (
            <div className="flex flex-col gap-2">
              {cuentasActivas.map((c) => {
                const s = TIPO_STYLES[c.tipo];
                const editando = editandoNombreId === c.id;
                return (
                  <div key={c.id} className="rounded-xl" style={{ background: COLORS.card, padding: "0.6rem 0.9rem" }}>
                  <div className="flex items-center justify-between">
                    <div style={{ flex: 1 }}>
                      {editando ? (
                        <div className="flex flex-col gap-1">
                          <input
                            autoFocus
                            style={{ ...inputStyle, maxWidth: "18rem" }}
                            placeholder="Nombre"
                            value={nombreTemp}
                            onChange={(e) => setNombreTemp(e.target.value)}
                            onKeyDown={(e) => { if (e.key === "Escape") setEditandoNombreId(null); }}
                          />
                          <input
                            style={{ ...inputStyle, maxWidth: "18rem" }}
                            placeholder="Descripción"
                            value={descripcionTemp}
                            onChange={(e) => setDescripcionTemp(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") { guardarEdicionCuenta(c.id, nombreTemp, descripcionTemp); setEditandoNombreId(null); }
                              if (e.key === "Escape") setEditandoNombreId(null);
                            }}
                          />
                        </div>
                      ) : (
                        <div style={{ fontSize: "0.9rem" }}>{c.nombre}</div>
                      )}
                      {!editando && c.descripcion && <div style={{ fontSize: "0.75rem", color: COLORS.sub }}>{c.descripcion}</div>}
                    </div>
                    <div className="flex items-center gap-3">
                      <span style={{ background: s.bg, color: s.text, borderRadius: "999px", padding: "0.15rem 0.6rem", fontSize: "0.75rem" }}>{s.label}</span>
                      {c.esMedio && <span style={{ background: "#EDEAE0", color: COLORS.sub, borderRadius: "999px", padding: "0.15rem 0.6rem", fontSize: "0.75rem" }}>Medio</span>}
                      {editando ? (
                        <>
                          <button onClick={() => { guardarEdicionCuenta(c.id, nombreTemp, descripcionTemp); setEditandoNombreId(null); }} style={{ background: "none", border: "none", cursor: "pointer", color: "#3B6E4F" }}>
                            <Check size={16} />
                          </button>
                          <button onClick={() => setEditandoNombreId(null)} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}>
                            <X size={16} />
                          </button>
                        </>
                      ) : (
                        <button onClick={() => { setEditandoNombreId(c.id); setNombreTemp(c.nombre); setDescripcionTemp(c.descripcion || ""); }} title="Editar nombre / descripción" style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}>
                          <Pencil size={16} />
                        </button>
                      )}
                      <button onClick={() => toggleMedio(c.id)} title={c.esMedio ? "Quitar como medio de pago" : "Marcar como medio de pago"} style={{ background: "none", border: "none", cursor: "pointer", color: c.esMedio ? COLORS.finanzas : COLORS.sub }}>
                        <CreditCard size={16} />
                      </button>
                      <button onClick={() => archivarCuenta(c.id)} title="Marcar como inactiva" style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}>
                        <Archive size={16} />
                      </button>
                      <button onClick={() => removeCuenta(c.id)} title="Eliminar" style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}>
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </div>
                  {errorEliminarId?.id === c.id && (
                    <p style={{ color: "#9C4A2E", fontSize: "0.78rem", marginTop: "0.4rem" }}>
                      Hay {errorEliminarId.cantidad} movimiento(s) asociados a esta cuenta. Puede marcar la cuenta como inactiva, pero no se puede eliminar.
                    </p>
                  )}
                  </div>
                );
              })}
            </div>
          )}
        </>
      )}

      {subview === "fijos" && <MontosFijosTab montosFijos={montosFijos} setMontosFijos={setMontosFijos} indicadores={indicadores} />}

      {subview === "pasadas" && (
        <>
          <p style={{ color: COLORS.sub, fontSize: "0.85rem", marginBottom: "1rem" }}>
            Cuentas marcadas como inactivas. No aparecen al registrar movimientos nuevos, pero se conservan para tu historial.
          </p>
          {cuentasPasadas.length === 0 ? (
            <p style={{ color: COLORS.sub, fontSize: "0.85rem" }}>No tienes cuentas pasadas.</p>
          ) : (
            <div className="flex flex-col gap-2">
              {cuentasPasadas.map((c) => {
                const s = TIPO_STYLES[c.tipo];
                const editando = editandoNombreId === c.id;
                return (
                  <div key={c.id} className="rounded-xl" style={{ background: COLORS.card, padding: "0.6rem 0.9rem" }}>
                  <div className="flex items-center justify-between">
                    <div style={{ flex: 1 }}>
                      {editando ? (
                        <div className="flex flex-col gap-1">
                          <input
                            autoFocus
                            style={{ ...inputStyle, maxWidth: "18rem" }}
                            placeholder="Nombre"
                            value={nombreTemp}
                            onChange={(e) => setNombreTemp(e.target.value)}
                            onKeyDown={(e) => { if (e.key === "Escape") setEditandoNombreId(null); }}
                          />
                          <input
                            style={{ ...inputStyle, maxWidth: "18rem" }}
                            placeholder="Descripción"
                            value={descripcionTemp}
                            onChange={(e) => setDescripcionTemp(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") { guardarEdicionCuenta(c.id, nombreTemp, descripcionTemp); setEditandoNombreId(null); }
                              if (e.key === "Escape") setEditandoNombreId(null);
                            }}
                          />
                        </div>
                      ) : (
                        <div style={{ fontSize: "0.9rem" }}>{c.nombre}</div>
                      )}
                      {!editando && c.descripcion && <div style={{ fontSize: "0.75rem", color: COLORS.sub }}>{c.descripcion}</div>}
                      {c.fechaInactivacion && <div style={{ fontSize: "0.72rem", color: COLORS.sub }}>Inactiva desde {c.fechaInactivacion}</div>}
                    </div>
                    <div className="flex items-center gap-3">
                      <span style={{ background: s.bg, color: s.text, borderRadius: "999px", padding: "0.15rem 0.6rem", fontSize: "0.75rem" }}>{s.label}</span>
                      {editando ? (
                        <>
                          <button onClick={() => { guardarEdicionCuenta(c.id, nombreTemp, descripcionTemp); setEditandoNombreId(null); }} style={{ background: "none", border: "none", cursor: "pointer", color: "#3B6E4F" }}>
                            <Check size={16} />
                          </button>
                          <button onClick={() => setEditandoNombreId(null)} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}>
                            <X size={16} />
                          </button>
                        </>
                      ) : (
                        <button onClick={() => { setEditandoNombreId(c.id); setNombreTemp(c.nombre); setDescripcionTemp(c.descripcion || ""); }} title="Editar nombre / descripción" style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}>
                          <Pencil size={16} />
                        </button>
                      )}
                      <button onClick={() => reactivarCuenta(c.id)} title="Reactivar" style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}>
                        <RotateCcw size={16} />
                      </button>
                      <button onClick={() => removeCuenta(c.id)} title="Eliminar definitivamente" style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}>
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </div>
                  {errorEliminarId?.id === c.id && (
                    <p style={{ color: "#9C4A2E", fontSize: "0.78rem", marginTop: "0.4rem" }}>
                      Hay {errorEliminarId.cantidad} movimiento(s) asociados a esta cuenta. Puede marcar la cuenta como inactiva, pero no se puede eliminar.
                    </p>
                  )}
                  </div>
                );
              })}
            </div>
          )}
        </>
      )}
    </div>
  );
}

/* ---- 1b. Montos Fijos ---- */
function MontosFijosTab({ montosFijos, setMontosFijos, indicadores }) {
  const [nombre, setNombre] = useState("");
  const [monto, setMonto] = useState("");
  const [moneda, setMoneda] = useState("CLP");
  const [editId, setEditId] = useState(null);
  const [editTemp, setEditTemp] = useState(null);

  function addFijo() {
    if (!nombre.trim() || !monto) return;
    setMontosFijos([...montosFijos, { id: Date.now().toString(), nombre: nombre.trim(), monto: Number(monto), moneda }]);
    setNombre("");
    setMonto("");
  }
  function removeFijo(id) {
    setMontosFijos(montosFijos.filter((m) => m.id !== id));
  }
  function startEdit(m) {
    setEditId(m.id);
    setEditTemp({ ...m });
  }
  function saveEdit() {
    setMontosFijos(montosFijos.map((m) => (m.id === editId ? { ...editTemp, monto: Number(editTemp.monto) } : m)));
    setEditId(null);
    setEditTemp(null);
  }
  function cancelEdit() {
    setEditId(null);
    setEditTemp(null);
  }

  return (
    <div>
      <p style={{ color: COLORS.sub, fontSize: "0.85rem", marginBottom: "1rem" }}>
        Montos recurrentes (arriendo, suscripciones, cuotas, etc.), en CLP, UF o USD. Editables si cambian de valor.
      </p>
      <div className="flex gap-2 mb-4 flex-wrap">
        <input style={{ ...inputStyle, flex: 2 }} placeholder="Nombre (ej. Arriendo)" value={nombre} onChange={(e) => setNombre(e.target.value)} />
        <input type="number" style={{ ...inputStyle, flex: 1 }} placeholder="Monto" value={monto} onChange={(e) => setMonto(e.target.value)} />
        <select style={{ ...inputStyle, flex: 1 }} value={moneda} onChange={(e) => setMoneda(e.target.value)}>
          <option value="CLP">CLP</option>
          <option value="UF">UF</option>
          <option value="USD">USD</option>
        </select>
        <button style={btnStyle(COLORS.finanzas)} onClick={addFijo}>
          <Plus size={16} /> Agregar
        </button>
      </div>

      {montosFijos.length === 0 ? (
        <p style={{ color: COLORS.sub, fontSize: "0.85rem" }}>Aún no tienes montos fijos registrados.</p>
      ) : (
        <div className="flex flex-col gap-2">
          {montosFijos.map((m) => {
            const editando = editId === m.id;
            const equivalente = m.moneda !== "CLP" ? convertirACLP(m.monto, m.moneda, indicadores) : null;
            return (
              <div key={m.id} className="flex items-center justify-between rounded-xl flex-wrap gap-2" style={{ background: COLORS.card, padding: "0.6rem 0.9rem" }}>
                {editando ? (
                  <div className="flex items-center gap-2 flex-wrap" style={{ flex: 1 }}>
                    <input style={{ ...inputStyle, flex: 2 }} value={editTemp.nombre} onChange={(e) => setEditTemp({ ...editTemp, nombre: e.target.value })} />
                    <input type="number" style={{ ...inputStyle, width: "7rem" }} value={editTemp.monto} onChange={(e) => setEditTemp({ ...editTemp, monto: e.target.value })} />
                    <select style={{ ...inputStyle, width: "6rem" }} value={editTemp.moneda} onChange={(e) => setEditTemp({ ...editTemp, moneda: e.target.value })}>
                      <option value="CLP">CLP</option>
                      <option value="UF">UF</option>
                      <option value="USD">USD</option>
                    </select>
                    <button onClick={saveEdit} style={{ background: "none", border: "none", cursor: "pointer", color: "#3B6E4F" }}><Check size={17} /></button>
                    <button onClick={cancelEdit} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}><X size={17} /></button>
                  </div>
                ) : (
                  <>
                    <div>
                      <div style={{ fontSize: "0.9rem" }}>{m.nombre}</div>
                      <div style={{ fontSize: "0.75rem", color: COLORS.sub }}>
                        {m.moneda === "CLP" ? fmtMoney(m.monto) : `${m.monto.toLocaleString("es-CL")} ${m.moneda}`}
                        {equivalente ? ` · ≈ ${fmtMoney(equivalente)}` : ""}
                      </div>
                    </div>
                    <div className="flex items-center gap-3">
                      <button onClick={() => startEdit(m)} title="Editar" style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}>
                        <Pencil size={16} />
                      </button>
                      <button onClick={() => removeFijo(m.id)} title="Eliminar" style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}>
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ---- 2. Detalle ---- */
const TIPOS_MOVIMIENTO = [
  { id: "gasto_ingreso", label: "Gasto / Ingreso", campos: ["medio", "cuentaEERR"] },
  { id: "transferencia", label: "Transferencia entre cuentas", campos: ["medio", "cuentaBalance"] },
  { id: "ajuste", label: "Ajuste sin movimiento de dinero", campos: ["cuentaBalance", "cuentaEERR"] },
  { id: "saldo_inicial", label: "Saldo inicial de cuenta", campos: ["cuentaUnica"] },
];

// Una cuenta inactiva no puede recibir movimientos nuevos desde el día en
// que quedó inhabilitada — pero sí admite movimientos atrasados con fecha
// anterior a esa inactivación (por eso se filtra dinámicamente según "fecha").
function disponiblesPorFecha(lista, fecha) {
  return lista.filter((c) => c.estado !== "inactiva" || (c.fechaInactivacion && fecha < c.fechaInactivacion));
}

function DetalleTab({ cuentas, setCuentas, movimientos, setMovimientos, indicadores, fechasFacturacion }) {
  const [tipoMov, setTipoMov] = useState("gasto_ingreso");
  const [fecha, setFecha] = useState(new Date().toISOString().slice(0, 10));
  const [mesImpacto, setMesImpacto] = useState(new Date().toISOString().slice(0, 7));
  const [mesImpactoManual, setMesImpactoManual] = useState(false);
  const [divisa, setDivisa] = useState("CLP");
  const [medio, setMedio] = useState("");
  const [cuentaBalance, setCuentaBalance] = useState("");
  const [cuentaEERR, setCuentaEERR] = useState("");
  const [cuentaUnica, setCuentaUnica] = useState("");
  const [monto, setMonto] = useState("");
  const [usdPagados, setUsdPagados] = useState("");
  const usdSinPagarTarjetaUSD = useMemo(
    () => movimientos.filter((m) => m.medio === "Tarjeta USD" && m.montoUSD != null && !m.pagado).reduce((a, m) => a + m.montoUSD, 0),
    [movimientos]
  );
  const [descripcion, setDescripcion] = useState("");
  const [error, setError] = useState("");
  const [enCuotas, setEnCuotas] = useState(false);
  const [numCuotas, setNumCuotas] = useState("");
  const [mesPrimeraCuota, setMesPrimeraCuota] = useState("");
  const [mesPrimeraCuotaManual, setMesPrimeraCuotaManual] = useState(false);
  const [esComision, setEsComision] = useState(false);
  const [esRefinanciamiento, setEsRefinanciamiento] = useState(false);
  const [facturacionTipo, setFacturacionTipo] = useState("");
  const [facturacionMesDestino, setFacturacionMesDestino] = useState("");

  useEffect(() => {
    if (!CUENTAS_CUOTAS.includes(medio)) {
      setEnCuotas(false);
      setNumCuotas("");
      setMesPrimeraCuota("");
      setMesPrimeraCuotaManual(false);
      setEsComision(false);
      setEsRefinanciamiento(false);
    }
  }, [medio]);

  // Igual que Mes Impacto: si el usuario no toca "Mes primera cuota" a mano,
  // se sugiere según la fecha de facturación de la tarjeta (o el mes de la
  // fecha si no hay una definida), mientras "Pago en cuotas" esté activo.
  useEffect(() => {
    if (!enCuotas || mesPrimeraCuotaManual || !CUENTAS_CUOTAS.includes(medio)) return;
    setMesPrimeraCuota(mesImpactoPorFacturacion(medio, fecha, fechasFacturacion));
  }, [enCuotas, mesPrimeraCuotaManual, medio, fecha, fechasFacturacion]);

  useEffect(() => {
    if (!CUENTAS_CUOTAS.includes(cuentaUnica)) {
      setFacturacionTipo("");
      setFacturacionMesDestino("");
    }
  }, [cuentaUnica]);

  // Si el usuario no toca "Mes impacto" a mano, se sigue el mes de la fecha
  // — salvo que el Medio sea una tarjeta con Fecha de Facturación definida
  // para ese mes, en cuyo caso se usa esa fecha de corte para decidir si el
  // movimiento cae en ese mes o en el siguiente.
  useEffect(() => {
    if (mesImpactoManual) return;
    if (CUENTAS_CUOTAS.includes(medio)) setMesImpacto(mesImpactoPorFacturacion(medio, fecha, fechasFacturacion));
    else setMesImpacto(fecha.slice(0, 7));
  }, [fecha, mesImpactoManual, medio, fechasFacturacion]);

  const medios = useMemo(() => disponiblesPorFecha(cuentas.filter((c) => c.esMedio), fecha), [cuentas, fecha]);
  const balances = useMemo(() => disponiblesPorFecha(cuentas.filter((c) => c.tipo === "activo" || c.tipo === "pasivo"), fecha), [cuentas, fecha]);
  const eerr = useMemo(() => disponiblesPorFecha(cuentas.filter((c) => c.tipo === "ingreso" || c.tipo === "gasto"), fecha), [cuentas, fecha]);
  const poolUnica = useMemo(() => {
    const nombres = new Set();
    const lista = [];
    [...medios, ...balances].forEach((c) => {
      if (!nombres.has(c.nombre)) { nombres.add(c.nombre); lista.push(c); }
    });
    return lista;
  }, [medios, balances]);

  useEffect(() => { if (!medios.find((c) => c.nombre === medio)) setMedio(""); }, [medios]);
  useEffect(() => { if (!(tipoMov === "transferencia" && cuentaBalance === "Tarjeta USD")) setUsdPagados(""); }, [tipoMov, cuentaBalance]);
  useEffect(() => { if (!balances.find((c) => c.nombre === cuentaBalance)) setCuentaBalance(""); }, [balances]);
  useEffect(() => { if (!eerr.find((c) => c.nombre === cuentaEERR)) setCuentaEERR(""); }, [eerr]);
  useEffect(() => { if (!poolUnica.find((c) => c.nombre === cuentaUnica)) setCuentaUnica(""); }, [poolUnica]);

  const campos = TIPOS_MOVIMIENTO.find((t) => t.id === tipoMov).campos;

  function addMov() {
    setError("");
    if (!fecha || !monto) return;
    if (campos.includes("medio") && !medio) { setError("Falta seleccionar Medio."); return; }
    if (campos.includes("cuentaBalance") && !cuentaBalance) { setError("Falta seleccionar Cuenta Balance."); return; }
    if (campos.includes("cuentaEERR") && !cuentaEERR) { setError("Falta seleccionar Cuenta EERR."); return; }
    if (campos.includes("cuentaUnica") && !cuentaUnica) { setError("Falta seleccionar la cuenta del saldo inicial."); return; }
    if (enCuotas && (!numCuotas || Number(numCuotas) < 2 || !mesPrimeraCuota)) { setError("Para pago en cuotas, indica la cantidad de cuotas (2 o más) y el mes de la primera cuota."); return; }
    if (facturacionTipo === "cuotasFuturas" && (!facturacionMesDestino || facturacionMesDestino < mesImpacto)) { setError("Indica un Mes Impacto igual o posterior al de este saldo inicial para el saldo de cuotas próximas."); return; }
    if (tipoMov === "transferencia" && cuentaBalance === "Tarjeta USD" && usdPagados && Number(usdPagados) <= 0) { setError("El monto de USD que se están pagando debe ser mayor a 0."); return; }

    let cb = campos.includes("cuentaBalance") ? cuentaBalance : "";
    let cuentaUnicaObj = null;
    if (campos.includes("cuentaUnica")) {
      cuentaUnicaObj = poolUnica.find((c) => c.nombre === cuentaUnica);
      // Si la cuenta única es Medio de pago (ej. una tarjeta), se guarda solo
      // en el campo Medio; si no, se guarda solo en Cuenta Balance — nunca en
      // los dos, para que no aparezca duplicada.
      if (cuentaUnicaObj && !cuentaUnicaObj.esMedio && (cuentaUnicaObj.tipo === "activo" || cuentaUnicaObj.tipo === "pasivo")) cb = cuentaUnicaObj.nombre;
    }

    const montoUSD = divisa === "USD" ? Number(monto) : null;
    const montoCLP = divisa === "USD" ? Number(monto) * (indicadores?.usd || 0) : Number(monto);

    const nombrePrimaria = campos.includes("cuentaUnica")
      ? (cuentaUnicaObj && (cuentaUnicaObj.tipo === "activo" || cuentaUnicaObj.tipo === "pasivo") ? cuentaUnicaObj.nombre : cuentaUnica)
      : (campos.includes("medio") ? medio : cuentaBalance);
    const nombreSecundaria = campos.includes("cuentaEERR") ? cuentaEERR : (campos.includes("cuentaBalance") && campos.includes("medio") ? cuentaBalance : null);
    const tipoPrimaria = cuentas.find((c) => c.nombre === nombrePrimaria)?.tipo;
    const tipoSecundaria = nombreSecundaria ? cuentas.find((c) => c.nombre === nombreSecundaria)?.tipo : null;
    const montoSecundario = nombreSecundaria && tipoPrimaria && tipoSecundaria ? montoCLP * multiplicador(tipoPrimaria, tipoSecundaria) : null;
    const refiId = Date.now().toString() + "-refi";

    const base = {
      tipoMov,
      fecha,
      divisa,
      medio: campos.includes("medio") ? medio : (campos.includes("cuentaUnica") && cuentaUnicaObj?.esMedio ? cuentaUnica : ""),
      cuentaBalance: cb,
      cuentaEERR: campos.includes("cuentaEERR") ? cuentaEERR : "",
      esComision: CUENTAS_CUOTAS.includes(medio) ? esComision : false,
      esRefinanciamiento: CUENTAS_CUOTAS.includes(medio) ? esRefinanciamiento : false,
      refinanciamientoId: esRefinanciamiento && CUENTAS_CUOTAS.includes(medio) ? refiId : null,
      facturacionTipo: campos.includes("cuentaUnica") && CUENTAS_CUOTAS.includes(cuentaUnica) && facturacionTipo ? facturacionTipo : null,
      facturacionMesDestino: campos.includes("cuentaUnica") && CUENTAS_CUOTAS.includes(cuentaUnica) && facturacionTipo === "cuotasFuturas" ? facturacionMesDestino : null,
    };

    // Refinanciamiento: repactación de la deuda de la tarjeta. Se registra
    // el movimiento principal (con el monto ingresado) y UN ajuste agregado
    // (la suma de todo lo que Saldo Cuotas ya proyectaba para los meses
    // posteriores, en negativo) — este ajuste es un movimiento real, porque
    // reduce el pasivo de la tarjeta de inmediato. Guarda también el detalle
    // mes a mes en "cronogramaRefinanciamiento", que es lo que usa
    // Facturación para mostrarlo en la columna que corresponda, en la misma
    // fila del Refinanciamiento (vinculados por refinanciamientoId).
    let ajusteRefinanciamiento = null;
    if (esRefinanciamiento && CUENTAS_CUOTAS.includes(medio)) {
      const proyeccionActual = calcularProyeccionCuotas(medio, mesImpacto, movimientos);
      const cronogramaRefinanciamiento = proyeccionActual.slice(1).filter((p) => p.monto !== 0).map((p) => ({ mes: p.mes, monto: -p.monto }));
      if (cronogramaRefinanciamiento.length) {
        const totalAjuste = cronogramaRefinanciamiento.reduce((a, c) => a + c.monto, 0);
        const montoSecundarioAjuste = tipoPrimaria && tipoSecundaria ? totalAjuste * multiplicador(tipoPrimaria, tipoSecundaria) : null;
        ajusteRefinanciamiento = {
          id: refiId + "-ajuste",
          tipoMov: "gasto_ingreso",
          fecha,
          mesImpacto,
          divisa: "CLP",
          medio,
          cuentaBalance: "",
          cuentaEERR: base.cuentaEERR,
          monto: totalAjuste,
          montoSecundario: montoSecundarioAjuste,
          montoUSD: null,
          descripcion: "Refinanciamiento. ajuste Saldo Cuotas",
          esCuota: false,
          cuotaPlanId: null,
          cuotaTotal: null,
          cronograma: null,
          esComision: false,
          esRefinanciamiento: false,
          esAjusteRefinanciamiento: true,
          refinanciamientoId: refiId,
          cronogramaRefinanciamiento,
          facturacionTipo: null,
          facturacionMesDestino: null,
        };
      }
    }
    const extras = ajusteRefinanciamiento ? [ajusteRefinanciamiento] : [];

    if (enCuotas) {
      // Pago en cuotas: se registra UN solo movimiento con el MONTO TOTAL en
      // el Mes Impacto de la compra (así el Balance reconoce la deuda
      // completa de inmediato, no de a poco mes a mes). El calendario de
      // facturación (cuánto se cobra en cada mes futuro) queda guardado
      // como "cronograma" en ese mismo movimiento, y es lo que usa
      // Facturación para mostrar cada cuota en su columna correspondiente.
      const n = Math.round(Number(numCuotas));
      const planId = Date.now().toString();
      const partesCLP = splitEnCuotas(montoCLP, n);
      const cronograma = partesCLP.map((parte, i) => ({ mes: sumarMeses(mesPrimeraCuota, i), monto: parte }));
      setMovimientos([
        ...movimientos,
        {
          ...base,
          id: planId,
          mesImpacto,
          monto: montoCLP,
          montoSecundario,
          montoUSD,
          pagado: montoUSD != null ? false : null,
          descripcion,
          esCuota: true,
          cuotaPlanId: planId,
          cuotaTotal: n,
          cronograma,
        },
        ...extras,
      ]);
    } else {
      // Pago de Tarjeta USD indicando cuántos USD se están pagando: se
      // marcan como pagados (de más antiguo a más nuevo) los movimientos en
      // USD sin pagar hasta completar exactamente ese monto, re-precificados
      // al tipo de cambio real implícito de este pago (ver liquidarPagoUSD).
      if (tipoMov === "transferencia" && cuentaBalance === "Tarjeta USD" && usdPagados) {
        const liquidacionId = Date.now().toString() + "-liq";
        const resultado = liquidarPagoUSD(movimientos, cuentas, "Tarjeta USD", Number(usdPagados), montoCLP, fecha, liquidacionId);
        if (resultado.error) { setError(resultado.error); return; }
        setMovimientos([
          ...resultado.movimientos,
          { ...base, id: Date.now().toString(), mesImpacto, monto: montoCLP, montoSecundario, montoUSD, pagado: null, liquidacionId, descripcion, esCuota: false, cuotaPlanId: null, cuotaTotal: null, cronograma: null },
        ]);
      } else {
        setMovimientos([
          ...movimientos,
          { ...base, id: Date.now().toString(), mesImpacto, monto: montoCLP, montoSecundario, montoUSD, pagado: montoUSD != null ? false : null, descripcion, esCuota: false, cuotaPlanId: null, cuotaTotal: null, cronograma: null },
          ...extras,
        ]);
      }
    }
    setMonto("");
    setUsdPagados("");
    setDescripcion("");
    setEnCuotas(false);
    setNumCuotas("");
    setMesPrimeraCuota("");
    setMesPrimeraCuotaManual(false);
    setEsComision(false);
    setEsRefinanciamiento(false);
    setFacturacionTipo("");
    setFacturacionMesDestino("");
  }
  function removeMov(id) {
    const mov = movimientos.find((m) => m.id === id);
    const base = mov?.liquidacionId ? revertirLiquidacionUSD(movimientos, mov.liquidacionId) : movimientos;
    setMovimientos(base.filter((m) => m.id !== id));
  }

  // Ordenar y filtrar son operaciones caras si el detalle crece mucho — se
  // memorizan para que solo se recalculen cuando cambian los movimientos o
  // los filtros, no en cada render de la app.
  const ordenados = useMemo(() => [...movimientos].sort((a, b) => (a.fecha < b.fecha ? 1 : -1)), [movimientos]);

  // El rendimiento se resuelve con paginación (más abajo), no ocultando
  // datos por defecto — así el filtro parte vacío y muestra todo, como antes.
  const [filtros, setFiltros] = useState({ fecha: "", mesImpacto: "", cuentas: "", descripcion: "" });

  // Valores disponibles para los filtros desplegables de Fecha, Mes Impacto
  // y Cuentas, en base a lo que realmente existe en el detalle.
  const fechasDisponibles = useMemo(() => [...new Set(movimientos.map((m) => m.fecha).filter(Boolean))].sort((a, b) => (a < b ? 1 : -1)), [movimientos]);
  const mesesDisponibles = useMemo(() => [...new Set(movimientos.map((m) => m.mesImpacto).filter(Boolean))].sort((a, b) => (a < b ? 1 : -1)), [movimientos]);
  const cuentasDisponibles = useMemo(
    () => [...new Set(movimientos.flatMap((m) => [m.medio, m.cuentaBalance, m.cuentaEERR]).filter(Boolean))].sort(),
    [movimientos]
  );

  const filtrados = useMemo(
    () =>
      ordenados.filter(
        (m) =>
          (!filtros.fecha || m.fecha === filtros.fecha) &&
          (!filtros.mesImpacto || m.mesImpacto === filtros.mesImpacto) &&
          (!filtros.cuentas || [m.medio, m.cuentaBalance, m.cuentaEERR].includes(filtros.cuentas)) &&
          (!filtros.descripcion || (m.descripcion || "").toLowerCase().includes(filtros.descripcion.toLowerCase()))
      ),
    [ordenados, filtros]
  );

  // Se muestran de a 50 movimientos por página (no acumulados), para que la
  // pantalla no se ponga lenta aunque el detalle crezca mucho.
  const [pagina, setPagina] = useState(0);
  useEffect(() => { setPagina(0); }, [filtros]);
  const totalPaginas = Math.max(1, Math.ceil(filtrados.length / 50));
  useEffect(() => { if (pagina > totalPaginas - 1) setPagina(0); }, [totalPaginas]);
  const visibles = filtrados.slice(pagina * 50, pagina * 50 + 50);

  // Selección múltiple para borrar en lote — pensada para usarse junto con
  // los filtros (por ejemplo, filtrar por Cuenta o Descripción y borrar todo
  // lo que calzó). La selección se limpia si cambian los filtros.
  const [seleccionados, setSeleccionados] = useState(() => new Set());
  const [confirmandoBorrado, setConfirmandoBorrado] = useState(false);
  useEffect(() => { setSeleccionados(new Set()); setConfirmandoBorrado(false); }, [filtros]);

  function toggleSeleccion(id) {
    setSeleccionados((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }
  const todosFiltradosSeleccionados = filtrados.length > 0 && filtrados.every((m) => seleccionados.has(m.id));
  function toggleSeleccionarTodos() {
    setSeleccionados(todosFiltradosSeleccionados ? new Set() : new Set(filtrados.map((m) => m.id)));
  }
  // window.confirm queda bloqueado dentro del sandbox de artifacts (no
  // muestra el diálogo nativo), así que la confirmación se maneja acá mismo
  // con un paso extra en la propia interfaz, no con un diálogo del navegador.
  function borrarSeleccionados() {
    let base = movimientos;
    movimientos
      .filter((m) => seleccionados.has(m.id) && m.liquidacionId)
      .forEach((m) => { base = revertirLiquidacionUSD(base, m.liquidacionId); });
    setMovimientos(base.filter((m) => !seleccionados.has(m.id)));
    setSeleccionados(new Set());
    setConfirmandoBorrado(false);
  }

  // Pools de cuentas por tipo, calculados UNA sola vez (no por cada fila de
  // la tabla) — cada fila solo aplica el filtro liviano por fecha sobre
  // estas listas ya reducidas.
  const mediosBase = useMemo(() => cuentas.filter((c) => c.esMedio), [cuentas]);
  const balancesBase = useMemo(() => cuentas.filter((c) => c.tipo === "activo" || c.tipo === "pasivo"), [cuentas]);
  const eerrBase = useMemo(() => cuentas.filter((c) => c.tipo === "ingreso" || c.tipo === "gasto"), [cuentas]);

  const [mostrarExportar, setMostrarExportar] = useState(false);
  const [exportDesde, setExportDesde] = useState(new Date().toISOString().slice(0, 7));
  const [exportHasta, setExportHasta] = useState(new Date().toISOString().slice(0, 7));
  const [errorExportar, setErrorExportar] = useState("");

  // Template vacío con las mismas columnas exactas que reconoce Importe, para
  // que el usuario lo llene afuera y lo suba después en ese mismo formato.
  function descargarTemplate() {
    const ws = XLSX.utils.aoa_to_sheet([["Fecha", "Mes Impacto", "Medio", "Cuenta Balance", "Cuenta EERR", "Descripción", "Monto", "Numero Cuotas", "Primera Cuota", "Comisiones", "Facturación", "Facturación Mes Destino"]]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Movimientos");
    XLSX.writeFile(wb, "template_movimientos.xlsx");
  }

  function exportarPeriodo() {
    const rango = movimientos
      .filter((m) => m.mesImpacto >= exportDesde && m.mesImpacto <= exportHasta)
      .sort((a, b) => (a.fecha < b.fecha ? -1 : 1))
      .map((m) => ({
        Fecha: m.fecha,
        "Mes Impacto": m.mesImpacto,
        Medio: m.medio || "",
        "Cuenta Balance": m.cuentaBalance || "",
        "Cuenta EERR": m.cuentaEERR || "",
        Descripción: m.descripcion || "",
        Monto: m.monto,
      }));
    if (!rango.length) { setErrorExportar("No hay movimientos en ese rango de Mes Impacto."); return; }
    setErrorExportar("");
    const ws = XLSX.utils.json_to_sheet(rango);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Movimientos");
    XLSX.writeFile(wb, `movimientos_${exportDesde}_a_${exportHasta}.xlsx`);
  }

  // Importar movimientos desde Excel/CSV (antes era la hoja "Importe",
  // ahora vive aquí mismo en Detalle). Misma lógica: detecta columnas por
  // nombre exacto o palabra clave, exige que de Medio/Cuenta Balance/Cuenta
  // EERR quede exactamente una vacía por fila, y crea automáticamente las
  // cuentas nuevas que aparezcan (clasificadas por Claude).
  const [mostrarImportar, setMostrarImportar] = useState(false);
  const [importRawJson, setImportRawJson] = useState(null);
  const [importHeaders, setImportHeaders] = useState([]);
  const [importRows, setImportRows] = useState(null);
  const [errorImportar, setErrorImportar] = useState("");

  function handleFileImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    setErrorImportar("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (!json.length) { setErrorImportar("El archivo no tiene filas."); return; }
        const hdrs = Object.keys(json[0]);
        setImportRawJson(json);
        setImportHeaders(hdrs);
        setImportRows(buildPreview(json, detectarMapa(hdrs)));
      } catch (err) {
        setErrorImportar("No se pudo leer el archivo. Verifica que sea un Excel o CSV válido.");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function updateRowImport(id, field, value) {
    setImportRows(importRows.map((r) => (r._id === id ? { ...r, [field]: value } : r)));
  }
  function removeRowImport(id) {
    setImportRows(importRows.filter((r) => r._id !== id));
  }

  // Regla: de Medio / Cuenta Balance / Cuenta EERR, exactamente UNA puede
  // venir vacía por fila — eso determina el tipo de movimiento de esa fila.
  // Regla: de Medio / Cuenta Balance / Cuenta EERR, exactamente UNA puede
  // venir vacía por fila (2 llenas) — eso determina el tipo de movimiento de
  // esa fila. Si viene llena SOLO UNA de las 3 (las otras 2 vacías), se toma
  // como "Saldo inicial de cuenta" — pero solo tiene sentido si esa única
  // columna es Medio o Cuenta Balance (cuentas de activo/pasivo); una Cuenta
  // EERR sola no puede representar un saldo inicial, así que esa fila queda
  // inválida.
  function tipoMovDeFilaImport(r) {
    const llenos = [r.medio, r.cuentaBalance, r.cuentaEERR].filter(Boolean).length;
    if (llenos === 2) {
      if (!r.cuentaBalance) return "gasto_ingreso";
      if (!r.cuentaEERR) return "transferencia";
      return "ajuste";
    }
    if (llenos === 1 && (r.medio || r.cuentaBalance)) return "saldo_inicial";
    return null;
  }

  // Valida las 3 columnas nuevas de una fila (solo aplican si el Medio es
  // una tarjeta): si vienen vacías, el movimiento no es en cuotas / no es
  // comisión. Si viene alguno de Numero Cuotas o Primera Cuota, deben venir
  // ambos y ser válidos; Comisiones, si viene, debe ser "SI".
  function cuotasInfoDeFilaImport(r) {
    const esTarjeta = CUENTAS_CUOTAS.includes(r.medio);
    const numRaw = (r.numeroCuotas || "").trim();
    const primeraRaw = (r.primeraCuota || "").trim();
    const comisionRaw = (r.comisiones || "").trim().toUpperCase();
    if (!esTarjeta) return { ok: true, tieneCuotas: false, esComision: false };

    const esComision = comisionRaw === "SI" || comisionRaw === "SÍ";
    if (comisionRaw && !esComision) return { ok: false };

    const tieneAlguno = numRaw !== "" || primeraRaw !== "";
    if (!tieneAlguno) return { ok: true, tieneCuotas: false, esComision };

    const n = Number(numRaw);
    if (!numRaw || !primeraRaw || !Number.isInteger(n) || n < 2) return { ok: false };
    return { ok: true, tieneCuotas: true, numeroCuotas: n, primeraCuota: primeraRaw, esComision };
  }

  // Valida la columna "Facturación" (solo aplica a filas de Saldo inicial
  // cuya cuenta es una tarjeta): vacía = no corresponde; "Mes Pasado"; o
  // "Cuotas Futuras" (requiere Facturación Mes Destino, un mes igual o
  // posterior al Mes Impacto de la fila).
  function facturacionInfoDeFilaImport(r, tipoMov) {
    if (tipoMov !== "saldo_inicial") return { ok: true, tipo: null, mesDestino: null };
    const primaria = primariaDe({ tipoMov, medio: r.medio, cuentaBalance: r.cuentaBalance, cuentaEERR: r.cuentaEERR });
    if (!CUENTAS_CUOTAS.includes(primaria)) return { ok: true, tipo: null, mesDestino: null };
    const raw = (r.facturacion || "").trim().toLowerCase();
    if (!raw) return { ok: true, tipo: null, mesDestino: null };
    if (raw === "mes pasado" || raw === "mespasado") return { ok: true, tipo: "mesPasado", mesDestino: null };
    if (raw === "cuotas futuras" || raw === "cuotasfuturas") {
      const destino = (r.facturacionMesDestino || "").trim();
      if (!destino || destino < r.mesImpacto) return { ok: false };
      return { ok: true, tipo: "cuotasFuturas", mesDestino: destino };
    }
    return { ok: false };
  }

  // Avisos calculados EN VIVO sobre la vista previa (antes de confirmar):
  // simula qué movimientos resultarían de cada fila válida, sin guardarlos
  // todavía, y revisa (1) cuentas de saldo inicial con movimientos previos y
  // (2) coincidencias exactas con un movimiento ya existente.
  const avisoPreview = useMemo(() => {
    if (!importRows) return { mensaje: "", filasDuplicadas: new Set() };
    const cuentasSaldoInicial = new Set();
    const filasDuplicadas = new Set();
    let duplicados = 0;
    importRows.forEach((r) => {
      const tipoMov = tipoMovDeFilaImport(r);
      const info = cuotasInfoDeFilaImport(r);
      const infoFact = facturacionInfoDeFilaImport(r, tipoMov);
      if (!tipoMov || !info.ok || !infoFact.ok || !r.fecha || !r.monto) return;
      const cuentaPrimariaNombre = primariaDe({ tipoMov, medio: r.medio, cuentaBalance: r.cuentaBalance, cuentaEERR: r.cuentaEERR });
      if (tipoMov === "saldo_inicial" && cuentaPrimariaNombre) {
        const tieneAnteriores = movimientos.some((m) => m.fecha <= r.fecha && [m.medio, m.cuentaBalance, m.cuentaEERR].includes(cuentaPrimariaNombre));
        if (tieneAnteriores) cuentasSaldoInicial.add(cuentaPrimariaNombre);
      }
      const montoCLP = Number(r.monto);
      const candidatas = info.tieneCuotas
        ? splitEnCuotas(montoCLP, info.numeroCuotas).map((monto, i) => ({
            fecha: r.fecha,
            mesImpacto: sumarMeses(info.primeraCuota, i),
            medio: r.medio,
            cuentaBalance: r.cuentaBalance,
            cuentaEERR: r.cuentaEERR,
            monto,
            descripcion: r.descripcion ? `${r.descripcion} (cuota ${i + 1}/${info.numeroCuotas})` : `Cuota ${i + 1}/${info.numeroCuotas}`,
          }))
        : [{ fecha: r.fecha, mesImpacto: r.mesImpacto || r.fecha.slice(0, 7), medio: r.medio, cuentaBalance: r.cuentaBalance, cuentaEERR: r.cuentaEERR, monto: montoCLP, descripcion: r.descripcion }];
      candidatas.forEach((nm) => {
        const esDuplicado = movimientos.some(
          (m) =>
            m.fecha === nm.fecha &&
            m.mesImpacto === nm.mesImpacto &&
            (m.medio || "") === (nm.medio || "") &&
            (m.cuentaBalance || "") === (nm.cuentaBalance || "") &&
            (m.cuentaEERR || "") === (nm.cuentaEERR || "") &&
            (m.descripcion || "") === (nm.descripcion || "") &&
            Math.round(m.monto) === Math.round(nm.monto)
        );
        if (esDuplicado) {
          duplicados++;
          filasDuplicadas.add(r._id);
        }
      });
    });
    const partes = [];
    if (cuentasSaldoInicial.size) partes.push(`Ojo: ${[...cuentasSaldoInicial].join(", ")} ya tenía(n) movimientos anteriores a la fecha del saldo inicial que estás por importar — revisa si corresponde.`);
    if (duplicados) partes.push(`Ojo: ${duplicados} movimiento(s) de este archivo son exactamente iguales a uno ya existente (misma fecha, mes impacto, cuentas, descripción y monto) — marcados con ⚠️ abajo, revisa antes de confirmar si no quedará duplicado.`);
    return { mensaje: partes.join(" "), filasDuplicadas };
  }, [importRows, movimientos]);

  // Cancela la importación en curso: limpia el archivo cargado y cierra el
  // panel sin guardar nada.
  function cancelarImport() {
    setImportRawJson(null);
    setImportHeaders([]);
    setImportRows(null);
    setErrorImportar("");
    setMostrarImportar(false);
  }

  async function confirmarImport() {
    setErrorImportar("");
    const candidatas = importRows.filter((r) => r.fecha && r.monto);
    const invalidas = candidatas.filter((r) => !tipoMovDeFilaImport(r));
    const conTipoMov = candidatas.filter((r) => tipoMovDeFilaImport(r));
    const validas = [];
    const invalidasCuotas = [];
    const invalidasFacturacion = [];
    conTipoMov.forEach((r) => {
      const tipoMov = tipoMovDeFilaImport(r);
      const info = cuotasInfoDeFilaImport(r);
      const infoFact = facturacionInfoDeFilaImport(r, tipoMov);
      if (!info.ok) { invalidasCuotas.push(r); return; }
      if (!infoFact.ok) { invalidasFacturacion.push(r); return; }
      validas.push({ ...r, _cuotasInfo: info, _facturacionInfo: infoFact });
    });

    const nombresExistentes = cuentas.map((c) => c.nombre);
    const nuevosBalance = [...new Set(validas.flatMap((r) => [r.medio, r.cuentaBalance]).filter((n) => n && !nombresExistentes.includes(n)))];
    const nuevosEERR = [...new Set(validas.map((r) => r.cuentaEERR).filter((n) => n && !nombresExistentes.includes(n)))];
    let cuentasNuevas = [];

    if (nuevosBalance.length || nuevosEERR.length) {
      // Sin IA: las cuentas nuevas detectadas se crean con un tipo por
      // defecto razonable (activo / gasto) — se pueden corregir después en
      // Datos si corresponde a pasivo o ingreso.
      cuentasNuevas = [
        ...nuevosBalance.map((n) => ({
          id: Date.now().toString() + Math.random().toString(36).slice(2, 6),
          nombre: n,
          tipo: "activo",
          estado: "activa",
          esMedio: true,
        })),
        ...nuevosEERR.map((n) => ({
          id: Date.now().toString() + Math.random().toString(36).slice(2, 6),
          nombre: n,
          tipo: "gasto",
          estado: "activa",
          esMedio: false,
        })),
      ];
      setCuentas((prev) => [...prev, ...cuentasNuevas]);
    }

    const cuentasCompletas = [...cuentas, ...cuentasNuevas];
    const bloqueadas = [];
    const nuevosMov = [];
    validas.forEach((r) => {
      const tipoMov = tipoMovDeFilaImport(r);
      const cuentaPrimariaNombre = primariaDe({ tipoMov, medio: r.medio, cuentaBalance: r.cuentaBalance, cuentaEERR: r.cuentaEERR });
      const cuentaObj = cuentasCompletas.find((c) => c.nombre === cuentaPrimariaNombre);
      if (cuentaObj && cuentaObj.estado === "inactiva" && cuentaObj.fechaInactivacion && r.fecha >= cuentaObj.fechaInactivacion) {
        bloqueadas.push(r);
        return;
      }
      const montoCLP = Number(r.monto);
      const info = r._cuotasInfo;

      if (info.tieneCuotas) {
        // Pago en cuotas detectado en el archivo: se registra UN solo
        // movimiento con el MONTO TOTAL en el Mes Impacto de la fila (para
        // que el Balance reconozca la deuda completa de inmediato), con un
        // cronograma de facturación (mismo reparto que al agregarlo a mano,
        // desde Primera Cuota) para que Facturación muestre cada cuota.
        const n = info.numeroCuotas;
        const planId = Date.now().toString() + "-" + r._id;
        const partesCLP = splitEnCuotas(montoCLP, n);
        const cronograma = partesCLP.map((parte, i) => ({ mes: sumarMeses(info.primeraCuota, i), monto: parte }));
        const { monto: montoCuota, montoSecundario: montoSecundarioCuota } = calcularMontos(tipoMov, r.medio, r.cuentaBalance, r.cuentaEERR, montoCLP, cuentasCompletas);
        nuevosMov.push({
          id: planId,
          tipoMov,
          fecha: r.fecha,
          mesImpacto: r.mesImpacto || r.fecha.slice(0, 7),
          divisa: "CLP",
          medio: r.medio,
          cuentaBalance: r.cuentaBalance,
          cuentaEERR: r.cuentaEERR,
          monto: montoCuota,
          montoSecundario: montoSecundarioCuota,
          montoUSD: null,
          descripcion: r.descripcion,
          esCuota: true,
          cuotaPlanId: planId,
          cuotaTotal: n,
          cronograma,
          esComision: info.esComision,
          facturacionTipo: r._facturacionInfo.tipo,
          facturacionMesDestino: r._facturacionInfo.mesDestino,
        });
      } else {
        const { monto, montoSecundario } = calcularMontos(tipoMov, r.medio, r.cuentaBalance, r.cuentaEERR, montoCLP, cuentasCompletas);
        nuevosMov.push({
          id: Date.now().toString() + r._id,
          tipoMov,
          fecha: r.fecha,
          mesImpacto: r.mesImpacto || r.fecha.slice(0, 7),
          divisa: "CLP",
          medio: r.medio,
          cuentaBalance: r.cuentaBalance,
          cuentaEERR: r.cuentaEERR,
          monto,
          montoSecundario,
          montoUSD: null,
          descripcion: r.descripcion,
          esCuota: false,
          cuotaPlanId: null,
          cuotaTotal: null,
          cronograma: null,
          esComision: info.esComision,
          facturacionTipo: r._facturacionInfo.tipo,
          facturacionMesDestino: r._facturacionInfo.mesDestino,
        });
      }
    });

    if (invalidas.length || invalidasCuotas.length || invalidasFacturacion.length || bloqueadas.length) {
      const partes = [];
      if (invalidas.length) partes.push(`${invalidas.length} fila(s) sin importar: de Medio/Cuenta Balance/Cuenta EERR debe venir llena al menos Medio o Cuenta Balance (sola, como saldo inicial) o exactamente 2 de las 3 (Cuenta EERR sola no es válida).`);
      if (invalidasCuotas.length) partes.push(`${invalidasCuotas.length} fila(s) sin importar: Numero Cuotas/Primera Cuota deben venir juntos y válidos (Numero Cuotas natural mayor a 1), y Comisiones debe venir vacío o con SI.`);
      if (invalidasFacturacion.length) partes.push(`${invalidasFacturacion.length} fila(s) sin importar: Facturación debe venir vacío, "Mes Pasado" o "Cuotas Futuras" (con Facturación Mes Destino, un mes igual o posterior al Mes Impacto).`);
      if (bloqueadas.length) partes.push(`${bloqueadas.length} fila(s) sin importar: la cuenta ya estaba inactiva en esa fecha.`);
      setErrorImportar(partes.join(" "));
    }

    setMovimientos((prev) => [...prev, ...nuevosMov]);
    setImportRawJson(null);
    setImportRows(null);
    // Si hay algo que avisar, se deja el panel abierto para que el mensaje
    // sea visible (si se cierra de inmediato, nunca se alcanza a ver).
    if (!(invalidas.length || invalidasCuotas.length || invalidasFacturacion.length || bloqueadas.length)) {
      setMostrarImportar(false);
    }
  }

  // Edición por fila en la tabla: se activa con el botón de lápiz, muestra
  // los campos editables solo mientras esa fila está en modo edición; el
  // resto del tiempo se ve como texto/etiquetas de solo lectura.
  const [editId, setEditId] = useState(null);
  const [editTemp, setEditTemp] = useState(null);

  function startEdit(m) {
    setEditId(m.id);
    setEditTemp({ ...m });
  }
  function cancelEdit() {
    setEditId(null);
    setEditTemp(null);
  }
  function saveEdit() {
    let montoUSD = editTemp.montoUSD;
    let montoCLP;
    if (editTemp.divisa === "USD") {
      montoUSD = Number(editTemp.montoUSD) || 0;
      montoCLP = montoUSD * (indicadores?.usd || 0);
    } else {
      montoCLP = Number(editTemp.monto) || 0;
    }
    const { monto, montoSecundario } = calcularMontos(editTemp.tipoMov, editTemp.medio, editTemp.cuentaBalance, editTemp.cuentaEERR, montoCLP, cuentas);
    setMovimientos(movimientos.map((m) => (m.id === editId ? { ...editTemp, monto, montoSecundario, montoUSD } : m)));
    setEditId(null);
    setEditTemp(null);
  }

  function badgeCuenta(nombre) {
    const tipo = (cuentas.find((c) => c.nombre === nombre) || {}).tipo;
    return (
      <span key={nombre} style={{ background: TIPO_STYLES[tipo]?.bg || "#EDEAE0", color: TIPO_STYLES[tipo]?.text || COLORS.sub, borderRadius: "999px", padding: "0.1rem 0.55rem", fontSize: "0.72rem", display: "inline-block" }}>
        {nombre}
      </span>
    );
  }

  return (
    <div>
      <div className="flex gap-2 mb-3 flex-wrap">
        <select style={{ ...inputStyle, flex: "1 1 220px" }} value={tipoMov} onChange={(e) => setTipoMov(e.target.value)}>
          {TIPOS_MOVIMIENTO.map((t) => (<option key={t.id} value={t.id}>{t.label}</option>))}
        </select>
        <select style={{ ...inputStyle, width: "6rem" }} value={divisa} onChange={(e) => setDivisa(e.target.value)}>
          <option value="CLP">CLP</option>
          <option value="USD">USD</option>
        </select>
      </div>

      <div className="flex gap-2 mb-2 flex-wrap">
        <input type="date" style={{ ...inputStyle, flex: 1 }} value={fecha} onChange={(e) => setFecha(e.target.value)} />
        <input type="month" style={{ ...inputStyle, flex: 1 }} value={mesImpacto} onChange={(e) => { setMesImpacto(e.target.value); setMesImpactoManual(true); }} title="Mes impacto" />
        {campos.includes("medio") && (
          <select style={{ ...inputStyle, flex: 1 }} value={medio} onChange={(e) => setMedio(e.target.value)}>
            <option value="">Medio...</option>
            {medios.map((c) => (<option key={c.id} value={c.nombre}>{c.nombre}</option>))}
          </select>
        )}
        {campos.includes("cuentaBalance") && (
          <select style={{ ...inputStyle, flex: 1 }} value={cuentaBalance} onChange={(e) => setCuentaBalance(e.target.value)}>
            <option value="">Cuenta Balance...</option>
            {balances.map((c) => (<option key={c.id} value={c.nombre}>{c.nombre}</option>))}
          </select>
        )}
        {tipoMov === "transferencia" && cuentaBalance === "Tarjeta USD" && (
          <input
            type="number"
            style={{ ...inputStyle, flex: 1 }}
            placeholder="USD que se están pagando"
            value={usdPagados}
            onChange={(e) => setUsdPagados(e.target.value)}
            title={`Sin pagar en Tarjeta USD: US$${usdSinPagarTarjetaUSD.toLocaleString("es-CL")}`}
          />
        )}
        {campos.includes("cuentaEERR") && (
          <select style={{ ...inputStyle, flex: 1 }} value={cuentaEERR} onChange={(e) => setCuentaEERR(e.target.value)}>
            <option value="">Cuenta EERR...</option>
            {eerr.map((c) => (<option key={c.id} value={c.nombre}>{c.nombre}</option>))}
          </select>
        )}
        {campos.includes("cuentaUnica") && (
          <select style={{ ...inputStyle, flex: 1 }} value={cuentaUnica} onChange={(e) => setCuentaUnica(e.target.value)}>
            <option value="">Cuenta Balance o Medio...</option>
            {poolUnica.map((c) => (<option key={c.id} value={c.nombre}>{c.nombre}</option>))}
          </select>
        )}
      </div>
      {tipoMov === "transferencia" && cuentaBalance === "Tarjeta USD" && (
        <p style={{ fontSize: "0.75rem", color: COLORS.sub, marginTop: "-0.5rem", marginBottom: "0.5rem" }}>
          Sin pagar en Tarjeta USD: US${usdSinPagarTarjetaUSD.toLocaleString("es-CL")}
        </p>
      )}
      {campos.includes("medio") && CUENTAS_CUOTAS.includes(medio) && (
        <div className="flex gap-2 mb-2 flex-wrap items-center">
          <button style={enCuotas ? btnStyle(COLORS.finanzas) : ghostBtnStyle(COLORS.finanzas)} onClick={() => setEnCuotas((v) => !v)}>
            Pago en cuotas
          </button>
          {enCuotas && (
            <>
              <input type="number" min="2" style={{ ...inputStyle, width: "9rem" }} placeholder="Cantidad de cuotas" value={numCuotas} onChange={(e) => setNumCuotas(e.target.value)} />
              <input type="month" style={{ ...inputStyle, width: "10rem" }} value={mesPrimeraCuota} onChange={(e) => { setMesPrimeraCuota(e.target.value); setMesPrimeraCuotaManual(true); }} title="Mes primera cuota" />
            </>
          )}
          <button style={esComision ? btnStyle(COLORS.finanzas) : ghostBtnStyle(COLORS.finanzas)} onClick={() => setEsComision((v) => !v)}>
            Comisiones
          </button>
          <button style={esRefinanciamiento ? btnStyle(COLORS.finanzas) : ghostBtnStyle(COLORS.finanzas)} onClick={() => setEsRefinanciamiento((v) => !v)}>
            Refinanciamiento
          </button>
        </div>
      )}
      {campos.includes("cuentaUnica") && CUENTAS_CUOTAS.includes(cuentaUnica) && (
        <div className="flex gap-2 mb-2 flex-wrap items-center">
          <button
            style={facturacionTipo ? btnStyle(COLORS.finanzas) : ghostBtnStyle(COLORS.finanzas)}
            onClick={() => setFacturacionTipo((v) => (v ? "" : "mesPasado"))}
          >
            Corresponde facturación
          </button>
          {facturacionTipo && (
            <>
              <select style={{ ...inputStyle, width: "13rem" }} value={facturacionTipo} onChange={(e) => setFacturacionTipo(e.target.value)}>
                <option value="mesPasado">Facturación mes pasado</option>
                <option value="cuotasFuturas">Saldo de cuotas próximos meses</option>
              </select>
              {facturacionTipo === "cuotasFuturas" && (
                <input type="month" style={{ ...inputStyle, width: "10rem" }} value={facturacionMesDestino} onChange={(e) => setFacturacionMesDestino(e.target.value)} title="Mes impacto al que corresponde este saldo" />
              )}
            </>
          )}
        </div>
      )}
      <div className="flex gap-2 mb-2 flex-wrap">
        <input type="number" style={{ ...inputStyle, flex: 1 }} placeholder={`Monto (${divisa})`} value={monto} onChange={(e) => setMonto(e.target.value)} />
        <input style={{ ...inputStyle, flex: 2 }} placeholder="Descripción" value={descripcion} onChange={(e) => setDescripcion(e.target.value)} />
        <button style={btnStyle(COLORS.finanzas)} onClick={addMov}>
          <Plus size={16} /> Agregar
        </button>
      </div>
      {error && <p style={{ color: "#9C4A2E", fontSize: "0.8rem", marginBottom: "0.8rem" }}>{error}</p>}

      <div className="flex gap-2 mb-1 flex-wrap">
        <button style={ghostBtnStyle(COLORS.finanzas)} onClick={descargarTemplate}>
          <Download size={15} /> Descargar template
        </button>
        <button style={ghostBtnStyle(COLORS.finanzas)} onClick={() => setMostrarExportar((v) => !v)}>
          <Download size={15} /> Exportar periodo
        </button>
        <button style={ghostBtnStyle(COLORS.finanzas)} onClick={() => setMostrarImportar((v) => !v)}>
          <Upload size={15} /> Importar
        </button>
      </div>
      <div style={{ marginBottom: "1rem" }}>
        {mostrarExportar && (
          <div className="flex items-center gap-2 flex-wrap" style={{ marginTop: "0.6rem", background: COLORS.card, borderRadius: "0.8rem", padding: "0.7rem 0.9rem" }}>
            <span style={{ fontSize: "0.8rem", color: COLORS.sub }}>Mes Impacto desde</span>
            <input type="month" style={{ ...inputStyle, width: "9rem" }} value={exportDesde} onChange={(e) => setExportDesde(e.target.value)} />
            <span style={{ fontSize: "0.8rem", color: COLORS.sub }}>hasta</span>
            <input type="month" style={{ ...inputStyle, width: "9rem" }} value={exportHasta} onChange={(e) => setExportHasta(e.target.value)} />
            <button style={btnStyle(COLORS.finanzas)} onClick={exportarPeriodo}>
              <Download size={15} /> Descargar Excel
            </button>
          </div>
        )}
        {errorExportar && <p style={{ color: "#9C4A2E", fontSize: "0.8rem", marginTop: "0.4rem" }}>{errorExportar}</p>}

        {mostrarImportar && (
          <div style={{ marginTop: "0.6rem", background: COLORS.card, borderRadius: "0.8rem", padding: "0.9rem" }}>
            <p style={{ color: COLORS.sub, fontSize: "0.82rem", marginBottom: "0.8rem" }}>
              Sube un Excel o CSV con Fecha, Mes Impacto, Medio, Cuenta Balance, Cuenta EERR, Descripción y Monto — si vienen con esos nombres exactos, se reconocen todos automáticamente. Las columnas que falten se completan aquí manualmente. Por fila, solo uno de Medio / Cuenta Balance / Cuenta EERR puede quedar vacío.
            </p>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileImport} style={{ marginBottom: "0.9rem" }} />
            {errorImportar && <p style={{ color: "#9C4A2E", fontSize: "0.82rem", marginBottom: "0.8rem" }}>{errorImportar}</p>}
            {avisoPreview.mensaje && <p style={{ color: "#8C6E2F", fontSize: "0.82rem", marginBottom: "0.8rem" }}>{avisoPreview.mensaje}</p>}

            {importRows && (
              <>


                <div className="flex flex-col gap-1 mb-3" style={{ maxHeight: "24rem", overflowY: "auto" }}>
                  {importRows.map((r) => {
                    const tipoMov = tipoMovDeFilaImport(r);
                    const infoCuotas = cuotasInfoDeFilaImport(r);
                    const infoFact = facturacionInfoDeFilaImport(r, tipoMov);
                    const esDuplicada = avisoPreview.filasDuplicadas.has(r._id);
                    const primariaFila = tipoMov ? primariaDe({ tipoMov, medio: r.medio, cuentaBalance: r.cuentaBalance, cuentaEERR: r.cuentaEERR }) : null;
                    return (
                      <div key={r._id} className="rounded-xl" style={{ background: "#fff", padding: "0.6rem 0.7rem", border: tipoMov && infoCuotas.ok && infoFact.ok ? (esDuplicada ? "1px solid #C9A24A" : "none") : "1px solid #C97B4A" }}>
                        <div className="flex items-center gap-2 flex-wrap mb-1">
                          {esDuplicada && <AlertTriangle size={16} color="#8C6E2F" title="Coincide con un movimiento ya existente" />}
                          <input type="date" style={{ ...inputStyle, width: "8.5rem" }} value={r.fecha} onChange={(e) => updateRowImport(r._id, "fecha", e.target.value)} title="Fecha" />
                          <input type="month" style={{ ...inputStyle, width: "8rem" }} value={r.mesImpacto} onChange={(e) => updateRowImport(r._id, "mesImpacto", e.target.value)} title="Mes impacto" />
                          <input type="number" style={{ ...inputStyle, width: "7.5rem" }} value={r.monto} onChange={(e) => updateRowImport(r._id, "monto", e.target.value)} placeholder="Monto" />
                          <input style={{ ...inputStyle, flex: 1 }} value={r.descripcion} onChange={(e) => updateRowImport(r._id, "descripcion", e.target.value)} placeholder="Descripción" />
                          <button onClick={() => removeRowImport(r._id)} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}>
                            <Trash2 size={16} />
                          </button>
                        </div>
                        <div className="flex items-center gap-2 flex-wrap">
                          <input list="import-medio-list" style={{ ...inputStyle, flex: 1 }} value={r.medio} onChange={(e) => updateRowImport(r._id, "medio", e.target.value)} placeholder="Medio" />
                          <input list="import-balance-list" style={{ ...inputStyle, flex: 1 }} value={r.cuentaBalance} onChange={(e) => updateRowImport(r._id, "cuentaBalance", e.target.value)} placeholder="Cuenta Balance" />
                          <input list="import-eerr-list" style={{ ...inputStyle, flex: 1 }} value={r.cuentaEERR} onChange={(e) => updateRowImport(r._id, "cuentaEERR", e.target.value)} placeholder="Cuenta EERR" />
                        </div>
                        {CUENTAS_CUOTAS.includes(r.medio) && (
                          <div className="flex items-center gap-2 flex-wrap" style={{ marginTop: "0.4rem" }}>
                            <input type="number" min="2" style={{ ...inputStyle, width: "9rem" }} value={r.numeroCuotas} onChange={(e) => updateRowImport(r._id, "numeroCuotas", e.target.value)} placeholder="Numero Cuotas" />
                            <input type="month" style={{ ...inputStyle, width: "9rem" }} value={r.primeraCuota} onChange={(e) => updateRowImport(r._id, "primeraCuota", e.target.value)} title="Primera Cuota" />
                            <input style={{ ...inputStyle, width: "7rem" }} value={r.comisiones} onChange={(e) => updateRowImport(r._id, "comisiones", e.target.value)} placeholder="Comisiones" />
                          </div>
                        )}
                        {tipoMov === "saldo_inicial" && CUENTAS_CUOTAS.includes(primariaFila) && (
                          <div className="flex items-center gap-2 flex-wrap" style={{ marginTop: "0.4rem" }}>
                            <select style={{ ...inputStyle, width: "11rem" }} value={r.facturacion} onChange={(e) => updateRowImport(r._id, "facturacion", e.target.value)}>
                              <option value="">Facturación: sin marcar</option>
                              <option value="Mes Pasado">Facturación mes pasado</option>
                              <option value="Cuotas Futuras">Saldo de cuotas próximas</option>
                            </select>
                            {r.facturacion.trim().toLowerCase() === "cuotas futuras" && (
                              <input type="month" style={{ ...inputStyle, width: "9rem" }} value={r.facturacionMesDestino} onChange={(e) => updateRowImport(r._id, "facturacionMesDestino", e.target.value)} title="Facturación Mes Destino" />
                            )}
                          </div>
                        )}
                        {!tipoMov && <p style={{ color: "#9C4A2E", fontSize: "0.72rem", marginTop: "0.3rem" }}>Debe quedar exactamente 1 de estos 3 campos vacío.</p>}
                        {tipoMov && !infoCuotas.ok && <p style={{ color: "#9C4A2E", fontSize: "0.72rem", marginTop: "0.3rem" }}>Numero Cuotas y Primera Cuota deben venir juntos (Numero Cuotas natural mayor a 1); Comisiones debe venir vacío o "SI".</p>}
                        {tipoMov && infoCuotas.ok && !infoFact.ok && <p style={{ color: "#9C4A2E", fontSize: "0.72rem", marginTop: "0.3rem" }}>Facturación debe venir vacío, "Mes Pasado" o "Cuotas Futuras" (con Facturación Mes Destino igual o posterior al Mes Impacto).</p>}
                      </div>
                    );
                  })}
                </div>
                <datalist id="import-medio-list">{mediosBase.filter((c) => c.estado !== "inactiva").map((c) => (<option key={c.id} value={c.nombre} />))}</datalist>
                <datalist id="import-balance-list">{balancesBase.filter((c) => c.estado !== "inactiva").map((c) => (<option key={c.id} value={c.nombre} />))}</datalist>
                <datalist id="import-eerr-list">{eerrBase.filter((c) => c.estado !== "inactiva").map((c) => (<option key={c.id} value={c.nombre} />))}</datalist>
                <div className="flex gap-2">
                  <button style={btnStyle(COLORS.finanzas)} onClick={confirmarImport}>
                    <Upload size={16} /> {`Confirmar importación (${importRows.length})`}
                  </button>
                  <button style={ghostBtnStyle(COLORS.sub)} onClick={cancelarImport}>
                    Cancelar
                  </button>
                </div>
              </>
            )}
          </div>
        )}
      </div>

      {ordenados.length === 0 ? (
        <p style={{ color: COLORS.sub, fontSize: "0.85rem" }}>Sin movimientos registrados todavía.</p>
      ) : (
        <div style={{ overflowX: "auto" }}>
          {seleccionados.size > 0 && (
            <div className="flex items-center gap-3 flex-wrap" style={{ marginBottom: "0.6rem", background: "#F3E3DE", borderRadius: "0.6rem", padding: "0.5rem 0.8rem" }}>
              <span style={{ fontSize: "0.82rem", color: "#9C4A2E" }}>{seleccionados.size} seleccionado(s)</span>
              {confirmandoBorrado ? (
                <>
                  <span style={{ fontSize: "0.8rem", color: "#9C4A2E" }}>¿Borrar {seleccionados.size} movimiento(s)? No se puede deshacer.</span>
                  <button style={{ ...ghostBtnStyle("#9C4A2E") }} onClick={borrarSeleccionados}>
                    <Trash2 size={15} /> Sí, borrar
                  </button>
                  <button style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub, fontSize: "0.8rem" }} onClick={() => setConfirmandoBorrado(false)}>
                    Cancelar
                  </button>
                </>
              ) : (
                <>
                  <button style={{ ...ghostBtnStyle("#9C4A2E") }} onClick={() => setConfirmandoBorrado(true)}>
                    <Trash2 size={15} /> Borrar seleccionados
                  </button>
                  <button style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub, fontSize: "0.8rem" }} onClick={() => setSeleccionados(new Set())}>
                    Quitar selección
                  </button>
                </>
              )}
            </div>
          )}
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.85rem" }}>
            <thead>
              <tr style={{ textAlign: "left" }}>
                <th style={thStyle}>
                  <input type="checkbox" checked={todosFiltradosSeleccionados} onChange={toggleSeleccionarTodos} title="Seleccionar todos los filtrados" />
                </th>
                <th style={thStyle}>Fecha</th>
                <th style={thStyle}>Mes Impacto</th>
                <th style={thStyle}>Cuentas</th>
                <th style={thStyle}>Descripción</th>
                <th style={{ ...thStyle, textAlign: "right" }}>Monto</th>
                <th style={thStyle}></th>
              </tr>
              <tr>
                <th style={thFiltroStyle}></th>
                <th style={thFiltroStyle}>
                  <select style={filtroInputStyle} value={filtros.fecha} onChange={(e) => setFiltros({ ...filtros, fecha: e.target.value })}>
                    <option value="">Todas</option>
                    {fechasDisponibles.map((f) => (<option key={f} value={f}>{f}</option>))}
                  </select>
                </th>
                <th style={thFiltroStyle}>
                  <select style={filtroInputStyle} value={filtros.mesImpacto} onChange={(e) => setFiltros({ ...filtros, mesImpacto: e.target.value })}>
                    <option value="">Todos</option>
                    {mesesDisponibles.map((mi) => (<option key={mi} value={mi}>{fmtMesImpacto(mi)}</option>))}
                  </select>
                </th>
                <th style={thFiltroStyle}>
                  <select style={filtroInputStyle} value={filtros.cuentas} onChange={(e) => setFiltros({ ...filtros, cuentas: e.target.value })}>
                    <option value="">Todas</option>
                    {cuentasDisponibles.map((c) => (<option key={c} value={c}>{c}</option>))}
                  </select>
                </th>
                <th style={thFiltroStyle}><input style={filtroInputStyle} placeholder="Filtrar..." value={filtros.descripcion} onChange={(e) => setFiltros({ ...filtros, descripcion: e.target.value })} /></th>
                <th style={thFiltroStyle}></th>
                <th style={thFiltroStyle}></th>
              </tr>
            </thead>
            <tbody>
              {filtrados.length === 0 ? (
                <tr>
                  <td colSpan={7} style={{ padding: "0.9rem 0.5rem", color: COLORS.sub, fontSize: "0.85rem" }}>
                    Ningún movimiento coincide con los filtros (tienes {ordenados.length} en total). Cambia o limpia los filtros de arriba para ver otros movimientos.
                  </td>
                </tr>
              ) : (
                visibles.map((m) => {
                const editando = editId === m.id;
                const medios = disponiblesPorFecha(mediosBase, editando ? editTemp.fecha : m.fecha);
                const balancesRow = disponiblesPorFecha(balancesBase, editando ? editTemp.fecha : m.fecha);
                const eerrRow = disponiblesPorFecha(eerrBase, editando ? editTemp.fecha : m.fecha);

                if (editando) {
                  return (
                    <tr key={m.id} style={{ borderTop: `1px solid ${COLORS.line}` }}>
                      <td style={tdStyle}></td>
                      <td style={tdStyle}>
                        <input type="date" style={{ ...inputStyle, ...cellInputStyle }} value={editTemp.fecha} onChange={(e) => setEditTemp({ ...editTemp, fecha: e.target.value })} />
                      </td>
                      <td style={tdStyle}>
                        <input type="month" style={{ ...inputStyle, ...cellInputStyle }} value={editTemp.mesImpacto} onChange={(e) => setEditTemp({ ...editTemp, mesImpacto: e.target.value })} />
                      </td>
                      <td style={tdStyle}>
                        <div className="flex flex-col gap-1">
                          {editTemp.medio && (
                            <select style={{ ...inputStyle, ...cellInputStyle }} value={editTemp.medio} onChange={(e) => setEditTemp({ ...editTemp, medio: e.target.value })}>
                              {medios.map((c) => (<option key={c.id} value={c.nombre}>{c.nombre}</option>))}
                            </select>
                          )}
                          {editTemp.cuentaBalance && (
                            <select style={{ ...inputStyle, ...cellInputStyle }} value={editTemp.cuentaBalance} onChange={(e) => setEditTemp({ ...editTemp, cuentaBalance: e.target.value })}>
                              {balancesRow.map((c) => (<option key={c.id} value={c.nombre}>{c.nombre}</option>))}
                            </select>
                          )}
                          {editTemp.cuentaEERR && (
                            <select style={{ ...inputStyle, ...cellInputStyle }} value={editTemp.cuentaEERR} onChange={(e) => setEditTemp({ ...editTemp, cuentaEERR: e.target.value })}>
                              {eerrRow.map((c) => (<option key={c.id} value={c.nombre}>{c.nombre}</option>))}
                            </select>
                          )}
                        </div>
                      </td>
                      <td style={tdStyle}>
                        <input style={{ ...inputStyle, ...cellInputStyle }} value={editTemp.descripcion} onChange={(e) => setEditTemp({ ...editTemp, descripcion: e.target.value })} />
                      </td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>
                        {editTemp.divisa === "USD" ? (
                          <input
                            type="number"
                            style={{ ...inputStyle, ...cellInputStyle, textAlign: "right" }}
                            value={editTemp.montoUSD}
                            onChange={(e) => setEditTemp({ ...editTemp, montoUSD: e.target.value })}
                            placeholder="Monto USD"
                          />
                        ) : (
                          <input type="number" style={{ ...inputStyle, ...cellInputStyle, textAlign: "right" }} value={editTemp.monto} onChange={(e) => setEditTemp({ ...editTemp, monto: e.target.value })} />
                        )}
                      </td>
                      <td style={tdStyle}>
                        <div className="flex items-center gap-2">
                          <button onClick={saveEdit} style={{ background: "none", border: "none", cursor: "pointer", color: "#3B6E4F" }}><Check size={16} /></button>
                          <button onClick={cancelEdit} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}><X size={16} /></button>
                        </div>
                      </td>
                    </tr>
                  );
                }

                return (
                  <tr key={m.id} style={{ borderTop: `1px solid ${COLORS.line}` }}>
                    <td style={tdStyle}>
                      <input type="checkbox" checked={seleccionados.has(m.id)} onChange={() => toggleSeleccion(m.id)} />
                    </td>
                    <td style={tdStyle}>
                      <span style={{ color: COLORS.sub }}>{m.fecha}</span>
                    </td>
                    <td style={tdStyle}>
                      <span style={{ color: COLORS.sub }}>{fmtMesImpacto(m.mesImpacto)}</span>
                    </td>
                    <td style={tdStyle}>
                      <div className="flex flex-col gap-1" style={{ alignItems: "flex-start" }}>
                        {m.medio && badgeCuenta(m.medio)}
                        {m.cuentaBalance && badgeCuenta(m.cuentaBalance)}
                        {m.cuentaEERR && badgeCuenta(m.cuentaEERR)}
                      </div>
                    </td>
                    <td style={tdStyle}>
                      {m.descripcion}
                      {m.esCuota && <span style={{ marginLeft: "0.4rem", fontSize: "0.68rem", color: COLORS.sub }}>({m.cuotaTotal} cuotas)</span>}
                      {m.esComision && <span style={{ marginLeft: "0.4rem", fontSize: "0.68rem", color: "#9C4A2E" }}>Comisión</span>}
                      {m.esRefinanciamiento && <span style={{ marginLeft: "0.4rem", fontSize: "0.68rem", color: "#534AB7" }}>Refinanciamiento</span>}
                    </td>
                    <td style={{ ...tdStyle, textAlign: "right" }}>
                      <span style={{ fontWeight: 500 }}>{fmtMoney(m.monto)}</span>
                      {m.divisa === "USD" && m.montoUSD ? (
                        <div style={{ fontSize: "0.68rem", color: COLORS.sub }}>
                          US${m.montoUSD.toLocaleString("es-CL")}{" "}
                          <span style={{ color: m.pagado ? "#3B6E4F" : "#9C4A2E" }}>({m.pagado ? "Pagado" : "No pagado"})</span>
                        </div>
                      ) : null}
                    </td>
                    <td style={tdStyle}>
                      <div className="flex items-center gap-2">
                        <button onClick={() => startEdit(m)} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}><Pencil size={15} /></button>
                        <button onClick={() => removeMov(m.id)} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub }}><Trash2 size={16} /></button>
                      </div>
                    </td>
                  </tr>
                );
              })
              )}
            </tbody>
          </table>
          {filtrados.length > 0 && (
          <div className="flex items-center justify-between flex-wrap gap-2" style={{ marginTop: "0.7rem" }}>
            <span style={{ fontSize: "0.78rem", color: COLORS.sub }}>
              Mostrando {filtrados.length === 0 ? 0 : pagina * 50 + 1}-{pagina * 50 + visibles.length} de {filtrados.length}{filtrados.length !== ordenados.length ? ` (de ${ordenados.length} en total)` : ""}
            </span>
            {totalPaginas > 1 && (
              <div className="flex items-center gap-2">
                <button style={{ ...ghostBtnStyle(COLORS.finanzas), opacity: pagina === 0 ? 0.4 : 1 }} onClick={() => setPagina((p) => Math.max(0, p - 1))} disabled={pagina === 0}>
                  Anterior
                </button>
                <span style={{ fontSize: "0.78rem", color: COLORS.sub }}>Página {pagina + 1} de {totalPaginas}</span>
                <button style={{ ...ghostBtnStyle(COLORS.finanzas), opacity: pagina >= totalPaginas - 1 ? 0.4 : 1 }} onClick={() => setPagina((p) => Math.min(totalPaginas - 1, p + 1))} disabled={pagina >= totalPaginas - 1}>
                  Siguiente
                </button>
              </div>
            )}
          </div>
          )}
        </div>
      )}
    </div>
  );
}


/* ---- 3. Resultados ---- */
function ResultadosTab({ cuentas, movimientos, fechasFacturacion }) {
  const currentYear = new Date().getFullYear();
  const [tipoPeriodo, setTipoPeriodo] = useState("año");
  const [anio, setAnio] = useState(currentYear);
  const [sub, setSub] = useState(new Date().getMonth() + 1);

  const subOptions = useMemo(() => {
    if (tipoPeriodo === "mes") return Array.from({ length: 12 }, (_, i) => i + 1);
    if (tipoPeriodo === "trimestre") return [1, 2, 3, 4];
    if (tipoPeriodo === "semestre") return [1, 2];
    return [1];
  }, [tipoPeriodo]);

  const cuentaTipo = useMemo(() => {
    const map = {};
    cuentas.forEach((c) => (map[c.nombre] = c.tipo));
    return map;
  }, [cuentas]);

  const { start, end } = getPeriodRange(tipoPeriodo, anio, sub);
  // Para Balance y Estado de Resultados, lo que define a qué periodo
  // pertenece un movimiento es su "Mes Impacto" (no la fecha real del
  // movimiento, que es solo referencial). Se comparan como texto
  // "YYYY-MM", que ordena igual que una fecha.
  const startMes = `${start.getFullYear()}-${String(start.getMonth() + 1).padStart(2, "0")}`;
  const endMes = `${end.getFullYear()}-${String(end.getMonth() + 1).padStart(2, "0")}`;

  const [showActivos, setShowActivos] = useState(false);
  const [showPasivos, setShowPasivos] = useState(false);
  const [showIngresos, setShowIngresos] = useState(false);
  const [showGastos, setShowGastos] = useState(false);

  // Desglose por cuenta de Ingresos y Gastos del periodo seleccionado
  // (mismo criterio de mes impacto que "resultado": entre startMes y endMes).
  const resultadoDetalle = useMemo(() => {
    const porCuenta = {};
    movimientos.forEach((m) => {
      const mesEfectivo = mesImpactoEfectivo(m, fechasFacturacion);
      if (mesEfectivo < startMes || mesEfectivo > endMes) return;
      const s = secundariaDe(m);
      const ts = s ? cuentaTipo[s] : null;
      if (ts === "ingreso" || ts === "gasto") porCuenta[s] = (porCuenta[s] || 0) + m.montoSecundario;
    });
    const ingresos = cuentas.filter((c) => c.tipo === "ingreso").map((c) => ({ nombre: c.nombre, monto: porCuenta[c.nombre] || 0 })).filter((r) => r.monto !== 0);
    const gastos = cuentas.filter((c) => c.tipo === "gasto").map((c) => ({ nombre: c.nombre, monto: porCuenta[c.nombre] || 0 })).filter((r) => r.monto !== 0);
    return { ingresos, gastos };
  }, [movimientos, cuentas, cuentaTipo, startMes, endMes, fechasFacturacion]);

  // Desglose de Activos y Pasivos por cuenta, y Resultados Acumulados
  // (Ingresos - Gastos): en los 3 casos se suma TODO movimiento desde el
  // inicio de los datos (no hay límite inferior) hasta el fin del periodo
  // seleccionado, según su Mes Impacto — así el balance de junio incluye
  // lo acumulado desde el primer mes con datos hasta junio.
  // Primer Mes Impacto presente en todo el detalle. En ese mes en particular,
  // los movimientos suelen ser solo "Saldo inicial de cuenta" (sin Ingresos/
  // Gastos asociados), así que Ingresos-Gastos acumulado daría $0 aunque sí
  // exista un patrimonio inicial real. Por eso, para ese primer mes, el
  // Resultado Acumulado se calcula como Activos - Pasivos de ese mismo mes.
  const primerMes = useMemo(() => {
    const meses = movimientos.map((m) => m.mesImpacto).filter(Boolean).sort();
    return meses.length ? meses[0] : null;
  }, [movimientos]);

  const balanceDetalle = useMemo(() => {
    const porCuenta = {};
    movimientos.forEach((m) => {
      if (m.mesImpacto > endMes) return;
      const p = primariaDe(m);
      const tp = p ? cuentaTipo[p] : null;
      if (tp === "activo" || tp === "pasivo") porCuenta[p] = (porCuenta[p] || 0) + m.monto;
      const s = secundariaDe(m);
      const ts = s ? cuentaTipo[s] : null;
      if (ts === "activo" || ts === "pasivo") porCuenta[s] = (porCuenta[s] || 0) + m.montoSecundario;
    });
    const activos = cuentas.filter((c) => c.tipo === "activo").map((c) => ({ nombre: c.nombre, monto: porCuenta[c.nombre] || 0 })).filter((r) => r.monto !== 0);
    const pasivos = cuentas.filter((c) => c.tipo === "pasivo").map((c) => ({ nombre: c.nombre, monto: porCuenta[c.nombre] || 0 })).filter((r) => r.monto !== 0);
    const totalActivos = activos.reduce((a, r) => a + r.monto, 0);
    const totalPasivos = pasivos.reduce((a, r) => a + r.monto, 0);

    // Patrimonio inicial = Activos - Pasivos al cierre del primer mes con
    // datos (normalmente movimientos de "Saldo inicial de cuenta", que no
    // pasan por Ingresos/Gastos). A partir de ahí, el acumulado suma el
    // flujo real de Ingresos - Gastos mes a mes hasta el periodo analizado
    // — así el resultado acumulado sí refleja todo el historial, no solo
    // el mes que se está mirando.
    const porCuentaPrimerMes = {};
    movimientos.forEach((m) => {
      if (m.mesImpacto !== primerMes) return;
      const p = primariaDe(m);
      const tp = p ? cuentaTipo[p] : null;
      if (tp === "activo" || tp === "pasivo") porCuentaPrimerMes[p] = (porCuentaPrimerMes[p] || 0) + m.monto;
      const s = secundariaDe(m);
      const ts = s ? cuentaTipo[s] : null;
      if (ts === "activo" || ts === "pasivo") porCuentaPrimerMes[s] = (porCuentaPrimerMes[s] || 0) + m.montoSecundario;
    });
    const patrimonioInicial = cuentas.reduce((acc, c) => {
      const v = porCuentaPrimerMes[c.nombre] || 0;
      if (c.tipo === "activo") return acc + v;
      if (c.tipo === "pasivo") return acc - v;
      return acc;
    }, 0);

    let flujoPosterior = 0;
    movimientos.forEach((m) => {
      if (m.mesImpacto <= primerMes || m.mesImpacto > endMes) return;
      const s = secundariaDe(m);
      const ts = s ? cuentaTipo[s] : null;
      if (ts === "ingreso") flujoPosterior += m.montoSecundario;
      if (ts === "gasto") flujoPosterior -= m.montoSecundario;
    });

    return {
      activos,
      pasivos,
      totalActivos,
      totalPasivos,
      resultadosAcumulados: patrimonioInicial + flujoPosterior,
    };
  }, [movimientos, cuentas, cuentaTipo, endMes, primerMes]);

  const resultado = useMemo(() => {
    // Estado de resultados: solo flujo del periodo seleccionado (ingresos/gastos).
    // Balance: acumulado de activos/pasivos hasta el fin del periodo, no solo
    // lo que entra ese mes — así "Sueldo" y "MUST" reflejan su saldo real
    // a esa fecha, aunque su movimiento haya ocurrido antes.
    // El periodo se determina por Mes Impacto, no por la fecha real del
    // movimiento (esa es solo referencial) — salvo que haya una Fecha de
    // Facturación definida y la fecha real sea posterior a ella, en cuyo
    // caso Ingresos/Gastos (no Balance) usan el mes impacto siguiente.
    let ingresos = 0, gastos = 0, activos = 0, pasivos = 0, sinClasificar = 0;
    movimientos.forEach((m) => {
      const mesEfectivo = mesImpactoEfectivo(m, fechasFacturacion);
      const p = primariaDe(m);
      const tp = p ? cuentaTipo[p] : null;
      if (p && !tp) sinClasificar++;
      if (tp === "activo" && m.mesImpacto <= endMes) activos += m.monto;
      if (tp === "pasivo" && m.mesImpacto <= endMes) pasivos += m.monto;

      const s = secundariaDe(m);
      const ts = s ? cuentaTipo[s] : null;
      if (s && !ts) sinClasificar++;
      if (ts === "ingreso" && mesEfectivo >= startMes && mesEfectivo <= endMes) ingresos += m.montoSecundario;
      if (ts === "gasto" && mesEfectivo >= startMes && mesEfectivo <= endMes) gastos += m.montoSecundario;
      if (ts === "activo" && m.mesImpacto <= endMes) activos += m.montoSecundario;
      if (ts === "pasivo" && m.mesImpacto <= endMes) pasivos += m.montoSecundario;
    });
    return { ingresos, gastos, activos, pasivos, sinClasificar, neto: ingresos - gastos, patrimonio: activos - pasivos };
  }, [movimientos, cuentaTipo, startMes, endMes, fechasFacturacion]);

  return (
    <div>
      <div className="flex gap-2 mb-6 flex-wrap items-center">
        <select style={inputStyle} value={tipoPeriodo} onChange={(e) => { setTipoPeriodo(e.target.value); setSub(1); }}>
          <option value="mes">Mes</option>
          <option value="trimestre">Trimestre (Q)</option>
          <option value="semestre">Semestre</option>
          <option value="año">Año</option>
        </select>
        {tipoPeriodo !== "año" && (
          <select style={inputStyle} value={sub} onChange={(e) => setSub(Number(e.target.value))}>
            {subOptions.map((o) => (
              <option key={o} value={o}>
                {tipoPeriodo === "mes"
                  ? new Date(2000, o - 1, 1).toLocaleDateString("es-CL", { month: "long" })
                  : `${tipoPeriodo === "trimestre" ? "Q" : "S"}${o}`}
              </option>
            ))}
          </select>
        )}
        <input type="number" style={{ ...inputStyle, width: "6rem" }} value={anio} onChange={(e) => setAnio(Number(e.target.value))} />
      </div>

      <div className="rounded-2xl mb-6" style={{ background: COLORS.card, padding: "1.2rem" }}>
        <p style={{ fontFamily: "'Fraunces', serif", fontSize: "1.05rem", marginBottom: "0.8rem" }}>Estado de resultados</p>
        <div className="flex items-center justify-between" style={{ padding: "0.3rem 0" }}>
          <span style={{ fontFamily: "'Fraunces', serif", fontSize: "1.05rem" }}>Ingresos</span>
          <div className="flex items-center gap-3">
            <span style={{ fontWeight: 600 }}>{fmtMoney(resultado.ingresos)}</span>
            <button onClick={() => setShowIngresos((v) => !v)} style={{ background: "none", border: `1px solid ${COLORS.line}`, borderRadius: "999px", width: "1.4rem", height: "1.4rem", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", color: COLORS.ink }}>
              {showIngresos ? <Minus size={12} /> : <Plus size={12} />}
            </button>
          </div>
        </div>
        {showIngresos && resultadoDetalle.ingresos.map((r) => <Row key={r.nombre} label={r.nombre} value={fmtMoney(r.monto)} />)}

        <div className="flex items-center justify-between" style={{ padding: "0.3rem 0" }}>
          <span style={{ fontFamily: "'Fraunces', serif", fontSize: "1.05rem" }}>Gastos</span>
          <div className="flex items-center gap-3">
            <span style={{ fontWeight: 600 }}>{fmtMoney(resultado.gastos)}</span>
            <button onClick={() => setShowGastos((v) => !v)} style={{ background: "none", border: `1px solid ${COLORS.line}`, borderRadius: "999px", width: "1.4rem", height: "1.4rem", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", color: COLORS.ink }}>
              {showGastos ? <Minus size={12} /> : <Plus size={12} />}
            </button>
          </div>
        </div>
        {showGastos && resultadoDetalle.gastos.map((r) => <Row key={r.nombre} label={r.nombre} value={fmtMoney(r.monto)} />)}

        <Row label="Resultado neto" value={fmtMoney(resultado.neto)} strong color={resultado.neto >= 0 ? "#3B6E4F" : "#9C4A2E"} />
      </div>

      <div className="rounded-2xl mb-6" style={{ background: COLORS.card, display: "grid", gridTemplateColumns: "1fr 1fr" }}>
        <div style={{ padding: "1.2rem", borderRight: `1px solid ${COLORS.line}` }}>
          <div className="flex items-center justify-between" style={{ marginBottom: showActivos ? "0.7rem" : 0 }}>
            <p style={{ fontFamily: "'Fraunces', serif", fontSize: "1.05rem" }}>Activos</p>
            <div className="flex items-center gap-3">
              <span style={{ fontWeight: 600 }}>{fmtMoney(balanceDetalle.totalActivos)}</span>
              <button onClick={() => setShowActivos((v) => !v)} style={{ background: "none", border: `1px solid ${COLORS.line}`, borderRadius: "999px", width: "1.6rem", height: "1.6rem", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", color: COLORS.ink }}>
                {showActivos ? <Minus size={14} /> : <Plus size={14} />}
              </button>
            </div>
          </div>
          {showActivos && balanceDetalle.activos.map((r) => <Row key={r.nombre} label={r.nombre} value={fmtMoney(r.monto)} />)}
        </div>

        <div>
          <div style={{ padding: "1.2rem" }}>
            <div className="flex items-center justify-between" style={{ marginBottom: showPasivos ? "0.7rem" : 0 }}>
              <p style={{ fontFamily: "'Fraunces', serif", fontSize: "1.05rem" }}>Pasivos</p>
              <div className="flex items-center gap-3">
                <span style={{ fontWeight: 600 }}>{fmtMoney(balanceDetalle.totalPasivos)}</span>
                <button onClick={() => setShowPasivos((v) => !v)} style={{ background: "none", border: `1px solid ${COLORS.line}`, borderRadius: "999px", width: "1.6rem", height: "1.6rem", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", color: COLORS.ink }}>
                  {showPasivos ? <Minus size={14} /> : <Plus size={14} />}
                </button>
              </div>
            </div>
            {showPasivos && balanceDetalle.pasivos.map((r) => <Row key={r.nombre} label={r.nombre} value={fmtMoney(r.monto)} />)}
          </div>

          <div style={{ padding: "1.2rem", borderTop: `1px solid ${COLORS.line}` }}>
            <p style={{ fontFamily: "'Fraunces', serif", fontSize: "1.05rem", marginBottom: "0.5rem" }}>Resultados acumulados</p>
            <p style={{ fontSize: "1.3rem", fontWeight: 600, color: balanceDetalle.resultadosAcumulados >= 0 ? "#3B6E4F" : "#9C4A2E" }}>
              {fmtMoney(balanceDetalle.resultadosAcumulados)}
            </p>
          </div>
        </div>
      </div>
      {Math.round(balanceDetalle.totalActivos) !== Math.round(balanceDetalle.totalPasivos + balanceDetalle.resultadosAcumulados) && (
        <p style={{ color: "#9C4A2E", fontSize: "0.85rem", marginBottom: "1rem" }}>
          La regla del balance no se está cumpliendo, revisar el detalle.
        </p>
      )}
      {resultado.sinClasificar > 0 && (
        <p style={{ color: COLORS.sub, fontSize: "0.8rem", marginBottom: "1rem" }}>
          {resultado.sinClasificar} movimiento(s) con cuenta no clasificada — revisa el módulo Datos.
        </p>
      )}
    </div>
  );
}
function Row({ label, value, strong, color }) {
  return (
    <div className="flex justify-between" style={{ padding: "0.3rem 0", borderTop: strong ? "1px solid #DAD6C9" : "none", marginTop: strong ? "0.4rem" : 0 }}>
      <span style={{ fontSize: "0.9rem", color: strong ? COLORS.ink : COLORS.sub, fontWeight: strong ? 600 : 400 }}>{label}</span>
      <span style={{ fontSize: "0.9rem", fontWeight: strong ? 600 : 500, color: color || COLORS.ink }}>{value}</span>
    </div>
  );
}

/* ---- Análisis de Gastos ---- */
function AnalisisTab({ cuentas, movimientos }) {
  const currentYear = new Date().getFullYear();
  // Por defecto se abre por Año = año actual. Si se cambia a Mes/Trimestre/
  // Semestre, se ajusta automáticamente al último periodo con movimientos
  // reales (no necesariamente el actual, que puede no tener datos aún).
  const maxFechaInicial = movimientos.length ? movimientos.reduce((max, m) => (m.fecha > max ? m.fecha : max), movimientos[0].fecha) : null;
  const [tipoPeriodo, setTipoPeriodo] = useState("año");
  const [anio, setAnio] = useState(currentYear);
  const [sub, setSub] = useState(new Date().getMonth() + 1);
  const [seleccionada, setSeleccionada] = useState(null);

  function subDeFecha(tipo, fechaStr) {
    const d = new Date(fechaStr);
    const mes = d.getMonth() + 1;
    if (tipo === "trimestre") return Math.ceil(mes / 3);
    if (tipo === "semestre") return Math.ceil(mes / 6);
    return mes;
  }

  // Al cambiar el tipo de apertura (Mes/Trimestre/Semestre), se salta al
  // último periodo de ESE tipo que tiene movimientos reales.
  useEffect(() => {
    if (tipoPeriodo === "año" || !maxFechaInicial) return;
    setAnio(Number(maxFechaInicial.slice(0, 4)));
    setSub(subDeFecha(tipoPeriodo, maxFechaInicial));
  }, [tipoPeriodo]);

  const subOptions = useMemo(() => {
    if (tipoPeriodo === "mes") return Array.from({ length: 12 }, (_, i) => i + 1);
    if (tipoPeriodo === "trimestre") return [1, 2, 3, 4];
    if (tipoPeriodo === "semestre") return [1, 2];
    return [1];
  }, [tipoPeriodo]);

  const { start, end } = getPeriodRange(tipoPeriodo, anio, sub);
  const gastoCuentas = useMemo(() => cuentas.filter((c) => c.tipo === "gasto"), [cuentas]);

  // Composición: total gastado por cada cuenta de tipo "gasto" en el periodo.
  const composicion = useMemo(() => {
    const totales = {};
    movimientos.forEach((m) => {
      const f = new Date(m.fecha);
      if (f < start || f > end) return;
      if (!m.cuentaEERR || !gastoCuentas.find((c) => c.nombre === m.cuentaEERR)) return;
      totales[m.cuentaEERR] = (totales[m.cuentaEERR] || 0) + (m.montoSecundario ?? m.monto);
    });
    return Object.entries(totales)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
  }, [movimientos, gastoCuentas, start, end]);

  const totalGasto = composicion.reduce((acc, d) => acc + d.value, 0);

  // Apertura: al seleccionar una cuenta de gasto, se agrupan sus movimientos
  // del periodo por descripción (ej. peajes, bencina, seguro dentro de "Auto")
  // para ver qué compone ese gasto.
  const desglose = useMemo(() => {
    if (!seleccionada) return [];
    const totales = {};
    movimientos.forEach((m) => {
      const f = new Date(m.fecha);
      if (f < start || f > end) return;
      if (m.cuentaEERR !== seleccionada) return;
      // Se agrupa solo por lo que viene antes de ":" o "." (lo que aparezca
      // primero), para que variantes de un mismo concepto (ej.
      // "Refinanciamiento: pago en cuotas", "Refinanciamiento. ajuste Saldo
      // Cuotas") queden en una sola fila.
      const desc = m.descripcion?.trim() || "Sin descripción";
      const corte = [desc.indexOf(":"), desc.indexOf(".")].filter((i) => i !== -1).sort((a, b) => a - b)[0];
      const key = corte !== undefined ? desc.slice(0, corte).trim() : desc;
      totales[key] = (totales[key] || 0) + (m.montoSecundario ?? m.monto);
    });
    return Object.entries(totales)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
  }, [movimientos, seleccionada, start, end]);

  // Evolución: suma mensual en los últimos 12 meses (terminando en el mes de
  // fin del periodo seleccionado). Si hay una cuenta seleccionada, es la
  // evolución de esa cuenta; si no, es el total de todos los gastos del mes.
  const evolucion = useMemo(() => {
    const meses = [];
    for (let i = 11; i >= 0; i--) {
      const d = new Date(end.getFullYear(), end.getMonth() - i, 1);
      meses.push({ anio: d.getFullYear(), mes: d.getMonth(), label: d.toLocaleDateString("es-CL", { month: "short", year: "2-digit" }) });
    }
    return meses.map(({ anio: a, mes, label }) => {
      const total = movimientos
        .filter((m) => {
          if (seleccionada) {
            if (m.cuentaEERR !== seleccionada) return false;
          } else if (!m.cuentaEERR || !gastoCuentas.find((c) => c.nombre === m.cuentaEERR)) {
            return false;
          }
          const f = new Date(m.fecha);
          return f.getFullYear() === a && f.getMonth() === mes;
        })
        .reduce((acc, m) => acc + (m.montoSecundario ?? m.monto), 0);
      return { mes: label, total };
    });
  }, [movimientos, seleccionada, end, gastoCuentas]);

  return (
    <div>
      <div className="flex gap-2 mb-6 flex-wrap items-center">
        <select style={inputStyle} value={tipoPeriodo} onChange={(e) => { setTipoPeriodo(e.target.value); setSub(1); }}>
          <option value="mes">Mes</option>
          <option value="trimestre">Trimestre (Q)</option>
          <option value="semestre">Semestre</option>
          <option value="año">Año</option>
        </select>
        {tipoPeriodo !== "año" && (
          <select style={inputStyle} value={sub} onChange={(e) => setSub(Number(e.target.value))}>
            {subOptions.map((o) => (
              <option key={o} value={o}>
                {tipoPeriodo === "mes"
                  ? new Date(2000, o - 1, 1).toLocaleDateString("es-CL", { month: "long" })
                  : `${tipoPeriodo === "trimestre" ? "Q" : "S"}${o}`}
              </option>
            ))}
          </select>
        )}
        <input type="number" style={{ ...inputStyle, width: "6rem" }} value={anio} onChange={(e) => setAnio(Number(e.target.value))} />
      </div>

      {composicion.length === 0 ? (
        <p style={{ color: COLORS.sub, fontSize: "0.85rem" }}>No hay gastos registrados en este periodo.</p>
      ) : (
        <div className="rounded-2xl mb-5" style={{ background: COLORS.card, padding: "1.2rem" }}>
          <div className="flex items-center justify-between" style={{ marginBottom: "0.6rem" }}>
            <p style={{ fontFamily: "'Fraunces', serif", fontSize: "1.05rem" }}>Composición de gastos</p>
            <span style={{ fontSize: "0.9rem", fontWeight: 600, color: COLORS.finanzas }}>{fmtMoney(totalGasto)}</span>
          </div>
          <div className="flex flex-col md:flex-row gap-4 items-center">
            <div style={{ width: "100%", maxWidth: "260px", height: "220px" }}>
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={composicion} dataKey="value" nameKey="name" innerRadius={45} outerRadius={90} onClick={(d) => setSeleccionada(d.name)}>
                    {composicion.map((d, i) => (
                      <Cell key={d.name} fill={PIE_COLORS[i % PIE_COLORS.length]} cursor="pointer" />
                    ))}
                  </Pie>
                  <Tooltip formatter={(v) => fmtMoney(v)} />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div className="flex flex-col gap-1" style={{ flex: 1, width: "100%" }}>
              {composicion.map((d, i) => (
                <button
                  key={d.name}
                  onClick={() => setSeleccionada(d.name)}
                  className="flex items-center justify-between rounded-lg"
                  style={{
                    padding: "0.4rem 0.6rem",
                    background: seleccionada === d.name ? "#F2ECDD" : "transparent",
                    border: "none",
                    cursor: "pointer",
                    textAlign: "left",
                  }}
                >
                  <span className="flex items-center gap-2" style={{ fontSize: "0.85rem" }}>
                    <span style={{ width: "10px", height: "10px", borderRadius: "999px", background: PIE_COLORS[i % PIE_COLORS.length], display: "inline-block" }} />
                    {d.name}
                  </span>
                  <span style={{ fontSize: "0.85rem", color: COLORS.sub }}>
                    {fmtMoney(d.value)} · {totalGasto ? Math.round((d.value / totalGasto) * 100) : 0}%
                  </span>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {seleccionada && (
        <>
          <div className="flex items-center justify-between mb-2">
            <p style={{ fontFamily: "'Fraunces', serif", fontSize: "1.1rem", color: COLORS.finanzas }}>{seleccionada}</p>
            <button onClick={() => setSeleccionada(null)} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.sub, fontSize: "0.8rem", display: "flex", alignItems: "center", gap: "0.2rem" }}>
              <X size={14} /> cerrar
            </button>
          </div>

          <div className="rounded-2xl mb-5" style={{ background: COLORS.card, padding: "1.2rem" }}>
            <div className="flex items-center justify-between" style={{ marginBottom: "0.8rem" }}>
              <p style={{ fontSize: "0.9rem", color: COLORS.sub }}>Movimientos más representativos en el periodo</p>
              <span style={{ fontSize: "0.9rem", fontWeight: 600, color: COLORS.finanzas }}>{fmtMoney(desglose.reduce((a, d) => a + d.value, 0))}</span>
            </div>
            {desglose.length === 0 ? (
              <p style={{ color: COLORS.sub, fontSize: "0.85rem" }}>Sin movimientos en este periodo.</p>
            ) : (
              <div className="flex flex-col md:flex-row gap-4 items-center">
                <div style={{ width: "100%", maxWidth: "220px", height: "190px" }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={desglose} dataKey="value" nameKey="name" outerRadius={80}>
                        {desglose.map((d, i) => (
                          <Cell key={d.name} fill={PIE_COLORS[i % PIE_COLORS.length]} />
                        ))}
                      </Pie>
                      <Tooltip formatter={(v) => fmtMoney(v)} />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
                <div className="flex flex-col gap-1" style={{ flex: 1, width: "100%" }}>
                  {desglose.map((d, i) => (
                    <div key={d.name} className="flex items-center justify-between" style={{ padding: "0.3rem 0.2rem" }}>
                      <span className="flex items-center gap-2" style={{ fontSize: "0.85rem" }}>
                        <span style={{ width: "10px", height: "10px", borderRadius: "999px", background: PIE_COLORS[i % PIE_COLORS.length], display: "inline-block" }} />
                        {d.name}
                      </span>
                      <span style={{ fontSize: "0.85rem", color: COLORS.sub }}>{fmtMoney(d.value)}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </>
      )}

      <div className="rounded-2xl" style={{ background: COLORS.card, padding: "1.2rem" }}>
        <p style={{ fontSize: "0.9rem", color: COLORS.sub, marginBottom: "0.8rem" }}>
          {seleccionada ? `Evolución de ${seleccionada} — últimos 12 meses` : "Evolución total de gastos — últimos 12 meses"}
        </p>
        <div style={{ width: "100%", height: "220px" }}>
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={evolucion}>
              <CartesianGrid stroke="#E3E0D5" vertical={false} />
              <XAxis dataKey="mes" tick={{ fontSize: 11, fill: COLORS.sub }} />
              <YAxis tick={{ fontSize: 11, fill: COLORS.sub }} tickFormatter={(v) => (v >= 1000 ? `${Math.round(v / 1000)}k` : v)} />
              <Tooltip formatter={(v) => fmtMoney(v)} />
              <Line type="monotone" dataKey="total" stroke={COLORS.finanzas} strokeWidth={2} dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* ---- 4. Importe ---- */
const FIELD_KEYWORDS = {
  fecha: ["fecha", "date"],
  mesImpacto: ["mes impacto", "mesimpacto"],
  medio: ["medio"],
  cuentaBalance: ["cuenta balance", "balance"],
  cuentaEERR: ["cuenta eerr", "eerr", "categoria", "categoría"],
  monto: ["monto", "valor", "amount", "importe"],
  descripcion: ["descripcion", "descripción", "detalle", "glosa", "concepto"],
  numeroCuotas: ["numero cuotas", "número cuotas", "n cuotas", "cantidad cuotas"],
  primeraCuota: ["primera cuota"],
  comisiones: ["comisiones", "comision", "comisión"],
  facturacion: ["facturación", "facturacion"],
  facturacionMesDestino: ["facturación mes destino", "facturacion mes destino", "mes destino"],
};

function findExact(headers, nombre) {
  return headers.find((h) => h.trim().toLowerCase() === nombre.toLowerCase());
}
function detectColumn(headers, field) {
  const kws = FIELD_KEYWORDS[field];
  return headers.find((h) => kws.some((k) => h.toLowerCase().includes(k)));
}
// Si el archivo trae exactamente estas columnas (como en el formato que ya
// usas), se reconocen todas directo. Si no, se cae a detección por palabra clave.
function detectarMapa(headers) {
  return {
    fecha: findExact(headers, "fecha") || detectColumn(headers, "fecha"),
    mesImpacto: findExact(headers, "mes impacto") || detectColumn(headers, "mesImpacto"),
    medio: findExact(headers, "medio") || detectColumn(headers, "medio"),
    cuentaBalance: findExact(headers, "cuenta balance") || detectColumn(headers, "cuentaBalance"),
    cuentaEERR: findExact(headers, "cuenta eerr") || detectColumn(headers, "cuentaEERR"),
    descripcion: findExact(headers, "descripción") || findExact(headers, "descripcion") || detectColumn(headers, "descripcion"),
    monto: findExact(headers, "monto") || detectColumn(headers, "monto"),
    numeroCuotas: findExact(headers, "numero cuotas") || findExact(headers, "número cuotas") || detectColumn(headers, "numeroCuotas"),
    primeraCuota: findExact(headers, "primera cuota") || detectColumn(headers, "primeraCuota"),
    comisiones: findExact(headers, "comisiones") || detectColumn(headers, "comisiones"),
    facturacion: findExact(headers, "facturación") || findExact(headers, "facturacion") || detectColumn(headers, "facturacion"),
    facturacionMesDestino: findExact(headers, "facturación mes destino") || findExact(headers, "facturacion mes destino") || detectColumn(headers, "facturacionMesDestino"),
  };
}

// Excel puede entregar la fecha como objeto Date (con cellDates:true), como
// número de serie (si la celda no viene con formato de fecha), o como texto.
// Esta función normaliza cualquiera de los tres casos a "YYYY-MM-DD".
function excelFechaAISO(v) {
  if (!v && v !== 0) return "";
  if (v instanceof Date) {
    const y = v.getUTCFullYear();
    const m = String(v.getUTCMonth() + 1).padStart(2, "0");
    const d = String(v.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof v === "number") {
    const parsed = XLSX.SSF.parse_date_code(v);
    if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }

  // Texto: reconoce varios formatos comunes, sin importar cómo vengan desde
  // el Excel (algunas celdas quedan como texto en vez de fecha real).
  const str = String(v).trim();

  // YYYY-MM-DD o YYYY/MM/DD (con o sin hora pegada al final)
  let m = str.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;

  // DD/MM/YYYY, DD-MM-YYYY o DD.MM.YYYY (formato habitual en Chile)
  m = str.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})/);
  if (m) {
    let [, a, b, y] = m;
    a = Number(a);
    b = Number(b);
    // Si uno de los dos números no puede ser un mes (>12), ese es el día.
    // Si ambos son válidos como mes, se asume DD/MM (formato chileno).
    let dia = a, mes = b;
    if (a > 12 && b <= 12) { dia = a; mes = b; }
    else if (b > 12 && a <= 12) { dia = b; mes = a; }
    return `${y}-${String(mes).padStart(2, "0")}-${String(dia).padStart(2, "0")}`;
  }

  // DD/MM/YY, DD-MM-YY o DD.MM.YY (año de 2 dígitos, se asume 20YY)
  m = str.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2})$/);
  if (m) {
    let [, a, b, y] = m;
    a = Number(a);
    b = Number(b);
    let dia = a, mes = b;
    if (a > 12 && b <= 12) { dia = a; mes = b; }
    else if (b > 12 && a <= 12) { dia = b; mes = a; }
    return `20${y}-${String(mes).padStart(2, "0")}-${String(dia).padStart(2, "0")}`;
  }

  return str.slice(0, 10);
}

function buildPreview(rawJson, map) {
  return rawJson.map((r, i) => {
    const fecha = map.fecha && r[map.fecha] !== undefined ? excelFechaAISO(r[map.fecha]) : "";
    const mesImpactoRaw = map.mesImpacto && r[map.mesImpacto] !== undefined ? excelFechaAISO(r[map.mesImpacto]).slice(0, 7) : "";
    const primeraCuotaRaw = map.primeraCuota && r[map.primeraCuota] !== undefined && r[map.primeraCuota] !== "" ? excelFechaAISO(r[map.primeraCuota]).slice(0, 7) : "";
    const facturacionMesDestinoRaw = map.facturacionMesDestino && r[map.facturacionMesDestino] !== undefined && r[map.facturacionMesDestino] !== "" ? excelFechaAISO(r[map.facturacionMesDestino]).slice(0, 7) : "";
    return {
      _id: i,
      fecha,
      mesImpacto: mesImpactoRaw || fecha.slice(0, 7),
      medio: map.medio && r[map.medio] !== undefined ? String(r[map.medio]) : "",
      cuentaBalance: map.cuentaBalance && r[map.cuentaBalance] !== undefined ? String(r[map.cuentaBalance]) : "",
      cuentaEERR: map.cuentaEERR && r[map.cuentaEERR] !== undefined ? String(r[map.cuentaEERR]) : "",
      monto: map.monto && r[map.monto] !== undefined ? Number(r[map.monto]) || 0 : 0,
      descripcion: map.descripcion && r[map.descripcion] !== undefined ? String(r[map.descripcion]) : "",
      numeroCuotas: map.numeroCuotas && r[map.numeroCuotas] !== undefined && r[map.numeroCuotas] !== "" ? String(r[map.numeroCuotas]).trim() : "",
      primeraCuota: primeraCuotaRaw,
      comisiones: map.comisiones && r[map.comisiones] !== undefined ? String(r[map.comisiones]).trim() : "",
      facturacion: map.facturacion && r[map.facturacion] !== undefined ? String(r[map.facturacion]).trim() : "",
      facturacionMesDestino: facturacionMesDestinoRaw,
    };
  });
}

/* ---- Sueldo (placeholder) ---- */
function SueldoTab() {
  return <p style={{ color: COLORS.sub, fontSize: "0.9rem" }}>Esta hoja todavía no tiene contenido — la definimos en el próximo paso.</p>;
}

/* ---- Facturación ---- */
function FacturacionTab({ cuentas, movimientos, setMovimientos, fechasFacturacion, setFechasFacturacion }) {
  const [tarjeta, setTarjeta] = useState("Tarjeta");

  // Para Facturación, si hay Fecha de Facturación definida y la fecha real
  // de un movimiento de tarjeta es posterior a ella, se considera que cae
  // en el Mes Impacto siguiente (ver mesImpactoEfectivo) — Balance sigue
  // usando el Mes Impacto real (movimientos, sin ajustar), por eso esta
  // copia se usa solo acá, no para escribir cambios de vuelta.
  const movimientosFact = useMemo(() => conMesImpactoEfectivo(movimientos, fechasFacturacion), [movimientos, fechasFacturacion]);

  // Por defecto se muestra el último Mes Impacto con movimientos en
  // cualquiera de las 3 tarjetas (no necesariamente el mes calendario actual).
  const ultimoMes = useMemo(() => {
    const meses = movimientosFact.filter((m) => CUENTAS_CUOTAS.includes(m.medio)).map((m) => m.mesImpacto).filter(Boolean).sort();
    return meses.length ? meses[meses.length - 1] : new Date().toISOString().slice(0, 7);
  }, [movimientosFact]);
  const [mes, setMes] = useState(ultimoMes);

  // Clasificación de cada movimiento en los 4 tipos de transacción de la
  // tarjeta seleccionada:
  // - Pago: transferencia donde la tarjeta es la Cuenta Balance (se pagó
  //   desde la cuenta corriente hacia la tarjeta).
  // - Cuotas: gasto con la tarjeta como Medio, generado por "Pago en cuotas".
  // - Comisiones: gasto con la tarjeta como Medio, marcado como Comisiones.
  // - Gasto: cualquier otro gasto con la tarjeta como Medio.
  function clasificar(m) {
    return clasificarFacturacion(m, tarjeta);
  }

  // Cada compra en cuotas es UN solo movimiento (con su cronograma), así que
  // ya no hace falta separar "filas" de "cuotas ocultas" — cada fila
  // simplemente aparece en el Mes Impacto con que se ingresó.
  const movsDelMes = useMemo(() => {
    return movimientosFact
      .filter((m) => m.mesImpacto === mes && clasificar(m))
      .map((m) => ({ ...m, _transaccion: clasificar(m) }))
      .sort((a, b) => (a.fecha < b.fecha ? -1 : 1));
  }, [movimientosFact, tarjeta, mes]);

  // Filtros de la tabla de movimientos, iguales en espíritu a los de Detalle:
  // desplegables con los valores realmente disponibles ese mes, más texto
  // libre para Descripción. Se limpian al cambiar de tarjeta o de mes.
  const [filtrosFact, setFiltrosFact] = useState({ fecha: "", transaccion: "", descripcion: "" });
  useEffect(() => { setFiltrosFact({ fecha: "", transaccion: "", descripcion: "" }); }, [tarjeta, mes]);

  const fechasDisponiblesFact = useMemo(() => [...new Set(movsDelMes.map((m) => m.fecha))].sort((a, b) => (a < b ? 1 : -1)), [movsDelMes]);
  const transaccionesDisponibles = useMemo(() => ["Gasto", "Cuotas", "Comisiones", "Pago", "Refinanciamiento"].filter((t) => movsDelMes.some((m) => m._transaccion === t)), [movsDelMes]);

  const movsFiltrados = useMemo(
    () =>
      movsDelMes.filter(
        (m) =>
          (!filtrosFact.fecha || m.fecha === filtrosFact.fecha) &&
          (!filtrosFact.transaccion || m._transaccion === filtrosFact.transaccion) &&
          (!filtrosFact.descripcion || (m.descripcion || "").toLowerCase().includes(filtrosFact.descripcion.toLowerCase()))
      ),
    [movsDelMes, filtrosFact]
  );

  // Facturado Mes Pasado se resuelve automáticamente: si hay un saldo
  // inicial manual marcado como "mes pasado" para este mes, se usa ese valor
  // (para cuando se empieza a trackear la tarjeta a mitad de camino); si no,
  // y el mes anterior ya tiene datos en la app, el Facturado Mes Pasado de
  // este mes es directamente el Total del mes anterior (encadenado).
  const totales = useMemo(() => calcularFacturacionMes(tarjeta, mes, movimientosFact), [tarjeta, mes, movimientosFact]);
  const { facturadoMesPasado, totalGasto, totalComisiones, totalCuotas, totalPago, totalRefinanciamiento, total: totalFacturado } = totales;

  // Proyección (Saldo Cuotas): saldos manuales + cronograma de compras en
  // cuotas de meses anteriores — el refinanciamiento y las cuotas de ESTE
  // mes NO la modifican, ver calcularProyeccionCuotas.
  const proyeccion = useMemo(() => calcularProyeccionCuotas(tarjeta, mes, movimientosFact), [movimientosFact, mes, tarjeta]);

  // Ajuste de cada refinanciamiento (refinanciamientoId -> cronograma):
  // el ajuste es un movimiento aparte (para que afecte el Balance), pero se
  // muestra en la misma fila del Refinanciamiento principal, vinculado por
  // refinanciamientoId.
  const ajustesRefinanciamiento = useMemo(() => {
    const mapa = {};
    movimientosFact.filter((m) => m.refinanciamientoId && m.cronogramaRefinanciamiento).forEach((m) => {
      mapa[m.refinanciamientoId] = m.cronogramaRefinanciamiento;
    });
    return mapa;
  }, [movimientosFact]);

  // Columnas de mes de la tabla: Saldo Cuotas, más cualquier mes al que
  // apunte el cronograma (cuotas o ajuste de refinanciamiento) de una fila
  // de este mes, aunque no tenga saldo en el cuadro de Saldo Cuotas.
  const monthColumns = useMemo(() => {
    const meses = new Set(proyeccion.length ? proyeccion.map((p) => p.mes) : [mes]);
    meses.add(mes);
    movsDelMes.forEach((m) => {
      if (m.esCuota && m.cronograma) m.cronograma.forEach((c) => meses.add(c.mes));
      const cronoRefi = ajustesRefinanciamiento[m.refinanciamientoId];
      if (cronoRefi) cronoRefi.forEach((c) => meses.add(c.mes));
    });
    const maxMes = [...meses].reduce((max, mm) => (mm > max ? mm : max), mes);
    const totalMeses = diffMeses(mes, maxMes) + 1;
    return Array.from({ length: totalMeses }, (_, k) => sumarMeses(mes, k));
  }, [proyeccion, mes, movsDelMes, ajustesRefinanciamiento]);
  const proyeccionPorMes = useMemo(() => Object.fromEntries(proyeccion.map((p) => [p.mes, p.monto])), [proyeccion]);

  // Valor que le corresponde a un movimiento en la columna de un mes dado:
  // para una cuota o un refinanciamiento, lo que indique su propio
  // cronograma; si no, su monto cuando ese mes es su Mes Impacto.
  function valorCelda(m, mesCol) {
    if (m.esCuota && m.cronograma) {
      const entry = m.cronograma.find((c) => c.mes === mesCol);
      return entry ? entry.monto : null;
    }
    const cronoRefi = ajustesRefinanciamiento[m.refinanciamientoId];
    if (cronoRefi) {
      const entry = cronoRefi.find((c) => c.mes === mesCol);
      if (entry) return entry.monto;
    }
    if (mesCol === m.mesImpacto) return m.tipoMov === "transferencia" ? (m.montoSecundario ?? m.monto) : m.monto;
    return null;
  }

  // Cambia la Transacción de un movimiento entre Gasto/Comisiones/Cuotas/Pago
  // — todos comparten la misma estructura (Gasto/Ingreso con la tarjeta como
  // Medio), así que es solo cuestión de ajustar las banderas. "Pago" acá no
  // implica una transferencia real desde otra cuenta: es para gastos que
  // funcionan como crédito contra la facturación (ej. un reembolso).
  function cambiarTransaccion(m, valor) {
    if (valor === m._transaccion) return;
    setMovimientos(
      movimientos.map((mm) =>
        mm.id === m.id
          ? {
              ...mm,
              tipoMov: "gasto_ingreso",
              medio: tarjeta,
              cuentaBalance: "",
              esComision: valor === "Comisiones",
              esCuota: valor === "Cuotas",
              esPago: valor === "Pago",
              esRefinanciamiento: false,
            }
          : mm
      )
    );
  }

  const coloresTransaccion = {
    Gasto: { bg: "#F2E8D8", text: "#8C6E2F" },
    Cuotas: { bg: "#E3EAF0", text: "#33587A" },
    Comisiones: { bg: "#F3E3DE", text: "#9C4A2E" },
    Pago: { bg: "#E4EDE3", text: "#3B6E4F" },
    Refinanciamiento: { bg: "#EEEBF6", text: "#534AB7" },
  };

  return (
    <div>
      <div className="flex gap-2 mb-3 flex-wrap">
        {CUENTAS_CUOTAS.map((t) => (
          <button
            key={t}
            onClick={() => setTarjeta(t)}
            className="flex items-center gap-2 rounded-full"
            style={{
              padding: "0.45rem 0.9rem",
              border: `1px solid ${tarjeta === t ? COLORS.finanzas : "#DAD6C9"}`,
              background: tarjeta === t ? COLORS.finanzas : "#fff",
              color: tarjeta === t ? "#fff" : COLORS.ink,
              fontFamily: "'Work Sans', sans-serif",
              fontSize: "0.85rem",
              cursor: "pointer",
            }}
          >
            {t}
          </button>
        ))}
      </div>

      <div className="flex items-center gap-2 mb-4 flex-wrap">
        <span style={{ fontSize: "0.85rem", color: COLORS.sub }}>Mes Impacto</span>
        <input type="month" style={{ ...inputStyle, width: "10rem" }} value={mes} onChange={(e) => setMes(e.target.value)} />
        <span style={{ fontSize: "0.85rem", color: COLORS.sub, marginLeft: "0.6rem" }}>Fecha de facturación</span>
        <input
          type="date"
          style={{ ...inputStyle, width: "10rem" }}
          value={fechasFacturacion[`${tarjeta}|${mes}`] || ""}
          onChange={(e) => setFechasFacturacion({ ...fechasFacturacion, [`${tarjeta}|${mes}`]: e.target.value })}
        />
      </div>

      <div className="rounded-2xl mb-4" style={{ background: COLORS.card, padding: "1.2rem" }}>
        <p style={{ fontFamily: "'Fraunces', serif", fontSize: "1.05rem", marginBottom: "0.6rem" }}>
          Facturación {tarjeta} — {fmtMesImpacto(mes)}
          {fechasFacturacion[`${tarjeta}|${mes}`] && <span style={{ fontSize: "0.8rem", color: COLORS.sub, fontFamily: "'Work Sans', sans-serif", marginLeft: "0.6rem" }}>({fechasFacturacion[`${tarjeta}|${mes}`]})</span>}
        </p>
        {facturadoMesPasado !== 0 && <Row label="Facturado Mes Pasado" value={fmtMoney(facturadoMesPasado)} />}
        <Row label="Gasto" value={fmtMoney(totalGasto)} />
        <Row label="Cuotas" value={fmtMoney(totalCuotas)} />
        <Row label="Comisiones" value={fmtMoney(totalComisiones)} />
        <Row label="Pago" value={fmtMoney(totalPago)} />
        {totalRefinanciamiento !== 0 && <Row label="Refinanciamiento" value={fmtMoney(totalRefinanciamiento)} />}
        <Row label="Total" value={fmtMoney(totalFacturado)} strong color={COLORS.finanzas} />
      </div>

      {movsDelMes.length === 0 && proyeccion.length === 0 ? (
        <p style={{ color: COLORS.sub, fontSize: "0.85rem" }}>Sin movimientos de {tarjeta} en ese Mes Impacto.</p>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.85rem" }}>
            <thead>
              {monthColumns.length > 1 && (
                <tr>
                  <th style={thStyle}></th>
                  <th style={thStyle}></th>
                  <th style={{ ...thStyle, textAlign: "left" }}>Total</th>
                  {monthColumns.map((mesCol, i) => (
                    <th key={mesCol} style={{ ...thStyle, textAlign: "right" }}>
                      {i > 0 && (
                        <span style={{ background: "#F2E8D8", color: "#8C6E2F", fontSize: "0.78rem", fontWeight: 600, padding: "0.15rem 0.6rem", borderRadius: "999px", display: "inline-block" }}>
                          {fmtMoney((proyeccionPorMes[mesCol] || 0) + movsDelMes.reduce((a, m) => a + (valorCelda(m, mesCol) || 0), 0))}
                        </span>
                      )}
                    </th>
                  ))}
                </tr>
              )}
              {proyeccion.length > 0 && (
                <tr>
                  <th style={thStyle}></th>
                  <th style={thStyle}></th>
                  <th style={{ ...thStyle, textAlign: "left" }}>Saldo Cuotas</th>
                  {monthColumns.map((mesCol) => (
                    <th key={mesCol} style={{ ...thStyle, textAlign: "right" }}>
                      <span style={{ background: "#E3EAF0", color: "#33587A", fontSize: "0.78rem", fontWeight: 600, padding: "0.15rem 0.6rem", borderRadius: "999px", display: "inline-block" }}>
                        {fmtMoney(proyeccionPorMes[mesCol] || 0)}
                      </span>
                    </th>
                  ))}
                </tr>
              )}
              <tr style={{ textAlign: "left" }}>
                <th style={thStyle}>Fecha</th>
                <th style={thStyle}>Transacción</th>
                <th style={thStyle}>Descripción</th>
                {monthColumns.map((mesCol) => (
                  <th key={mesCol} style={{ ...thStyle, textAlign: "right" }}>{fmtMesImpacto(mesCol)}</th>
                ))}
              </tr>
              <tr>
                <th style={thFiltroStyle}>
                  <select style={filtroInputStyle} value={filtrosFact.fecha} onChange={(e) => setFiltrosFact({ ...filtrosFact, fecha: e.target.value })}>
                    <option value="">Todas</option>
                    {fechasDisponiblesFact.map((f) => (<option key={f} value={f}>{f}</option>))}
                  </select>
                </th>
                <th style={thFiltroStyle}>
                  <select style={filtroInputStyle} value={filtrosFact.transaccion} onChange={(e) => setFiltrosFact({ ...filtrosFact, transaccion: e.target.value })}>
                    <option value="">Todas</option>
                    {transaccionesDisponibles.map((t) => (<option key={t} value={t}>{t}</option>))}
                  </select>
                </th>
                <th style={thFiltroStyle}><input style={filtroInputStyle} placeholder="Filtrar..." value={filtrosFact.descripcion} onChange={(e) => setFiltrosFact({ ...filtrosFact, descripcion: e.target.value })} /></th>
                {monthColumns.map((mesCol) => (<th key={mesCol} style={thFiltroStyle}></th>))}
              </tr>
            </thead>
            <tbody>
              {movsFiltrados.length === 0 ? (
                <tr>
                  <td colSpan={3 + monthColumns.length} style={{ padding: "0.9rem 0.5rem", color: COLORS.sub, fontSize: "0.85rem" }}>
                    {movsDelMes.length === 0
                      ? `Sin movimientos de ${tarjeta} en ese Mes Impacto (más allá de Saldo Cuotas).`
                      : `Ningún movimiento coincide con los filtros (tienes ${movsDelMes.length} en total).`}
                  </td>
                </tr>
              ) : (
                movsFiltrados.map((m) => (
                  <tr key={m.id} style={{ borderTop: `1px solid ${COLORS.line}` }}>
                    <td style={tdStyle}><span style={{ color: COLORS.sub }}>{m.fecha}</span></td>
                    <td style={tdStyle}>
                      <select
                        value={m._transaccion}
                        onChange={(e) => cambiarTransaccion(m, e.target.value)}
                        style={{
                          ...inputStyle,
                          width: "8.5rem",
                          fontSize: "0.78rem",
                          padding: "0.2rem 0.4rem",
                          background: coloresTransaccion[m._transaccion]?.bg,
                          color: coloresTransaccion[m._transaccion]?.text,
                          border: "none",
                          borderRadius: "999px",
                        }}
                      >
                        {m._transaccion === "Refinanciamiento" && <option value="Refinanciamiento">Refinanciamiento</option>}
                        <option value="Gasto">Gasto</option>
                        <option value="Comisiones">Comisiones</option>
                        <option value="Cuotas">Cuotas</option>
                        <option value="Pago">Pago</option>
                      </select>
                    </td>
                    <td style={tdStyle}>{m.descripcion}</td>
                    {monthColumns.map((mesCol) => {
                      const valor = valorCelda(m, mesCol);
                      return (
                        <td key={mesCol} style={{ ...tdStyle, textAlign: "right", fontWeight: 500 }}>
                          {valor !== null ? fmtMoney(valor) : ""}
                        </td>
                      );
                    })}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}


/* ================= DEPORTES (placeholder) ================= */
function Deportes({ onBack }) {
  return (
    <div className="w-full max-w-md">
      <button onClick={onBack} className="flex items-center gap-1 mb-4" style={{ color: COLORS.sub, background: "none", border: "none", cursor: "pointer", fontFamily: "'Work Sans', sans-serif" }}>
        <ArrowLeft size={18} /> Inicio
      </button>
      <p style={{ fontFamily: "'Fraunces', serif", fontSize: "1.4rem", fontWeight: 600, color: COLORS.deportes }}>Deportes</p>
      <p style={{ color: COLORS.sub, fontSize: "0.9rem", marginTop: "0.6rem" }}>Este módulo todavía no tiene contenido — lo definimos en el próximo paso.</p>
    </div>
  );
}

/* ================= APP ================= */
function Login() {
  const [modo, setModo] = useState("entrar"); // "entrar" | "crear"
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [cargando, setCargando] = useState(false);
  const [avisoConfirmacion, setAvisoConfirmacion] = useState(false);

  async function enviar() {
    setError("");
    if (!email.trim() || !password) { setError("Completa email y contraseña."); return; }
    setCargando(true);
    if (modo === "crear") {
      const { error: err } = await supabase.auth.signUp({ email: email.trim(), password });
      if (err) setError(err.message);
      else setAvisoConfirmacion(true);
    } else {
      const { error: err } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
      if (err) setError(err.message);
    }
    setCargando(false);
  }

  return (
    <div style={{ background: COLORS.bg, minHeight: "100vh", fontFamily: "'Work Sans', sans-serif", color: COLORS.ink }} className="flex flex-col items-center justify-center px-6">
      <div className="rounded-2xl w-full max-w-sm" style={{ background: COLORS.card, padding: "2rem" }}>
        <p style={{ fontFamily: "'Fraunces', serif", fontSize: "1.3rem", marginBottom: "1.2rem", textAlign: "center" }}>Mis Finanzas</p>
        {avisoConfirmacion ? (
          <p style={{ fontSize: "0.9rem", textAlign: "center" }}>
            Te enviamos un correo a <strong>{email}</strong> para confirmar tu cuenta. Ábrelo y luego vuelve a iniciar sesión acá.
          </p>
        ) : (
          <>
            <div className="flex flex-col gap-2 mb-3">
              <input type="email" placeholder="Email" value={email} onChange={(e) => setEmail(e.target.value)} style={{ ...inputStyle }} />
              <input
                type="password"
                placeholder="Contraseña"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                style={{ ...inputStyle }}
                onKeyDown={(e) => e.key === "Enter" && enviar()}
              />
            </div>
            {error && <p style={{ color: "#9C4A2E", fontSize: "0.82rem", marginBottom: "0.6rem" }}>{error}</p>}
            <button onClick={enviar} disabled={cargando} className="w-full" style={{ ...btnStyle(COLORS.finanzas), justifyContent: "center", marginBottom: "0.8rem" }}>
              {cargando ? "Un momento..." : modo === "crear" ? "Crear cuenta" : "Iniciar sesión"}
            </button>
            <p style={{ fontSize: "0.82rem", color: COLORS.sub, textAlign: "center" }}>
              {modo === "crear" ? "¿Ya tienes cuenta? " : "¿Primera vez? "}
              <button onClick={() => { setModo(modo === "crear" ? "entrar" : "crear"); setError(""); }} style={{ background: "none", border: "none", color: COLORS.finanzas, cursor: "pointer", textDecoration: "underline", padding: 0 }}>
                {modo === "crear" ? "Inicia sesión" : "Crea una"}
              </button>
            </p>
          </>
        )}
      </div>
    </div>
  );
}

export default function App() {
  const [view, setView] = useState("home");
  const [session, setSession] = useState(undefined); // undefined = cargando, null = sin sesión

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => listener.subscription.unsubscribe();
  }, []);

  if (session === undefined) {
    return (
      <div style={{ background: COLORS.bg, minHeight: "100vh" }} className="flex items-center justify-center">
        <p style={{ color: COLORS.sub }}>Cargando...</p>
      </div>
    );
  }
  if (!session) return <Login />;

  return (
    <div style={{ background: COLORS.bg, minHeight: "100vh", fontFamily: "'Work Sans', sans-serif", color: COLORS.ink }} className="flex flex-col items-center px-6 py-12">
      <div className="w-full max-w-2xl flex justify-end mb-2">
        <button onClick={() => supabase.auth.signOut()} style={{ background: "none", border: "none", color: COLORS.sub, cursor: "pointer", fontSize: "0.8rem" }}>
          Cerrar sesión ({session.user.email})
        </button>
      </div>
      {view === "home" && <Home onOpen={setView} />}
      {view === "finanzas" && <Finanzas onBack={() => setView("home")} userId={session.user.id} />}
      {view === "deportes" && <Deportes onBack={() => setView("home")} />}
    </div>
  );
}
