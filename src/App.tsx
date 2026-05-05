import React, { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid,
  Legend,
  Cell,
  LabelList,
} from "recharts";

type ExportRow = Record<string, any>;
type RankingMode = "TOTAL_H" | "PCT";
type TopReason = { motivo: string; horas: number; pct01: number };
type MachineAgg = {
  maquina: string;
  trabAvg: number;
  transpAvg: number;
  ociosoAvg: number;
  canceladoH: number;
  ignoradoH: number;
  desligadoH: number;
  indefinidoH: number;
  naoApontadoH: number;
  topReasons: TopReason[];
  totalReasonsH: number;
};

const COLORS = {
  TRAB: "#16a34a",
  TRANSP: "#6b7280",
  OCIOSO: "#f97316",
  CANCELADO: "#ef4444",
  IGNORADO: "#111827",
  BROWN_1: "#8b451c",
  BROWN_2: "#bc5d2a",
  ORANGE: "#ff6b2c",
};

const LABEL_STYLE = { fill: "#111827", fontSize: 12, fontWeight: 700 as const };
const TICK_STYLE = { fill: "#111827", fontSize: 11 };

function num(v: any): number {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v === "string") {
    const s = v.replace("%", "").replace(/\s/g, "").replace(",", ".");
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
}
function fmtPercent01(v01: number) {
  return `${(v01 * 100).toFixed(2)}%`;
}
function fmtHours(h: number) {
  return h.toFixed(2).replace(".", ",");
}
function labelHours(v: any) {
  return `${fmtHours(num(v))}`;
}
function labelPercent(v: any) {
  return fmtPercent01(num(v));
}
function uniqueSorted(arr: string[]) {
  return Array.from(new Set(arr)).sort((a, b) => a.localeCompare(b));
}
function sumCols(row: ExportRow, a: string, b: string) {
  return num(row[a]) + num(row[b]);
}
function heatColor(t: number): string {
  const clamp = (x: number) => Math.max(0, Math.min(1, x));
  const tt = clamp(t);
  const green = { r: 34, g: 197, b: 94 };
  const yellow = { r: 250, g: 204, b: 21 };
  const red = { r: 239, g: 68, b: 68 };
  let a = green;
  let b = yellow;
  let u = 0;
  if (tt <= 0.5) {
    a = green;
    b = yellow;
    u = tt / 0.5;
  } else {
    a = yellow;
    b = red;
    u = (tt - 0.5) / 0.5;
  }
  const r = Math.round(a.r + (b.r - a.r) * u);
  const g = Math.round(a.g + (b.g - a.g) * u);
  const bb = Math.round(a.b + (b.b - a.b) * u);
  return `rgb(${r}, ${g}, ${bb})`;
}

function getReasonTotals(rows: ExportRow[]): { all: Array<{ motivo: string; horas: number }>; totalAll: number } {
  if (!rows.length) return { all: [], totalAll: 0 };
  const cols = Object.keys(rows[0] ?? {});
  const periodCols = cols.filter((c) => c.endsWith("Período (h)"));
  const exclude = new Set(["Tempo Desligado", "Indefinido", "Cancelado", "Ignorado"]);
  const roots = uniqueSorted(
    periodCols
      .map((c) => c.replace(" Período (h)", ""))
      .filter((r) => r.trim().length > 0 && !exclude.has(r))
  );
  const all = roots.map((root) => {
    const p = `${root} Período (h)`;
    const d = `${root} Desligado (h)`;
    let total = 0;
    for (const r of rows) total += num(r[p]) + num(r[d]);
    return { motivo: root, horas: total };
  });
  all.sort((a, b) => b.horas - a.horas);
  const totalAll = all.reduce((acc, r) => acc + r.horas, 0);
  return { all, totalAll };
}

function buildReasonRanking(rows: ExportRow[], topN = 10): { top: TopReason[]; totalAll: number } {
  const { all, totalAll } = getReasonTotals(rows);
  return {
    top: all.slice(0, topN).map((r) => ({ motivo: r.motivo, horas: r.horas, pct01: totalAll > 0 ? r.horas / totalAll : 0 })),
    totalAll,
  };
}

function groupByMachine(rows: ExportRow[]): MachineAgg[] {
  const map = new Map<string, ExportRow[]>();
  for (const r of rows) {
    const m = String(r["Máquina"] ?? "").trim();
    if (!m) continue;
    if (!map.has(m)) map.set(m, []);
    map.get(m)!.push(r);
  }
  const out: MachineAgg[] = [];
  for (const [maquina, items] of map.entries()) {
    const trabAvg = items.reduce((acc, x) => acc + num(x["Utilização Trabalho (%)"]), 0) / items.length;
    const transpAvg = items.reduce((acc, x) => acc + num(x["Utilização Transporte (%)"]), 0) / items.length;
    const ociosoAvg = items.reduce((acc, x) => acc + num(x["Utilização Ocioso (%)"]), 0) / items.length;
    const canceladoH = items.reduce((acc, x) => acc + sumCols(x, "Cancelado Período (h)", "Cancelado Desligado (h)"), 0);
    const indefinidoH = items.reduce((acc, x) => acc + sumCols(x, "Indefinido Período (h)", "Indefinido Desligado (h)"), 0);
    const ignoradoH = items.reduce((acc, x) => acc + sumCols(x, "Ignorado Período (h)", "Ignorado Desligado (h)"), 0);
    const desligadoH = items.reduce((acc, x) => acc + num(x["Tempo Desligado Período (h)"]), 0);
    const { top, totalAll } = buildReasonRanking(items, 3);
    out.push({
      maquina,
      trabAvg,
      transpAvg,
      ociosoAvg,
      canceladoH,
      ignoradoH,
      desligadoH,
      indefinidoH,
      naoApontadoH: indefinidoH + ignoradoH,
      topReasons: top,
      totalReasonsH: totalAll,
    });
  }
  out.sort((a, b) => b.trabAvg - a.trabAvg);
  return out;
}

function ProgressBar({ value01, color }: { value01: number; color: string }) {
  return (
    <div style={{ height: 15, borderRadius: 999, background: "#eee6dc", overflow: "hidden" }}>
      <div style={{ width: `${Math.max(3, Math.min(100, value01 * 100))}%`, height: "100%", borderRadius: 999, background: color }} />
    </div>
  );
}

export default function App() {
  const [fileName, setFileName] = useState("");
  const [rows, setRows] = useState<ExportRow[]>([]);
  const [error, setError] = useState("");
  const [machineFilter, setMachineFilter] = useState("__ALL__");
  const [rankingMode, setRankingMode] = useState<RankingMode>("TOTAL_H");
  const [isExporting, setIsExporting] = useState(false);
  const [goalWork, setGoalWork] = useState(0.5);
  const [goalTransport, setGoalTransport] = useState(0.25);
  const [goalIdle, setGoalIdle] = useState(0.25);
  const pdfRef = useRef<HTMLDivElement | null>(null);

  async function handleFile(file: File) {
    setError("");
    setFileName(file.name);
    setMachineFilter("__ALL__");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets["Exportar"] ?? wb.Sheets[wb.SheetNames[0]];
      if (!ws) {
        setError("Não encontrei nenhuma aba no arquivo.");
        setRows([]);
        return;
      }
      const json = XLSX.utils.sheet_to_json<ExportRow>(ws, { defval: null });
      if (!json.length) {
        setError("A aba está vazia (sem linhas).");
        setRows([]);
        return;
      }
      const needed = ["Máquina", "Utilização Trabalho (%)", "Utilização Transporte (%)", "Utilização Ocioso (%)"];
      const missing = needed.filter((c) => !(c in (json[0] ?? {})));
      if (missing.length) {
        setError(`Arquivo não parece ser o export padrão. Colunas ausentes: ${missing.join(", ")}`);
        setRows(json);
        return;
      }
      setRows(json);
    } catch (e: any) {
      setError(e?.message ?? "Falha ao ler o arquivo .xlsx");
      setRows([]);
    }
  }

  async function exportPDFSingleA4() {
    if (!pdfRef.current || !fileName || rows.length === 0) {
      alert("Faça upload do Excel antes de exportar o PDF.");
      return;
    }
    setIsExporting(true);
    (document.activeElement as HTMLElement | null)?.blur?.();
    try {
      await new Promise((r) => requestAnimationFrame(() => r(null)));
      await new Promise((r) => setTimeout(r, 500));
      const canvas = await html2canvas(pdfRef.current, { backgroundColor: "#ffffff", scale: 4, useCORS: true, logging: false });
      const imgData = canvas.toDataURL("image/jpeg", 0.95);
      const pdf = new jsPDF("p", "pt", "a4");
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const scale = Math.min(pageW / canvas.width, pageH / canvas.height);
      const drawW = canvas.width * scale;
      const drawH = canvas.height * scale;
      pdf.addImage(imgData, "JPEG", (pageW - drawW) / 2, (pageH - drawH) / 2, drawW, drawH);
      const base = fileName.replace(/\.[^/.]+$/, "");
      pdf.save(`${base}_RELATORIO_OPERACIONAL_A4.pdf`);
    } catch (e) {
      console.error(e);
      alert("Não foi possível gerar o PDF. Abra o console (F12) para ver o erro.");
    } finally {
      setIsExporting(false);
    }
  }

  const perMachineAll = useMemo(() => groupByMachine(rows), [rows]);
  const machineOptions = useMemo(() => uniqueSorted(perMachineAll.map((x) => x.maquina)), [perMachineAll]);
  const filteredMachines = useMemo(() => machineFilter === "__ALL__" ? perMachineAll : perMachineAll.filter((m) => m.maquina === machineFilter), [perMachineAll, machineFilter]);

  const operationAvg = useMemo(() => {
    if (!filteredMachines.length) return { trab: 0, transp: 0, ocioso: 0 };
    return {
      trab: filteredMachines.reduce((a, m) => a + m.trabAvg, 0) / filteredMachines.length,
      transp: filteredMachines.reduce((a, m) => a + m.transpAvg, 0) / filteredMachines.length,
      ocioso: filteredMachines.reduce((a, m) => a + m.ociosoAvg, 0) / filteredMachines.length,
    };
  }, [filteredMachines]);

  const deviations = useMemo(() => {
    const arr = [
      { key: "trab", label: "Trabalho abaixo da meta", value: Math.max(0, goalWork - operationAvg.trab) },
      { key: "transp", label: "Transporte acima da meta", value: Math.max(0, operationAvg.transp - goalTransport) },
      { key: "ocioso", label: "Ocioso acima da meta", value: Math.max(0, operationAvg.ocioso - goalIdle) },
    ];
    arr.sort((a, b) => b.value - a.value);
    return arr;
  }, [operationAvg, goalWork, goalTransport, goalIdle]);

  const statusGeral = deviations.some((d) => d.value > 0) ? "Operação fora dos parâmetros" : "Operação dentro dos parâmetros";
  const principalDesvio = deviations[0]?.value > 0 ? deviations[0].label : "Sem desvio relevante";

  const criticalEquipment = useMemo(() => {
    if (!filteredMachines.length) return "-";
    const scored = filteredMachines.map((m) => ({
      maquina: m.maquina,
      score: Math.max(0, goalWork - m.trabAvg) + Math.max(0, m.transpAvg - goalTransport) + Math.max(0, m.ociosoAvg - goalIdle) + m.canceladoH / 100 + m.ignoradoH / 100,
    }));
    scored.sort((a, b) => b.score - a.score);
    return scored[0]?.maquina ?? "-";
  }, [filteredMachines, goalWork, goalTransport, goalIdle]);

  const rankingBundle = useMemo(() => {
    const baseRows = machineFilter === "__ALL__" ? rows : rows.filter((r) => String(r["Máquina"] ?? "").trim() === machineFilter);
    return buildReasonRanking(baseRows, 10);
  }, [rows, machineFilter]);
  const top3General = rankingBundle.top.slice(0, 3);

  const utilChart = useMemo(() => [
    { name: "Trabalho", value: operationAvg.trab, color: COLORS.TRAB },
    { name: "Transporte", value: operationAvg.transp, color: COLORS.TRANSP },
    { name: "Ocioso", value: operationAvg.ocioso, color: COLORS.OCIOSO },
  ], [operationAvg]);

  const cancelVsIgnoradoChart = useMemo(() => {
    const base = filteredMachines.slice();
    if (machineFilter === "__ALL__") return base.slice(0, 14);
    return base;
  }, [filteredMachines, machineFilter]);

  const rankingChartData = useMemo(() => rankingBundle.top.map((r) => ({ ...r, valor: rankingMode === "PCT" ? r.pct01 : r.horas })), [rankingBundle, rankingMode]);

  const rankingColors = useMemo(() => {
    if (!rankingChartData.length) return [] as string[];
    const values = rankingChartData.map((r) => num(r.valor));
    const min = Math.min(...values);
    const max = Math.max(...values);
    const denom = max - min || 1;
    return rankingChartData.map((r) => heatColor((num(r.valor) - min) / denom));
  }, [rankingChartData]);

  const utilYAxisMax = useMemo(() => {
    const maxVal = Math.max(...utilChart.map((d) => num(d.value)), goalWork, goalTransport, goalIdle, 0);
    return maxVal <= 0 ? 1 : maxVal * 1.1;
  }, [utilChart, goalWork, goalTransport, goalIdle]);

  const cancelIgnoradoYAxisMax = useMemo(() => {
    if (!cancelVsIgnoradoChart.length) return 5;
    const maxVal = Math.max(...cancelVsIgnoradoChart.flatMap((d) => [num(d.canceladoH), num(d.ignoradoH)]), 0);
    return maxVal <= 0 ? 5 : maxVal + 5;
  }, [cancelVsIgnoradoChart]);

  const top3ByEquipment = useMemo(() => {
    const base = machineFilter === "__ALL__" ? perMachineAll.slice(0, 8) : filteredMachines;
    return base.map((m) => ({ maquina: m.maquina, top: m.topReasons, total: m.totalReasonsH }));
  }, [perMachineAll, filteredMachines, machineFilter]);

  const show = rows.length > 0;

  return (
    <div style={{ padding: 18, maxWidth: 1280, margin: "0 auto", background: "#f7f2ec", minHeight: "100vh" }}>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
        <div>
          <h2 style={{ margin: 0 }}>Relatório Operacional — Gráficos e Resumo</h2>
          <p style={{ marginTop: 8, opacity: 0.75 }}>Upload do Excel com aba <b>Exportar</b>.</p>
        </div>
        <button onClick={exportPDFSingleA4} disabled={!show || isExporting} style={{ padding: "10px 14px", borderRadius: 12, border: "1px solid #e5e7eb", background: !show || isExporting ? "#f3f4f6" : "#111827", color: !show || isExporting ? "#6b7280" : "#fff", cursor: !show || isExporting ? "not-allowed" : "pointer", fontWeight: 800 }}>
          {isExporting ? "Gerando PDF..." : "Exportar PDF A4"}
        </button>
      </div>

      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center", marginTop: 12, background: "#fff", padding: 12, borderRadius: 14, border: "1px solid #e5d8ca" }}>
        <input type="file" accept=".xlsx" onChange={(e) => { const f = e.target.files?.[0]; if (f) void handleFile(f); }} />
        {fileName ? <span>Arquivo: <b>{fileName}</b></span> : null}
        {machineOptions.length ? (
          <>
            <span>Filtro de máquina:</span>
            <select value={machineFilter} onChange={(e) => setMachineFilter(e.target.value)} style={{ padding: "6px 10px", borderRadius: 10 }}>
              <option value="__ALL__">Todas</option>
              {machineOptions.map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
          </>
        ) : null}
        <span>Ranking em:</span>
        <select value={rankingMode} onChange={(e) => setRankingMode(e.target.value as RankingMode)} style={{ padding: "6px 10px", borderRadius: 10 }}>
          <option value="TOTAL_H">Total (h)</option>
          <option value="PCT">Porcentagem (%)</option>
        </select>
      </div>

      <div style={{ marginTop: 12, background: "#fff", padding: 14, borderRadius: 14, border: "1px solid #e5d8ca" }}>
        <h3 style={{ margin: "0 0 10px" }}>Metas de Rendimento Operacional</h3>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 12 }}>
          <label>Tempo em Trabalho (%)<input type="number" value={(goalWork * 100).toFixed(0)} onChange={(e) => setGoalWork(num(e.target.value) / 100)} style={{ width: "100%", marginTop: 6, padding: 8, borderRadius: 10, border: "1px solid #ddd" }} /></label>
          <label>Tempo em Transporte (%)<input type="number" value={(goalTransport * 100).toFixed(0)} onChange={(e) => setGoalTransport(num(e.target.value) / 100)} style={{ width: "100%", marginTop: 6, padding: 8, borderRadius: 10, border: "1px solid #ddd" }} /></label>
          <label>Tempo Ocioso (%)<input type="number" value={(goalIdle * 100).toFixed(0)} onChange={(e) => setGoalIdle(num(e.target.value) / 100)} style={{ width: "100%", marginTop: 6, padding: 8, borderRadius: 10, border: "1px solid #ddd" }} /></label>
        </div>
      </div>

      {error && <div style={{ marginTop: 12, padding: 12, background: "#ffe5e5", borderRadius: 10 }}><b>Erro:</b> {error}</div>}
      {!show && !error && <div style={{ marginTop: 16, padding: 14, border: "1px dashed #d1d5db", borderRadius: 12, background: "#fff" }}>Faça upload do Excel para gerar o relatório.</div>}

      {show && (
        <div ref={pdfRef} style={{ marginTop: 16, background: "#fff", border: "1px solid #e5d8ca", borderRadius: 14, padding: 16, width: "min(1050px, 100%)", marginLeft: "auto", marginRight: "auto" }}>
          <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "baseline" }}>
            <div style={{ fontSize: 19, fontWeight: 900 }}>Relatório de Gráficos — {machineFilter === "__ALL__" ? "Todas as máquinas" : machineFilter}</div>
            <div style={{ fontSize: 11, opacity: 0.7 }}>Fonte: {fileName} • {new Date().toLocaleString()}</div>
          </div>

          <section style={{ marginTop: 12, border: statusGeral.includes("fora") ? "1px solid #ef4444" : "1px solid #16a34a", background: statusGeral.includes("fora") ? "#fff1f2" : "#f0fdf4", borderRadius: 14, padding: 14 }}>
            <div style={{ fontSize: 12, letterSpacing: 1 }}>STATUS EXECUTIVO</div>
            <div style={{ color: statusGeral.includes("fora") ? "#c82828" : "#15803d", fontSize: 25, fontWeight: 900, marginTop: 4 }}>{statusGeral}</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 8, marginTop: 10, fontSize: 13 }}>
              <div><b>Principal desvio:</b> {principalDesvio}</div>
              <div><b>Equipamento crítico:</b> {criticalEquipment}</div>
              <div><b>Metas:</b> Trabalho {fmtPercent01(goalWork)} | Transporte {fmtPercent01(goalTransport)} | Ocioso {fmtPercent01(goalIdle)}</div>
            </div>
          </section>

          <section style={{ marginTop: 12, border: "1px solid #e5d8ca", borderRadius: 14, padding: 14, background: "#fff" }}>
            <h2 style={{ margin: 0 }}>Resumo da Operação</h2>
            <div style={{ marginTop: 8, background: "#fbf8f4", border: "1px solid #e5d8ca", borderRadius: 14, padding: 12, fontSize: 15, lineHeight: 1.65 }}>
              <div><b>Status geral:</b> {statusGeral}</div>
              <div><b>Principal desvio:</b> {principalDesvio}</div>
              <div><b>Equipamento crítico:</b> {criticalEquipment}</div>
              <div><b>Principais motivos de parada:</b> {top3General.map((r) => r.motivo).join("; ") || "-"}</div>
            </div>
          </section>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginTop: 12 }}>
            <section>
              <h3 style={{ margin: "0 0 6px" }}>Utilização (média)</h3>
              <div style={{ height: 235, border: "1px solid #eee", borderRadius: 12, padding: 8 }}>
                <ResponsiveContainer>
                  <BarChart data={utilChart} margin={{ top: 26, right: 16, left: 8, bottom: 8 }}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="name" tick={TICK_STYLE} interval={0} />
                    <YAxis tick={TICK_STYLE} domain={[0, utilYAxisMax]} tickFormatter={(v: any) => fmtPercent01(num(v))} />
                    <Tooltip formatter={(v: any, n: any) => [fmtPercent01(num(v)), n]} />
                    <Bar dataKey="value" name="Utilização" isAnimationActive={false}>
                      {utilChart.map((entry, i) => <Cell key={`u-${i}`} fill={entry.color} />)}
                      <LabelList dataKey="value" position="top" offset={6} formatter={(v: any) => labelPercent(v)} {...LABEL_STYLE} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </section>

            <section>
              <h3 style={{ margin: "0 0 6px" }}>Cancelado x Ignorado (h)</h3>
              <div style={{ height: 235, border: "1px solid #eee", borderRadius: 12, padding: 8 }}>
                <ResponsiveContainer>
                  <BarChart data={cancelVsIgnoradoChart} margin={{ top: 26, right: 16, left: 8, bottom: 24 }}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="maquina" interval={0} angle={0} height={30} tick={TICK_STYLE} />
                    <YAxis tick={TICK_STYLE} domain={[0, cancelIgnoradoYAxisMax]} tickFormatter={(v: any) => fmtHours(num(v))} />
                    <Tooltip formatter={(v: any, n: any) => [fmtHours(num(v)), n]} />
                    <Legend />
                    <Bar dataKey="canceladoH" name="Cancelado (h)" fill={COLORS.CANCELADO} isAnimationActive={false}>
                      <LabelList dataKey="canceladoH" position="top" offset={6} formatter={(v: any) => labelHours(v)} {...LABEL_STYLE} />
                    </Bar>
                    <Bar dataKey="ignoradoH" name="Ignorado (h)" fill={COLORS.IGNORADO} isAnimationActive={false}>
                      <LabelList dataKey="ignoradoH" position="top" offset={6} formatter={(v: any) => labelHours(v)} {...LABEL_STYLE} />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </section>
          </div>

          <section style={{ marginTop: 12 }}>
            <h3 style={{ margin: "0 0 6px" }}>Ranking — Motivos de Parada (Top 10)</h3>
            <div style={{ height: 310, border: "1px solid #eee", borderRadius: 12, padding: 8 }}>
              <ResponsiveContainer>
                <BarChart data={rankingChartData} layout="vertical" margin={{ top: 10, left: 20, right: 45, bottom: 10 }}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis type="number" tick={TICK_STYLE} tickFormatter={(v: any) => rankingMode === "PCT" ? fmtPercent01(num(v)) : fmtHours(num(v))} />
                  <YAxis dataKey="motivo" type="category" width={260} tick={TICK_STYLE} />
                  <Tooltip formatter={(v: any) => rankingMode === "PCT" ? [fmtPercent01(num(v)), "Participação (%)"] : [fmtHours(num(v)), "Tempo total (h)"]} labelFormatter={(l) => `Motivo: ${l}`} />
                  <Bar dataKey="valor" name={rankingMode === "PCT" ? "Participação (%)" : "Tempo total (h)"} isAnimationActive={false}>
                    {rankingChartData.map((_, i) => <Cell key={`r-${i}`} fill={rankingColors[i]} />)}
                    <LabelList dataKey="valor" position="right" formatter={(v: any) => rankingMode === "PCT" ? labelPercent(v) : labelHours(v)} {...LABEL_STYLE} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </section>

          <section style={{ marginTop: 12, border: "1px solid #e5d8ca", borderRadius: 14, padding: 14, background: "#fffaf5" }}>
            <h2 style={{ margin: 0 }}>Top 3 Motivos por Equipamento</h2>
            <div style={{ fontSize: 13, marginTop: 4 }}>Visão individual por equipamento para priorização operacional.</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginTop: 12 }}>
              {top3ByEquipment.map((eq) => {
                const max = Math.max(...eq.top.map((x) => rankingMode === "PCT" ? x.pct01 : x.horas), 0.0001);
                return (
                  <div key={eq.maquina} style={{ border: "1px solid #e5d8ca", borderRadius: 14, padding: 12, background: "#fff" }}>
                    <div style={{ color: COLORS.BROWN_1, fontWeight: 900, fontSize: 16, marginBottom: 10 }}>{eq.maquina}</div>
                    {eq.top.map((r, idx) => {
                      const value = rankingMode === "PCT" ? r.pct01 : r.horas;
                      return (
                        <div key={r.motivo} style={{ marginBottom: 9 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", gap: 8, fontSize: 12 }}>
                            <span>{idx + 1}. {r.motivo}</span>
                            <b>{rankingMode === "PCT" ? fmtPercent01(r.pct01) : `${fmtHours(r.horas)} h`}</b>
                          </div>
                          <div style={{ marginTop: 5 }}><ProgressBar value01={value / max} color={idx === 0 ? COLORS.BROWN_1 : idx === 1 ? COLORS.BROWN_2 : COLORS.ORANGE} /></div>
                        </div>
                      );
                    })}
                  </div>
                );
              })}
            </div>
          </section>
        </div>
      )}
    </div>
  );
}
