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

type MachineAgg = {
  maquina: string;
  trabAvg: number;
  transpAvg: number;
  ociosoAvg: number;
  canceladoH: number;
  ignoradoH: number;
  indefinidoH: number;
  naoApontadoH: number;
};

type RankingRow = {
  motivo: string;
  horas: number;
  pct01: number;
};

type TopReason = {
  motivo: string;
  horas: number;
  pct01: number;
};

function num(v: any): number {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v === "string") {
    const s = v.replace("%", "").replace(/\s/g, "").replace(",", ".");
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
}

function pct01(v: any): number {
  const n = num(v);
  return n > 1 ? n / 100 : n;
}

function fmtPercent01(v01: number) {
  return `${(v01 * 100).toFixed(2)}%`;
}

function fmtHours(h: number) {
  return h.toFixed(2).replace(".", ",");
}

function uniqueSorted(arr: string[]) {
  return Array.from(new Set(arr)).sort((a, b) => a.localeCompare(b));
}

function sumCols(row: ExportRow, a: string, b: string) {
  return num(row[a]) + num(row[b]);
}

const LABEL_STYLE = {
  fill: "#111827",
  fontSize: 12,
  fontWeight: 700 as const,
};

const TICK_STYLE = {
  fill: "#111827",
  fontSize: 11,
};

function labelHours(v: any) {
  return fmtHours(num(v));
}

function labelPercent(v: any) {
  return fmtPercent01(num(v));
}

function getReasonRoots(rows: ExportRow[]) {
  if (!rows.length) return [];

  const cols = Object.keys(rows[0] ?? {});
  const periodCols = cols.filter((c) => c.endsWith("Período (h)"));

  const exclude = new Set([
    "Tempo Desligado",
    "Indefinido",
    "Cancelado",
    "Ignorado",
  ]);

  return uniqueSorted(
    periodCols
      .map((c) => c.replace(" Período (h)", ""))
      .filter((r) => r.trim().length > 0 && !exclude.has(r))
  );
}

function getReasonTotal(row: ExportRow, root: string) {
  return num(row[`${root} Período (h)`]) + num(row[`${root} Desligado (h)`]);
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
    const trabAvg =
      items.reduce((acc, x) => acc + pct01(x["Utilização Trabalho (%)"]), 0) /
      items.length;

    const transpAvg =
      items.reduce((acc, x) => acc + pct01(x["Utilização Transporte (%)"]), 0) /
      items.length;

    const ociosoAvg =
      items.reduce((acc, x) => acc + pct01(x["Utilização Ocioso (%)"]), 0) /
      items.length;

    const canceladoH = items.reduce(
      (acc, x) => acc + sumCols(x, "Cancelado Período (h)", "Cancelado Desligado (h)"),
      0
    );

    const ignoradoH = items.reduce(
      (acc, x) => acc + sumCols(x, "Ignorado Período (h)", "Ignorado Desligado (h)"),
      0
    );

    const indefinidoH = items.reduce(
      (acc, x) => acc + sumCols(x, "Indefinido Período (h)", "Indefinido Desligado (h)"),
      0
    );

    out.push({
      maquina,
      trabAvg,
      transpAvg,
      ociosoAvg,
      canceladoH,
      ignoradoH,
      indefinidoH,
      naoApontadoH: indefinidoH + ignoradoH,
    });
  }

  out.sort((a, b) => b.trabAvg - a.trabAvg);
  return out;
}

function buildReasonRanking(rows: ExportRow[], topN = 10): { top: RankingRow[]; totalAll: number } {
  if (!rows.length) return { top: [], totalAll: 0 };

  const roots = getReasonRoots(rows);

  const all = roots.map((root) => {
    let total = 0;
    for (const r of rows) total += getReasonTotal(r, root);
    return { motivo: root, horas: total };
  });

  const totalAll = all.reduce((acc, r) => acc + r.horas, 0);

  all.sort((a, b) => b.horas - a.horas);

  return {
    top: all.slice(0, topN).map((r) => ({
      motivo: r.motivo,
      horas: r.horas,
      pct01: totalAll > 0 ? r.horas / totalAll : 0,
    })),
    totalAll,
  };
}

function buildTop3ByMachine(rows: ExportRow[]) {
  const roots = getReasonRoots(rows);
  const machineMap = new Map<string, ExportRow[]>();

  for (const r of rows) {
    const maquina = String(r["Máquina"] ?? "").trim();
    if (!maquina) continue;

    if (!machineMap.has(maquina)) machineMap.set(maquina, []);
    machineMap.get(maquina)!.push(r);
  }

  const out: Array<{ maquina: string; top3: TopReason[]; total: number }> = [];

  for (const [maquina, items] of machineMap.entries()) {
    const motivos = roots.map((root) => {
      let total = 0;
      for (const r of items) total += getReasonTotal(r, root);
      return { motivo: root, horas: total };
    });

    const total = motivos.reduce((acc, x) => acc + x.horas, 0);

    const top3 = motivos
      .filter((x) => x.horas > 0)
      .sort((a, b) => b.horas - a.horas)
      .slice(0, 3)
      .map((x) => ({
        motivo: x.motivo,
        horas: x.horas,
        pct01: total > 0 ? x.horas / total : 0,
      }));

    if (top3.length) out.push({ maquina, top3, total });
  }

  out.sort((a, b) => {
    const ah = a.top3[0]?.horas ?? 0;
    const bh = b.top3[0]?.horas ?? 0;
    return bh - ah;
  });

  return out;
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

export default function App() {
  const [fileName, setFileName] = useState("");
  const [rows, setRows] = useState<ExportRow[]>([]);
  const [error, setError] = useState("");

  const [machineFilter, setMachineFilter] = useState<string>("__ALL__");
  const [rankingMode, setRankingMode] = useState<RankingMode>("TOTAL_H");

  const [metaTrabalho, setMetaTrabalho] = useState(50);
  const [metaTransporte, setMetaTransporte] = useState(25);
  const [metaOcioso, setMetaOcioso] = useState(25);

  const [isExporting, setIsExporting] = useState(false);

  const pdfRef = useRef<HTMLDivElement | null>(null);

  const COLORS = {
    TRAB: "#16a34a",
    TRANSP: "#6b7280",
    OCIOSO: "#f97316",
    CANCELADO: "#ef4444",
    IGNORADO: "#111827",
  };

  const TOP3_COLORS = ["#dc2626", "#f97316", "#facc15"];

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
        setError("A aba está vazia.");
        setRows([]);
        return;
      }

      const needed = [
        "Máquina",
        "Utilização Trabalho (%)",
        "Utilização Transporte (%)",
        "Utilização Ocioso (%)",
      ];

      const missing = needed.filter((c) => !(c in (json[0] ?? {})));

      if (missing.length) {
        setError(`Colunas ausentes: ${missing.join(", ")}`);
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
    if (!pdfRef.current) {
      alert("Não encontrei a área do relatório.");
      return;
    }

    if (!fileName || rows.length === 0) {
      alert("Faça upload do Excel antes de exportar o PDF.");
      return;
    }

    setIsExporting(true);
    (document.activeElement as HTMLElement | null)?.blur?.();

    try {
      await new Promise((r) => requestAnimationFrame(() => r(null)));
      await new Promise((r) => setTimeout(r, 500));

      const canvas = await html2canvas(pdfRef.current, {
        backgroundColor: "#ffffff",
        scale: 4,
        useCORS: true,
        logging: false,
      });

      const imgData = canvas.toDataURL("image/jpeg", 0.95);

      const pdf = new jsPDF("p", "pt", "a4");
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();

      const imgW = canvas.width;
      const imgH = canvas.height;

      const scale = Math.min(pageW / imgW, pageH / imgH);
      const drawW = imgW * scale;
      const drawH = imgH * scale;

      const x = (pageW - drawW) / 2;
      const y = (pageH - drawH) / 2;

      pdf.addImage(imgData, "JPEG", x, y, drawW, drawH);

      const base = fileName.replace(/\.[^/.]+$/, "");
      pdf.save(`${base}_RELATORIO_OPERACIONAL.pdf`);
    } catch (e) {
      console.error(e);
      alert("Não foi possível gerar o PDF.");
    } finally {
      setIsExporting(false);
    }
  }

  const perMachineAll = useMemo(() => groupByMachine(rows), [rows]);

  const machineOptions = useMemo(
    () => uniqueSorted(perMachineAll.map((x) => x.maquina)),
    [perMachineAll]
  );

  const filteredRows = useMemo(() => {
    if (machineFilter === "__ALL__") return rows;
    return rows.filter((r) => String(r["Máquina"] ?? "").trim() === machineFilter);
  }, [rows, machineFilter]);

  const filteredMachines = useMemo(() => {
    if (machineFilter === "__ALL__") return perMachineAll;
    return perMachineAll.filter((m) => m.maquina === machineFilter);
  }, [perMachineAll, machineFilter]);

  const avgResumo = useMemo(() => {
    if (!filteredMachines.length) {
      return { trabalho: 0, transporte: 0, ocioso: 0 };
    }

    return {
      trabalho:
        filteredMachines.reduce((acc, m) => acc + m.trabAvg, 0) / filteredMachines.length,
      transporte:
        filteredMachines.reduce((acc, m) => acc + m.transpAvg, 0) / filteredMachines.length,
      ocioso:
        filteredMachines.reduce((acc, m) => acc + m.ociosoAvg, 0) / filteredMachines.length,
    };
  }, [filteredMachines]);

  const utilChart = useMemo(() => {
    return [
      { name: "Trabalho", value: avgResumo.trabalho, color: COLORS.TRAB },
      { name: "Transporte", value: avgResumo.transporte, color: COLORS.TRANSP },
      { name: "Ocioso", value: avgResumo.ocioso, color: COLORS.OCIOSO },
    ];
  }, [avgResumo]);

  const cancelVsIgnoradoChart = useMemo(() => {
    const base = filteredMachines.slice();
    if (machineFilter === "__ALL__") return base.slice(0, 14);
    return base;
  }, [filteredMachines, machineFilter]);

  const rankingBundle = useMemo(() => {
    return buildReasonRanking(filteredRows, 10);
  }, [filteredRows]);

  const ranking = rankingBundle.top;

  const rankingColors = useMemo(() => {
    if (!ranking.length) return [] as string[];

    const values = ranking.map((r) => (rankingMode === "PCT" ? r.pct01 : r.horas));
    const min = Math.min(...values);
    const max = Math.max(...values);
    const denom = max - min || 1;

    return ranking.map((r) => {
      const v = rankingMode === "PCT" ? r.pct01 : r.horas;
      return heatColor((v - min) / denom);
    });
  }, [ranking, rankingMode]);

  const rankingChartData = useMemo(() => {
    return ranking.map((r) => ({
      motivo: r.motivo,
      valor: rankingMode === "PCT" ? r.pct01 : r.horas,
    }));
  }, [ranking, rankingMode]);

  const top3ByMachine = useMemo(() => {
    return buildTop3ByMachine(filteredRows).slice(0, 4);
  }, [filteredRows]);

  const utilYAxisMax = useMemo(() => {
    const maxVal = Math.max(...utilChart.map((d) => num(d.value)), 0);
    return maxVal <= 0 ? 1 : maxVal * 1.1;
  }, [utilChart]);

  const cancelIgnoradoYAxisMax = useMemo(() => {
    if (!cancelVsIgnoradoChart.length) return 5;

    const maxVal = Math.max(
      ...cancelVsIgnoradoChart.flatMap((d) => [num(d.canceladoH), num(d.ignoradoH)]),
      0
    );

    return maxVal <= 0 ? 5 : maxVal + 5;
  }, [cancelVsIgnoradoChart]);

  const rankingXAxisMax = useMemo(() => {
    if (!rankingChartData.length) return rankingMode === "PCT" ? 1 : 5;

    const maxVal = Math.max(...rankingChartData.map((d) => num(d.valor)), 0);

    if (rankingMode === "PCT") return maxVal <= 0 ? 1 : maxVal * 1.1;
    return maxVal <= 0 ? 5 : maxVal + 5;
  }, [rankingChartData, rankingMode]);

  const operationSummary = useMemo(() => {
    const metaT = metaTrabalho / 100;
    const metaTransp = metaTransporte / 100;
    const metaO = metaOcioso / 100;

    const desvios = [
      {
        nome: "Trabalho abaixo da meta",
        valor: Math.max(0, metaT - avgResumo.trabalho),
      },
      {
        nome: "Transporte acima da meta",
        valor: Math.max(0, avgResumo.transporte - metaTransp),
      },
      {
        nome: "Ocioso acima da meta",
        valor: Math.max(0, avgResumo.ocioso - metaO),
      },
    ];

    const principal = [...desvios].sort((a, b) => b.valor - a.valor)[0];

    const dentro =
      avgResumo.trabalho >= metaT &&
      avgResumo.transporte <= metaTransp &&
      avgResumo.ocioso <= metaO;

    let equipamentoCritico = "-";
    let piorScore = -1;

    for (const m of filteredMachines) {
      const score =
        Math.max(0, metaT - m.trabAvg) +
        Math.max(0, m.transpAvg - metaTransp) +
        Math.max(0, m.ociosoAvg - metaO) +
        m.canceladoH / 100 +
        m.ignoradoH / 100;

      if (score > piorScore) {
        piorScore = score;
        equipamentoCritico = m.maquina;
      }
    }

    return {
      status: dentro ? "Operação dentro dos parâmetros" : "Operação fora dos parâmetros",
      dentro,
      principalDesvio: principal?.valor > 0 ? principal.nome : "Sem desvio relevante",
      equipamentoCritico,
      top3Motivos: ranking.slice(0, 3).map((r) => r.motivo),
    };
  }, [
    avgResumo,
    metaTrabalho,
    metaTransporte,
    metaOcioso,
    filteredMachines,
    ranking,
  ]);

  const show = rows.length > 0;

  return (
    <div
      style={{
        fontFamily: "system-ui, Arial",
        padding: 18,
        maxWidth: 1200,
        margin: "0 auto",
        background: "#f6f2ec",
      }}
    >
      <div
        style={{
          background: "#fff",
          border: "1px solid #e5e7eb",
          borderRadius: 14,
          padding: 16,
          marginBottom: 14,
        }}
      >
        <h2 style={{ margin: 0 }}>Relatório Operacional — Máquinas</h2>
        <p style={{ marginTop: 6, opacity: 0.75 }}>
          Upload do Excel, metas operacionais e geração de PDF.
        </p>

        <div
          style={{
            display: "flex",
            gap: 12,
            alignItems: "center",
            flexWrap: "wrap",
            marginTop: 12,
          }}
        >
          <input
            type="file"
            accept=".xlsx"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) void handleFile(f);
            }}
          />

          {fileName ? (
            <span>
              Arquivo: <b>{fileName}</b>
            </span>
          ) : null}

          {machineOptions.length ? (
            <>
              <span>Filtro de máquina:</span>
              <select
                value={machineFilter}
                onChange={(e) => setMachineFilter(e.target.value)}
                style={{ padding: "6px 10px", borderRadius: 8 }}
              >
                <option value="__ALL__">Todas</option>
                {machineOptions.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
            </>
          ) : null}

          <span>Ranking em:</span>
          <select
            value={rankingMode}
            onChange={(e) => setRankingMode(e.target.value as RankingMode)}
            style={{ padding: "6px 10px", borderRadius: 8 }}
          >
            <option value="TOTAL_H">Total (h)</option>
            <option value="PCT">Porcentagem (%)</option>
          </select>

          <button
            onClick={exportPDFSingleA4}
            disabled={!show || isExporting}
            style={{
              padding: "9px 14px",
              borderRadius: 10,
              border: "1px solid #111827",
              background: !show || isExporting ? "#e5e7eb" : "#111827",
              color: !show || isExporting ? "#6b7280" : "#fff",
              cursor: !show || isExporting ? "not-allowed" : "pointer",
              fontWeight: 800,
            }}
          >
            {isExporting ? "Gerando PDF..." : "Exportar PDF"}
          </button>
        </div>

        <div
          style={{
            marginTop: 14,
            display: "grid",
            gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))",
            gap: 10,
          }}
        >
          <label>
            Meta Tempo em Trabalho (%)
            <input
              type="number"
              value={metaTrabalho}
              onChange={(e) => setMetaTrabalho(Number(e.target.value))}
              style={{ width: "100%", padding: 8, borderRadius: 8, border: "1px solid #d1d5db" }}
            />
          </label>

          <label>
            Meta Tempo em Transporte (%)
            <input
              type="number"
              value={metaTransporte}
              onChange={(e) => setMetaTransporte(Number(e.target.value))}
              style={{ width: "100%", padding: 8, borderRadius: 8, border: "1px solid #d1d5db" }}
            />
          </label>

          <label>
            Meta Tempo Ocioso (%)
            <input
              type="number"
              value={metaOcioso}
              onChange={(e) => setMetaOcioso(Number(e.target.value))}
              style={{ width: "100%", padding: 8, borderRadius: 8, border: "1px solid #d1d5db" }}
            />
          </label>
        </div>
      </div>

      {error && (
        <div style={{ background: "#fee2e2", padding: 12, borderRadius: 10, marginBottom: 12 }}>
          <b>Erro:</b> {error}
        </div>
      )}

      {!show && !error && (
        <div style={{ background: "#fff", padding: 14, borderRadius: 12 }}>
          Faça upload do Excel para gerar o relatório.
        </div>
      )}

      {show && (
        <div
          ref={pdfRef}
          style={{
            background: "#fff",
            border: "1px solid #e5e7eb",
            borderRadius: 16,
            padding: 14,
            width: "min(920px, 100%)",
            margin: "0 auto",
          }}
        >
          <div style={{ marginBottom: 12 }}>
            <h2 style={{ margin: 0 }}>
              Relatório de Gráficos — {machineFilter === "__ALL__" ? "Todas as máquinas" : machineFilter}
            </h2>
            <div style={{ fontSize: 12, opacity: 0.72 }}>
              Fonte: {fileName} • {new Date().toLocaleString()}
            </div>
          </div>

          <div
            style={{
              border: operationSummary.dentro ? "1px solid #16a34a" : "1px solid #ef4444",
              background: operationSummary.dentro ? "#f0fdf4" : "#fff1f2",
              borderRadius: 14,
              padding: 14,
              marginBottom: 12,
            }}
          >
            <div style={{ fontSize: 12, letterSpacing: 1, textTransform: "uppercase" }}>
              Resumo da Operação
            </div>

            <h2
              style={{
                margin: "6px 0 10px",
                color: operationSummary.dentro ? "#15803d" : "#dc2626",
              }}
            >
              {operationSummary.status}
            </h2>

            <ul style={{ margin: 0, paddingLeft: 20, lineHeight: 1.7 }}>
              <li>
                <b>Principal desvio:</b> {operationSummary.principalDesvio}
              </li>
              <li>
                <b>Equipamento crítico:</b> {operationSummary.equipamentoCritico}
              </li>
              <li>
                <b>Principais motivos de parada:</b>{" "}
                {operationSummary.top3Motivos.length
                  ? operationSummary.top3Motivos.join(", ")
                  : "-"}
              </li>
              <li>
                <b>Metas:</b> Trabalho ≥ {metaTrabalho}% | Transporte ≤ {metaTransporte}% | Ocioso ≤ {metaOcioso}%
              </li>
            </ul>
          </div>

          <div style={{ marginTop: 10 }}>
            <h3 style={{ margin: "0 0 6px" }}>Utilização (média)</h3>
            <div
              style={{
                width: "100%",
                height: 230,
                border: "1px solid #eee",
                borderRadius: 12,
                padding: 10,
                background: "#fff",
              }}
            >
              <ResponsiveContainer>
                <BarChart data={utilChart} margin={{ top: 28, right: 16, left: 8, bottom: 8 }}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="name" tick={TICK_STYLE} interval={0} />
                  <YAxis
                    tick={TICK_STYLE}
                    domain={[0, utilYAxisMax]}
                    tickFormatter={(v: any) => fmtPercent01(num(v))}
                  />
                  <Tooltip formatter={(v: any, n: any) => [fmtPercent01(num(v)), n]} />

                  <Bar dataKey="value" name="Utilização" isAnimationActive={false}>
                    {utilChart.map((entry, i) => (
                      <Cell key={`u-${i}`} fill={entry.color} />
                    ))}
                    <LabelList
                      dataKey="value"
                      position="top"
                      offset={6}
                      formatter={(v: any) => labelPercent(v)}
                      {...LABEL_STYLE}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div style={{ marginTop: 12 }}>
            <h3 style={{ margin: "0 0 6px" }}>Cancelado x Ignorado (h)</h3>
            <div
              style={{
                width: "100%",
                height: 240,
                border: "1px solid #eee",
                borderRadius: 12,
                padding: 10,
                background: "#fff",
              }}
            >
              <ResponsiveContainer>
                <BarChart
                  data={cancelVsIgnoradoChart}
                  margin={{ top: 28, right: 16, left: 8, bottom: 24 }}
                >
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis
                    dataKey="maquina"
                    interval={0}
                    angle={0}
                    height={30}
                    tick={TICK_STYLE}
                  />
                  <YAxis
                    tick={TICK_STYLE}
                    domain={[0, cancelIgnoradoYAxisMax]}
                    tickFormatter={(v: any) => fmtHours(num(v))}
                  />
                  <Tooltip formatter={(v: any, n: any) => [fmtHours(num(v)), n]} />
                  <Legend />

                  <Bar
                    dataKey="canceladoH"
                    name="Cancelado (h)"
                    fill={COLORS.CANCELADO}
                    isAnimationActive={false}
                  >
                    <LabelList
                      dataKey="canceladoH"
                      position="top"
                      offset={6}
                      formatter={(v: any) => labelHours(v)}
                      {...LABEL_STYLE}
                    />
                  </Bar>

                  <Bar
                    dataKey="ignoradoH"
                    name="Ignorado (h)"
                    fill={COLORS.IGNORADO}
                    isAnimationActive={false}
                  >
                    <LabelList
                      dataKey="ignoradoH"
                      position="top"
                      offset={6}
                      formatter={(v: any) => labelHours(v)}
                      {...LABEL_STYLE}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div style={{ marginTop: 12 }}>
            <h3 style={{ margin: "0 0 6px" }}>Ranking — Motivos de Parada (Top 10)</h3>
            <div
              style={{
                width: "100%",
                height: 300,
                border: "1px solid #eee",
                borderRadius: 12,
                padding: 10,
                background: "#fff",
              }}
            >
              <ResponsiveContainer>
                <BarChart
                  data={rankingChartData}
                  layout="vertical"
                  margin={{ top: 10, left: 20, right: 40, bottom: 10 }}
                >
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis
                    type="number"
                    tick={TICK_STYLE}
                    domain={[0, rankingXAxisMax]}
                    tickFormatter={(v: any) =>
                      rankingMode === "PCT" ? fmtPercent01(num(v)) : fmtHours(num(v))
                    }
                  />
                  <YAxis dataKey="motivo" type="category" width={260} tick={TICK_STYLE} />
                  <Tooltip
                    formatter={(v: any) =>
                      rankingMode === "PCT"
                        ? [fmtPercent01(num(v)), "Participação (%)"]
                        : [fmtHours(num(v)), "Tempo total (h)"]
                    }
                  />

                  <Bar
                    dataKey="valor"
                    name={rankingMode === "PCT" ? "Participação (%)" : "Tempo total (h)"}
                    isAnimationActive={false}
                  >
                    {rankingChartData.map((_, i) => (
                      <Cell key={`r-${i}`} fill={rankingColors[i]} />
                    ))}

                    <LabelList
                      dataKey="valor"
                      position="right"
                      formatter={(v: any) =>
                        rankingMode === "PCT" ? labelPercent(v) : labelHours(v)
                      }
                      {...LABEL_STYLE}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div
            style={{
              marginTop: 12,
              border: "1px solid #e5e7eb",
              borderRadius: 14,
              padding: 12,
              background: "#fff",
            }}
          >
            <h3 style={{ margin: "0 0 2px" }}>Top 3 Motivos por Equipamento</h3>
            <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 10 }}>
              Visão individual por equipamento para priorização operacional.
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
                gap: 10,
              }}
            >
              {top3ByMachine.map((eq) => {
                const maxValue = Math.max(
                  ...eq.top3.map((r) => (rankingMode === "PCT" ? r.pct01 : r.horas)),
                  0
                );

                return (
                  <div
                    key={eq.maquina}
                    style={{
                      border: "1px solid #e5e7eb",
                      borderRadius: 10,
                      padding: 10,
                      background: "#fff",
                    }}
                  >
                    <div
                      style={{
                        fontSize: 15,
                        fontWeight: 900,
                        marginBottom: 8,
                        color: "#111827",
                      }}
                    >
                      {eq.maquina}
                    </div>

                    {eq.top3.map((r, idx) => {
                      const value = rankingMode === "PCT" ? r.pct01 : r.horas;
                      const width = maxValue > 0 ? `${Math.max(4, (value / maxValue) * 100)}%` : "0%";

                      return (
                        <div key={r.motivo} style={{ marginBottom: 8 }}>
                          <div
                            style={{
                              display: "flex",
                              justifyContent: "space-between",
                              gap: 8,
                              fontSize: 12,
                              marginBottom: 4,
                            }}
                          >
                            <span>
                              {idx + 1}. {r.motivo}
                            </span>
                            <b>
                              {rankingMode === "PCT"
                                ? fmtPercent01(r.pct01)
                                : `${fmtHours(r.horas)} h`}
                            </b>
                          </div>

                          <div
                            style={{
                              height: 10,
                              background: "#f3f4f6",
                              overflow: "hidden",
                            }}
                          >
                            <div
                              style={{
                                height: "100%",
                                width,
                                background: TOP3_COLORS[idx] ?? "#9ca3af",
                              }}
                            />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
