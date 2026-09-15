"use strict";

const EMBEDDED_DATA = {
  source: "가상기업 3개 5개년매출데이터.xlsx",
  headers: ["기준월", "한빛테크 매출(억원)", "미래솔루션 매출(억원)", "새봄산업 매출(억원)"],
  rows: [["2021-01-01",8.59,7.1,10.15],["2021-02-01",9.08,7.79,10.32],["2021-03-01",10.22,7.44,9.75],["2021-04-01",9.83,7.35,9.32],["2021-05-01",9.77,7.02,8.94],["2021-06-01",9.78,6.66,8.55],["2021-07-01",9.36,6.22,8.56],["2021-08-01",8.64,6.23,8.25],["2021-09-01",8.35,6.28,8.52],["2021-10-01",7.69,6.82,9.34],["2021-11-01",9.47,8.07,10.89],["2021-12-01",10.36,8.41,11.25],["2022-01-01",10.07,8.11,10.39],["2022-02-01",10.79,8.45,9.58],["2022-03-01",11.19,8.22,9.25],["2022-04-01",11.54,7.92,8.93],["2022-05-01",11.84,7.74,8.54],["2022-06-01",10.98,7.27,8.58],["2022-07-01",10.66,7.11,7.95],["2022-08-01",9.99,6.68,8.2],["2022-09-01",9.66,6.99,8.64],["2022-10-01",9.31,7.41,9.24],["2022-11-01",10.43,8.3,10.45],["2022-12-01",11.31,9.02,11.33],["2023-01-01",11.66,8.6,9.98],["2023-02-01",12.27,9.34,9.67],["2023-03-01",12.95,8.91,9.39],["2023-04-01",13.35,8.96,9.16],["2023-05-01",13.41,8.32,8.76],["2023-06-01",12.98,7.79,7.8],["2023-07-01",11.78,7.74,8.15],["2023-08-01",11.25,7.2,8.01],["2023-09-01",11.15,7.54,8.71],["2023-10-01",10.89,7.54,8.98],["2023-11-01",12.48,9.4,10.16],["2023-12-01",13.45,9.73,10.62],["2024-01-01",13.01,9.84,9.91],["2024-02-01",14,9.83,9.23],["2024-03-01",15.73,10.18,8.98],["2024-04-01",15.88,9.76,8.43],["2024-05-01",15.92,9.28,8.4],["2024-06-01",15.01,8.41,7.79],["2024-07-01",13.49,8.62,8.01],["2024-08-01",13.41,7.99,7.61],["2024-09-01",12.75,8.47,7.93],["2024-10-01",12.2,8.24,8.79],["2024-11-01",14.5,9.89,10.09],["2024-12-01",15.4,10.71,10.74],["2025-01-01",14.98,10.08,9.42],["2025-02-01",16.77,10.43,9.11],["2025-03-01",18.27,10.88,8.86],["2025-04-01",18.14,10.2,8.28],["2025-05-01",17.76,10.13,7.85],["2025-06-01",16.89,9.19,7.79],["2025-07-01",15.84,9.35,7.81],["2025-08-01",14.55,8.63,7.81],["2025-09-01",13.95,8.63,8.28],["2025-10-01",14.17,9.25,8.6],["2025-11-01",16.78,10.8,9.55],["2025-12-01",17.61,11.8,10.15]]
};

const COLORS = ["#2358d8", "#16a3b6", "#f07b63", "#7b61c9", "#2f8c62", "#d59b32"];
const $ = (selector) => document.querySelector(selector);
const el = (tag, className, html = "") => { const node = document.createElement(tag); if (className) node.className = className; node.innerHTML = html; return node; };
const fmt = (value, digits = 1) => new Intl.NumberFormat("ko-KR", { maximumFractionDigits: digits, minimumFractionDigits: digits }).format(value || 0);
const pct = (value) => `${value >= 0 ? "+" : ""}${fmt(value, 1)}%`;
const sum = (values) => values.reduce((a, b) => a + (Number.isFinite(b) ? b : 0), 0);
const mean = (values) => values.length ? sum(values) / values.length : 0;

let state = { dataset: normalizeDataset(EMBEDDED_DATA), year: "all", selected: [] };

function normalizeDataset(raw) {
  if (!raw.headers || raw.headers.length < 2 || !raw.rows?.length) throw new Error("날짜 열과 하나 이상의 숫자 열이 필요합니다.");
  const companies = raw.headers.slice(1).map((h, i) => String(h || `기업 ${i + 1}`).replace(/\s*매출\s*\([^)]*\)\s*| /gi, "").replace(/\s*연매출\s*\([^)]*\)\s*/gi, "").trim());
  const rows = raw.rows.map(row => {
    const date = parseDate(row[0]);
    return { date, values: row.slice(1, companies.length + 1).map(v => v === "" || v == null ? null : Number(v)) };
  }).filter(row => row.date instanceof Date && !Number.isNaN(row.date) && row.values.some(Number.isFinite)).sort((a, b) => a.date - b.date);
  if (!rows.length) throw new Error("분석할 수 있는 월간 데이터를 찾지 못했습니다.");
  return { source: raw.source || "사용자 파일", companies, rows };
}

function parseDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === "number" && value > 20000) return new Date(Date.UTC(1899, 11, 30 + value));
  const text = String(value ?? "").trim().replace(/[./]/g, "-");
  const match = text.match(/^(\d{4})-(\d{1,2})/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, 1);
  const d = new Date(text);
  return Number.isNaN(d.getTime()) ? null : d;
}

function initialize(resetSelection = true) {
  const { companies, rows, source } = state.dataset;
  if (resetSelection) state.selected = companies.map((_, i) => i);
  const years = [...new Set(rows.map(r => r.date.getFullYear()))];
  $("#yearSelect").innerHTML = `<option value="all">전체 기간</option>${years.map(y => `<option value="${y}">${y}년</option>`).join("")}`;
  state.year = "all";
  $("#companyFilters").innerHTML = companies.map((name, i) => `<label class="company-chip" style="--chip-color:${COLORS[i % COLORS.length]}"><input type="checkbox" value="${i}" checked><span>${escapeHtml(name)}</span></label>`).join("");
  const missing = rows.reduce((n, r) => n + r.values.filter(v => !Number.isFinite(v)).length, 0);
  $("#sourceName").textContent = source;
  $("#sourceMeta").textContent = `${rows.length}개월 · ${companies.length}개 기업 · 결측치 ${missing}건`;
  bindFilterInputs();
  render();
}

function bindFilterInputs() {
  document.querySelectorAll("#companyFilters input").forEach(input => input.addEventListener("change", () => {
    const checked = [...document.querySelectorAll("#companyFilters input:checked")].map(i => Number(i.value));
    if (!checked.length) { input.checked = true; showToast("하나 이상의 기업을 선택해 주세요."); return; }
    state.selected = checked;
    render();
  }));
}

function filteredRows() {
  return state.year === "all" ? state.dataset.rows : state.dataset.rows.filter(r => r.date.getFullYear() === Number(state.year));
}

function totalsByCompany(rows) {
  return state.dataset.companies.map((_, i) => sum(rows.map(r => r.values[i]).filter(Number.isFinite)));
}

function render() {
  const rows = filteredRows();
  renderKpis(rows);
  renderTrend(rows);
  renderMix(rows);
  renderAnnual();
  renderSeason(rows);
  renderInsights(rows);
  renderTable();
}

function renderKpis(rows) {
  const totals = totalsByCompany(rows);
  const selectedTotals = state.selected.map(i => ({ i, value: totals[i] }));
  const total = sum(selectedTotals.map(x => x.value));
  const avg = total / Math.max(rows.length, 1);
  const leader = selectedTotals.sort((a, b) => b.value - a.value)[0];
  const prevRows = previousComparableRows(rows);
  const prevTotal = sum(state.selected.map(i => totalsByCompany(prevRows)[i]));
  const latestYear = rows[rows.length - 1].date.getFullYear();
  const compareRows = state.year === "all" ? rows.filter(r => r.date.getFullYear() === latestYear) : rows;
  const compareTotal = sum(state.selected.map(i => totalsByCompany(compareRows)[i]));
  const growth = prevTotal ? (compareTotal / prevTotal - 1) * 100 : null;
  const growthLabel = state.year === "all"
    ? `${rows[rows.length - 1].date.getFullYear()}년 전년 대비`
    : "전년 동일 기간 대비";
  const monthly = rows.map(r => ({ date: r.date, value: sum(state.selected.map(i => r.values[i]).filter(Number.isFinite)) }));
  const peak = monthly.sort((a, b) => b.value - a.value)[0];
  const cards = [
    { label: "총 매출", value: `${fmt(total, 1)}<small>억원</small>`, foot: growth == null ? `선택 기간 ${rows.length}개월` : `<span class="delta ${growth < 0 ? "negative" : ""}">${pct(growth)}</span> ${growthLabel}`, icon: "Σ" },
    { label: "월평균 매출", value: `${fmt(avg, 1)}<small>억원</small>`, foot: `${state.selected.length}개 기업 합산 기준`, icon: "ø" },
    { label: "매출 1위", value: `${escapeHtml(state.dataset.companies[leader.i])}`, foot: `선택 기간 ${fmt(leader.value, 1)}억원`, icon: "1" },
    { label: "최고 매출월", value: `${peak.date.getFullYear()}.${String(peak.date.getMonth()+1).padStart(2,"0")}`, foot: `${fmt(peak.value, 1)}억원 · 선택 기업 합산`, icon: "↑" }
  ];
  $("#kpiGrid").innerHTML = cards.map(c => `<article class="kpi-card"><div class="kpi-top"><span class="kpi-label">${c.label}</span><span class="kpi-icon">${c.icon}</span></div><div class="kpi-value">${c.value}</div><div class="kpi-foot">${c.foot}</div></article>`).join("");
}

function previousComparableRows(rows) {
  if (!rows.length) return [];
  if (state.year === "all") {
    const latestYear = rows[rows.length - 1].date.getFullYear();
    return state.dataset.rows.filter(r => r.date.getFullYear() === latestYear - 1);
  }
  const start = rows[0].date;
  const end = rows[rows.length - 1].date;
  const startKey = (start.getFullYear() - 1) * 12 + start.getMonth();
  const endKey = (end.getFullYear() - 1) * 12 + end.getMonth();
  return state.dataset.rows.filter(r => { const key = r.date.getFullYear() * 12 + r.date.getMonth(); return key >= startKey && key <= endKey; });
}

function renderTrend(rows) {
  const W = 760, H = 288, pad = { l: 50, r: 16, t: 18, b: 35 };
  const values = rows.flatMap(r => state.selected.map(i => r.values[i])).filter(Number.isFinite);
  const max = Math.ceil(Math.max(...values, 1) / 5) * 5;
  const min = Math.max(0, Math.floor(Math.min(...values, 0) / 5) * 5);
  const x = i => pad.l + i * (W - pad.l - pad.r) / Math.max(rows.length - 1, 1);
  const y = v => pad.t + (max - v) * (H - pad.t - pad.b) / Math.max(max - min, 1);
  const grid = Array.from({ length: 5 }, (_, i) => {
    const gy = pad.t + i * (H - pad.t - pad.b) / 4;
    const label = max - i * (max - min) / 4;
    return `<line class="grid-line" x1="${pad.l}" y1="${gy}" x2="${W-pad.r}" y2="${gy}"/><text class="axis-label" x="${pad.l-9}" y="${gy+3}" text-anchor="end">${fmt(label,0)}</text>`;
  }).join("");
  const labelEvery = Math.max(1, Math.ceil(rows.length / 7));
  const labels = rows.map((r, i) => i % labelEvery === 0 || i === rows.length - 1 ? `<text class="axis-label" x="${x(i)}" y="${H-8}" text-anchor="middle">${state.year === "all" ? `${String(r.date.getFullYear()).slice(2)}.${String(r.date.getMonth()+1).padStart(2,"0")}` : `${r.date.getMonth()+1}월`}</text>` : "").join("");
  const series = state.selected.map(idx => {
    const points = rows.map((r, i) => Number.isFinite(r.values[idx]) ? [x(i), y(r.values[idx])] : null).filter(Boolean);
    const path = points.map((p, i) => `${i ? "L" : "M"}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join(" ");
    return `<path class="trend-path" d="${path}" stroke="${COLORS[idx % COLORS.length]}"/>`;
  }).join("");
  const targets = rows.map((r, i) => `<rect class="hover-target" data-index="${i}" x="${x(i)-(W-pad.l-pad.r)/Math.max(rows.length-1,1)/2}" y="${pad.t}" width="${Math.max(9,(W-pad.l-pad.r)/Math.max(rows.length-1,1))}" height="${H-pad.t-pad.b}"/>`).join("");
  $("#trendChart").innerHTML = `<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" role="img">${grid}${labels}${series}${targets}<g id="hoverDots"></g></svg>`;
  $("#trendLegend").innerHTML = state.selected.map(i => `<span><i style="background:${COLORS[i % COLORS.length]}"></i>${escapeHtml(state.dataset.companies[i])}</span>`).join("");
  document.querySelectorAll(".hover-target").forEach(target => {
    target.addEventListener("mousemove", e => showTrendTooltip(e, rows[Number(target.dataset.index)], x(Number(target.dataset.index)), y));
    target.addEventListener("mouseleave", hideTooltip);
  });
}

function showTrendTooltip(event, row, px, yFn) {
  const dateLabel = `${row.date.getFullYear()}년 ${row.date.getMonth()+1}월`;
  const lines = state.selected.map(i => `<span style="color:${COLORS[i % COLORS.length]}">●</span> ${escapeHtml(state.dataset.companies[i])} <b>${fmt(row.values[i], 2)}</b>`).join("<br>");
  const tip = $("#tooltip");
  tip.innerHTML = `<strong>${dateLabel}</strong>${lines}`;
  tip.style.display = "block";
  tip.style.left = `${Math.min(event.clientX + 14, window.innerWidth - 155)}px`;
  tip.style.top = `${Math.max(8, event.clientY - 60)}px`;
  $("#hoverDots").innerHTML = state.selected.filter(i => Number.isFinite(row.values[i])).map(i => `<circle class="hover-dot" cx="${px}" cy="${yFn(row.values[i])}" r="4" fill="${COLORS[i % COLORS.length]}"/>`).join("");
}
function hideTooltip() { $("#tooltip").style.display = "none"; const dots = $("#hoverDots"); if (dots) dots.innerHTML = ""; }

function renderMix(rows) {
  const totals = totalsByCompany(rows);
  const chosen = state.selected.map(i => ({ i, value: totals[i] }));
  const total = sum(chosen.map(x => x.value));
  let at = 0;
  const stops = chosen.map(x => { const start = at; at += total ? x.value / total * 100 : 0; return `${COLORS[x.i % COLORS.length]} ${start}% ${at}%`; });
  $("#mixDonut").style.background = `conic-gradient(${stops.join(",")})`;
  $("#mixTotal").textContent = fmt(total, 0);
  $("#mixList").innerHTML = chosen.sort((a,b) => b.value-a.value).map(x => `<div class="mix-item"><i style="background:${COLORS[x.i % COLORS.length]}"></i><span>${escapeHtml(state.dataset.companies[x.i])}</span><strong>${fmt(total ? x.value/total*100 : 0,1)}%</strong><small>${fmt(x.value,1)}억원</small></div>`).join("");
}

function annualData() {
  const years = [...new Set(state.dataset.rows.map(r => r.date.getFullYear()))];
  return years.map(year => ({ year, totals: totalsByCompany(state.dataset.rows.filter(r => r.date.getFullYear() === year)) }));
}

function renderAnnual() {
  const data = annualData();
  const max = Math.max(...data.flatMap(d => state.selected.map(i => d.totals[i])), 1);
  $("#annualChart").innerHTML = data.map(d => `<div class="year-group"><div class="bar-set">${state.selected.map(i => `<div class="annual-bar" title="${escapeHtml(state.dataset.companies[i])}: ${fmt(d.totals[i],1)}억원" data-value="${fmt(d.totals[i],0)}" style="height:${d.totals[i]/max*100}%;background:${COLORS[i % COLORS.length]}"></div>`).join("")}</div><span class="year-label">${d.year}</span></div>`).join("");
}

function renderSeason(rows) {
  const months = Array.from({ length: 12 }, (_, month) => {
    const values = rows.filter(r => r.date.getMonth() === month).flatMap(r => state.selected.map(i => r.values[i])).filter(Number.isFinite);
    return mean(values);
  });
  const max = Math.max(...months, 1), min = Math.min(...months.filter(v => v > 0));
  const maxIndex = months.indexOf(max), minIndex = months.indexOf(min);
  $("#seasonChart").innerHTML = months.map((value, i) => `<div class="month-bar-wrap"><div class="month-bar ${i === maxIndex ? "hot" : ""} ${i === minIndex ? "low" : ""}" data-value="${fmt(value,1)}" title="${i+1}월 평균 ${fmt(value,2)}억원" style="height:${value/max*82}%"></div><span class="month-label">${i+1}월</span></div>`).join("");
}

function renderInsights(rows) {
  const annual = annualData();
  const first = annual[0], last = annual[annual.length - 1];
  const span = Math.max(annual.length - 1, 1);
  const cagrs = state.selected.map(i => ({ i, v: first.totals[i] ? (Math.pow(last.totals[i] / first.totals[i], 1/span) - 1) * 100 : 0 })).sort((a,b)=>b.v-a.v);
  const totals = totalsByCompany(rows); const selectedTotals = state.selected.map(i => ({i,v:totals[i]})).sort((a,b)=>b.v-a.v);
  const monthlyAvg = Array.from({length:12},(_,m)=>mean(rows.filter(r=>r.date.getMonth()===m).flatMap(r=>state.selected.map(i=>r.values[i])).filter(Number.isFinite)));
  const peakM = monthlyAvg.indexOf(Math.max(...monthlyAvg)); const lowM = monthlyAvg.indexOf(Math.min(...monthlyAvg));
  const volatility = state.selected.map(i => { const vals=rows.map(r=>r.values[i]).filter(Number.isFinite); const av=mean(vals); const sd=Math.sqrt(mean(vals.map(v=>(v-av)**2))); return {i,v:av?sd/av*100:0}; }).sort((a,b)=>a.v-b.v)[0];
  const latestGrowth = last && annual.length > 1 ? state.selected.map(i => ({i,v:annual[annual.length-2].totals[i] ? (last.totals[i]/annual[annual.length-2].totals[i]-1)*100 : 0})).sort((a,b)=>b.v-a.v)[0] : null;
  const items = [
    { title: `${state.dataset.companies[cagrs[0].i]} 장기 성장 우위`, body: `${first.year}~${last.year}년 연평균 성장률은 ${pct(cagrs[0].v)}로 선택 기업 중 가장 높습니다.` },
    { title: `${peakM+1}월이 가장 강한 달`, body: `월 평균 ${fmt(monthlyAvg[peakM],1)}억원으로, 저점인 ${lowM+1}월보다 ${fmt(monthlyAvg[lowM] ? (monthlyAvg[peakM]/monthlyAvg[lowM]-1)*100 : 0,1)}% 높습니다.` },
    { title: `${state.dataset.companies[selectedTotals[0].i]} 비중 ${fmt(sum(selectedTotals.map(x=>x.v)) ? selectedTotals[0].v/sum(selectedTotals.map(x=>x.v))*100 : 0,1)}%`, body: `선택한 기간의 총매출을 기준으로 가장 큰 기업입니다.` },
    { title: `${state.dataset.companies[volatility.i]} 변동성 최저`, body: `매출 변동계수 ${fmt(volatility.v,1)}%로 선택 기업 중 가장 안정적입니다.${latestGrowth ? ` ${last.year}년은 ${state.dataset.companies[latestGrowth.i]}의 성장률이 ${pct(latestGrowth.v)}로 가장 높습니다.` : ""}` }
  ];
  $("#insights").innerHTML = items.map((x,i)=>`<div class="insight"><span class="insight-index">0${i+1}</span><strong>${escapeHtml(x.title)}</strong><p>${escapeHtml(x.body)}</p></div>`).join("");
}

function renderTable() {
  const data = annualData(); const names = state.selected.map(i=>state.dataset.companies[i]);
  $("#summaryHead").innerHTML = `<tr><th>연도</th>${names.map(n=>`<th>${escapeHtml(n)}</th>`).join("")}<th>합계</th><th>전년비</th></tr>`;
  $("#summaryBody").innerHTML = data.map((d, rowIndex) => {
    const values=state.selected.map(i=>d.totals[i]); const total=sum(values); const prev=rowIndex?sum(state.selected.map(i=>data[rowIndex-1].totals[i])):null; const g=prev?(total/prev-1)*100:null;
    return `<tr><td>${d.year}</td>${values.map(v=>`<td>${fmt(v,1)}</td>`).join("")}<td><strong>${fmt(total,1)}</strong></td><td class="${g == null ? "" : g >= 0 ? "up" : "down"}">${g == null ? "—" : pct(g)}</td></tr>`;
  }).join("");
}

async function readFile(file) {
  if (!file) return;
  try {
    const ext = file.name.split(".").pop().toLowerCase();
    let raw;
    if (ext === "csv") raw = parseCsv(await file.text(), file.name);
    else if (ext === "xlsx") raw = await parseXlsx(await file.arrayBuffer(), file.name);
    else throw new Error(".xlsx 또는 .csv 파일만 지원합니다.");
    state.dataset = normalizeDataset(raw);
    initialize(true);
    showToast(`${file.name}을 로컬에서 분석했습니다.`);
  } catch (error) { console.error(error); showToast(error.message || "파일을 읽지 못했습니다.", 4200); }
  $("#fileInput").value = "";
}

function parseCsv(text, source) {
  const rows=[]; let row=[], cell="", quoted=false;
  for(let i=0;i<text.length;i++){ const c=text[i], n=text[i+1]; if(c==='"'&&quoted&&n==='"'){cell+='"';i++;} else if(c==='"'){quoted=!quoted;} else if(c===','&&!quoted){row.push(cell.trim());cell="";} else if((c==='\n'||c==='\r')&&!quoted){if(c==='\r'&&n==='\n')i++;row.push(cell.trim());if(row.some(Boolean))rows.push(row);row=[];cell="";} else cell+=c; }
  if(cell||row.length){row.push(cell.trim());rows.push(row);}
  return { source, headers: rows[0], rows: rows.slice(1) };
}

async function parseXlsx(buffer, source) {
  if (typeof JSZip === "undefined") throw new Error("로컬 엑셀 읽기 모듈을 불러오지 못했습니다.");
  const zip = await JSZip.loadAsync(buffer);
  const xml = async path => { const entry=zip.file(path); if(!entry) throw new Error(`엑셀 구조에서 ${path}을 찾지 못했습니다.`); return new DOMParser().parseFromString(await entry.async("text"), "application/xml"); };
  const workbook = await xml("xl/workbook.xml");
  const rels = await xml("xl/_rels/workbook.xml.rels");
  const relMap = Object.fromEntries([...rels.getElementsByTagNameNS("*", "Relationship")].map(r => [r.getAttribute("Id"), r.getAttribute("Target")]));
  const sharedFile = zip.file("xl/sharedStrings.xml");
  let shared=[];
  if(sharedFile){ const doc=new DOMParser().parseFromString(await sharedFile.async("text"),"application/xml"); shared=[...doc.getElementsByTagNameNS("*","si")].map(si=>[...si.getElementsByTagNameNS("*","t")].map(t=>t.textContent).join("")); }
  const sheets=[...workbook.getElementsByTagNameNS("*","sheet")];
  for(const sheet of sheets){
    const rid=sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships","id") || sheet.getAttribute("r:id");
    let target=relMap[rid]; if(!target) continue; target=target.replace(/^\//,""); if(!target.startsWith("xl/")) target=`xl/${target.replace(/^\.\//,"")}`;
    const doc=await xml(target); const parsed=[];
    for(const rowNode of doc.getElementsByTagNameNS("*","row")){
      const row=[];
      for(const c of rowNode.getElementsByTagNameNS("*","c")){
        const ref=c.getAttribute("r")||"A1"; const col=columnIndex(ref); const type=c.getAttribute("t"); const v=c.getElementsByTagNameNS("*","v")[0]?.textContent ?? "";
        let value=v;
        if(type==="s") value=shared[Number(v)] ?? "";
        else if(type==="inlineStr") value=[...c.getElementsByTagNameNS("*","t")].map(t=>t.textContent).join("");
        else if(type==="b") value=v==="1";
        else if(v!=="" && Number.isFinite(Number(v))) value=Number(v);
        row[col]=value;
      }
      parsed.push(row);
    }
    const compact=parsed.filter(r=>r.some(v=>v!==""&&v!=null));
    if(compact.length>1 && compact[0].length>1) return { source, headers:compact[0], rows:compact.slice(1) };
  }
  throw new Error("분석할 표 데이터를 찾지 못했습니다.");
}

function columnIndex(ref) { const letters=(ref.match(/[A-Z]+/i)||["A"])[0].toUpperCase(); return [...letters].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1; }
function escapeHtml(value) { return String(value ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c])); }
function showToast(message, duration=2600) { const toast=$("#toast"); toast.textContent=message; toast.classList.add("show"); clearTimeout(showToast.timer); showToast.timer=setTimeout(()=>toast.classList.remove("show"),duration); }

function downloadCsv() {
  const data=annualData(); const headers=["연도",...state.selected.map(i=>state.dataset.companies[i]),"합계"];
  const rows=data.map(d=>{const vals=state.selected.map(i=>d.totals[i]);return[d.year,...vals.map(v=>v.toFixed(2)),sum(vals).toFixed(2)];});
  const csv="\ufeff"+[headers,...rows].map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(",")).join("\r\n");
  const url=URL.createObjectURL(new Blob([csv],{type:"text/csv;charset=utf-8"})); const a=document.createElement("a"); a.href=url; a.download="연간_매출_요약.csv"; a.click(); setTimeout(()=>URL.revokeObjectURL(url),1000);
}

$("#yearSelect").addEventListener("change", e => { state.year=e.target.value; render(); });
$("#resetButton").addEventListener("click", () => { initialize(true); showToast("필터를 초기화했습니다."); });
$("#fileInput").addEventListener("change", e => readFile(e.target.files[0]));
$("#downloadCsv").addEventListener("click", downloadCsv);
const dropZone=$("#dropZone");
["dragenter","dragover"].forEach(name=>dropZone.addEventListener(name,e=>{e.preventDefault();dropZone.classList.add("dragover");}));
["dragleave","drop"].forEach(name=>dropZone.addEventListener(name,e=>{e.preventDefault();dropZone.classList.remove("dragover");}));
dropZone.addEventListener("drop",e=>readFile(e.dataTransfer.files[0]));
window.addEventListener("resize",()=>{clearTimeout(window.__resizeTimer);window.__resizeTimer=setTimeout(()=>renderTrend(filteredRows()),100);});

initialize();
