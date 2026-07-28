"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  AlertTriangle, ArrowDownUp, ChevronDown, CircleDollarSign,
  FileSpreadsheet, Gem, RefreshCw, Search, Settings2, SlidersHorizontal,
  Store, Upload, X
} from "lucide-react";

type StoreName = string;
type StoreRecord = { id: string; name: string; active: boolean };
type Metal = "9ct" | "10ct" | "14ct" | "18ct" | "22ct" | "Sterling Silver" | "Platinum" | "Unknown";
type Item = {
  stockCode: string; description: string; ticketPrice: number; originalShelfDate: string;
  currentShelfDate: string; metal: Metal; weight: number | null; store: StoreName;
  reviewReason?: string;
};
type Prices = { gold: number; silver: number; updatedAt: string; source: "daily" | "manual" };
type Thresholds = { red: number; orange: number; yellow: number };
type SortKey = "ratio" | "ticket" | "melt" | "difference" | "days" | "stock";

const DEFAULT_STORES: StoreRecord[] = [
  { id: "palmerston-north", name: "Palmerston North", active: true },
  { id: "new-plymouth", name: "New Plymouth", active: true },
  { id: "wanganui", name: "Wanganui", active: true },
];
const DEFAULT_SELLING_FEE = 0.15;
const SAMPLE_ITEMS: Item[] = [
  { stockCode: "DEMO-1001", description: "9CT YG CHAIN TW 12.40GMS", ticketPrice: 649, originalShelfDate: "2025-03-14", currentShelfDate: "2026-05-02", metal: "9ct", weight: 12.4, store: "Palmerston North" },
  { stockCode: "DEMO-1002", description: "18CT WG DIAMOND RING TW 4.20GMS", ticketPrice: 899, originalShelfDate: "2025-10-08", currentShelfDate: "2026-04-19", metal: "18ct", weight: 4.2, store: "Palmerston North" },
  { stockCode: "DEMO-1003", description: "STERLING SILVER BRACELET TW 38.60GMS", ticketPrice: 139, originalShelfDate: "2026-01-21", currentShelfDate: "2026-01-21", metal: "Sterling Silver", weight: 38.6, store: "New Plymouth" },
  { stockCode: "DEMO-1004", description: "14CT YG BANGLE TW 10.80GMS", ticketPrice: 1199, originalShelfDate: "2024-11-03", currentShelfDate: "2026-02-11", metal: "14ct", weight: 10.8, store: "Wanganui" },
  { stockCode: "DEMO-1005", description: "10CT RG RING TW 2.45GMS", ticketPrice: 275, originalShelfDate: "2026-04-07", currentShelfDate: "2026-04-07", metal: "10ct", weight: 2.45, store: "New Plymouth" },
  { stockCode: "DEMO-1006", description: "22CT YG PENDANT TW 5.10GMS", ticketPrice: 875, originalShelfDate: "2025-07-17", currentShelfDate: "2026-03-04", metal: "22ct", weight: 5.1, store: "Wanganui" },
  { stockCode: "DEMO-1007", description: "PT RING TW 5.30GMS", ticketPrice: 1100, originalShelfDate: "2025-12-12", currentShelfDate: "2025-12-12", metal: "Platinum", weight: 5.3, store: "Palmerston North", reviewReason: "Platinum melt calculation not supported" },
];

const money = (n: number) => new Intl.NumberFormat("en-NZ", { style: "currency", currency: "NZD", maximumFractionDigits: 0 }).format(n || 0);
const todayIso = () => new Date().toISOString().slice(0, 10);
const daysOnSale = (iso: string) => iso ? Math.max(0, Math.floor((new Date(todayIso()).getTime() - new Date(iso).getTime()) / 86400000)) : 0;
const parseDate = (value: unknown): string => {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const d = XLSX.SSF.parse_date_code(value);
    return d ? `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}` : "";
  }
  const s = String(value).trim();
  const m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);
  if (m) {
    const y = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]);
    return `${y}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  }
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? "" : d.toISOString().slice(0, 10);
};
const parseCurrency = (v: unknown) => Number(String(v ?? "0").replace(/[^0-9.-]/g, "")) || 0;
const leadingTicket = (text: string) => {
  const match = text.trim().match(/^\$?\s*(\d{1,3}(?:,\d{3})+|\d+)(?:\.(\d{2}))?(?=\s|$)/);
  if (!match) return null;
  const value = Number(`${match[1].replace(/,/g, "")}.${match[2] ?? "00"}`);
  return Number.isFinite(value) && value > 0
    ? { value, description: text.slice(match[0].length).trim() }
    : null;
};
const repairSavedTicket = (item: Item): Item => {
  if (item.ticketPrice > 0) return item;
  const recovered = leadingTicket(item.description);
  return recovered ? { ...item, ticketPrice: recovered.value, description: recovered.description } : item;
};
const detectMetal = (text: string): Metal => {
  const t = text.toUpperCase();
  const k = t.match(/\b(9|10|14|18|22)\s*(?:CT|KT|K)\b/);
  if (k) return `${k[1]}ct` as Metal;
  if (/\b(SS|STERLING|925)\b/.test(t)) return "Sterling Silver";
  if (/\b(PT|PLATINUM)\b/.test(t)) return "Platinum";
  return "Unknown";
};
const detectWeight = (text: string) => {
  const m = text.toUpperCase().match(/(\d+(?:\.\d+)?)\s*GMS?\b/);
  return m ? Number(m[1]) : null;
};
const meltValue = (item: Item, prices: Prices) => {
  if (!item.weight) return 0;
  if (item.metal === "Sterling Silver") return prices.silver * 0.925 * 0.97 * item.weight * 1.15;
  const k = Number(item.metal.replace("ct", ""));
  return k ? (prices.gold * 0.97 / 24) * k * item.weight * 1.15 : 0;
};
const netSaleReturn = (ticketPrice: number, sellingFee: number) =>
  Math.max(0, ticketPrice * (1 - sellingFee));
const priority = (ratio: number, t: Thresholds) => ratio > t.red ? "red" : ratio >= t.orange ? "orange" : ratio >= t.yellow ? "yellow" : "green";

function parseReport(buffer: ArrayBuffer, store: StoreName): Item[] {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });
  const out: Item[] = [];
  for (const sheetName of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[sheetName], { header: 1, defval: null, raw: true });
    let active = false;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i] ?? [];
      const joined = row.filter(Boolean).join(" ").toUpperCase().replace(/\s+/g, " ");
      if (joined.includes("NOT FOUND DURING STOCKTAKE")) {
        active = joined.includes("ON SALE ACCORDING TO SYSTEM") && !joined.includes("ON SALE (ON HOLD");
        continue;
      }
      if (!active) continue;
      const stockCode = String(row[3] ?? "").trim();
      const department = String(row[5] ?? "").trim().toUpperCase();
      if (!stockCode || department !== "JEWELLERY") continue;
      const descriptionParts: string[] = [];
      const ticketCandidates: number[] = [];
      let j = i + 1;
      for (; j < rows.length; j++) {
        const next = rows[j] ?? [];
        const nextText = next.filter(Boolean).join(" ").toUpperCase();
        if (String(next[3] ?? "").trim() || nextText.includes("NOT FOUND DURING STOCKTAKE") || nextText.includes("CASH CONVERTERS")) break;
        if (next[0]) descriptionParts.push(String(next[0]).trim());
        // Ticket prices are stored in a separate cell on a description row
        // beneath the stock code (normally column I in the stocktake report).
        for (let cellIndex = 1; cellIndex < next.length; cellIndex++) {
          const cell = next[cellIndex];
          if (typeof cell === "string" && cell.includes("$")) {
            const amount = parseCurrency(cell);
            if (amount > 0) ticketCandidates.push(amount);
          } else if (cellIndex === 8 && typeof cell === "number" && cell > 0) {
            ticketCandidates.push(cell);
          }
        }
        if (j - i > 5) break;
      }
      let description = descriptionParts.join(" ").replace(/\s+/g, " ").trim();
      if (/\bWATCH(?:ES)?\b/i.test(description)) { i = Math.max(i, j - 1); continue; }
      const recoveredTicket = leadingTicket(description);
      const ticketPrice = parseCurrency(row[8]) || ticketCandidates[0] || recoveredTicket?.value || 0;
      if (recoveredTicket) description = recoveredTicket.description;
      const metal = detectMetal(description);
      const weight = detectWeight(description);
      const reasons: string[] = [];
      if (metal === "Platinum") reasons.push("Platinum melt calculation not supported");
      if (metal === "Unknown") reasons.push("Metal type not recognised");
      if (!weight) reasons.push("Weight in GMS not found");
      out.push({
        stockCode, description: description || "Description not found", ticketPrice,
        originalShelfDate: parseDate(row[9]), currentShelfDate: parseDate(row[12]), metal, weight, store,
        reviewReason: reasons.join("; ") || undefined
      });
      i = Math.max(i, j - 1);
    }
  }
  return [...new Map(out.map(x => [x.stockCode, x])).values()];
}

export default function Home() {
  const fileRef = useRef<HTMLInputElement>(null);
  const [items, setItems] = useState<Item[]>([]);
  const [stores, setStores] = useState<StoreRecord[]>(DEFAULT_STORES);
  const [ready, setReady] = useState(false);
  const [storeFilter, setStoreFilter] = useState<"All stores" | StoreName>("All stores");
  const [uploadStore, setUploadStore] = useState<StoreName>("Palmerston North");
  const [query, setQuery] = useState("");
  const [metalFilter, setMetalFilter] = useState<"All metals" | Metal>("All metals");
  const [priorityFilter, setPriorityFilter] = useState("All priorities");
  const [sortKey, setSortKey] = useState<SortKey>("ratio");
  const [showSettings, setShowSettings] = useState(false);
  const [showStores, setShowStores] = useState(false);
  const [showUpload, setShowUpload] = useState(false);
  const [showReview, setShowReview] = useState(false);
  const [notice, setNotice] = useState("");
  const [prices, setPrices] = useState<Prices>({ gold: 185, silver: 2.25, updatedAt: todayIso(), source: "daily" });
  const [dailyPrices, setDailyPrices] = useState<Prices>({ gold: 185, silver: 2.25, updatedAt: todayIso(), source: "daily" });
  const [thresholds, setThresholds] = useState<Thresholds>({ red: 1, orange: 0.75, yellow: 0.5 });
  const [sellingFee, setSellingFee] = useState(DEFAULT_SELLING_FEE);
  const [newStoreName, setNewStoreName] = useState("");
  const activeStores = useMemo(() => stores.filter(store => store.active), [stores]);

  useEffect(() => {
    try {
      const savedItems = localStorage.getItem("ccv-jewellery-items");
      const savedPrices = localStorage.getItem("ccv-jewellery-prices");
      const savedThresholds = localStorage.getItem("ccv-jewellery-thresholds");
      const savedSellingFee = localStorage.getItem("ccv-jewellery-selling-fee");
      const savedStores = localStorage.getItem("ccv-jewellery-stores");
      setItems(savedItems ? (JSON.parse(savedItems) as Item[]).map(repairSavedTicket) : SAMPLE_ITEMS);
      if (savedStores) setStores(JSON.parse(savedStores));
      if (savedPrices) { const p = JSON.parse(savedPrices); setPrices(p); setDailyPrices({ ...p, source: "daily" }); }
      if (savedThresholds) setThresholds(JSON.parse(savedThresholds));
      if (savedSellingFee !== null) setSellingFee(Number(savedSellingFee));
    } finally { setReady(true); }
  }, []);
  useEffect(() => {
    if (!ready || prices.source === "manual") return;
    const cachedToday = prices.updatedAt === todayIso();
    if (cachedToday && localStorage.getItem("ccv-daily-price-checked") === todayIso()) return;
    fetch("/api/metal-prices").then(r => r.ok ? r.json() : Promise.reject()).then((p: Prices) => {
      setDailyPrices(p); setPrices(p); localStorage.setItem("ccv-daily-price-checked", todayIso());
    }).catch(() => {
      setNotice("Daily metal prices could not refresh. The last saved prices are still active.");
      window.setTimeout(() => setNotice(""), 6000);
    });
  }, [ready, prices.source, prices.updatedAt]);
  useEffect(() => { if (ready) localStorage.setItem("ccv-jewellery-items", JSON.stringify(items)); }, [items, ready]);
  useEffect(() => { if (ready) localStorage.setItem("ccv-jewellery-prices", JSON.stringify(prices)); }, [prices, ready]);
  useEffect(() => { if (ready) localStorage.setItem("ccv-jewellery-thresholds", JSON.stringify(thresholds)); }, [thresholds, ready]);
  useEffect(() => { if (ready) localStorage.setItem("ccv-jewellery-selling-fee", String(sellingFee)); }, [sellingFee, ready]);
  useEffect(() => { if (ready) localStorage.setItem("ccv-jewellery-stores", JSON.stringify(stores)); }, [stores, ready]);

  const storeItems = useMemo(() => {
    const activeNames = new Set(activeStores.map(store => store.name));
    return items.filter(i => storeFilter === "All stores" ? activeNames.has(i.store) : i.store === storeFilter);
  }, [items, storeFilter, activeStores]);
  const reviewItems = storeItems.filter(i => i.reviewReason);
  const validItems = storeItems.filter(i => !i.reviewReason);
  const totals = useMemo(() => ({
    count: storeItems.length,
    ticket: storeItems.reduce((s, i) => s + i.ticketPrice, 0),
    net: storeItems.reduce((s, i) => s + netSaleReturn(i.ticketPrice, sellingFee), 0),
    melt: validItems.reduce((s, i) => s + meltValue(i, prices), 0),
    opportunities: validItems.filter(i => meltValue(i, prices) > netSaleReturn(i.ticketPrice, sellingFee)).length
  }), [storeItems, validItems, prices, sellingFee]);
  const metals = useMemo(() => (["9ct", "10ct", "14ct", "18ct", "22ct", "Sterling Silver"] as Metal[]).map(m => ({
    label: m, count: validItems.filter(i => i.metal === m).length,
    weight: validItems.filter(i => i.metal === m).reduce((s, i) => s + (i.weight || 0), 0)
  })), [validItems]);
  const rows = useMemo(() => validItems.filter(i => {
    const m = meltValue(i, prices); const net = netSaleReturn(i.ticketPrice, sellingFee); const r = net ? m / net : 0;
    return (!query || `${i.stockCode} ${i.description}`.toLowerCase().includes(query.toLowerCase()))
      && (metalFilter === "All metals" || i.metal === metalFilter)
      && (priorityFilter === "All priorities" || priority(r, thresholds) === priorityFilter);
  }).sort((a, b) => {
    const am = meltValue(a, prices), bm = meltValue(b, prices);
    if (sortKey === "ticket") return b.ticketPrice - a.ticketPrice;
    if (sortKey === "melt") return bm - am;
    if (sortKey === "difference") {
      const aDifference = am - netSaleReturn(a.ticketPrice, sellingFee);
      const bDifference = bm - netSaleReturn(b.ticketPrice, sellingFee);
      return bDifference - aDifference;
    }
    if (sortKey === "days") return daysOnSale(b.originalShelfDate) - daysOnSale(a.originalShelfDate);
    if (sortKey === "stock") return a.stockCode.localeCompare(b.stockCode);
    return (bm / (netSaleReturn(b.ticketPrice, sellingFee) || 1)) - (am / (netSaleReturn(a.ticketPrice, sellingFee) || 1));
  }), [validItems, query, metalFilter, priorityFilter, sortKey, prices, thresholds, sellingFee]);
  const filteredTicketTotal = useMemo(
    () => rows.reduce((sum, item) => sum + item.ticketPrice, 0),
    [rows]
  );

  async function handleUpload(file?: File) {
    if (!file) return;
    try {
      const imported = parseReport(await file.arrayBuffer(), uploadStore);
      if (!imported.length) throw new Error("No jewellery was found under the required ON SALE section.");
      setItems(prev => [...prev.filter(i => i.store !== uploadStore && !i.stockCode.startsWith("DEMO-")), ...imported]);
      setStoreFilter(uploadStore);
      setShowUpload(false);
      setNotice(`${imported.length} jewellery items loaded for ${uploadStore}. Existing store stock was replaced.`);
      window.setTimeout(() => setNotice(""), 7000);
    } catch (e) { setNotice(e instanceof Error ? e.message : "The report could not be read."); }
    if (fileRef.current) fileRef.current.value = "";
  }

  function resetDemo() {
    setItems(SAMPLE_ITEMS); setStores(DEFAULT_STORES); setStoreFilter("All stores"); setUploadStore("Palmerston North"); setNotice("Demo data restored.");
  }

  function addStore() {
    const name = newStoreName.trim().replace(/\s+/g, " ");
    if (!name) return;
    if (stores.some(store => store.name.toLowerCase() === name.toLowerCase())) {
      setNotice("A store with that name already exists.");
      return;
    }
    setStores(current => [...current, { id: `${Date.now()}-${name.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`, name, active: true }]);
    setUploadStore(name); setNewStoreName(""); setNotice(`${name} was added.`);
  }

  function renameStore(id: string, nextName: string) {
    const store = stores.find(candidate => candidate.id === id);
    const name = nextName.trim().replace(/\s+/g, " ");
    if (!store || !name || name === store.name) return;
    if (stores.some(candidate => candidate.id !== id && candidate.name.toLowerCase() === name.toLowerCase())) {
      setNotice("A store with that name already exists.");
      return;
    }
    setStores(current => current.map(candidate => candidate.id === id ? { ...candidate, name } : candidate));
    setItems(current => current.map(item => item.store === store.name ? { ...item, store: name } : item));
    if (storeFilter === store.name) setStoreFilter(name);
    if (uploadStore === store.name) setUploadStore(name);
    setNotice(`${store.name} was renamed to ${name}.`);
  }

  function toggleStore(id: string) {
    const store = stores.find(candidate => candidate.id === id);
    if (!store) return;
    setStores(current => current.map(candidate => candidate.id === id ? { ...candidate, active: !candidate.active } : candidate));
    if (store.active && storeFilter === store.name) setStoreFilter("All stores");
    if (store.active && uploadStore === store.name) {
      const replacement = activeStores.find(candidate => candidate.id !== id);
      if (replacement) setUploadStore(replacement.name);
    }
    setNotice(`${store.name} was ${store.active ? "deactivated" : "reactivated"}.`);
  }

  function deleteStore(id: string) {
    const store = stores.find(candidate => candidate.id === id);
    if (!store || !window.confirm(`Permanently delete ${store.name} and all of its saved stock? This cannot be undone.`)) return;
    setStores(current => current.filter(candidate => candidate.id !== id));
    setItems(current => current.filter(item => item.store !== store.name));
    if (storeFilter === store.name) setStoreFilter("All stores");
    const replacement = activeStores.find(candidate => candidate.id !== id);
    if (uploadStore === store.name && replacement) setUploadStore(replacement.name);
    setNotice(`${store.name} and its stock were permanently deleted.`);
  }

  return (
    <main>
      <header className="topbar">
        <div className="brand"><span className="brandMark">CCV</span><div><b>Jewellery Melt</b><small>Decision Dashboard</small></div></div>
        <div className="headerActions">
          <button className="iconButton" onClick={() => setShowStores(true)} aria-label="Manage stores"><Store size={20}/></button>
          <button className="iconButton" onClick={() => setShowSettings(true)} aria-label="Open settings"><Settings2 size={20}/></button>
          <button className="uploadButton" onClick={() => setShowUpload(true)}><Upload size={18}/> Upload report</button>
        </div>
      </header>

      <section className="workspace">
        {notice && <div className="notice">{notice}<button onClick={() => setNotice("")}><X size={16}/></button></div>}
        <div className="pageHead">
          <div><p className="eyebrow">JEWELLERY PERFORMANCE</p><h1>What is worth more melted?</h1><p className="subhead">Compare refinery values with estimated net sale returns after selling fees.</p></div>
          <label className="storeSelect"><Store size={18}/><select value={storeFilter} onChange={e => setStoreFilter(e.target.value as typeof storeFilter)}><option>All stores</option>{activeStores.map(s => <option key={s.id}>{s.name}</option>)}</select><ChevronDown size={16}/></label>
        </div>

        <div className="priceStrip">
          <div className="priceStatus"><span className="liveDot"/><div><small>{prices.source === "manual" ? "SCENARIO PRICES ACTIVE" : "DAILY METAL PRICES"}</small><b>Gold {money(prices.gold)} <em>/g</em> · Silver ${prices.silver.toFixed(2)} <em>/g</em></b></div></div>
          <div className="priceMeta">Updated {prices.updatedAt} <button onClick={() => setShowSettings(true)}><SlidersHorizontal size={15}/> Adjust scenario</button></div>
        </div>

        <section className="metrics">
          <article><span className="metricIcon burgundy"><Gem size={21}/></span><div><small>ITEMS ON SALE</small><strong>{totals.count}</strong><p>{storeFilter}</p></div></article>
          <article><span className="metricIcon gold"><CircleDollarSign size={21}/></span><div><small>TOTAL TICKET PRICE</small><strong>{money(filteredTicketTotal)}</strong><p>{rows.length} items in current search</p></div></article>
          <article><span className="metricIcon dark"><RefreshCw size={21}/></span><div><small>EST. MELT VALUE</small><strong>{money(totals.melt)}</strong><p>{totals.ticket ? Math.round(totals.melt / totals.ticket * 100) : 0}% of ticket value</p></div></article>
          <article className="dangerCard"><span className="metricIcon red"><AlertTriangle size={21}/></span><div><small>MELT OPPORTUNITIES</small><strong>{totals.opportunities}</strong><p>Melt exceeds net sale return</p></div></article>
        </section>

        <section className="caratGrid">
          {metals.map((m, idx) => <article key={m.label}>
            <div className={`caratToken c${idx}`}>{m.label === "Sterling Silver" ? "925" : m.label.replace("ct","")}</div>
            <div><b>{m.label}</b><span>{m.count} items</span></div><strong>{m.weight.toFixed(1)}<small> g</small></strong>
          </article>)}
        </section>

        <section className="tableCard">
          <div className="tableTitle"><div><h2>Jewellery stock</h2><span>{rows.length} items shown</span></div>
            <button className="reviewButton" onClick={() => setShowReview(true)}><AlertTriangle size={17}/>{reviewItems.length} Needs review</button>
          </div>
          <div className="filters">
            <label className="search"><Search size={18}/><input value={query} onChange={e => setQuery(e.target.value)} placeholder="Search stock code or description…"/></label>
            <select value={metalFilter} onChange={e => setMetalFilter(e.target.value as typeof metalFilter)}><option>All metals</option>{metals.map(m => <option key={m.label}>{m.label}</option>)}</select>
            <select value={priorityFilter} onChange={e => setPriorityFilter(e.target.value)}><option>All priorities</option><option value="red">Red — melt exceeds net return</option><option value="orange">Orange — 75% to 100%</option><option value="yellow">Yellow — 50% to 75%</option><option value="green">Green — below 50%</option></select>
            <label className="sort"><ArrowDownUp size={16}/><select aria-label="Sort jewellery items" value={sortKey} onChange={e => setSortKey(e.target.value as SortKey)}><option value="ratio">Melt risk % — highest to lowest</option><option value="days">On sale days — highest to lowest</option><option value="ticket">Ticket — highest to lowest</option><option value="melt">Melt value — highest to lowest</option><option value="difference">Difference — highest to lowest</option><option value="stock">Stock code — A to Z</option></select></label>
          </div>
          <div className="tableWrap"><table><thead><tr><th>Melt risk</th><th>Stock item</th><th>Store</th><th>Metal / weight</th><th>On sale</th><th>Ticket</th><th>Net sale return</th><th>Melt value</th><th>Difference</th></tr></thead>
            <tbody>{rows.map(item => {
              const melt = meltValue(item, prices), net = netSaleReturn(item.ticketPrice, sellingFee), difference = melt - net, ratio = net ? melt / net : 0, level = priority(ratio, thresholds);
              return <tr key={`${item.store}-${item.stockCode}`}>
                <td><span className={`priority ${level}`}>{Math.round(ratio * 100)}%</span></td>
                <td><b>{item.stockCode}</b><small>{item.description}</small></td><td>{item.store}</td>
                <td><b>{item.metal}</b><small>{item.weight?.toFixed(2)} g</small></td>
                <td><b>{daysOnSale(item.originalShelfDate)} days</b><small>Repriced {item.currentShelfDate || "—"}</small></td>
                <td className="numeric">{money(item.ticketPrice)}</td><td className="numeric"><b>{money(net)}</b><small>after selling fee</small></td><td className="numeric"><b>{money(melt)}</b></td>
                <td className={`numeric diff ${difference > 0 ? "negative" : "positive"}`}><b>{difference > 0 ? "+" : "−"}{money(Math.abs(difference))}</b><small>{difference > 0 ? "melt advantage" : "net sale advantage"}</small></td>
              </tr>;
            })}</tbody></table>{!rows.length && <div className="empty">No items match the selected filters.</div>}</div>
        </section>
        <p className="dataNote">Version 7 stores uploaded stock, stores and settings in this browser. Use “Reset demo” in Settings to restore the example dashboard.</p>
      </section>

      {showUpload && <div className="modalBackdrop"><section className="modal">
        <button className="modalClose" onClick={() => setShowUpload(false)}><X/></button><span className="modalIcon"><FileSpreadsheet/></span>
        <p className="eyebrow">STORE STOCK UPDATE</p><h2>Upload jewellery report</h2><p>Select the store, then upload its Excel report. This replaces that store’s current list; new codes are added and missing codes are removed.</p>
        <label className="field"><span>Store</span><select value={uploadStore} onChange={e => setUploadStore(e.target.value as StoreName)}>{activeStores.map(s => <option key={s.id}>{s.name}</option>)}</select></label>
        <button className="dropzone" onClick={() => fileRef.current?.click()}><Upload size={26}/><b>Choose Excel report</b><span>.xls or .xlsx</span></button>
        <input ref={fileRef} hidden type="file" accept=".xls,.xlsx" onChange={e => handleUpload(e.target.files?.[0])}/>
        <div className="importRules"><b>Version 7 import rules</b><span>✓ Reads “NOT FOUND DURING STOCKTAKE, ON SALE ACCORDING TO SYSTEM”</span><span>✓ Ignores watches completely</span><span>✓ Flags missing carat, weight and platinum</span></div>
      </section></div>}

      {showStores && <div className="modalBackdrop"><section className="modal storesModal">
        <button className="modalClose" onClick={() => setShowStores(false)}><X/></button><span className="modalIcon"><Store/></span>
        <p className="eyebrow">STORE MANAGEMENT</p><h2>Manage stores</h2>
        <p>Active stores are included in combined totals and available for report uploads. Deactivate a store to keep its stock without including it.</p>
        <div className="addStore">
          <label className="field"><span>New store name</span><input value={newStoreName} onChange={e => setNewStoreName(e.target.value)} onKeyDown={e => { if (e.key === "Enter") addStore(); }} placeholder="Enter franchise store name"/></label>
          <button className="primary" onClick={addStore}>Add store</button>
        </div>
        <div className="storeList">{stores.map(store => <article key={store.id} className={!store.active ? "inactive" : ""}>
          <div className="storeIdentity"><span className={`statusDot ${store.active ? "active" : ""}`}/><div><b>{store.name}</b><small>{store.active ? "Active · included in combined totals" : "Inactive · stock retained but excluded"}</small></div></div>
          <div className="storeActions">
            <button className="secondary" onClick={() => { const name = window.prompt("Rename store", store.name); if (name !== null) renameStore(store.id, name); }}>Rename</button>
            <button className="secondary" onClick={() => toggleStore(store.id)}>{store.active ? "Deactivate" : "Reactivate"}</button>
            <button className="deleteButton" onClick={() => deleteStore(store.id)}>Delete</button>
          </div>
        </article>)}</div>
      </section></div>}

      {showSettings && <div className="modalBackdrop"><section className="modal settingsModal">
        <button className="modalClose" onClick={() => setShowSettings(false)}><X/></button><span className="modalIcon"><Settings2/></span><p className="eyebrow">CALCULATION SETTINGS</p><h2>Melt price scenario</h2>
        <p>Enter NZD spot prices per gram to test what happens when metal prices change.</p>
        <div className="twoCols">
          <label className="field"><span>Gold spot price (NZD/g)</span><input type="number" step="0.01" value={prices.gold} onChange={e => setPrices(p => ({...p, gold: Number(e.target.value), source: "manual", updatedAt: todayIso()}))}/></label>
          <label className="field"><span>Silver spot price (NZD/g)</span><input type="number" step="0.01" value={prices.silver} onChange={e => setPrices(p => ({...p, silver: Number(e.target.value), source: "manual", updatedAt: todayIso()}))}/></label>
        </div>
        <h3>Sale deductions</h3>
        <label className="field"><span>Selling fee (editable)</span><input type="number" min="0" max="100" step="0.1" value={sellingFee * 100} onChange={e => setSellingFee(Math.min(1, Math.max(0, Number(e.target.value) / 100)))}/><i>%</i></label>
        <h3>Priority thresholds</h3><div className="threeCols">
          <label className="field"><span>Red above</span><input type="number" value={thresholds.red * 100} onChange={e => setThresholds(t => ({...t, red: Number(e.target.value)/100}))}/><i>%</i></label>
          <label className="field"><span>Orange from</span><input type="number" value={thresholds.orange * 100} onChange={e => setThresholds(t => ({...t, orange: Number(e.target.value)/100}))}/><i>%</i></label>
          <label className="field"><span>Yellow from</span><input type="number" value={thresholds.yellow * 100} onChange={e => setThresholds(t => ({...t, yellow: Number(e.target.value)/100}))}/><i>%</i></label>
        </div>
        <div className="formula"><b>Net sale return:</b> ticket × (1 − selling fee)<br/><b>Melt risk:</b> melt value ÷ net sale return × 100<br/><b>Gold:</b> ((97% × spot ÷ 24) × carat × weight) × 1.15<br/><b>Silver:</b> spot × 92.5% × 97% × weight × 1.15</div>
        <div className="modalActions"><button className="secondary" onClick={() => { resetDemo(); setSellingFee(DEFAULT_SELLING_FEE); }}>Reset demo</button><button className="secondary" onClick={() => setPrices(dailyPrices)}>Use daily prices</button><button className="primary" onClick={() => setShowSettings(false)}>Apply settings</button></div>
      </section></div>}

      {showReview && <div className="modalBackdrop"><section className="modal reviewModal">
        <button className="modalClose" onClick={() => setShowReview(false)}><X/></button><span className="modalIcon warning"><AlertTriangle/></span><p className="eyebrow">MANUAL CHECKS</p><h2>Needs review</h2><p>These jewellery items were imported but cannot receive a reliable melt value yet.</p>
        <div className="reviewList">{reviewItems.length ? reviewItems.map(i => <article key={`${i.store}-${i.stockCode}`}><div><b>{i.stockCode}</b><span>{i.description}</span><small>{i.store}</small></div><strong>{i.reviewReason}</strong></article>) : <div className="empty">No items need review.</div>}</div>
      </section></div>}
    </main>
  );
}
