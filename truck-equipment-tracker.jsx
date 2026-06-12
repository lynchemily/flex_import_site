import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

export default function TruckEquipmentTracker() {
  const [entries, setEntries] = useState([]);
  const [loading, setLoading] = useState(true);
  const [name, setName] = useState("");
  const [serial, setSerial] = useState("");
  const [truck, setTruck] = useState("");
  const [photo, setPhoto] = useState(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [filterTruck, setFilterTruck] = useState("All");
  const [catalogue, setCatalogue] = useState([]);
  const [regCategory, setRegCategory] = useState("");
  const [importMsg, setImportMsg] = useState("");
  const fileInputRef = useRef(null);
  const importInputRef = useRef(null);

  // ---- load saved entries + catalogue on start ----
  useEffect(() => {
    (async () => {
      try {
        const result = await window.storage.get("equipment-entries");
        if (result && result.value) setEntries(JSON.parse(result.value));
      } catch (e) {}
      try {
        const cat = await window.storage.get("asset-catalogue");
        if (cat && cat.value) setCatalogue(JSON.parse(cat.value));
      } catch (e) {}
      setLoading(false);
    })();
  }, []);

  const persist = async (list) => {
    try {
      await window.storage.set("equipment-entries", JSON.stringify(list));
    } catch (e) {
      setError("Couldn't save to storage. The list may be too large — try deleting old entries.");
    }
  };

  // ---- photo handling: resize down so storage stays small ----
  const handlePhoto = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const img = new Image();
      img.onload = () => {
        const MAX = 640;
        let { width, height } = img;
        if (width > height && width > MAX) {
          height = Math.round((height * MAX) / width);
          width = MAX;
        } else if (height > MAX) {
          width = Math.round((width * MAX) / height);
          height = MAX;
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        setPhoto(canvas.toDataURL("image/jpeg", 0.65));
      };
      img.src = ev.target.result;
    };
    reader.readAsDataURL(file);
  };

  // ---- add entry ----
  const addEntry = async () => {
    setError("");
    if (!name.trim() || !serial.trim() || !truck.trim()) {
      setError("Name, serial number and truck are all required.");
      return;
    }
    setSaving(true);
    const entry = {
      id: Date.now().toString(36),
      name: name.trim(),
      serial: serial.trim(),
      truck: truck.trim(),
      photo,
      added: new Date().toISOString(),
    };
    const next = [entry, ...entries];
    setEntries(next);
    await persist(next);
    setName("");
    setSerial("");
    setTruck("");
    setPhoto(null);
    if (fileInputRef.current) fileInputRef.current.value = "";
    setSaving(false);
  };

  const deleteEntry = async (id) => {
    const next = entries.filter((e) => e.id !== id);
    setEntries(next);
    await persist(next);
  };

  // ---- export to Excel ----
  const exportToExcel = () => {
    const list = filterTruck === "All" ? entries : entries.filter((e) => e.truck === filterTruck);
    const rows = list.map((e) => ({
      "Equipment name": e.name,
      "Serial number": e.serial,
      "Assigned truck": e.truck,
      "Photo on file": e.photo ? "Yes" : "No",
      "Date added": new Date(e.added).toLocaleDateString("en-GB"),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 28 }, { wch: 18 }, { wch: 16 }, { wch: 12 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Equipment");
    const fileName =
      filterTruck === "All"
        ? "equipment-log.xlsx"
        : `equipment-log-${filterTruck.replace(/\s+/g, "-")}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  // ---- import asset register (.xlsx) as the reference catalogue ----
  const importRegister = (file) => {
    if (!file) return;
    setImportMsg("");
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const items = rows
          .filter((r) => String(r["Name"] || "").trim())
          .map((r, i) => ({
            id: "cat-" + i,
            name: String(r["Name"]).trim(),
            serial:
              String(r["Asset: Number"] || "").trim() ||
              String(r["Serial Number"] || "").trim(),
            category: String(r["Asset: Category"] || "").trim(),
            status: String(r["Status"] || "").trim(),
          }));
        if (items.length === 0) {
          setImportMsg("No assets found — the file needs 'Name' and 'Asset: Number' columns.");
          return;
        }
        setCatalogue(items);
        await window.storage.set("asset-catalogue", JSON.stringify(items));
        setImportMsg(`Imported ${items.length} assets from the register.`);
      } catch (err) {
        setImportMsg("Couldn't read that file. Make sure it's the asset register .xlsx.");
      }
      if (importInputRef.current) importInputRef.current.value = "";
    };
    reader.readAsArrayBuffer(file);
  };

  // dropdown data from the register
  const regCategories = [...new Set(catalogue.map((a) => a.category).filter(Boolean))].sort();
  const regNames = catalogue.filter((a) => !regCategory || a.category === regCategory);

  const pickAsset = (assetId) => {
    const a = catalogue.find((x) => x.id === assetId);
    if (!a) return;
    setName(a.name);
    setSerial(a.serial || "No tag/serial on register");
  };

  const trucks = [...new Set(entries.map((e) => e.truck))].sort();
  const visible = filterTruck === "All" ? entries : entries.filter((e) => e.truck === filterTruck);

  // shared input style: text-base (16px) stops iPhones auto-zooming on focus,
  // py-3 gives a proper thumb-sized tap target
  const field =
    "w-full border border-zinc-300 rounded-lg px-3 py-3 text-base bg-white focus:outline-none focus:ring-2 focus:ring-amber-400";

  return (
    <div className="min-h-screen bg-zinc-100 text-zinc-900 pb-24">
      {/* header */}
      <header className="bg-zinc-900 text-white px-4 py-4 border-b-4 border-amber-400 sticky top-0 z-20">
        <div className="max-w-3xl mx-auto">
          <h1 className="text-xl font-extrabold tracking-tight uppercase">Equipment Log</h1>
        </div>
      </header>

      <main className="max-w-3xl mx-auto px-3 sm:px-4 py-4 space-y-6">
        {/* ---- add form ---- */}
        <section className="bg-white rounded-xl shadow-sm border border-zinc-200 p-4 sm:p-5">
          <h2 className="font-bold uppercase text-sm tracking-wider text-zinc-500 mb-3">
            Add equipment
          </h2>

          {/* ---- asset register dropdowns ---- */}
          <div className="bg-zinc-50 border border-zinc-200 rounded-lg p-3 mb-4">
            <div className="flex items-center justify-between flex-wrap gap-2">
              <span className="text-xs font-bold uppercase tracking-wider text-zinc-500">
                Asset register{" "}
                {catalogue.length > 0 && (
                  <span className="text-zinc-400 font-normal normal-case">
                    — {catalogue.length} loaded
                  </span>
                )}
              </span>
              <label className="text-sm text-zinc-700 underline cursor-pointer hover:text-zinc-900 py-1">
                {catalogue.length > 0 ? "Re-import" : "Import register (.xlsx)"}
                <input
                  ref={importInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  className="hidden"
                  onChange={(e) => importRegister(e.target.files[0])}
                />
              </label>
            </div>
            {importMsg && <p className="text-xs text-zinc-600 mt-2">{importMsg}</p>}
            {catalogue.length > 0 && (
              <div className="mt-3 space-y-3">
                <div>
                  <label className="block text-xs font-medium text-zinc-600 mb-1">Category</label>
                  <select
                    value={regCategory}
                    onChange={(e) => setRegCategory(e.target.value)}
                    className={field}
                  >
                    <option value="">All categories</option>
                    {regCategories.map((c) => (
                      <option key={c} value={c}>{c}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-zinc-600 mb-1">
                    Asset name (from register)
                  </label>
                  <select value="" onChange={(e) => pickAsset(e.target.value)} className={field}>
                    <option value="" disabled>
                      Select an asset… ({regNames.length})
                    </option>
                    {regNames.map((a) => (
                      <option key={a.id} value={a.id}>
                        {a.name}{a.serial ? ` — ${a.serial}` : ""}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            )}
          </div>

          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium mb-1">Equipment name</label>
              <input
                className={field}
                placeholder="e.g. Holmatro spreader"
                value={name}
                onChange={(e) => setName(e.target.value)}
              />
            </div>

            <div>
              <label className="block text-sm font-medium mb-1">Serial number</label>
              <input
                className={field + " font-mono"}
                placeholder="e.g. SN-04421-A"
                value={serial}
                onChange={(e) => setSerial(e.target.value)}
              />
            </div>

            <div>
              <label className="block text-sm font-medium mb-1">Assigned truck</label>
              <input
                list="truck-options"
                className={field}
                placeholder="e.g. LK11A1"
                value={truck}
                onChange={(e) => setTruck(e.target.value)}
              />
              <datalist id="truck-options">
                {trucks.map((t) => (
                  <option key={t} value={t} />
                ))}
              </datalist>
            </div>

            <div>
              <label className="block text-sm font-medium mb-1">Photo</label>
              <input
                ref={fileInputRef}
                type="file"
                accept="image/*"
                capture="environment"
                onChange={(e) => handlePhoto(e.target.files[0])}
                className="hidden"
                id="photo-input"
              />
              <div className="flex items-center gap-3">
                <label
                  htmlFor="photo-input"
                  className="flex-1 text-center bg-zinc-900 text-white font-medium px-4 py-3 rounded-lg cursor-pointer active:bg-zinc-700"
                >
                  📷 {photo ? "Retake photo" : "Take / choose photo"}
                </label>
                {photo && (
                  <img
                    src={photo}
                    alt="Preview"
                    className="w-14 h-14 object-cover rounded-lg border border-zinc-300"
                  />
                )}
              </div>
            </div>

            {error && (
              <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2">
                {error}
              </p>
            )}

            <button
              onClick={addEntry}
              disabled={saving}
              className="w-full bg-amber-400 active:bg-amber-300 disabled:opacity-50 text-zinc-900 font-bold uppercase tracking-wide px-6 py-4 rounded-lg text-base"
            >
              {saving ? "Saving…" : "Add to log"}
            </button>
          </div>
        </section>

        {/* ---- list ---- */}
        <section>
          <h2 className="font-bold uppercase text-sm tracking-wider text-zinc-500 mb-2">
            Logged equipment ({visible.length})
          </h2>
          <div className="flex items-center gap-2 mb-3">
            {trucks.length > 0 && (
              <select
                value={filterTruck}
                onChange={(e) => setFilterTruck(e.target.value)}
                className="flex-1 border border-zinc-300 rounded-lg px-2 py-2.5 text-base bg-white"
              >
                <option>All</option>
                {trucks.map((t) => (
                  <option key={t}>{t}</option>
                ))}
              </select>
            )}
            {entries.length > 0 && (
              <button
                onClick={exportToExcel}
                className="flex-1 bg-zinc-900 active:bg-zinc-700 text-white text-base font-medium px-3 py-2.5 rounded-lg"
              >
                Export to Excel
              </button>
            )}
          </div>

          {loading ? (
            <p className="text-zinc-500 text-sm">Loading saved entries…</p>
          ) : visible.length === 0 ? (
            <div className="bg-white border border-dashed border-zinc-300 rounded-xl p-8 text-center text-zinc-500 text-sm">
              No equipment logged yet. Add your first item above.
            </div>
          ) : (
            <ul className="space-y-2">
              {visible.map((e) => (
                <li
                  key={e.id}
                  className="bg-white rounded-xl border border-zinc-200 shadow-sm p-3 flex gap-3 items-start"
                >
                  {e.photo ? (
                    <img
                      src={e.photo}
                      alt={e.name}
                      className="w-16 h-16 object-cover rounded-lg border border-zinc-200 shrink-0"
                    />
                  ) : (
                    <div className="w-16 h-16 rounded-lg bg-zinc-100 border border-zinc-200 shrink-0 flex items-center justify-center text-zinc-400 text-[10px] text-center px-1">
                      No photo
                    </div>
                  )}
                  <div className="min-w-0 flex-1">
                    <div className="font-semibold text-sm leading-snug">{e.name}</div>
                    <div className="font-mono text-xs text-zinc-600 mt-0.5 break-all">
                      {e.serial}
                    </div>
                    <span className="inline-block mt-1.5 bg-zinc-900 text-amber-400 text-[11px] font-bold uppercase tracking-wide px-2 py-0.5 rounded">
                      {e.truck}
                    </span>
                  </div>
                  <button
                    onClick={() => deleteEntry(e.id)}
                    className="text-zinc-400 active:text-red-600 shrink-0 p-2 -m-1 text-base"
                    title="Delete entry"
                  >
                    ✕
                  </button>
                </li>
              ))}
            </ul>
          )}
        </section>
      </main>
    </div>
  );
}
