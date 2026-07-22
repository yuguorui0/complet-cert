import { useState, useRef } from "react";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import logo from "../logo.png";
import sigDefault from "../signatures/sig_default.png";
import { useOptions } from "../hooks/useOptions";
import SettingsModal from "../components/SettingsModal";

const SIGNATURES: Record<string, { label: string; src: string }> = {
  sig_default: { label: "默认签名", src: sigDefault },
};

type TimelineRow = { bjTime: string; localTime: string; action: string };
type TaskItem = string;

interface CertData {
  projectCode: string;
  projectName: string;
  requester: string;
  requesterContact: string;
  endUser: string;
  deliveryAddress: string;
  completionDate: string;
  serviceEngineer: string[];
  serviceLevel: "urgent" | "non-urgent" | "";
  serviceTypes: string[];
  taskItems: TaskItem[];
  hasTimeline: boolean;
  timeline: TimelineRow[];
  result: "completed" | "";
  notes: string;
  confirmedWithRemote: boolean;
  remoteContact: string;
  enriginSig: string;
}

const SERVICE_TYPES = [
  "Problem Handling",
  "Software Upgrade",
  "Circuit Migration",
  "Circuit Testing",
  "New Circuit Patching",
  "Integration Services",
  "Daily Maintenance",
  "Logistics",
  "Other",
];

const defaultData: CertData = {
  projectCode: "",
  projectName: "",
  requester: "",
  requesterContact: "",
  endUser: "",
  deliveryAddress: "",
  completionDate: new Date().toISOString().split("T")[0],
  serviceEngineer: [],
  serviceLevel: "",
  serviceTypes: [],
  taskItems: [],
  hasTimeline: false,
  timeline: [
    { bjTime: "", localTime: "", action: "" },
    { bjTime: "", localTime: "", action: "" },
  ],
  result: "",
  notes: "",
  confirmedWithRemote: false,
  remoteContact: "",
  enriginSig: "",
};

function today() {
  return new Date().toLocaleDateString("zh-CN");
}

export default function CertificatePage() {
  const [data, setData] = useState<CertData>(defaultData);
  const [mode, setMode] = useState<"form" | "preview">("form");
  const [showSettings, setShowSettings] = useState(false);
  const printRef = useRef<HTMLDivElement>(null);
  const {
    options,
    addCustomer, updateCustomer, removeCustomer,
    addListItem, updateListItem, removeListItem,
  } = useOptions();

  function set<K extends keyof CertData>(key: K, val: CertData[K]) {
    setData((d) => ({ ...d, [key]: val }));
  }

  function setTaskItem(i: number, val: string) {
    const items = [...data.taskItems];
    items[i] = val;
    setData((d) => ({ ...d, taskItems: items }));
  }

  function addTaskItem() {
    setData((d) => ({ ...d, taskItems: [...d.taskItems, ""] }));
  }

  function removeTaskItem(i: number) {
    setData((d) => ({ ...d, taskItems: d.taskItems.filter((_, idx) => idx !== i) }));
  }

  function setTimelineRow(i: number, field: keyof TimelineRow, val: string) {
    const rows = [...data.timeline];
    rows[i] = { ...rows[i], [field]: val };
    setData((d) => ({ ...d, timeline: rows }));
  }

  function addTimelineRow() {
    setData((d) => ({
      ...d,
      timeline: [...d.timeline, { bjTime: "", localTime: "", action: "" }],
    }));
  }

  function removeTimelineRow(i: number) {
    setData((d) => ({ ...d, timeline: d.timeline.filter((_, idx) => idx !== i) }));
  }

  function toggleEngineer(name: string) {
    setData((d) => ({
      ...d,
      serviceEngineer: d.serviceEngineer.includes(name)
        ? d.serviceEngineer.filter((x) => x !== name)
        : [...d.serviceEngineer, name],
    }));
  }

  function addCustomEngineer(name: string) {
    const trimmed = name.trim();
    if (!trimmed) return;
    setData((d) => ({
      ...d,
      serviceEngineer: d.serviceEngineer.includes(trimmed)
        ? d.serviceEngineer
        : [...d.serviceEngineer, trimmed],
    }));
  }

  function removeEngineer(name: string) {
    setData((d) => ({
      ...d,
      serviceEngineer: d.serviceEngineer.filter((x) => x !== name),
    }));
  }

  function toggleServiceType(t: string) {
    setData((d) => ({
      ...d,
      serviceTypes: d.serviceTypes.includes(t)
        ? d.serviceTypes.filter((x) => x !== t)
        : [...d.serviceTypes, t],
    }));
  }

  function handlePrint() {
    window.print();
  }

  async function handleDownloadPDF() {
    if (!printRef.current) return;
    const el = printRef.current;
    const canvas = await html2canvas(el, {
      scale: 2,
      useCORS: true,
      allowTaint: true,
      backgroundColor: "#ffffff",
      onclone: (clonedDoc) => {
        // html2canvas can't parse oklch() used by Tailwind v4.
        // The certificate uses only inline styles so stripping all
        // global stylesheets from the clone is safe and fixes the crash.
        clonedDoc.querySelectorAll("style, link[rel='stylesheet']").forEach((s) => s.remove());
      },
    });
    const imgData = canvas.toDataURL("image/jpeg", 0.98);
    const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
    const pdfW = pdf.internal.pageSize.getWidth();
    const pdfH = (canvas.height * pdfW) / canvas.width;
    pdf.addImage(imgData, "JPEG", 0, 0, pdfW, pdfH);
    const safeName = (data.projectName || data.projectCode || "certificate").replace(/[/\\?%*:|"<>]/g, "_");
    pdf.save(`完工确认书_${safeName}.pdf`);
  }

  async function handleDownloadWord() {
    if (!printRef.current) return;
    const el = printRef.current;

    // Inline all images as base64 for Word compatibility
    const cloned = el.cloneNode(true) as HTMLElement;
    const imgs = cloned.querySelectorAll("img");
    await Promise.all(Array.from(imgs).map(async (img) => {
      try {
        const res = await fetch(img.src);
        const blob = await res.blob();
        const b64 = await new Promise<string>((resolve) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result as string);
          reader.readAsDataURL(blob);
        });
        img.src = b64;
      } catch { /* keep original src */ }
    }));

    const htmlContent = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office'
            xmlns:w='urn:schemas-microsoft-com:office:word'
            xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
          <meta charset='utf-8'>
          <title>Completion Certificate</title>
          <!--[if gte mso 9]><xml>
            <w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom>
            <w:DoNotOptimizeForBrowser/></w:WordDocument>
          </xml><![endif]-->
          <style>
            @page { size: A4 portrait; margin: 12mm 14mm 10mm 14mm; }
            body { font-family: Arial, 'Microsoft YaHei', sans-serif; font-size: 8pt; color: #1a1a1a; }
            table { border-collapse: collapse; width: 100%; }
            td, th { border: 1px solid #c8d9e8; padding: 1.2mm 2.5mm; font-size: 7.5pt; }
            ol { padding-left: 20px; }
          </style>
        </head>
        <body>${cloned.innerHTML}</body>
      </html>`;

    const blob = new Blob(["\ufeff", htmlContent], {
      type: "application/msword;charset=utf-8",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `completion-cert-${data.projectCode || "certificate"}.doc`;
    a.click();
    URL.revokeObjectURL(url);
  }

  function handleReset() {
    if (confirm("确认清空所有内容并重新填写？")) {
      setData(defaultData);
      setMode("form");
    }
  }

  function handleCustomerChange(val: string) {
    set("requester", val);
    const match = options.customers.find(
      (c) => c.name.toLowerCase() === val.toLowerCase()
    );
    if (match) {
      set("requesterContact", match.contact);
    }
  }

  return (
    <div className="min-h-screen bg-[#f0f4f8]">
      {/* Top bar */}
      <div className="no-print sticky top-0 z-50 bg-white border-b border-[#dde4ed] shadow-sm flex items-center justify-between px-6 py-3">
        <div className="flex items-center gap-3">
          <img src={logo} alt="Enrigin" className="h-8 w-auto" />
          <div>
            <div className="font-semibold text-[#1e3a52] text-sm leading-none">完工确认书模板</div>
            <div className="text-xs text-[#6b8099] mt-0.5">Completion &amp; Acceptance Certificate</div>
          </div>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => setShowSettings(true)}
            className="px-4 py-1.5 rounded text-sm font-medium border border-gray-300 text-gray-500 hover:bg-gray-50 transition-colors"
            title="配置下拉选项"
          >
            ⚙ 选项配置
          </button>
          <button
            onClick={() => setMode(mode === "form" ? "preview" : "form")}
            className="px-4 py-1.5 rounded text-sm font-medium border border-[#2B5F8B] text-[#2B5F8B] hover:bg-[#2B5F8B] hover:text-white transition-colors"
          >
            {mode === "form" ? "预览 / Preview" : "返回编辑 / Edit"}
          </button>
          {mode === "preview" && (
            <>
              <button
                onClick={handleDownloadPDF}
                className="px-4 py-1.5 rounded text-sm font-medium bg-[#2B5F8B] text-white hover:bg-[#1e4a6e] transition-colors"
              >
                ↓ PDF
              </button>
              <button
                onClick={handleDownloadWord}
                className="px-4 py-1.5 rounded text-sm font-medium border border-[#2B5F8B] text-[#2B5F8B] hover:bg-[#2B5F8B] hover:text-white transition-colors"
              >
                ↓ Word
              </button>
              <button
                onClick={handlePrint}
                className="px-4 py-1.5 rounded text-sm font-medium border border-gray-300 text-gray-500 hover:bg-gray-50 transition-colors"
              >
                打印 / Print
              </button>
            </>
          )}
          <button
            onClick={handleReset}
            className="px-4 py-1.5 rounded text-sm font-medium border border-gray-300 text-gray-500 hover:bg-gray-50 transition-colors"
          >
            重置
          </button>
        </div>
      </div>
      {/* Form mode */}
      {mode === "form" && (
        <div className="max-w-3xl mx-auto py-8 px-4">
          <div className="bg-white rounded-xl shadow-sm border border-[#dde4ed] overflow-hidden">
            <div className="bg-[#2B5F8B] px-6 py-4">
              <h2 className="text-white font-semibold text-base">快速填写 / Quick Fill</h2>
              <p className="text-[#a8c4d9] text-xs mt-0.5">填写完成后点击「预览」查看证书效果并打印</p>
            </div>

            <div className="p-6 space-y-6">
              {/* Section 1 */}
              <section>
                <h3 className="text-[#2B5F8B] font-semibold text-sm uppercase tracking-wider mb-4 pb-2 border-b border-[#dde4ed]">
                  SECTION 1 · 项目信息
                </h3>
                <div className="grid grid-cols-2 gap-4">
                  <Field label="项目编号 / Project Code">
                    <input className={inputCls} value={data.projectCode} onChange={e => set("projectCode", e.target.value)} placeholder="SV-BJ20250604-01" />
                  </Field>
                  <Field label="完工日期 / Completion Date">
                    <input className={inputCls} type="date" value={data.completionDate} onChange={e => set("completionDate", e.target.value)} />
                  </Field>
                  <Field label="项目名称 / Project Name" className="col-span-2">
                    <input className={inputCls} value={data.projectName} onChange={e => set("projectName", e.target.value)} placeholder="" />
                  </Field>
                  <Field label="客户 / Customer">
                    <input
                      className={inputCls}
                      list="dl-customers"
                      value={data.requester}
                      onChange={e => handleCustomerChange(e.target.value)}
                      placeholder=""
                    />
                    <datalist id="dl-customers">
                      {options.customers.map((c, i) => (
                        <option key={i} value={c.name} />
                      ))}
                    </datalist>
                  </Field>
                  <Field label="联系方式 / Contact">
                    <input className={inputCls} value={data.requesterContact} onChange={e => set("requesterContact", e.target.value)} placeholder="+86 138... · email@example.com" />
                  </Field>
                  <Field label="最终用户 / End User">
                    <input
                      className={inputCls}
                      list="dl-endusers"
                      value={data.endUser}
                      onChange={e => set("endUser", e.target.value)}
                      placeholder=""
                    />
                    <datalist id="dl-endusers">
                      {options.endUsers.map((u, i) => (
                        <option key={i} value={u} />
                      ))}
                    </datalist>
                  </Field>
                  <Field label="交付地址 / Delivery Address" className="col-span-2">
                    <input className={inputCls} value={data.deliveryAddress} onChange={e => set("deliveryAddress", e.target.value)} placeholder="" />
                  </Field>
                  <Field label="服务工程师 / Service Engineer(s)" className="col-span-2">
                    <EngineerMultiSelect
                      selected={data.serviceEngineer}
                      presets={options.engineers}
                      onToggle={toggleEngineer}
                      onAdd={addCustomEngineer}
                      onRemove={removeEngineer}
                    />
                  </Field>
                </div>

                <div className="mt-4">
                  <label className={labelCls}>服务级别 / Service Level</label>
                  <div className="flex gap-4 mt-1">
                    {(["urgent", "non-urgent"] as const).map(v => (
                      <label key={v} className="flex items-center gap-2 cursor-pointer">
                        <input type="radio" name="serviceLevel" checked={data.serviceLevel === v} onChange={() => set("serviceLevel", v)} className="accent-[#2B5F8B]" />
                        <span className="text-sm text-gray-700">{v === "urgent" ? "紧急 / Urgent" : "非紧急 / Non-urgent"}</span>
                      </label>
                    ))}
                  </div>
                </div>

                <div className="mt-4">
                  <label className={labelCls}>服务类型 / Service Type（可多选）</label>
                  <div className="grid grid-cols-3 gap-2 mt-2">
                    {SERVICE_TYPES.map(t => (
                      <label key={t} className="flex items-center gap-2 cursor-pointer">
                        <input type="checkbox" checked={data.serviceTypes.includes(t)} onChange={() => toggleServiceType(t)} className="accent-[#2B5F8B]" />
                        <span className="text-sm text-gray-700">{t}</span>
                      </label>
                    ))}
                  </div>
                </div>
              </section>

              {/* Section 2 */}
              <section>
                <h3 className="text-[#2B5F8B] font-semibold text-sm uppercase tracking-wider mb-4 pb-2 border-b border-[#dde4ed]">
                  SECTION 2 · 任务详情 / Task Details
                </h3>
                <div className="space-y-2">
                  {data.taskItems.map((item, i) => (
                    <div key={i} className="flex items-start gap-2">
                      <span className="text-[#2B5F8B] font-medium text-sm mt-2 min-w-[20px]">{i + 1}.</span>
                      <textarea
                        className={`${inputCls} resize-none min-h-[40px]`}
                        value={item}
                        onChange={e => setTaskItem(i, e.target.value)}
                        placeholder={`任务项 ${i + 1}`}
                        rows={2}
                      />
                      {data.taskItems.length > 1 && (
                        <button onClick={() => removeTaskItem(i)} className="mt-1.5 text-gray-400 hover:text-red-500 transition-colors text-lg leading-none">×</button>
                      )}
                    </div>
                  ))}
                  <button onClick={addTaskItem} className="text-[#2B5F8B] text-sm hover:underline mt-1">
                    + 添加任务项
                  </button>
                </div>
                {options.taskPresets.length > 0 && (
                  <div className="mt-3 pt-3 border-t border-[#dde4ed]">
                    <div className="text-xs text-gray-500 mb-2">快速添加预设任务 / Quick add preset:</div>
                    <div className="flex flex-wrap gap-1.5">
                      {options.taskPresets.map((preset, i) => (
                        <button
                          key={i}
                          onClick={() => setData(d => ({ ...d, taskItems: [...d.taskItems, preset] }))}
                          className="px-2.5 py-1 text-xs border border-[#2B5F8B]/40 rounded-full text-[#2B5F8B] hover:bg-[#2B5F8B] hover:text-white transition-colors"
                        >
                          + {preset}
                        </button>
                      ))}
                    </div>
                  </div>
                )}
              </section>

              {/* Section 3 - Timeline (optional) */}
              <section>
                <div className="flex items-center justify-between mb-4 pb-2 border-b border-[#dde4ed]">
                  <h3 className="text-[#2B5F8B] font-semibold text-sm uppercase tracking-wider">
                    SECTION 3 · 时间轴 / Timeline（可选）
                  </h3>
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" checked={data.hasTimeline} onChange={e => set("hasTimeline", e.target.checked)} className="accent-[#2B5F8B]" />
                    <span className="text-sm text-gray-600">包含时间轴</span>
                  </label>
                </div>

                {data.hasTimeline && (
                  <div className="space-y-3">
                    <div className="grid grid-cols-[1fr_1fr_2fr_auto] gap-2 text-xs text-gray-500 font-medium px-1">
                      <span>北京时间 BJ Time</span>
                      <span>当地时间 Local Time</span>
                      <span>行动 Action</span>
                      <span></span>
                    </div>
                    {data.timeline.map((row, i) => (
                      <div key={i} className="grid grid-cols-[1fr_1fr_2fr_auto] gap-2 items-center">
                        <input className={inputCls} value={row.bjTime} onChange={e => setTimelineRow(i, "bjTime", e.target.value)} placeholder="2025-12-18 23:00" />
                        <input className={inputCls} value={row.localTime} onChange={e => setTimelineRow(i, "localTime", e.target.value)} placeholder="2025-12-18 23:00" />
                        <input className={inputCls} value={row.action} onChange={e => setTimelineRow(i, "action", e.target.value)} placeholder="工程师到达机房" />
                        <button onClick={() => removeTimelineRow(i)} className="text-gray-400 hover:text-red-500 text-lg leading-none">×</button>
                      </div>
                    ))}
                    <button onClick={addTimelineRow} className="text-[#2B5F8B] text-sm hover:underline">
                      + 添加时间轴行
                    </button>
                  </div>
                )}
              </section>

              {/* Section 4 - Result */}
              <section>
                <h3 className="text-[#2B5F8B] font-semibold text-sm uppercase tracking-wider mb-4 pb-2 border-b border-[#dde4ed]">
                  SECTION 4 · 服务结果 / Result
                </h3>
                <div className="flex gap-6 mb-4">
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="radio" name="result" checked={data.result === "completed"} onChange={() => set("result", "completed")} className="accent-[#2B5F8B]" />
                    <span className="text-sm text-gray-700">✓ 已完成 / Completed</span>
                  </label>
                </div>

                <div className="mt-4">
                  <label className={labelCls}>是否已与客户远程人员确认 / Confirmed with CMI Remote</label>
                  <div className="flex items-center gap-4 mt-1">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input type="checkbox" checked={data.confirmedWithRemote} onChange={e => set("confirmedWithRemote", e.target.checked)} className="accent-[#2B5F8B]" />
                      <span className="text-sm text-gray-700">是 / Yes</span>
                    </label>
                    {data.confirmedWithRemote && (
                      <input className={`${inputCls} flex-1`} value={data.remoteContact} onChange={e => set("remoteContact", e.target.value)} placeholder="CMI 远程配合人员姓名" />
                    )}
                  </div>
                </div>

                <div className="mt-4">
                  <Field label="备注 / Notes（如有）">
                    <textarea className={`${inputCls} resize-none`} rows={2} value={data.notes} onChange={e => set("notes", e.target.value)} placeholder="其他需要说明的问题..." />
                  </Field>
                </div>

                <div className="mt-4">
                  <label className={labelCls}>服务方签名 / Enrigin Signature</label>
                  <div className="mt-2 flex flex-wrap gap-3">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="radio"
                        name="enriginSig"
                        checked={data.enriginSig === ""}
                        onChange={() => set("enriginSig", "")}
                        className="accent-[#2B5F8B]"
                      />
                      <span className="text-sm text-gray-600">无签名 / None</span>
                    </label>
                    {Object.entries(SIGNATURES).map(([key, sig]) => (
                      <label key={key} className="flex items-center gap-2 cursor-pointer">
                        <input
                          type="radio"
                          name="enriginSig"
                          checked={data.enriginSig === key}
                          onChange={() => set("enriginSig", key)}
                          className="accent-[#2B5F8B]"
                        />
                        <img src={sig.src} alt={sig.label} className="h-8 w-auto border border-[#dde4ed] rounded px-1 bg-white" />
                        <span className="text-sm text-gray-600">{sig.label}</span>
                      </label>
                    ))}
                  </div>
                </div>
              </section>
            </div>

            <div className="border-t border-[#dde4ed] px-6 py-4 bg-gray-50 flex justify-end">
              <button
                onClick={() => setMode("preview")}
                className="px-6 py-2 bg-[#2B5F8B] text-white rounded font-medium text-sm hover:bg-[#1e4a6e] transition-colors"
              >
                生成预览 / Preview Certificate →
              </button>
            </div>
          </div>
        </div>
      )}
      {/* Preview / Print mode */}
      {mode === "preview" && (
        <div className="py-6 px-4">
          <div
            ref={printRef}
            className="print-page bg-white mx-auto shadow-lg border border-gray-200"
            style={{
              width: "210mm",
              minHeight: "297mm",
              padding: "12mm 14mm 10mm 14mm",
              fontFamily: "'Arial', 'Helvetica Neue', 'Microsoft YaHei', sans-serif",
              fontSize: "8pt",
              color: "#1a1a1a",
              boxSizing: "border-box",
            }}
          >
            {/* Header — logo left, title centered, date right */}
            <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: "5mm", paddingBottom: "4mm", borderBottom: "1.5px solid #2B5F8B" }}>
              {/* Left: new logo (ENRIGIN LIMITED already in image) + address */}
              <div style={{ width: "30%" }}>
                <img src={logo} alt="Enrigin Limited" style={{ height: "56px", width: "auto", display: "block", marginBottom: "3px" }} />
                <div style={{ fontSize: "6.5pt", color: "#7a94a8", lineHeight: 1.5 }}>
                  英源國際有限公司<br />
                  9th Floor, Amtel Building, 148 Des Voeux Road, Central, Hong Kong
                </div>
              </div>
              {/* Center: document title */}
              <div style={{ flex: 1, textAlign: "center", padding: "4px 6mm 0" }}>
                <div style={{ fontSize: "15pt", fontWeight: 700, color: "#1a2e3d", letterSpacing: "0.5px", lineHeight: 1.2 }}>
                  完工确认书
                </div>
                <div style={{ fontSize: "10pt", fontWeight: 700, color: "#2B5F8B", letterSpacing: "1px", lineHeight: 1.4 }}>
                  COMPLETION CERTIFICATE
                </div>
              </div>
              {/* Right: date + project code */}
              <div style={{ width: "30%", textAlign: "right", paddingTop: "38px" }}>
                <div style={{ display: "inline-grid", gridTemplateColumns: "auto auto", columnGap: "4mm", rowGap: "3px" }}>
                  <span style={{ fontSize: "6.5pt", color: "#6b8099" }}>Date / 日期</span>
                  <span style={{ fontSize: "7pt", fontWeight: 700 }}>{data.completionDate || today()}</span>
                  <span style={{ fontSize: "6.5pt", color: "#6b8099" }}>Project Code / 项目编号</span>
                  <span style={{ fontSize: "7pt", fontWeight: 700 }}>{data.projectCode || "—"}</span>
                </div>
              </div>
            </div>

            {/* SECTION 1 */}
            <PSecTitle>SECTION 1 &nbsp;·&nbsp; 项目信息 / Project Information</PSecTitle>
            <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: "4mm", fontSize: "7.5pt" }}>
              <colgroup>
                <col style={{ width: "22%" }} />
                <col style={{ width: "28%" }} />
                <col style={{ width: "22%" }} />
                <col style={{ width: "28%" }} />
              </colgroup>
              <tbody>
                <tr>
                  <PCellL>项目名称 / Project Name</PCellL>
                  <PCellV colSpan={3}>{data.projectName || "\u00a0"}</PCellV>
                </tr>
                <tr>
                  <PCellL>客户 / Customer</PCellL>
                  <PCellV>{data.requester || "\u00a0"}</PCellV>
                  <PCellL>联系方式 / Contact</PCellL>
                  <PCellV>{data.requesterContact || "\u00a0"}</PCellV>
                </tr>
                <tr>
                  <PCellL>最终用户 / End User</PCellL>
                  <PCellV>{data.endUser || "\u00a0"}</PCellV>
                  <PCellL>完工日期 / Completion Date</PCellL>
                  <PCellV>{data.completionDate || "\u00a0"}</PCellV>
                </tr>
                <tr>
                  <PCellL>交付地址 / Delivery Address</PCellL>
                  <PCellV colSpan={3}>{data.deliveryAddress || "\u00a0"}</PCellV>
                </tr>
                <tr>
                  <PCellL>服务工程师 / Engineer</PCellL>
                  <PCellV>{data.serviceEngineer.length > 0 ? data.serviceEngineer.join("、") : "\u00a0"}</PCellV>
                  <PCellL>服务级别 / Level</PCellL>
                  <PCellV>
                    {data.serviceLevel === "urgent" ? "紧急 / Urgent" : data.serviceLevel === "non-urgent" ? "非紧急 / Non-urgent" : "\u00a0"}
                  </PCellV>
                </tr>
                <tr>
                  <PCellL>服务类型 / Service Type</PCellL>
                  <PCellV colSpan={3}>{data.serviceTypes.length > 0 ? data.serviceTypes.join("  ·  ") : "\u00a0"}</PCellV>
                </tr>
              </tbody>
            </table>

            {/* SECTION 2 */}
            <PSecTitle>SECTION 2 &nbsp;·&nbsp; 任务详情 / Task Details</PSecTitle>
            <div style={{ border: "1px solid #c5d5e5", padding: "2.5mm 3.5mm", marginBottom: "4mm", backgroundColor: "#fafcff", fontSize: "7.5pt" }}>
              {data.taskItems.filter(Boolean).length > 0 ? (
                <ol style={{ margin: 0, paddingLeft: "20px", lineHeight: 1.75, listStyleType: "decimal" }}>
                  {data.taskItems.filter(Boolean).map((item, i) => (
                    <li key={i} style={{ listStyleType: "decimal", display: "list-item" }}>{item}</li>
                  ))}
                </ol>
              ) : (
                <div style={{ color: "#bbb" }}>（任务详情 / Task details）</div>
              )}
            </div>

            {/* SECTION 3 - Timeline (optional) */}
            {data.hasTimeline && (
              <>
                <PSecTitle>SECTION 3 &nbsp;·&nbsp; 服务时间轴 / Service Timeline</PSecTitle>
                <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: "4mm", fontSize: "7.5pt" }}>
                  <thead>
                    <tr style={{ backgroundColor: "#2B5F8B", color: "white" }}>
                      <th style={pThStyle}>北京时间 / BJ Time</th>
                      <th style={pThStyle}>当地时间 / Local Time</th>
                      <th style={{ ...pThStyle, width: "52%" }}>行动 / Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {data.timeline.filter(r => r.action || r.bjTime).map((row, i) => (
                      <tr key={i} style={{ backgroundColor: i % 2 === 0 ? "#fff" : "#f4f8fc" }}>
                        <td style={pTdStyle}>{row.bjTime || "—"}</td>
                        <td style={pTdStyle}>{row.localTime || "—"}</td>
                        <td style={pTdStyle}>{row.action || "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            )}

            {/* SECTION 3/4 - Result */}
            <PSecTitle>{data.hasTimeline ? "SECTION 4" : "SECTION 3"} &nbsp;·&nbsp; 服务结果 / Service Result</PSecTitle>
            <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: "5mm", fontSize: "7.5pt" }}>
              <tbody>
                <tr>
                  <PCellL w="19%">服务结果 / Result</PCellL>
                  <PCellV colSpan={3}>
                    <PCheckBox checked={data.result === "completed"} /> 已完成 / Completed
                  </PCellV>
                </tr>
                <tr>
                  <PCellL>已与客户确认 / Confirmed</PCellL>
                  <PCellV colSpan={3}>
                    <PCheckBox checked={data.confirmedWithRemote} /> 是 / Yes
                    {data.confirmedWithRemote && data.remoteContact ? `  —  ${data.remoteContact}` : ""}
                  </PCellV>
                </tr>
                {data.notes && (
                  <tr>
                    <PCellL>备注 / Notes</PCellL>
                    <PCellV colSpan={3}>{data.notes}</PCellV>
                  </tr>
                )}
              </tbody>
            </table>

            {/* Signature */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "6mm" }}>
              <PSignBox label="Enrigin Limited" sub="服务方签章 / Authorized Signature" sigSrc={data.enriginSig ? SIGNATURES[data.enriginSig]?.src : undefined} />
              <PSignBox label="客户 / Customer" sub="客户签章 / Authorized Signature or Chop" />
            </div>

            {/* Footer */}
            <div style={{ marginTop: "4mm", borderTop: "1px solid #dde7f0", paddingTop: "2mm", textAlign: "center", fontSize: "6pt", color: "#aaa" }}>ENRIGIN LIMITED 英源國際有限公司</div>
          </div>
        </div>
      )}

      {showSettings && (
        <SettingsModal
          options={options}
          onClose={() => setShowSettings(false)}
          onAddCustomer={addCustomer}
          onUpdateCustomer={updateCustomer}
          onRemoveCustomer={removeCustomer}
          onAddListItem={addListItem}
          onUpdateListItem={updateListItem}
          onRemoveListItem={removeListItem}
        />
      )}
    </div>
  );
}

/* ---- small helpers ---- */

function Field({ label, children, className = "" }: { label: string; children: React.ReactNode; className?: string }) {
  return (
    <div className={className}>
      <label className={labelCls}>{label}</label>
      <div className="mt-1">{children}</div>
    </div>
  );
}

const labelCls = "block text-xs font-medium text-gray-500 uppercase tracking-wide";
const inputCls = "w-full border border-[#dde4ed] rounded px-3 py-1.5 text-sm text-gray-800 focus:outline-none focus:border-[#2B5F8B] focus:ring-1 focus:ring-[#2B5F8B]/20 bg-white";

/* ---- Print preview helpers ---- */

function PSecTitle({ children }: { children: React.ReactNode }) {
  return (
    <div style={{
      backgroundColor: "#2B5F8B",
      color: "white",
      fontWeight: 700,
      fontSize: "7pt",
      padding: "1.2mm 3mm",
      marginBottom: "1.5mm",
      letterSpacing: "0.3px",
    }}>
      {children}
    </div>
  );
}

const pCellBase: React.CSSProperties = {
  border: "1px solid #c8d9e8",
  padding: "1.2mm 2.5mm",
  verticalAlign: "middle",
  lineHeight: 1.4,
};

function PCellL({ children, w }: { children: React.ReactNode; w?: string }) {
  return (
    <td style={{ ...pCellBase, backgroundColor: "#edf3f9", fontWeight: 600, color: "#2B5F8B", whiteSpace: "nowrap", width: w || "19%", fontSize: "7pt" }}>
      {children}
    </td>
  );
}

function PCellV({ children, colSpan, w, style }: { children: React.ReactNode; colSpan?: number; w?: string; style?: React.CSSProperties }) {
  return (
    <td style={{ ...pCellBase, width: w, ...(style || {}) }} colSpan={colSpan}>
      {children}
    </td>
  );
}

const pThStyle: React.CSSProperties = {
  padding: "1.2mm 2.5mm",
  textAlign: "left",
  fontWeight: 600,
  fontSize: "7pt",
  border: "1px solid #1e4a6e",
};

const pTdStyle: React.CSSProperties = {
  padding: "1.2mm 2.5mm",
  border: "1px solid #c8d9e8",
  fontSize: "7pt",
  verticalAlign: "top",
  lineHeight: 1.4,
};

function PCheckBox({ checked }: { checked: boolean }) {
  return (
    <span style={{
      display: "inline-block",
      width: "8px",
      height: "8px",
      border: "1px solid #2B5F8B",
      marginRight: "3px",
      verticalAlign: "middle",
      backgroundColor: checked ? "#2B5F8B" : "transparent",
      position: "relative",
    }}>
      {checked && <span style={{ color: "white", fontSize: "7px", position: "absolute", top: "-2px", left: "0.5px", lineHeight: 1 }}>✓</span>}
    </span>
  );
}

function PSignBox({ label, sub, sigSrc }: { label: string; sub: string; sigSrc?: string }) {
  return (
    <div style={{ border: "1px solid #c8d9e8", padding: "2.5mm 3mm", minHeight: "30mm", display: "flex", flexDirection: "column" }}>
      <div style={{ fontSize: "7pt", fontWeight: 700, color: "#2B5F8B", marginBottom: "1mm" }}>{label}</div>
      <div style={{ fontSize: "6.5pt", color: "#999", marginBottom: "1mm" }}>{sub}</div>
      <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center" }}>
        {sigSrc && (
          <img
            src={sigSrc}
            alt="signature"
            style={{ height: "16mm", width: "auto", maxWidth: "88%", objectFit: "contain" }}
          />
        )}
      </div>
      <div style={{ borderTop: "1px dashed #c8d9e8", paddingTop: "1.5mm", fontSize: "6.5pt", color: "#ccc", textAlign: "center" }}>
        Authorized Signature
      </div>
    </div>
  );
}

/* ---- Engineer multi-select ---- */

function EngineerMultiSelect({
  selected,
  presets,
  onToggle,
  onAdd,
  onRemove,
}: {
  selected: string[];
  presets: string[];
  onToggle: (name: string) => void;
  onAdd: (name: string) => void;
  onRemove: (name: string) => void;
}) {
  const [custom, setCustom] = useState("");

  function handleCustomAdd() {
    if (!custom.trim()) return;
    onAdd(custom.trim());
    setCustom("");
  }

  return (
    <div className="space-y-2">
      {/* Selected chips */}
      {selected.length > 0 && (
        <div className="flex flex-wrap gap-1.5">
          {selected.map((name) => (
            <span
              key={name}
              className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-medium bg-[#2B5F8B] text-white"
            >
              {name}
              <button
                type="button"
                onClick={() => onRemove(name)}
                className="text-white/70 hover:text-white leading-none ml-0.5"
              >
                ×
              </button>
            </span>
          ))}
        </div>
      )}

      {/* Preset pills (from settings) */}
      {presets.length > 0 && (
        <div className="flex flex-wrap gap-1.5">
          {presets.map((eng) => {
            const active = selected.includes(eng);
            return (
              <button
                key={eng}
                type="button"
                onClick={() => onToggle(eng)}
                className={`px-2.5 py-0.5 rounded-full text-xs border transition-colors ${
                  active
                    ? "bg-[#2B5F8B]/10 border-[#2B5F8B] text-[#2B5F8B] font-medium"
                    : "border-[#dde4ed] text-gray-500 hover:border-[#2B5F8B] hover:text-[#2B5F8B]"
                }`}
              >
                {active ? "✓ " : "+ "}{eng}
              </button>
            );
          })}
        </div>
      )}

      {/* Custom name input */}
      <div className="flex gap-2">
        <input
          className={inputCls}
          value={custom}
          onChange={(e) => setCustom(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && handleCustomAdd()}
          placeholder={presets.length > 0 ? "或输入其他工程师姓名…" : "输入工程师姓名，回车添加"}
        />
        <button
          type="button"
          onClick={handleCustomAdd}
          disabled={!custom.trim()}
          className="px-3 py-1.5 text-sm border border-[#2B5F8B] text-[#2B5F8B] rounded hover:bg-[#2B5F8B] hover:text-white disabled:opacity-30 disabled:cursor-not-allowed transition-colors shrink-0"
        >
          添加
        </button>
      </div>
    </div>
  );
}
