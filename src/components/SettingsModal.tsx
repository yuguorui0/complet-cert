import { useState } from "react";
import type { OptionsData, CustomerEntry } from "../hooks/useOptions";

interface Props {
  options: OptionsData;
  onClose: () => void;
  onAddCustomer: (e: CustomerEntry) => void;
  onUpdateCustomer: (i: number, e: CustomerEntry) => void;
  onRemoveCustomer: (i: number) => void;
  onAddListItem: (field: "endUsers" | "engineers" | "taskPresets", val: string) => void;
  onUpdateListItem: (field: "endUsers" | "engineers" | "taskPresets", i: number, val: string) => void;
  onRemoveListItem: (field: "endUsers" | "engineers" | "taskPresets", i: number) => void;
}

type Tab = "customers" | "endUsers" | "engineers" | "taskPresets";

const TAB_LABELS: { key: Tab; label: string }[] = [
  { key: "customers", label: "客户 / Customers" },
  { key: "endUsers", label: "最终用户 / End Users" },
  { key: "engineers", label: "服务工程师 / Engineers" },
  { key: "taskPresets", label: "任务模板 / Task Presets" },
];

const inputCls =
  "w-full border border-[#dde4ed] rounded px-2 py-1 text-sm text-gray-800 focus:outline-none focus:border-[#2B5F8B] focus:ring-1 focus:ring-[#2B5F8B]/20 bg-white";

export default function SettingsModal({
  options,
  onClose,
  onAddCustomer,
  onUpdateCustomer,
  onRemoveCustomer,
  onAddListItem,
  onUpdateListItem,
  onRemoveListItem,
}: Props) {
  const [tab, setTab] = useState<Tab>("customers");

  const [newCustomer, setNewCustomer] = useState<CustomerEntry>({ name: "", contact: "" });
  const [editingCustomer, setEditingCustomer] = useState<number | null>(null);
  const [editCustomerVal, setEditCustomerVal] = useState<CustomerEntry>({ name: "", contact: "" });

  const [newItem, setNewItem] = useState("");
  const [editingItem, setEditingItem] = useState<number | null>(null);
  const [editItemVal, setEditItemVal] = useState("");

  function handleAddCustomer() {
    if (!newCustomer.name.trim()) return;
    onAddCustomer({ name: newCustomer.name.trim(), contact: newCustomer.contact.trim() });
    setNewCustomer({ name: "", contact: "" });
  }

  function handleSaveCustomer(i: number) {
    onUpdateCustomer(i, { name: editCustomerVal.name.trim(), contact: editCustomerVal.contact.trim() });
    setEditingCustomer(null);
  }

  function handleAddItem(field: "endUsers" | "engineers" | "taskPresets") {
    if (!newItem.trim()) return;
    onAddListItem(field, newItem.trim());
    setNewItem("");
  }

  function handleSaveItem(field: "endUsers" | "engineers" | "taskPresets", i: number) {
    onUpdateListItem(field, i, editItemVal.trim());
    setEditingItem(null);
  }

  const listField = tab as "endUsers" | "engineers" | "taskPresets";
  const listData = tab !== "customers" ? options[listField] : [];

  const listLabel: Record<string, string> = {
    endUsers: "最终用户名称",
    engineers: "工程师姓名",
    taskPresets: "任务描述",
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-2xl w-full max-w-2xl max-h-[85vh] flex flex-col overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="bg-[#2B5F8B] px-6 py-4 flex items-center justify-between">
          <div>
            <div className="text-white font-semibold text-base">选项配置 / Settings</div>
            <div className="text-[#a8c4d9] text-xs mt-0.5">管理下拉菜单选项，数据保存在本地浏览器</div>
          </div>
          <button onClick={onClose} className="text-white/70 hover:text-white text-2xl leading-none">×</button>
        </div>

        {/* Tabs */}
        <div className="flex border-b border-[#dde4ed] bg-gray-50">
          {TAB_LABELS.map((t) => (
            <button
              key={t.key}
              onClick={() => { setTab(t.key); setNewItem(""); setEditingItem(null); setEditingCustomer(null); }}
              className={`px-4 py-2.5 text-xs font-medium transition-colors border-b-2 ${
                tab === t.key
                  ? "border-[#2B5F8B] text-[#2B5F8B] bg-white"
                  : "border-transparent text-gray-500 hover:text-gray-700"
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>

        {/* Body */}
        <div className="flex-1 overflow-y-auto p-5">
          {tab === "customers" ? (
            <div className="space-y-2">
              {/* Existing customers */}
              {options.customers.map((c, i) => (
                <div key={i} className="border border-[#dde4ed] rounded-lg p-3 bg-[#fafcff]">
                  {editingCustomer === i ? (
                    <div className="space-y-2">
                      <div className="grid grid-cols-2 gap-2">
                        <div>
                          <label className="block text-xs text-gray-500 mb-1">客户名称 / Name</label>
                          <input
                            className={inputCls}
                            value={editCustomerVal.name}
                            onChange={(e) => setEditCustomerVal((v) => ({ ...v, name: e.target.value }))}
                            onKeyDown={(e) => e.key === "Enter" && handleSaveCustomer(i)}
                          />
                        </div>
                        <div>
                          <label className="block text-xs text-gray-500 mb-1">联系方式 / Contact</label>
                          <input
                            className={inputCls}
                            value={editCustomerVal.contact}
                            onChange={(e) => setEditCustomerVal((v) => ({ ...v, contact: e.target.value }))}
                            onKeyDown={(e) => e.key === "Enter" && handleSaveCustomer(i)}
                            placeholder="+86 138... · email@example.com"
                          />
                        </div>
                      </div>
                      <div className="flex gap-2 justify-end">
                        <button onClick={() => setEditingCustomer(null)} className="px-3 py-1 text-xs border border-gray-300 rounded text-gray-600 hover:bg-gray-50">取消</button>
                        <button onClick={() => handleSaveCustomer(i)} className="px-3 py-1 text-xs bg-[#2B5F8B] text-white rounded hover:bg-[#1e4a6e]">保存</button>
                      </div>
                    </div>
                  ) : (
                    <div className="flex items-start justify-between gap-2">
                      <div>
                        <div className="text-sm font-medium text-gray-800">{c.name}</div>
                        {c.contact && <div className="text-xs text-gray-500 mt-0.5">{c.contact}</div>}
                      </div>
                      <div className="flex gap-1 shrink-0">
                        <button
                          onClick={() => { setEditingCustomer(i); setEditCustomerVal({ ...c }); }}
                          className="px-2.5 py-1 text-xs border border-[#dde4ed] rounded text-[#2B5F8B] hover:bg-blue-50"
                        >编辑</button>
                        <button
                          onClick={() => onRemoveCustomer(i)}
                          className="px-2.5 py-1 text-xs border border-red-200 rounded text-red-500 hover:bg-red-50"
                        >删除</button>
                      </div>
                    </div>
                  )}
                </div>
              ))}

              {/* Add new customer */}
              <div className="border border-dashed border-[#2B5F8B]/40 rounded-lg p-3 bg-blue-50/30">
                <div className="text-xs font-medium text-[#2B5F8B] mb-2">+ 添加新客户</div>
                <div className="grid grid-cols-2 gap-2 mb-2">
                  <div>
                    <label className="block text-xs text-gray-500 mb-1">客户名称 / Name</label>
                    <input
                      className={inputCls}
                      value={newCustomer.name}
                      onChange={(e) => setNewCustomer((v) => ({ ...v, name: e.target.value }))}
                      onKeyDown={(e) => e.key === "Enter" && handleAddCustomer()}
                      placeholder="例：China Mobile International"
                    />
                  </div>
                  <div>
                    <label className="block text-xs text-gray-500 mb-1">联系方式 / Contact</label>
                    <input
                      className={inputCls}
                      value={newCustomer.contact}
                      onChange={(e) => setNewCustomer((v) => ({ ...v, contact: e.target.value }))}
                      onKeyDown={(e) => e.key === "Enter" && handleAddCustomer()}
                      placeholder="+86 138... · email@example.com"
                    />
                  </div>
                </div>
                <button
                  onClick={handleAddCustomer}
                  disabled={!newCustomer.name.trim()}
                  className="px-4 py-1.5 text-xs bg-[#2B5F8B] text-white rounded hover:bg-[#1e4a6e] disabled:opacity-40 disabled:cursor-not-allowed"
                >
                  添加
                </button>
              </div>

              {options.customers.length === 0 && (
                <p className="text-xs text-gray-400 text-center py-2">暂无客户，请先添加</p>
              )}
            </div>
          ) : (
            <div className="space-y-2">
              {listData.map((item, i) => (
                <div key={i} className="border border-[#dde4ed] rounded-lg px-3 py-2 bg-[#fafcff] flex items-center gap-2">
                  {editingItem === i ? (
                    <>
                      <input
                        className={`${inputCls} flex-1`}
                        value={editItemVal}
                        onChange={(e) => setEditItemVal(e.target.value)}
                        onKeyDown={(e) => e.key === "Enter" && handleSaveItem(listField, i)}
                        autoFocus
                      />
                      <button onClick={() => setEditingItem(null)} className="px-2 py-1 text-xs border border-gray-300 rounded text-gray-600 hover:bg-gray-50 shrink-0">取消</button>
                      <button onClick={() => handleSaveItem(listField, i)} className="px-2 py-1 text-xs bg-[#2B5F8B] text-white rounded hover:bg-[#1e4a6e] shrink-0">保存</button>
                    </>
                  ) : (
                    <>
                      <span className="flex-1 text-sm text-gray-800">{item}</span>
                      <button
                        onClick={() => { setEditingItem(i); setEditItemVal(item); }}
                        className="px-2 py-1 text-xs border border-[#dde4ed] rounded text-[#2B5F8B] hover:bg-blue-50 shrink-0"
                      >编辑</button>
                      <button
                        onClick={() => onRemoveListItem(listField, i)}
                        className="px-2 py-1 text-xs border border-red-200 rounded text-red-500 hover:bg-red-50 shrink-0"
                      >删除</button>
                    </>
                  )}
                </div>
              ))}

              {/* Add new item */}
              <div className="border border-dashed border-[#2B5F8B]/40 rounded-lg p-3 bg-blue-50/30">
                <label className="block text-xs font-medium text-[#2B5F8B] mb-1.5">+ 添加 {listLabel[tab]}</label>
                <div className="flex gap-2">
                  <input
                    className={`${inputCls} flex-1`}
                    value={newItem}
                    onChange={(e) => setNewItem(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && handleAddItem(listField)}
                    placeholder={
                      tab === "taskPresets"
                        ? "例：光纤线路端到端测试及验收"
                        : tab === "engineers"
                        ? "例：张三 / Zhang San"
                        : "例：CMI"
                    }
                  />
                  <button
                    onClick={() => handleAddItem(listField)}
                    disabled={!newItem.trim()}
                    className="px-4 py-1.5 text-xs bg-[#2B5F8B] text-white rounded hover:bg-[#1e4a6e] disabled:opacity-40 disabled:cursor-not-allowed shrink-0"
                  >
                    添加
                  </button>
                </div>
              </div>

              {listData.length === 0 && (
                <p className="text-xs text-gray-400 text-center py-2">暂无选项，请先添加</p>
              )}
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="border-t border-[#dde4ed] px-5 py-3 bg-gray-50 flex justify-end">
          <button onClick={onClose} className="px-5 py-1.5 text-sm bg-[#2B5F8B] text-white rounded hover:bg-[#1e4a6e]">完成 / Done</button>
        </div>
      </div>
    </div>
  );
}
