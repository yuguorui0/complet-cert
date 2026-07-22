import { useState, useCallback } from "react";

export interface CustomerEntry {
  name: string;
  contact: string;
}

export interface OptionsData {
  customers: CustomerEntry[];
  endUsers: string[];
  engineers: string[];
  taskPresets: string[];
}

const STORAGE_KEY = "enrigin_cert_options";

const defaultOptions: OptionsData = {
  customers: [],
  endUsers: [],
  engineers: [],
  taskPresets: [],
};

function load(): OptionsData {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultOptions;
    return { ...defaultOptions, ...JSON.parse(raw) };
  } catch {
    return defaultOptions;
  }
}

function save(data: OptionsData) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}

export function useOptions() {
  const [options, setOptions] = useState<OptionsData>(load);

  const update = useCallback((next: OptionsData) => {
    save(next);
    setOptions(next);
  }, []);

  const addCustomer = useCallback((entry: CustomerEntry) => {
    setOptions((prev) => {
      const next = { ...prev, customers: [...prev.customers, entry] };
      save(next);
      return next;
    });
  }, []);

  const updateCustomer = useCallback((i: number, entry: CustomerEntry) => {
    setOptions((prev) => {
      const customers = prev.customers.map((c, idx) => (idx === i ? entry : c));
      const next = { ...prev, customers };
      save(next);
      return next;
    });
  }, []);

  const removeCustomer = useCallback((i: number) => {
    setOptions((prev) => {
      const customers = prev.customers.filter((_, idx) => idx !== i);
      const next = { ...prev, customers };
      save(next);
      return next;
    });
  }, []);

  const addListItem = useCallback((field: "endUsers" | "engineers" | "taskPresets", val: string) => {
    setOptions((prev) => {
      const next = { ...prev, [field]: [...prev[field], val] };
      save(next);
      return next;
    });
  }, []);

  const updateListItem = useCallback((field: "endUsers" | "engineers" | "taskPresets", i: number, val: string) => {
    setOptions((prev) => {
      const arr = prev[field].map((v, idx) => (idx === i ? val : v));
      const next = { ...prev, [field]: arr };
      save(next);
      return next;
    });
  }, []);

  const removeListItem = useCallback((field: "endUsers" | "engineers" | "taskPresets", i: number) => {
    setOptions((prev) => {
      const arr = prev[field].filter((_, idx) => idx !== i);
      const next = { ...prev, [field]: arr };
      save(next);
      return next;
    });
  }, []);

  return {
    options,
    update,
    addCustomer,
    updateCustomer,
    removeCustomer,
    addListItem,
    updateListItem,
    removeListItem,
  };
}
