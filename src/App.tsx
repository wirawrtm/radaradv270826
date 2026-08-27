import React, {
  useState,
  useRef,
  useEffect,
  useMemo,
  useCallback,
} from "react";
import { Html5Qrcode } from "html5-qrcode";
import { AdvantaLogo } from "./AdvantaLogo";
import { UserIcon } from "./UserIcon";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  PieChart,
  Pie,
  Cell,
  AreaChart,
  Area,
  LineChart,
  Line,
  LabelList,
} from "recharts";


const ORIGINAL_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxUUPKhsEo-LencnYjex3gOhVl7w2tS154VCICVbqGfFSBLAwzv0P7XOu9oMTE1jTUg1g/exec";
const SCRIPT_URL = "/api";

const customFetch = async (input: RequestInfo | URL, init?: RequestInit) => {
  const originalFetch = window.fetch;
  let url = typeof input === 'string' ? input : (input instanceof Request ? input.url : '');
  
  // Determine primary target URL (Express Backend /api)
  let primaryUrl = url;
  if (!url.startsWith("http")) {
    if (url.startsWith("?")) {
      primaryUrl = SCRIPT_URL + url;
    } else if (!url.startsWith(SCRIPT_URL)) {
      primaryUrl = SCRIPT_URL + (url.startsWith("/") ? "" : "/") + url;
    } else {
      primaryUrl = url;
    }
  } else if (url.includes(ORIGINAL_SCRIPT_URL)) {
    // If direct Apps Script URL is passed, route it to local Express /api first
    primaryUrl = url.replace(ORIGINAL_SCRIPT_URL, SCRIPT_URL);
  }

  console.log("[API Request]:", primaryUrl, init);

  try {
    const res = await originalFetch(primaryUrl, init);
    const contentType = res.headers.get("content-type");
    if (!res.ok || (contentType && contentType.indexOf("application/json") === -1)) {
      throw new Error(`Invalid API response (Status: ${res.status})`);
    }
    return res;
  } catch (err) {
    console.warn("Primary Express API failed, falling back to Apps Script:", err);
    
    // Construct fallback Apps Script URL
    let fallbackUrl = primaryUrl;
    if (fallbackUrl.startsWith(SCRIPT_URL)) {
      fallbackUrl = fallbackUrl.replace(SCRIPT_URL, ORIGINAL_SCRIPT_URL);
    } else if (!fallbackUrl.startsWith("http")) {
      fallbackUrl = ORIGINAL_SCRIPT_URL + (fallbackUrl.startsWith("?") ? fallbackUrl : "?" + fallbackUrl);
    }
    
    console.log("[AppsScript Fallback Request]:", fallbackUrl, init);
    if (typeof input === 'string') {
      input = fallbackUrl;
    } else if (input instanceof Request) {
      input = new Request(fallbackUrl, init);
    }
    return originalFetch(input, init);
  }
};

const cleanForMatch = (s: any) => {
  if (s === undefined || s === null) return "";
  let str = String(s).toLowerCase().trim();
  str = str.replace(/^(ud\s+|ud\.|cv\s+|cv\.|pt\s+|pt\.|toko\s+|kiosk\s+)/gi, "");
  return str.replace(/[^a-z0-9]/g, "");
};

const INDO_MONTHS = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "Mei",
  "Jun",
  "Jul",
  "Ags",
  "Sep",
  "Okt",
  "Nov",
  "Des",
];

const OFFLINE_EMPLOYEES = [];
const OFFLINE_WORKING_DATA = [];
const OFFLINE_DR_SALES = [];
const OFFLINE_KIOSKS = [];

const getResolvedRole = (pos) => {
  if (!pos) return "";
  const cleaned = cleanForMatch(pos);
  if (cleaned.includes("businessanalyst") || cleaned.includes("ba")) {
    return "Business Analyst";
  }
  if (cleaned.includes("vegetablessalesmanager") || cleaned.includes("vsm")) {
    return "Vegetables Sales Manager";
  }
  if (cleaned.includes("commerciallead") || cleaned.includes("cl")) {
    return "Commercial Lead";
  }
  if (cleaned.includes("countryhead") || cleaned.includes("ch")) {
    return "Country Head";
  }
  if (cleaned.includes("areasalesmanager") || cleaned.includes("asm")) {
    return "Area Sales Manager";
  }
  if (cleaned.includes("salesagronomist") || cleaned.includes("sa")) {
    return "Sales Agronomist";
  }
  if (cleaned.includes("businesssolution") || cleaned.includes("bs")) {
    return "Business Solution";
  }
  return pos;
};

const getValidCategory = (
  itemCat?: any,
  kioskCat?: any,
  kioskName?: string,
  fallbackMap?: Record<string, string>,
) => {
  const cleanItem = String(itemCat || "").trim();
  const cleanKiosk = String(kioskCat || "").trim();

  const isInvalid = (val: string) =>
    !val ||
    val === "-" ||
    val.toLowerCase() === "uncategorized" ||
    val.toLowerCase() === "n/a" ||
    val.toLowerCase() === "unknown";

  let selected = "";
  if (!isInvalid(cleanItem)) {
    selected = cleanItem;
  } else if (!isInvalid(cleanKiosk)) {
    selected = cleanKiosk;
  } else if (fallbackMap && kioskName) {
    const cleanK = cleanForMatch(kioskName);
    const mapped = fallbackMap[cleanK];
    if (mapped && !isInvalid(mapped)) {
      selected = mapped;
    } else {
      for (const [mk, mv] of Object.entries(fallbackMap)) {
        if ((mk.includes(cleanK) || cleanK.includes(mk)) && !isInvalid(mv)) {
          selected = mv;
          break;
        }
      }
    }
  }

  // Fallback 1: Infer category from kioskName keywords
  if (!selected && kioskName) {
    const kLower = kioskName.toLowerCase();
    if (
      kLower.includes("distributor") ||
      kLower.includes("distr") ||
      kLower.includes("gudang") ||
      kLower.includes("warehouse") ||
      kLower.includes("subd")
    ) {
      selected = "Distributor";
    } else if (/\br1\b|toko r1|mitra r1|kiosk r1/i.test(kioskName)) {
      selected = "R1";
    } else if (/\br2\b|toko r2|mitra r2|kiosk r2/i.test(kioskName)) {
      selected = "R2";
    }
  }

  // Fallback 2: Default channel classification if unassigned
  if (!selected) {
    selected = "R1";
  }

  const lower = selected.toLowerCase();
  if (lower === "distributor") return "Distributor";
  if (lower === "r1") return "R1";
  if (lower === "r2") return "R2";
  return selected.charAt(0).toUpperCase() + selected.slice(1);
};

const safeParseFloat = (val: any): number => {
  if (val === undefined || val === null || val === "") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const cleaned = String(val).replace(/,/g, "").trim();
  const num = Number(cleaned);
  return isNaN(num) ? 0 : num;
};

const matchNames = (name1: any, name2: any) => {
  const c1 = cleanForMatch(name1);
  const c2 = cleanForMatch(name2);
  if (!c1 || !c2) return false;
  return c1 === c2;
};

// Helper function for crop fuzzy matching
const checkCropMatch = (itemCrop: string, filterCrop: string): boolean => {
  if (!filterCrop || filterCrop === "All") return true;
  const ic = String(itemCrop || "")
    .toLowerCase()
    .trim();
  const fc = String(filterCrop || "")
    .toLowerCase()
    .trim();
  return ic === fc;
};

const normalizePosition = (pos: string | undefined): string => {
  if (!pos) return "Unknown";
  const clean = cleanForMatch(pos);
  if (clean === "businessanalyst" || clean === "analyst")
    return "Business Analyst";
  if (clean === "areasalesmanager" || clean === "asm")
    return "Area Sales Manager";
  if (clean === "vegetablessalesmanager" || clean === "vsm")
    return "Vegetables Sales Manager";
  if (clean === "salesmanager" || clean === "sm") return "Sales Manager";
  if (clean === "salesagronomist" || clean === "sa") return "Sales Agronomist";
  if (clean === "businesssolution" || clean === "bs")
    return "Business Solution";
  if (clean === "countryhead") return "Country Head";
  if (clean === "commerciallead") return "Commercial Lead";

  // Custom casing logic for presentation
  if (clean.includes("businessanalyst")) return "Business Analyst";
  if (clean.includes("areasalesmanager") || clean.includes("asm"))
    return "Area Sales Manager";
  if (clean.includes("vegetablessalesmanager"))
    return "Vegetables Sales Manager";
  if (clean.includes("salesmanager") || clean.includes("sm"))
    return "Sales Manager";
  if (
    clean.includes("salesagronomist") ||
    clean.includes("sa") ||
    clean.includes("agronomist")
  )
    return "Sales Agronomist";
  if (clean.includes("businesssolution") || clean.includes("bs"))
    return "Business Solution";
  return pos;
};

const getPositionRank = (pos: string | undefined): number => {
  // Hirarki Posisi (dari yang tertinggi ke terendah):
  const norm = normalizePosition(pos);
  if (norm === "Country Head") return 1;
  if (norm === "Commercial Lead") return 1;
  if (norm === "Business Analyst") return 1;
  if (norm === "Vegetables Sales Manager") return 2;
  if (norm === "Sales Manager") return 2;
  if (norm === "Area Sales Manager") return 3;
  if (norm === "Sales Agronomist") return 4;
  if (norm === "Business Solution") return 5;
  
  const normLower = norm.toLowerCase();
  if (
    normLower.includes("head") || 
    normLower.includes("director") || 
    normLower.includes("vp") || 
    normLower.includes("lead") ||
    normLower.includes("business analyst")
  ) {
    return 1;
  }
  if (normLower.includes("manager")) return 2;
  return 5;
};

const parseLevelStr = (val: string | number | undefined | null): number => {
  if (val === undefined || val === null || val === "") return NaN;
  if (typeof val === "number") return val;
  const str = String(val).toUpperCase().trim();
  if (str === "ADMIN") return 4;
  // Prioritize explicit digit
  const match = str.match(/\d+/);
  if (match) return parseInt(match[0], 10);
  // Explicit roman numerals as fallback
  if (str.includes("IV")) return 4;
  if (str.includes("III")) return 3;
  if (str.includes("II")) return 2;
  if (str.includes("V")) return 5;
  if (str.includes("I")) return 1;
  return NaN;
};

const getFromRecord = <T,>(
  record: Record<string, T>,
  key: string | undefined,
): T | undefined => {
  if (!key) return undefined;
  const cleanKey = cleanForMatch(key);
  let foundKey = Object.keys(record).find((k) => cleanForMatch(k) === cleanKey);
  if (!foundKey) {
    foundKey = Object.keys(record).find((k) => matchNames(k, key));
  }
  return foundKey ? record[foundKey] : undefined;
};

const getMemberLevel = (
  name: string,
  teamLevels: Record<string, number>,
  teamPositions: Record<string, string>,
  userData: any,
): number => {
  const cleanName = cleanForMatch(name);

  // Check if it's the logged-in user:
  if (cleanName === cleanForMatch(userData?.name)) {
    if (
      userData?.level !== undefined &&
      userData.level !== null &&
      String(userData.level).trim() !== ""
    ) {
      const parsed = parseLevelStr(userData.level);
      if (!isNaN(parsed)) return parsed;
    }
  }

  // Check teamLevels state:
  const lvl = getFromRecord(teamLevels, name);
  if (lvl !== undefined && lvl !== null && !isNaN(lvl)) {
    return lvl;
  }



  // Fallback based on position name:
  const p =
    getFromRecord(teamPositions, name) ||
    (cleanName === cleanForMatch(userData?.name) ? userData?.position : "");
  const rank = getPositionRank(p);
  if (rank === 1) return 5;
  if (rank === 2) return 4;
  if (rank === 3) return 3;
  if (rank === 4) return 2;
  if (rank === 5) return 1;
  return 0;
};

const compareMembersByLevel = (
  a: string,
  b: string,
  teamLevels: Record<string, number>,
  teamPositions: Record<string, string>,
  userData: any,
): number => {
  const lvlA = getMemberLevel(a, teamLevels, teamPositions, userData);
  const lvlB = getMemberLevel(b, teamLevels, teamPositions, userData);
  if (lvlA !== lvlB) {
    return lvlB - lvlA; // Higher level (most senior) first
  }
  return a.localeCompare(b);
};

const getUplineInTeam = (
  member: string,
  teamMembers: string[],
  teamUpLines: Record<string, string>,
): string | null => {
  const cleanMember = cleanForMatch(member);

  // 1. Follow the direct upline path recursively to find the first ancestor who is in active teamMembers
  let currentUpline = getFromRecord(teamUpLines, member);
  const seen = new Set<string>([cleanMember]);
  while (currentUpline && currentUpline.trim() !== "") {
    const cleanUp = cleanForMatch(currentUpline);
    if (seen.has(cleanUp)) break; // Prevents circular loops
    seen.add(cleanUp);

    let found = teamMembers.find((tm) => cleanForMatch(tm) === cleanUp);
    if (!found) {
      found = teamMembers.find((tm) => matchNames(tm, currentUpline));
    }
    if (found) {
      return found; // Direct or transitive manager found from column!
    }
    currentUpline = getFromRecord(teamUpLines, currentUpline);
  }

  // 2. Strict mode: if no direct/transitive upline found in active teamMembers, return null (root)
  return null;
};

const formatNum = (num) => {
  if (!num) return "0";
  const absNum = Math.abs(num);
  let minFrac = 0;
  let maxFrac = 0;

  if (absNum >= 100) {
    minFrac = 0;
    maxFrac = 0;
  } else if (absNum >= 10) {
    minFrac = 1;
    maxFrac = 1;
  } else if (absNum > 0) {
    minFrac = 2;
    maxFrac = 2;
  }

  return Number(num).toLocaleString("en-US", {
    minimumFractionDigits: minFrac,
    maximumFractionDigits: maxFrac,
  });
};

const formatOverviewVal = (
  num: number | undefined | null,
  forceMt?: boolean
): { valueStr: string; unit: string } => {
  const n = num || 0;
  if (n === 0) {
    return { valueStr: "0", unit: forceMt ? "MT" : "Kg" };
  }
  const absVal = Math.abs(n);
  const isMt = forceMt !== undefined ? forceMt : (absVal >= 1000);

  if (isMt) {
    const mtVal = n / 1000;
    const absMtVal = Math.abs(mtVal);
    let dec = 0;
    if (absMtVal < 10) {
      dec = 2;
    } else if (absMtVal < 100) {
      dec = 1;
    } else {
      dec = 0;
    }
    return {
      valueStr: mtVal.toLocaleString("id-ID", {
        minimumFractionDigits: dec,
        maximumFractionDigits: dec,
      }),
      unit: "MT",
    };
  } else {
    const dec = absVal < 10 ? 2 : 1;
    return {
      valueStr: n.toLocaleString("id-ID", {
        minimumFractionDigits: dec,
        maximumFractionDigits: dec,
      }),
      unit: "Kg",
    };
  }
};

const formatOverviewWithUnit = (
  num: number | undefined | null,
  forceMt?: boolean
): string => {
  const formatted = formatOverviewVal(num, forceMt);
  return `${formatted.valueStr} ${formatted.unit}`;
};

const parseTaskDate = (timestamp: any) => {
  if (!timestamp) return null;
  let d;
  if (typeof timestamp === "string" && timestamp.includes("/")) {
    const parts = timestamp.split(/[\s/:]+/);
    if (parts.length >= 3) {
      d = new Date(
        `${parts[2]}-${parts[1]}-${parts[0]}T${parts[3] || "00"}:${parts[4] || "00"}:${parts[5] || "00"}`,
      );
    } else d = new Date(timestamp);
  } else {
    d = new Date(timestamp);
  }
  return d && !isNaN(d.getTime()) ? d : null;
};

const depthMapCache = new Map<string, Record<string, number>>();

const buildDepthMap = (
  rootName: string,
  teamProfiles: Record<string, any>,
): Record<string, number> => {
  const cacheKey = `${rootName}_${Object.keys(teamProfiles || {}).length}`;
  if (depthMapCache.has(cacheKey)) {
    return depthMapCache.get(cacheKey)!;
  }
  const depths: Record<string, number> = {};
  const cleanRoot = cleanForMatch(rootName);

  // Find if there's a profile in teamProfiles that matches the rootName by name or by key or email
  let realRootName = rootName;
  const foundProfileKey = Object.keys(teamProfiles).find(
    (k) =>
      cleanForMatch(k) === cleanRoot ||
      (teamProfiles[k]?.email &&
        cleanForMatch(teamProfiles[k].email) === cleanRoot),
  );
  if (foundProfileKey) {
    realRootName = foundProfileKey;
  }
  const cleanRealRoot = cleanForMatch(realRootName);

  depths[cleanRealRoot] = 0;
  if (cleanRoot !== cleanRealRoot) {
    depths[cleanRoot] = 0;
  }

  const queue: string[] = [cleanRealRoot];
  const visited = new Set<string>([cleanRealRoot]);

  while (queue.length > 0) {
    const currentClean = queue.shift()!;
    const currentDepth = depths[currentClean];

    Object.entries(teamProfiles).forEach(([name, p]: [string, any]) => {
      const cleanName = cleanForMatch(name);
      if (cleanName === cleanRealRoot || cleanName === cleanRoot) return;
      const cleanUpline = cleanForMatch(p.upline || "");
      if (cleanUpline === currentClean && !visited.has(cleanName)) {
        visited.add(cleanName);
        depths[cleanName] = currentDepth + 1;
        queue.push(cleanName);
      }
    });
  }

  // For any remaining nodes in teamProfiles, try to traverse up their upline to determine depth.
  Object.entries(teamProfiles).forEach(([name]) => {
    const cleanName = cleanForMatch(name);
    if (depths[cleanName] === undefined) {
      let current = cleanName;
      let climbVisited = new Set<string>();
      let path: string[] = [];
      while (
        current &&
        current !== cleanRealRoot &&
        current !== cleanRoot &&
        !climbVisited.has(current)
      ) {
        climbVisited.add(current);
        path.push(current);
        const currentProfile = Object.values(teamProfiles).find(
          (prof: any) => cleanForMatch(prof.name) === current,
        ) as any;
        if (currentProfile && currentProfile.upline) {
          current = cleanForMatch(currentProfile.upline);
        } else {
          break;
        }
      }
      if (current === cleanRealRoot || current === cleanRoot) {
        for (let i = 0; i < path.length; i++) {
          const node = path[i];
          depths[node] = path.length - i;
        }
      } else {
        depths[cleanName] = 99;
      }
    }
  });

  depthMapCache.set(cacheKey, depths);
  return depths;
};

const getDdaOfUserCache = new Map<string, string>();

const getDdaOfUser = (
  picName: string,
  rootName: string | undefined,
  teamProfiles: Record<string, any> | undefined,
): string => {
  if (!rootName || !teamProfiles) return picName;
  const cleanPic = cleanForMatch(picName);
  const cleanRoot = cleanForMatch(rootName);
  if (!cleanPic || cleanPic === "unknown") return picName;

  const cacheKey = `${cleanPic}_${cleanRoot}_${Object.keys(teamProfiles).length}`;
  if (getDdaOfUserCache.has(cacheKey)) {
    return getDdaOfUserCache.get(cacheKey)!;
  }

  const calculate = (): string => {
    let realRootName = rootName;
    const foundProfileKey = Object.keys(teamProfiles).find(
      (k) =>
        cleanForMatch(k) === cleanRoot ||
        (teamProfiles[k]?.email &&
          cleanForMatch(teamProfiles[k].email) === cleanRoot),
    );
    if (foundProfileKey) {
      realRootName = foundProfileKey;
    }
    const cleanRealRoot = cleanForMatch(realRootName);

    const rootProfile = Object.values(teamProfiles).find(
      (p: any) => cleanForMatch(p.name) === cleanRealRoot,
    ) as any;
    const rootPos = rootProfile?.position || "";
    const rootLevelClean = rootProfile?.level ? String(rootProfile.level).toLowerCase().trim() : "";
    const isBusinessAnalyst =
      cleanForMatch(rootPos) === "businessanalyst" ||
      cleanRealRoot === "adityawiratama" ||
      cleanRoot === "adityawiratama" ||
      cleanRealRoot === "aditya" ||
      cleanRoot === "aditya" ||
      rootLevelClean === "admin";

    if (isBusinessAnalyst) {
      return picName;
    }

    const maxThreshold = 5;

    const depths = { ...buildDepthMap(realRootName, teamProfiles) };

    // Also tag depths for the direct root alias so it functions correctly
    depths[cleanRoot] = 0;

    if ((depths[cleanPic] ?? 99) <= maxThreshold) {
      const matched = Object.keys(teamProfiles).find(
        (k) => cleanForMatch(k) === cleanPic,
      );
      return (matched && teamProfiles[matched]?.name) || picName;
    }

    let current = cleanPic;
    let visited = new Set<string>();

    while (current && current !== cleanRealRoot && current !== cleanRoot) {
      if (visited.has(current)) break;
      visited.add(current);

      const profile = Object.values(teamProfiles).find(
        (p: any) => cleanForMatch(p.name) === current,
      ) as any;
      if (!profile || !profile.upline) break;

      const parentClean = cleanForMatch(profile.upline);
      const parentDepth = depths[parentClean] ?? 99;

      if (parentDepth <= maxThreshold) {
        const matched = Object.keys(teamProfiles).find(
          (k) => cleanForMatch(k) === parentClean,
        );
        return (matched && teamProfiles[matched]?.name) || profile.upline;
      }

      current = parentClean;
    }

    return picName;
  };

  const result = calculate();
  getDdaOfUserCache.set(cacheKey, result);
  return result;
};

const EditModal = ({ isOpen, onClose, item, onSave, isSaving, allHybrids = [] }) => {
  const [newQty, setNewQty] = useState("");

  useEffect(() => {
    if (item) {
      setNewQty(item.stock);
    }
  }, [item]);

  if (!isOpen) return null;

  // Calculate prev month's quantity
  const monthsKeys = [
    "jan", "feb", "mar", "apr", "mei", "jun",
    "jul", "ags", "sep", "okt", "nov", "des",
  ];
  const currentMonthIdx = new Date().getMonth();
  const prevMonthIdx = currentMonthIdx === 0 ? 11 : currentMonthIdx - 1;
  const prevMonthKey = monthsKeys[prevMonthIdx];
  const prevMonthLabel = prevMonthKey.toUpperCase();
  const prevVal = item && item[prevMonthKey] !== undefined ? safeParseFloat(item[prevMonthKey]) : 0;

  const handleSaveClick = () => {
    onSave(item.id, newQty, {});
  };

  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-4 overflow-y-auto">
      <div className="bg-white w-full max-w-[380px] rounded-[24px] p-6 shadow-2xl border border-[#edecff] animate-in fade-in zoom-in-95 duration-200 my-auto">
        <div className="size-12 bg-[#edecff] rounded-full flex items-center justify-center text-primary mb-4">
          <span className="material-symbols-outlined text-[24px]">
            edit_note
          </span>
        </div>
        <h2 className="text-lg font-semibold text-[#181a2c] leading-tight mb-1">
          Edit Detail LOT & Stock
        </h2>
        <p className="text-[10px] text-[#8E94B7] font-semibold uppercase tracking-wider mb-5">
          Sesuaikan Qty untuk LOT ini
        </p>

        <div className="space-y-4 mb-6">
          {/* Hybrid / Varietas info block */}
          <div className="p-4 bg-[#fbf8ff] border border-[#edecff] rounded-2xl space-y-3">
            <div className="flex justify-between items-center gap-2">
              <span className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider">
                Hybrid / Varietas
              </span>
              <span className="text-xs font-bold text-primary text-right max-w-[180px] truncate" title={item?.hybrid}>
                {item?.hybrid}
              </span>
            </div>
            
            <div className="h-px bg-[#edecff]" />

            <div className="flex justify-between items-center">
              <span className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider">
                Nomor Lot
              </span>
              <span className="text-xs font-mono font-bold text-[#181a2c]">
                {item?.lot}
              </span>
            </div>

            <div className="h-px bg-[#edecff]" />

            <div className="flex justify-between items-center">
              <span className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider">
                Qty Bulan Lalu ({prevMonthLabel})
              </span>
              <span className="text-xs font-bold text-slate-700">
                {prevVal} Kg
              </span>
            </div>
          </div>

          {/* Quantity Input */}
          <div className="space-y-1">
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 block">
              Quantity Baru (Kg) <span className="text-red-500">*</span>
            </label>
            <input
              type="number"
              value={newQty}
              onChange={(e) => setNewQty(e.target.value)}
              placeholder="Masukkan quantity baru"
              className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all"
              required
              autoFocus
            />
          </div>
        </div>

        <div className="flex gap-3">
          <button
            onClick={onClose}
            className="flex-1 py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors"
          >
            Batal
          </button>
          <button
            onClick={handleSaveClick}
            disabled={isSaving || !newQty || isNaN(Number(newQty))}
            className={`flex-[2] py-3 rounded-full font-semibold text-xs uppercase tracking-wider transition-all ${
              isSaving || !newQty || isNaN(Number(newQty))
                ? "bg-[#e0e0fa] text-[#8E94B7] cursor-not-allowed"
                : "bg-gradient-to-r from-primary to-cyan-400 text-white shadow-[0_4px_12px_rgba(21,75,226,0.25)] active:scale-[0.98]"
            }`}
          >
            {isSaving ? "Saving..." : "Simpan"}
          </button>
        </div>
      </div>
    </div>
  );
};

const ConfirmModal = ({ isOpen, onClose, onConfirm, isProcessing }) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-6">
      <div className="bg-white w-full max-w-[340px] rounded-[24px] p-8 shadow-2xl border border-[#edecff] text-center animate-in fade-in zoom-in-95 duration-200">
        <div className="size-16 bg-red-50 rounded-full flex items-center justify-center text-[#ba1a1a] mx-auto mb-5">
          <span className="material-symbols-outlined text-[32px]">
            priority_high
          </span>
        </div>
        <h2 className="text-xl font-semibold text-[#181a2c] leading-tight mb-1">
          Konfirmasi Hapus
        </h2>
        <p className="text-xs font-semibold text-[#8E94B7] uppercase tracking-wider mb-6">
          Anda yakin ingin menghapus LOT ini?
        </p>
        <div className="space-y-2">
          <button
            onClick={onConfirm}
            disabled={isProcessing}
            className="w-full py-3 bg-gradient-to-r from-[#ba1a1a] to-rose-600 text-white rounded-full font-semibold text-xs uppercase tracking-wider active:scale-[0.98] shadow-md"
          >
            {isProcessing ? "Processing..." : "Ya, Hapus Data"}
          </button>
          <button
            onClick={onClose}
            className="w-full py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors"
          >
            Batal
          </button>
        </div>
      </div>
    </div>
  );
};

const LogoutConfirmModal = ({ isOpen, onClose, onConfirm }) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-6">
      <div className="bg-white w-full max-w-[340px] rounded-[24px] p-8 shadow-2xl border border-[#edecff] text-center animate-in fade-in zoom-in-95 duration-200">
        <div className="size-16 bg-red-50 rounded-full flex items-center justify-center text-[#ba1a1a] mx-auto mb-5">
          <span className="material-symbols-outlined text-[32px] text-red-500">
            logout
          </span>
        </div>
        <h2 className="text-xl font-semibold text-[#181a2c] leading-tight mb-1">
          Konfirmasi Keluar
        </h2>
        <p className="text-xs font-semibold text-[#8E94B7] uppercase tracking-wider mb-6">
          Apakah Anda yakin ingin keluar dari aplikasi?
        </p>
        <div className="space-y-2">
          <button
            onClick={onConfirm}
            className="w-full py-3 bg-gradient-to-r from-red-600 to-rose-600 text-white rounded-full font-semibold text-xs uppercase tracking-wider active:scale-[0.98] shadow-md cursor-pointer"
          >
            Ya, Keluar
          </button>
          <button
            onClick={onClose}
            className="w-full py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors cursor-pointer"
          >
            Batal
          </button>
        </div>
      </div>
    </div>
  );
};

const playBeep = () => {
  try {
    const ctx = new (
      window.AudioContext || (window as any).webkitAudioContext
    )();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = "sine";
    osc.frequency.setValueAtTime(1200, ctx.currentTime);
    gain.gain.setValueAtTime(0.2, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.12);
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    osc.stop(ctx.currentTime + 0.12);
  } catch (err) {
    console.warn("Could not play scan beep:", err);
  }
};

const QrScanModal = ({ isOpen, onClose, onScanSuccess }) => {
  const [scannerError, setScannerError] = useState<string | null>(null);
  const scannerRef = useRef<any>(null);

  useEffect(() => {
    if (!isOpen) {
      setScannerError(null);
      return;
    }

    let isScanningActive = false;
    let html5QrCode: any = null;

    const startCamera = async () => {
      try {
        html5QrCode = new Html5Qrcode("qr-camera-stream");
        scannerRef.current = html5QrCode;

        await html5QrCode.start(
          { facingMode: "environment" },
          {
            fps: 15,
            qrbox: (width, height) => {
              const size = Math.min(width, height) * 0.95;
              return { width: size, height: size };
            },
          },
          (decodedText) => {
            playBeep();
            onScanSuccess(decodedText);
            onClose();
          },
          () => {
            // silent frame error check
          },
        );
        isScanningActive = true;
      } catch (err) {
        console.error("Camera access error:", err);
        setScannerError(
          "Gagal mengakses kamera. Mohon berikan izin kamera pada browser Anda.",
        );
      }
    };

    const timer = setTimeout(() => {
      startCamera();
    }, 250);

    return () => {
      clearTimeout(timer);
      if (html5QrCode) {
        if (isScanningActive) {
          html5QrCode.stop().catch((e) => console.log("Stop failed:", e));
        }
      }
    };
  }, [isOpen, onClose, onScanSuccess]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[120] bg-[#181a2c]/80 backdrop-blur-md flex flex-col justify-center items-center p-6 animate-in fade-in duration-200">
      <style>{`
        #qr-camera-stream {
          background-color: transparent !important;
          border: none !important;
        }
        #qr-camera-stream video {
          object-fit: cover !important;
          width: 100% !important;
          height: 100% !important;
          border-radius: 22px !important;
        }
        #qr-camera-stream canvas {
          display: none !important;
        }
        #qr-camera-stream img {
          display: none !important;
        }
      `}</style>
      <div className="bg-white w-full max-w-[340px] rounded-[32px] overflow-hidden shadow-[0_24px_64px_rgba(24,26,44,0.12)] border border-[#f0edff] p-6 flex flex-col items-center relative animate-in fade-in zoom-in-95 duration-250">
        <button
          onClick={onClose}
          className="absolute top-4 right-4 size-8 rounded-full bg-[#f4f2ff] text-[#8E94B7] hover:text-red-500 hover:bg-red-50 transition-all flex items-center justify-center cursor-pointer z-10"
        >
          <span className="material-symbols-outlined text-lg">close</span>
        </button>

        <h3 className="text-base font-bold text-[#181a2c] tracking-wide mb-1 mt-2 text-center">
          Pindai QR / Barcode
        </h3>
        <p className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider text-center mb-6">
          Arahkan kamera ke kode LOT
        </p>

        {/* Thick elegant gradient container matching the website's theme */}
        <div className="relative w-full aspect-square max-w-[260px] p-[10px] rounded-[44px] bg-gradient-to-br from-primary to-cyan-400 shadow-[0_20px_48px_rgba(21,75,226,0.22)] flex items-center justify-center">
          <div className="w-full h-full bg-white rounded-[34px] p-[12px] flex items-center justify-center relative overflow-hidden">
            <div
              id="qr-camera-stream"
              className="w-full h-full object-cover rounded-[22px] overflow-hidden bg-slate-950"
            ></div>

            {!scannerError && (
              <div className="absolute inset-[16px] pointer-events-none flex items-center justify-center z-10">
                {/* Elegant scanning corners */}
                <div className="absolute top-0 left-0 w-5 h-5 border-t-[3px] border-l-[3px] border-primary rounded-tl-md"></div>
                <div className="absolute top-0 right-0 w-5 h-5 border-t-[3px] border-r-[3px] border-primary rounded-tr-md"></div>
                <div className="absolute bottom-0 left-0 w-5 h-5 border-b-[3px] border-l-[3px] border-cyan-400 rounded-bl-md"></div>
                <div className="absolute bottom-0 right-0 w-5 h-5 border-b-[3px] border-r-[3px] border-cyan-400 rounded-br-md"></div>

                {/* Scanning laser line in matching gradient */}
                <div
                  className="w-[95%] h-0.5 bg-gradient-to-r from-transparent via-primary to-transparent opacity-85 absolute animate-bounce"
                  style={{ top: "15%", animationDuration: "2.5s" }}
                ></div>
              </div>
            )}

            {scannerError && (
              <div className="absolute inset-0 bg-slate-950 text-white flex flex-col items-center justify-center p-4 text-center rounded-[22px] overflow-hidden">
                <span className="material-symbols-outlined text-red-500 text-3xl mb-2">
                  videocam_off
                </span>
                <p className="text-[11px] font-semibold select-none leading-relaxed text-slate-300">
                  {scannerError}
                </p>
              </div>
            )}
          </div>
        </div>

        <div className="mt-6 w-full">
          <button
            type="button"
            onClick={onClose}
            className="w-full h-11 bg-[#f4f2ff] hover:bg-[#edecff] text-[#8E94B7] hover:text-[#181a2c] rounded-full font-bold text-xs uppercase tracking-wider transition-all cursor-pointer"
          >
            Tutup Kamera
          </button>
        </div>
      </div>
    </div>
  );
};

const PartnerEditModal = ({
  isOpen,
  onClose,
  item,
  onSave,
  isSaving,
  activeEmployees,
  allProvinces,
  allCategories,
  userData,
  allGroups,
}) => {
  const [newPic, setNewPic] = useState("");
  const [partnerName, setPartnerName] = useState("");
  const [category, setCategory] = useState("");
  const [province, setProvince] = useState("");
  const [group, setGroup] = useState("");

  const categoriesToDisplay = useMemo(() => {
    const defaultCategories = ["Distributor", "R1", "R2", "RTL"];
    const merged = [...defaultCategories];
    if (allCategories && Array.isArray(allCategories)) {
      allCategories.forEach((cat) => {
        const trimmed = String(cat || "").trim();
        if (
          trimmed &&
          trimmed !== "Uncategorized" &&
          !merged.some((m) => m.toLowerCase() === trimmed.toLowerCase())
        ) {
          merged.push(trimmed);
        }
      });
    }
    return merged;
  }, [allCategories]);

  const groupsToDisplay = useMemo(() => {
    const list = new Set(["Advanta"]);
    if (allGroups && Array.isArray(allGroups)) {
      allGroups.forEach((g) => {
        const trimmed = String(g || "").trim();
        if (trimmed && trimmed !== "-") {
          list.add(trimmed);
        }
      });
    }
    return Array.from(list);
  }, [allGroups]);

  useEffect(() => {
    const userProv = String(userData?.province || userData?.area || "").trim();
    const userGroup = String(userData?.group || "").trim() || "Advanta";

    if (item) {
      const rawPic = String(item.pic || "").trim();
      const cleanRawPic = cleanForMatch(rawPic);
      const matchedPic = activeEmployees?.find((p) => cleanForMatch(p.name) === cleanRawPic)?.name || rawPic;
      setNewPic(matchedPic);
      setPartnerName(String(item.name || "").trim());
      setCategory(String(item.category || "").trim());
      setProvince(String(item.province || "").trim() || userProv);
      setGroup(String(item.group || "").trim() || userGroup);
    } else {
      setNewPic("");
      setPartnerName("");
      setCategory("");
      setProvince(userProv);
      setGroup(userGroup);
    }
  }, [item, activeEmployees, userData]);

  if (!isOpen) return null;

  const isAdd = !!item?.isAdd;

  const handleSave = () => {
    onSave(isAdd ? null : item.id, newPic, {
      isAdd,
      name: partnerName.trim(),
      category: category.trim(),
      province: province.trim(),
      group: group.trim(),
      originalName: isAdd ? "" : (item?.name || ""),
      originalPic: isAdd ? "" : (item?.pic || ""),
      originalProvince: isAdd ? "" : (item?.province || item?.area || ""),
      originalGroup: isAdd ? "" : (item?.group || ""),
      originalCategory: isAdd ? "" : (item?.category || ""),
    });
  };

  const isFormValid =
    partnerName.trim() !== "" &&
    category.trim() !== "" &&
    province.trim() !== "" &&
    group.trim() !== "" &&
    newPic.trim() !== "";

  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-6">
      <div className="bg-white w-full max-w-[380px] rounded-[24px] p-8 shadow-2xl border border-[#edecff] animate-in fade-in zoom-in-95 duration-200">
        <div className="size-14 bg-[#edecff] rounded-full flex items-center justify-center text-primary mb-5">
          <span className="material-symbols-outlined text-[28px]">
            {isAdd ? "add_business" : "manage_accounts"}
          </span>
        </div>
        <h2 className="text-xl font-semibold text-[#181a2c] leading-tight mb-1">
          {isAdd ? "Tambah Partner" : "Edit Partner"}
        </h2>
        <p className="text-[11px] text-[#8E94B7] font-semibold uppercase tracking-wider mb-6">
          {isAdd ? "Buat Partner Baru" : "Sesuaikan Data Partner"}
        </p>

        <div className="space-y-4 mb-6">
          {/* 1. Kiosk Name */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Nama Partner (Kiosk) <span className="text-red-500">*</span>
            </label>
            <input
              type="text"
              value={partnerName}
              onChange={(e) => setPartnerName(e.target.value)}
              disabled={!isAdd}
              className={`w-full h-11 border border-[#edecff] rounded-xl px-4 text-xs font-bold outline-none transition-all ${
                !isAdd
                  ? "bg-gray-100 cursor-not-allowed text-gray-400"
                  : "bg-[#fbf8ff] text-[#111] focus:border-primary focus:ring-1 focus:ring-primary/10"
              }`}
              placeholder="Contoh: Kios Mandiri Tani"
              required
            />
          </div>



          {/* 3. Category */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Kategori Partner <span className="text-red-500">*</span>
            </label>
            <div className="relative">
              <select
                value={category}
                onChange={(e) => setCategory(e.target.value)}
                className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-xl px-4 text-xs font-bold text-[#111] outline-none transition-all appearance-none pr-10"
                required
              >
                <option value="">-- Pilih Kategori --</option>
                {categoriesToDisplay.map((cat, idx) => (
                  <option key={idx} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
              <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                expand_more
              </span>
            </div>
          </div>



          {/* 5. PIC */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              PIC (Karyawan Aktif) <span className="text-red-500">*</span>
            </label>
            <div className="relative">
              <select
                value={newPic}
                onChange={(e) => setNewPic(e.target.value)}
                className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all appearance-none pr-10"
                required
              >
                <option value="">-- Pilih PIC --</option>
                {activeEmployees?.map((emp, idx) => (
                  <option key={idx} value={emp.name}>
                    {emp.name} ({emp.position || "Staff"})
                  </option>
                ))}
              </select>
              <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                expand_more
              </span>
            </div>
          </div>
        </div>

        <div className="flex gap-3">
          <button
            onClick={onClose}
            className="flex-1 py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors cursor-pointer"
          >
            Batal
          </button>
          <button
            onClick={handleSave}
            disabled={isSaving || !isFormValid}
            className={`flex-[2] py-3 text-white rounded-full font-semibold text-xs uppercase tracking-wider active:scale-[0.98] transition-all cursor-pointer ${
              isSaving || !isFormValid
                ? "bg-[#e0e0fa] text-[#8E94B7] cursor-not-allowed"
                : "bg-gradient-to-r from-primary to-cyan-400 shadow-[0_4px_12px_rgba(21,75,226,0.25)]"
            }`}
          >
            {isSaving ? "Saving..." : "Simpan"}
          </button>
        </div>
      </div>
    </div>
  );
};

const PartnerDeleteModal = ({
  isOpen,
  onClose,
  onConfirm,
  isProcessing,
  itemName,
}) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-6">
      <div className="bg-white w-full max-w-[340px] rounded-[24px] p-8 shadow-2xl border border-[#edecff] text-center animate-in fade-in zoom-in-95 duration-200">
        <div className="size-16 bg-red-50 rounded-full flex items-center justify-center text-[#ba1a1a] mx-auto mb-5">
          <span className="material-symbols-outlined text-[32px]">
            delete_forever
          </span>
        </div>
        <h2 className="text-xl font-semibold text-[#181a2c] leading-tight mb-1">
          Hapus Partner
        </h2>
        <p className="text-xs font-semibold text-[#181a2c] mb-6">
          Hapus partner{" "}
          <span className="font-bold text-red-700">{itemName}</span> dari
          database?
        </p>
        <div className="space-y-2">
          <button
            onClick={onConfirm}
            disabled={isProcessing}
            className="w-full py-3 bg-gradient-to-r from-[#ba1a1a] to-rose-600 text-white rounded-full font-semibold text-xs uppercase tracking-wider active:scale-[0.98] shadow-md"
          >
            {isProcessing ? "Processing..." : "Ya, Hapus"}
          </button>
          <button
            onClick={onClose}
            className="w-full py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors"
          >
            Batal
          </button>
        </div>
      </div>
    </div>
  );
};

const EmployeeEditModal = ({
  isOpen,
  onClose,
  item,
  onSave,
  isSaving,
  allEmployeeNames,
  userData,
  allProvinces,
  accessRules,
}) => {
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [position, setPosition] = useState("");
  const [province, setProvince] = useState("");
  const [password, setPassword] = useState("");
  const [upline, setUpline] = useState("");

  const isAdd = !!item?.isAdd;

  const userLevel = useMemo(() => {
    if (!userData) return 0;
    if (
      userData.level !== undefined &&
      userData.level !== null &&
      String(userData.level).trim() !== ""
    ) {
      const parsed = parseLevelStr(userData.level);
      if (!isNaN(parsed)) return parsed;
    }
    const rank = getPositionRank(userData.position);
    if (rank === 1) return 5;
    if (rank === 2) return 4;
    if (rank === 3) return 3;
    if (rank === 4) return 2;
    if (rank === 5) return 1;
    return 0;
  }, [userData]);

  const isLoginLevel2 = isAdd && userLevel === 2;

  useEffect(() => {
    if (item) {
      setName(String(item.name || ""));
      setEmail(String(item.user || item.email || ""));
      setPassword(String(item.password || ""));
      setUpline(String(item.upline || ""));

      if (isAdd && userLevel === 2) {
        setPosition("Business Solution");
        setProvince(String(userData?.province || "").trim() || "-");
      } else {
        setPosition(String(item.position || ""));
        setProvince(String(item.province || ""));
      }
    }
  }, [item, isAdd, userLevel, userData]);

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!name.trim()) return;
    const rank = getPositionRank(position);
    let levelVal = 0;
    if (rank === 1) levelVal = 5;
    else if (rank === 2) levelVal = 4;
    else if (rank === 3) levelVal = 3;
    else if (rank === 4) levelVal = 2;
    else if (rank === 5) levelVal = 1;
    onSave(
      isAdd ? "" : item?.name || "",
      {
        name,
        email,
        user: email,
        position,
        province,
        password,
        upline: upline || userData?.name,
        level: levelVal,
        group: userData?.group || "Advanta",
      },
      isAdd,
    );
  };

  const isSelf =
    !isAdd &&
    !!(
      userData?.name &&
      item?.name &&
      String(userData.name).toLowerCase().trim() ===
        String(item.name).toLowerCase().trim()
    );

  const loggedInRank = getPositionRank(userData?.position || "");
  const allPos = useMemo(() => {
    const list = Object.keys(accessRules || {});
    // Remove obsolete 'Sales Manager' if it's still lingering in rules
    const filteredList = list.filter(p => p !== "Sales Manager");
    const defaults = [
      "Business Analyst",
      "Vegetables Sales Manager",
      "Area Sales Manager",
      "Sales Agronomist",
      "Business Solution",
    ];
    defaults.forEach(d => {
      if (!filteredList.includes(d)) filteredList.push(d);
    });
    return filteredList;
  }, [accessRules]);

  const positions = allPos.filter((p) => {
    if (item && normalizePosition(item.position) === normalizePosition(p))
      return true;
    return getPositionRank(p) >= loggedInRank;
  });

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-[400px] max-h-[85vh] rounded-[28px] shadow-2xl border border-[#edecff] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        {/* Header */}
        <div className="p-6 pb-4 border-b border-[#edecff] flex items-center gap-3 bg-white">
          <div className="size-11 bg-primary/10 rounded-full flex items-center justify-center text-primary">
            <span className="material-symbols-outlined text-[22px]">
              {isAdd ? "person_add" : "face"}
            </span>
          </div>
          <div>
            <h2 className="text-base font-bold text-[#181a2c] leading-tight">
              {isAdd ? "Tambah Karyawan Baru" : "Edit Detail Karyawan"}
            </h2>
            <p className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider mt-0.5">
              {isAdd
                ? "Anggota Tim"
                : `${item?.name} ${isSelf ? "(Anda)" : ""}`}
            </p>
          </div>
        </div>

        {/* Scrollable Form */}
        <div className="flex-1 overflow-y-auto p-6 space-y-4 custom-scrollbar bg-white">
          {/* Nama */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Nama Lengkap
            </label>
            <input
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              required
              disabled={!isAdd}
              className={`w-full h-11 border border-[#edecff] rounded-full px-4 text-xs font-semibold outline-none transition-all ${
                !isAdd
                  ? "bg-gray-100 cursor-not-allowed text-gray-400"
                  : "bg-[#fbf8ff] text-[#111] focus:border-primary focus:ring-1 focus:ring-primary/10"
              }`}
              placeholder="Masukkan nama"
            />
          </div>

          {/* Username */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Username
            </label>
            <input
              type="text"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-full px-4 text-xs font-semibold text-[#111] outline-none transition-all"
              placeholder="Masukkan username"
            />
          </div>

          {/* Jabatan / Posisi */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Posisi / Jabatan
            </label>
            <div className="relative">
              <select
                value={position}
                onChange={(e) => setPosition(e.target.value)}
                disabled={isSelf || isLoginLevel2}
                className={`w-full h-11 border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-full px-4 text-xs font-semibold outline-none transition-all appearance-none pr-10 ${isSelf || isLoginLevel2 ? "bg-slate-100 text-[#8E94B7] cursor-not-allowed opacity-80" : "bg-[#fbf8ff] text-[#111]"}`}
              >
                <option value="">-- Pilih Posisi --</option>
                {positions.map((p) => (
                  <option key={p} value={p}>
                    {p}
                  </option>
                ))}
              </select>
              <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                expand_more
              </span>
            </div>
          </div>

          {/* Provinsi */}
          {!isLoginLevel2 && (
            <div>
              <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
                Provinsi
              </label>
              <div className="relative">
                <select
                  value={province}
                  onChange={(e) => setProvince(e.target.value)}
                  disabled={isSelf}
                  className={`w-full h-11 border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-full px-4 text-xs font-semibold outline-none transition-all appearance-none pr-10 ${isSelf ? "bg-slate-100 text-[#8E94B7] cursor-not-allowed opacity-80" : "bg-[#fbf8ff] text-[#111]"}`}
                >
                  <option value="">-- Pilih Provinsi --</option>
                  {(allProvinces || []).map((p) => (
                    <option key={p} value={p}>
                      {p}
                    </option>
                  ))}
                </select>
                <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                  expand_more
                </span>
              </div>
            </div>
          )}

          {/* Password */}
          <div>
            <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 mb-1.5 block">
              Password Akun
            </label>
            <input
              type="text"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] focus:border-primary focus:ring-1 focus:ring-primary/10 rounded-full px-4 text-xs font-semibold text-[#111] outline-none transition-all"
              placeholder="Password login"
            />
          </div>
        </div>

        {/* Action Buttons */}
        <div className="p-6 border-t border-[#edecff] flex gap-3 bg-slate-50/50">
          <button
            type="button"
            onClick={onClose}
            className="flex-1 h-11 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors cursor-pointer"
          >
            Batal
          </button>
          <button
            type="button"
            onClick={handleSubmit}
            disabled={isSaving}
            className="flex-[2] h-11 bg-gradient-to-r from-primary to-cyan-400 text-white rounded-full font-semibold text-xs uppercase tracking-wider shadow-[0_4px_12px_rgba(21,75,226,0.25)] active:scale-[0.98] transition-all flex items-center justify-center gap-1.5 cursor-pointer"
          >
            {isSaving ? "Saving..." : "Simpan"}
          </button>
        </div>
      </div>
    </div>
  );
};

const EmployeeDeleteModal = ({
  isOpen,
  onClose,
  onConfirm,
  isProcessing,
  itemName,
}) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[110] bg-[#181a2c]/50 backdrop-blur-md flex items-center justify-center p-6">
      <div className="bg-white w-full max-w-[340px] rounded-[24px] p-8 shadow-2xl border border-[#edecff] text-center animate-in fade-in zoom-in-95 duration-200">
        <div className="size-16 bg-red-50 rounded-full flex items-center justify-center text-[#ba1a1a] mx-auto mb-5">
          <span className="material-symbols-outlined text-[32px]">
            group_remove
          </span>
        </div>
        <h2 className="text-xl font-semibold text-[#181a2c] leading-tight mb-1">
          Hapus Karyawan
        </h2>
        <p className="text-xs font-semibold text-[#181a2c] mb-6">
          Hapus data karyawan{" "}
          <span className="font-bold text-red-700">{itemName}</span> dari
          database?
        </p>
        <div className="space-y-2">
          <button
            onClick={onConfirm}
            disabled={isProcessing}
            className="w-full py-3 bg-gradient-to-r from-[#ba1a1a] to-rose-600 text-white rounded-full font-semibold text-xs uppercase tracking-wider active:scale-[0.98] shadow-md cursor-pointer"
          >
            {isProcessing ? "Processing..." : "Ya, Hapus"}
          </button>
          <button
            onClick={onClose}
            className="w-full py-3 text-[#635b6e] font-semibold text-xs uppercase tracking-wider hover:bg-[#f4f2ff] rounded-full transition-colors cursor-pointer"
          >
            Batal
          </button>
        </div>
      </div>
    </div>
  );
};

const extractExpFromLot = (lotNo: string | undefined | null): string => {
  if (!lotNo) return "N/A";
  const raw = String(lotNo).trim().toUpperCase();
  if (!raw || raw === "N/A" || raw === "-") return "N/A";

  const months = [
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
    "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"
  ];

  // Clean to alphanumeric
  const clean = raw.replace(/[^A-Z0-9]/g, "");

  // 1. Check for 8 digits DDMMYYYY
  const digit8Matches = clean.match(/\d{8}/g);
  if (digit8Matches) {
    for (const mStr of digit8Matches) {
      let d = parseInt(mStr.slice(0, 2), 10);
      let m = parseInt(mStr.slice(2, 4), 10);
      let y = parseInt(mStr.slice(4, 8), 10);

      if (m > 12 && d >= 1 && d <= 12) {
        const temp = d; d = m; m = temp;
      }
      if (d >= 1 && d <= 31 && m >= 1 && m <= 12 && y >= 2020 && y <= 2050) {
        const dayStr = String(d).padStart(2, "0");
        const yearStr = String(y).slice(-2);
        return `${dayStr}/${months[m - 1]}/${yearStr}`;
      }
    }
  }

  // 2. Check for 6 digits (DDMMYY or YYMMDD)
  const digit6Matches = clean.match(/\d{6}/g);
  if (digit6Matches) {
    for (const mStr of digit6Matches) {
      let d = parseInt(mStr.slice(0, 2), 10);
      let m = parseInt(mStr.slice(2, 4), 10);
      let y = parseInt(mStr.slice(4, 6), 10);

      if (y >= 20 && y <= 50) {
        if (m > 12 && d >= 1 && d <= 12) {
          const temp = d; d = m; m = temp;
        }
        if (d >= 1 && d <= 31 && m >= 1 && m <= 12) {
          const dayStr = String(d).padStart(2, "0");
          const yearStr = String(y).padStart(2, "0");
          return `${dayStr}/${months[m - 1]}/${yearStr}`;
        }
      }

      let yy = parseInt(mStr.slice(0, 2), 10);
      let mm = parseInt(mStr.slice(2, 4), 10);
      let dd = parseInt(mStr.slice(4, 6), 10);
      if (yy >= 20 && yy <= 50 && mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
        const dayStr = String(dd).padStart(2, "0");
        const yearStr = String(yy).padStart(2, "0");
        return `${dayStr}/${months[mm - 1]}/${yearStr}`;
      }
    }
  }

  // 3. Fallback: match 6 digits in raw string
  const match = raw.match(/(\d{2})(\d{2})(\d{2})/);
  if (match) {
    let d = parseInt(match[1], 10);
    let m = parseInt(match[2], 10);
    let y = parseInt(match[3], 10);

    if (m > 12 && d >= 1 && d <= 12) {
      const temp = d; d = m; m = temp;
    }

    if (d >= 1 && d <= 31 && m >= 1 && m <= 12 && y >= 20 && y <= 50) {
      const dayStr = String(d).padStart(2, "0");
      const yearStr = String(y).padStart(2, "0");
      return `${dayStr}/${months[m - 1]}/${yearStr}`;
    }
  }

  return "N/A";
};

const DetailItemSection = ({
  items,
  onEdit,
  onDelete,
  onUploadActivity,
  isSyncing,
  hasChanges,
  title,
  subtitle,
  category,
  kiosks = [],
}) => {
  const getConditionBadge = (cond) => {
    const condition = String(cond).toLowerCase();
    if (condition === "new" || condition === "baru")
      return (
        <div className="px-2.5 py-0.5 rounded-full bg-cyan-100/60 border border-cyan-200 text-cyan-800 text-[8px] font-bold uppercase tracking-wider">
          BARU
        </div>
      );
    if (condition === "berkurang")
      return (
        <div className="px-2.5 py-0.5 rounded-full bg-emerald-100/60 border border-emerald-200 text-emerald-800 text-[8px] font-bold uppercase tracking-wider">
          BERKURANG
        </div>
      );
    if (condition === "bertambah")
      return (
        <div className="px-2.5 py-0.5 rounded-full bg-[#edecff] border border-[#c4c5d8] text-primary text-[8px] font-bold uppercase tracking-wider">
          BERTAMBAH
        </div>
      );
    if (condition === "habis")
      return (
        <div className="px-2.5 py-0.5 rounded-full bg-orange-100/60 border border-orange-200 text-orange-850 text-[8px] font-bold uppercase tracking-wider">
          AKAN DIHAPUS
        </div>
      );
    return (
      <div className="px-2.5 py-0.5 rounded-full bg-red-100/60 border border-red-200 text-red-800 text-[8px] font-bold uppercase tracking-wider">
        TETAP
      </div>
    );
  };

  return (
    <div className="-mx-5 bg-white overflow-hidden rounded-[48px] pt-6 px-5 pb-6 mb-8 shadow-[0_4px_44px_rgba(24,26,44,0.15)] border border-[#edecff] font-sans mt-8 relative">
      <div className="flex justify-between items-start mb-6">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <div className="size-2 bg-primary rounded-full animate-pulse" />
            <h3 className="text-base font-semibold text-[#181a2c] tracking-tight">
              {title}
            </h3>
          </div>
          <p className="text-[11px] text-[#8E94B7] font-semibold uppercase tracking-wider">
            {subtitle}
          </p>
          {category && (
            <p className="text-[9px] text-primary font-bold uppercase tracking-widest mt-0.5">
              {category}
            </p>
          )}
        </div>
        <div className="bg-white shadow-[0_4px_12px_rgba(21,75,226,0.08)] px-3 py-1.5 rounded-full text-right">
          <p className="text-[8px] font-bold text-[#8E94B7] uppercase leading-none mb-0.5">
            Total LOT
          </p>
          <p className="text-xs font-bold text-primary">{items.length}</p>
        </div>
      </div>

      <div className="space-y-3.5 max-h-[480px] overflow-y-auto px-5 py-5 -mx-5 -my-5 custom-scrollbar">
        {items.length === 0 ? (
          <div className="py-12 flex flex-col items-center justify-center text-center px-4 bg-white shadow-sm rounded-[24px]">
            <p className="text-[11px] font-semibold text-[#8E94B7] uppercase tracking-wider">
              Feed Kosong
            </p>
          </div>
        ) : (
          items.map((item, index) => (
            <div
              key={item.id || index}
              className="group bg-[#fbfaff] rounded-[18px] p-3.5 flex flex-col gap-3 shadow-[0_16px_36px_rgba(21,75,226,0.25)] hover:shadow-[0_20px_48px_rgba(21,75,226,0.32)] transition-all duration-300"
            >
              <div className="flex items-start justify-between w-full gap-2 font-sans">
                <div className="flex flex-col">
                  <div className="flex items-center gap-2 flex-wrap">
                    <p className="font-semibold text-[#181a2c] text-sm tracking-tight">
                      {item.lot}
                    </p>
                    <div className="px-2 py-0.5 rounded-full bg-red-100/50 border border-red-200 mt-0.5">
                      <p className="text-[8.5px] font-bold text-red-700 uppercase tracking-wide">
                        EXP: {(!item.expired || item.expired === "N/A" || item.expired === "-") ? extractExpFromLot(item.lot) : item.expired}
                      </p>
                    </div>
                  </div>
                  {(() => {
                    const monthsKeys = [
                      "jan",
                      "feb",
                      "mar",
                      "apr",
                      "mei",
                      "jun",
                      "jul",
                      "ags",
                      "sep",
                      "okt",
                      "nov",
                      "des",
                    ];
                    const currentMonthIdx = new Date().getMonth();
                    const prevMonthIdx =
                      currentMonthIdx === 0 ? 11 : currentMonthIdx - 1;
                    const prevMonthKey = monthsKeys[prevMonthIdx];
                    const prevVal =
                      item[prevMonthKey] !== undefined
                        ? safeParseFloat(item[prevMonthKey])
                        : 0;
                    return (
                      <p className="text-[10px] text-[#8E94B7] mt-0.5 font-medium">
                        Bulan lalu ({prevMonthKey.toUpperCase()}):{" "}
                        <span className="text-slate-700 font-semibold">
                          {prevVal} Kg
                        </span>
                      </p>
                    );
                  })()}
                </div>
                <div className="flex flex-col items-end gap-1 flex-shrink-0">
                  {getConditionBadge(item.condition)}
                  {(() => {
                    const kioskInfo = kiosks.find(
                      (k) =>
                        cleanForMatch(k.name) === cleanForMatch(item.kiosk || subtitle),
                    );
                    const picName = getResolvedPic(
                      item.kiosk || subtitle,
                      kioskInfo?.pic,
                      item.user,
                      item.lot,
                      item.pic,
                    );
                    if (!picName || picName === "Unknown") return null;
                    return (
                      <span className="text-[9px] text-[#8E94B7] font-semibold uppercase tracking-wider text-right">
                        PIC: <span className="text-primary font-bold">{picName}</span>
                      </span>
                    );
                  })()}
                </div>
              </div>

              <div className="flex items-center justify-between w-full gap-2">
                <div className="flex items-center gap-2 bg-primary px-3 py-1.5 rounded-full min-w-0 flex-1 shadow-[0_10px_24px_rgba(21,75,226,0.38)]">
                  <p className="text-[11px] text-white font-semibold uppercase tracking-tight whitespace-normal break-words leading-tight flex-1">
                    {item.hybrid}
                  </p>
                  <div className="w-px h-3.5 bg-white/25 flex-shrink-0" />
                  <p className="text-[11px] font-black text-white flex-shrink-0">
                    {(() => {
                      const num = safeParseFloat(item.stock);
                      if (!isNaN(num) && num < 1) {
                        return num.toFixed(2);
                      }
                      return item.stock;
                    })()}{" "}
                    <span className="text-[8.5px] text-white/75 font-medium">
                      Kg
                    </span>
                  </p>
                  <div className="w-px h-3.5 bg-white/25 flex-shrink-0" />
                  <p className="text-[10px] font-black text-amber-300 uppercase tracking-tight flex-shrink-0">
                    {item.aging} BLN
                  </p>
                </div>

                <div className="flex gap-1.5 flex-shrink-0">
                  <button
                    onClick={() => onEdit(item)}
                    className="size-9 rounded-full bg-[#edecff] text-primary hover:bg-[#e6e6ff] transition-all flex items-center justify-center border border-[#c4c5d8] shadow-sm"
                  >
                    <span className="material-symbols-outlined text-[18px]">
                      edit_square
                    </span>
                  </button>
                  <button
                    onClick={() => onDelete(item)}
                    className="size-9 rounded-full bg-red-50 text-red-500 hover:bg-red-100 transition-all flex items-center justify-center border border-red-100 shadow-sm"
                  >
                    <span className="material-symbols-outlined text-[18px]">
                      delete_sweep
                    </span>
                  </button>
                </div>
              </div>
            </div>
          ))
        )}
      </div>

      <button
        onClick={onUploadActivity}
        disabled={isSyncing || !hasChanges}
        className={`w-full h-14 mt-6 text-white rounded-full font-semibold text-xs uppercase tracking-wider shadow-none transition-all flex items-center justify-center gap-3 active:scale-[0.98] ${
          isSyncing || !hasChanges
            ? "bg-[#e0e0fa] text-[#8E94B7] border border-[#edecff] cursor-not-allowed"
            : "bg-gradient-to-r from-primary to-cyan-400 hover:opacity-95 shadow-[0_8px_20px_rgba(21,75,226,0.25)]"
        }`}
      >
        {isSyncing ? (
          <div className="size-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
        ) : (
          <span className="material-symbols-outlined text-sm">
            cloud_upload
          </span>
        )}
        {isSyncing
          ? "Uploading..."
          : hasChanges
            ? "Upload Activity"
            : "Data Tersinkron"}
      </button>
    </div>
  );
};

const Dashboard = ({
  userData,
  activeTab,
  onLogout,
  onUserSwitch,
  setUserData,
  setActiveTab,
  accessRules,
  setAccessRules,
  overviewMetricFilter,
  setOverviewMetricFilter,
  filterBelowMonth,
  setFilterBelowMonth,
  filterBelowChannel,
  setFilterBelowChannel,
  filterBelowMaterial,
  setFilterBelowMaterial,
  filterBelowTeam,
  setFilterBelowTeam,
  filterBelowArea,
  setFilterBelowArea,
  filterBelowCrop,
  setFilterBelowCrop,
}) => {
  const [isDropdownOpen, setIsDropdownOpen] = useState(false);
  const dropdownRef = useRef(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [isUserSwitching, setIsUserSwitching] = useState(false);
  const [historyChartType, setHistoryChartType] = useState<
    "opening" | "ending" | "stockIn" | "idle" | "pog"
  >("pog");

  const handleFullscreen = () => {
    try {
      const docEl = document.documentElement;
      if (
        !document.fullscreenElement &&
        !(document as any).webkitFullscreenElement &&
        !(document as any).mozFullScreenElement &&
        !(document as any).msFullscreenElement
      ) {
        if (docEl.requestFullscreen) {
          docEl.requestFullscreen();
        } else if ((docEl as any).webkitRequestFullscreen) {
          (docEl as any).webkitRequestFullscreen();
        } else if ((docEl as any).mozRequestFullScreen) {
          (docEl as any).mozRequestFullScreen();
        } else if ((docEl as any).msRequestFullscreen) {
          (docEl as any).msRequestFullscreen();
        }
      } else {
        if (document.exitFullscreen) {
          document.exitFullscreen();
        } else if ((document as any).webkitExitFullscreen) {
          (document as any).webkitExitFullscreen();
        } else if ((document as any).mozCancelFullScreen) {
          (document as any).mozCancelFullScreen();
        } else if ((document as any).msExitFullscreen) {
          (document as any).msExitFullscreen();
        }
      }
    } catch (e) {
      console.error("Error toggling fullscreen mode:", e);
    }
  };

  const [kiosks, setKiosks] = useState([]);
  const [workingData, setWorkingData] = useState([]);
  const [rawWorkingData, setRawWorkingData] = useState([]);
  const [drSalesData, setDrSalesData] = useState<any[]>([]);
  const [deletedItems, setDeletedItems] = useState([]);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isChannelsLoading, setIsChannelsLoading] = useState(true);
  const [isFetchingData, setIsFetchingData] = useState(true);
  const [employees, setEmployees] = useState<any[]>([]);
  const [isEmployeesLoading, setIsEmployeesLoading] = useState(false);

  const computedTeamProfiles = useMemo(() => {
    if (!employees || employees.length === 0) return undefined;
    const profiles: Record<string, any> = {};
    employees.forEach((emp) => {
      // Use clean names for keys to ensure successful lookup
      if (emp.name) profiles[cleanForMatch(emp.name)] = emp;
      if (emp.user) profiles[cleanForMatch(emp.user)] = emp;
      if (emp.email) profiles[cleanForMatch(emp.email)] = emp;
    });
    return profiles;
  }, [employees]);
  const [loadProgress, setLoadProgress] = useState(0);
  const [showLoader, setShowLoader] = useState(true);

  useEffect(() => {
    const isLoadingData =
      isFetchingData || isChannelsLoading || isEmployeesLoading;
    if (isLoadingData) {
      setShowLoader(true);
      setLoadProgress(0);
      let currentProgress = 0;
      const interval = setInterval(() => {
        if (currentProgress < 30) {
          currentProgress += Math.floor(Math.random() * 8) + 4;
        } else if (currentProgress < 60) {
          currentProgress += Math.floor(Math.random() * 5) + 2;
        } else if (currentProgress < 85) {
          currentProgress += Math.floor(Math.random() * 3) + 1;
        } else if (currentProgress < 98) {
          currentProgress += Math.random() > 0.6 ? 1 : 0;
        }
        if (currentProgress > 98) currentProgress = 98;
        setLoadProgress(currentProgress);
      }, 150);

      return () => {
        clearInterval(interval);
      };
    } else {
      setLoadProgress(100);
      const timeout = setTimeout(() => {
        setShowLoader(false);
      }, 500);
      return () => clearTimeout(timeout);
    }
  }, [isFetchingData, isChannelsLoading, isEmployeesLoading]);

  // Dimension filter for Executive Overview chart (Area, Province, Sales Agronomist, Business Solution, Material)
  const [overviewGroupDimension, setOverviewGroupDimension] = useState<
    "area" | "province" | "sales_agronomist" | "business_solution" | "material"
  >("area");

  // Dynamic parameters for Overview Bar Chart
  const isMovement = overviewMetricFilter === "movement";
  const isArea = overviewGroupDimension === "area";
  const isProvince = overviewGroupDimension === "province";

  const currentBarGap = 0;
  const currentBarCategoryGap = "10%";
  const currentMaxBarSize = 75;

  const isBusinessAnalyst = useMemo(() => {
    if (!userData) return false;
    const isBA = (userData.position &&
        cleanForMatch(userData.position) === "businessanalyst") ||
      cleanForMatch(userData.name || "") === "adityawiratama" ||
      cleanForMatch(userData.name || "") === "aditya";
    const isAdmin = userData.level && String(userData.level).toLowerCase().trim() === "admin";
    return isBA || isAdmin;
  }, [userData]);

  // State Tab Home
  const [selectedKiosk, setSelectedKiosk] = useState("Loading Kiosk...");

  const recommendedLots = useMemo(() => {
    if (!selectedKiosk) return [];
    const cleanKiosk = cleanForMatch(selectedKiosk);
    const set = new Set<string>();
    const list: { lot: string; hybrid: string }[] = [];
    
    // Find in drSalesData
    if (Array.isArray(drSalesData)) {
      drSalesData.forEach((dr) => {
        if (dr.lot && cleanForMatch(dr.channel) === cleanKiosk) {
          const uLot = String(dr.lot).toUpperCase().trim();
          if (!set.has(uLot)) {
            set.add(uLot);
            list.push({ lot: uLot, hybrid: dr.hybrid || "" });
          }
        }
      });
    }

    // Find in rawWorkingData
    if (Array.isArray(rawWorkingData)) {
      rawWorkingData.forEach((w) => {
        if (w.lot && cleanForMatch(w.kiosk) === cleanKiosk) {
          const uLot = String(w.lot).toUpperCase().trim();
          if (!set.has(uLot)) {
            set.add(uLot);
            list.push({ lot: uLot, hybrid: w.hybrid || "" });
          }
        }
      });
    }

    return list.slice(0, 8);
  }, [selectedKiosk, drSalesData, rawWorkingData]);

  const handleChannelClick = (channelName: string) => {
    const matchedKiosk = kiosks.find(
      (k) => cleanForMatch(k.name) === cleanForMatch(channelName),
    );
    if (matchedKiosk) {
      setSelectedKiosk(matchedKiosk.name);
      if (setActiveTab) {
        setActiveTab("home");
      }
    }
  };

  const renderMaybeChannelName = (name: string, defaultClass = "") => {
    const isChannel = kiosks.some(
      (k) => cleanForMatch(k.name) === cleanForMatch(name),
    );
    if (isChannel) {
      return (
        <span
          onClick={(e) => {
            e.stopPropagation();
            handleChannelClick(name);
          }}
          className={`${defaultClass} hover:underline decoration-primary hover:text-primary transition-all cursor-pointer inline-flex items-center gap-1 font-bold text-primary group`}
          title="Klik untuk input aktivitas partner ini"
        >
          {name}
          <span className="material-symbols-outlined text-[13px] inline opacity-60 group-hover:opacity-100 group-hover:scale-110 transition-all text-primary">
            edit_note
          </span>
        </span>
      );
    }
    return <span className={defaultClass}>{name}</span>;
  };
  const [lotNo, setLotNo] = useState("");
  const [qty, setQty] = useState("");
  const [editModal, setEditModal] = useState({ isOpen: false, item: null });
  const [deleteModal, setDeleteModal] = useState({ isOpen: false, item: null });
  const [isActionLoading, setIsActionLoading] = useState(false);
  const [lotIntel, setLotIntel] = useState(null);
  const [isLotChecking, setIsLotChecking] = useState(false);
  const [isLotNotFound, setIsLotNotFound] = useState(false);
  const [isScannerOpen, setIsScannerOpen] = useState(false);
  const [manualHybrid, setManualHybrid] = useState("");
  const [manualCrop, setManualCrop] = useState("Field Corn");
  const [manualDrDate, setManualDrDate] = useState("");
  const [manualExpDate, setManualExpDate] = useState("");
  const [manualHybridCustom, setManualHybridCustom] = useState("");

  // State Tab Partner
  const [isLogoutModalOpen, setIsLogoutModalOpen] = useState(false);

  const userLevel = useMemo(() => {
    if (!userData) return 0;
    if (
      userData.level !== undefined &&
      userData.level !== null &&
      String(userData.level).trim() !== ""
    ) {
      const parsed = parseLevelStr(userData.level);
      if (!isNaN(parsed)) return parsed;
    }
    const rank = getPositionRank(userData.position);
    if (rank === 1) return 5;
    if (rank === 2) return 4;
    if (rank === 3) return 3;
    if (rank === 4) return 2;
    if (rank === 5) return 1;
    return 0;
  }, [userData]);

  const [partnerSubTab, setPartnerSubTab] = useState(() => {
    const parsedLevel =
      userData?.level !== undefined &&
      userData?.level !== null &&
      String(userData?.level).trim() !== ""
        ? parseLevelStr(userData.level)
        : userData?.position
          ? getPositionRank(userData.position) === 5
            ? 1
            : 0
          : 0;
    return parsedLevel === 1 ? "channel" : "team";
  });

  useEffect(() => {
    if (userLevel === 1) {
      setPartnerSubTab("channel");
    } else {
      setPartnerSubTab("team");
    }
  }, [userData, userLevel]);

  const [mappingPic, setMappingPic] = useState("");
  const [mappingCategory, setMappingCategory] = useState("");
  const [partnerEditModal, setPartnerEditModal] = useState({
    isOpen: false,
    item: null,
  });
  const [partnerDeleteModal, setPartnerDeleteModal] = useState({
    isOpen: false,
    item: null,
  });
  const [channelsRefreshKey, setChannelsRefreshKey] = useState(0);
  const [isChartFocusedModalOpen, setIsChartFocusedModalOpen] = useState(false);

  // State Employee Modifying & Hirarki Expand/Collapse
  const [employeeEditModal, setEmployeeEditModal] = useState<{
    isOpen: boolean;
    item: any | null;
  }>({ isOpen: false, item: null });
  const [employeeDeleteModal, setEmployeeDeleteModal] = useState<{
    isOpen: boolean;
    item: any | null;
  }>({ isOpen: false, item: null });
  const [employeesRefreshKey, setEmployeesRefreshKey] = useState(0);



  const activeEmployees = useMemo(() => {
    return (employees || []).filter((emp: any) => {
      const status = String(emp.status || "").trim().toLowerCase();
      // If no status is specified, treat as active. Only filter out if explicit inactive status.
      return status === "aktif" || status === "active" || status === "" || status === "-";
    });
  }, [employees]);


  const [collapsedNodes, setCollapsedNodes] = useState<Record<string, boolean>>(
    {},
  );

  const availableProvinces = useMemo(() => {
    const list = new Set<string>();
    employees.forEach((emp) => {
      const prov = String(emp.province || "").trim();
      if (prov && prov !== "-") {
        list.add(prov);
      }
    });
    // Fallback defaults
    ["Jawa Timur", "Jawa Tengah", "Jawa Barat"].forEach((p) => list.add(p));
    return Array.from(list).sort();
  }, [employees]);

  const availableGroups = useMemo(() => {
    const list = new Set<string>();
    employees.forEach((emp) => {
      const g = String(emp.group || "").trim();
      if (g && g !== "-") {
        list.add(g);
      }
    });
    if (userData?.group) {
      list.add(String(userData.group).trim());
    }
    ["Advanta"].forEach((g) => list.add(g));
    return Array.from(list).sort();
  }, [employees, userData]);

  // State Tab Summary
  const [summaryGroupBy, setSummaryGroupBy] = useState("hybrid"); // Default changed to 'hybrid'
  const [summarySubGroupBy, setSummarySubGroupBy] = useState("channel"); // Sub category
  const [grandTotalViewBy, setGrandTotalViewBy] = useState<"hybrid" | "area">("hybrid");
  const [isSummaryFilterOpen, setIsSummaryFilterOpen] = useState(true);
  const [isOverviewFilterOpen, setIsOverviewFilterOpen] = useState(false);
  const enrichedSummaryDataRef = useRef<any[]>([]);

  // Access Rules
  const [isSavingAccess, setIsSavingAccess] = useState(false);
  const [accessSaveSuccess, setAccessSaveSuccess] = useState(false);

  const handleSaveAccessRules = async () => {
    setIsSavingAccess(true);
    const filteredRules: Record<string, Record<string, boolean>> = {};
    allPositionsList.forEach(pos => {
      if (accessRules[pos]) {
        filteredRules[pos] = {
          home: !!accessRules[pos].home,
          partner: !!accessRules[pos].partner,
          stock: !!accessRules[pos].stock,
          pog: !!accessRules[pos].pog,
          overview: !!accessRules[pos].overview,
          temp: !!accessRules[pos].temp,
          access: !!accessRules[pos].access,
        };
      } else {
        filteredRules[pos] = {
          home: true,
          partner: true,
          stock: true,
          pog: true,
          overview: pos === "Business Analyst",
          temp: pos === "Business Analyst",
          access: pos === "Business Analyst"
        };
      }
    });

    try {
      localStorage.setItem('appAccessRules', JSON.stringify(filteredRules));
      setAccessRules(filteredRules);
    } catch (e) {
      console.error('Failed to save access rules', e);
    }
    try {
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "saveAccessRules",
          rules: filteredRules
        })
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        throw new Error("Respon server tidak valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        setAccessSaveSuccess(true);
      } else {
        console.error("Failed to save access rules to spreadsheet", res.message);
      }
    } catch (err) {
      console.error("Error saving access rules to spreadsheet:", err);
    } finally {
      setIsSavingAccess(false);
      setTimeout(() => setAccessSaveSuccess(false), 3000);
    }
  };

  const allPositionsList = useMemo(() => {
    const list = Object.keys(accessRules || {});
    // Clean up obsolete 'Sales Manager'
    const filteredList = list.filter(p => p !== "Sales Manager");
    const defaults = [
      "Business Analyst",
      "Vegetables Sales Manager",
      "Area Sales Manager",
      "Sales Agronomist",
      "Business Solution"
    ];
    defaults.forEach(d => {
      if (!filteredList.includes(d)) filteredList.push(d);
    });
    return filteredList;
  }, [accessRules]);

  const removeAccessRule = (position: string) => {
    setAccessRules((prev: Record<string, Record<string, boolean>>) => {
      const currentRules = { ...prev };
      delete currentRules[position];
      return currentRules;
    });
  };

  const toggleAccessRule = (position: string, page: string) => {
    setAccessRules((prev: Record<string, Record<string, boolean>>) => {
      const currentRules = prev[position] || {
        home: true,
        partner: true,
        stock: true,
        pog: true,
        overview: false,
        temp: false,
        access: false
      };
      return {
        ...prev,
        [position]: {
          ...currentRules,
          [page]: !currentRules[page]
        }
      };
    });
  };

  const renderAccessCheckbox = (position: string, page: string) => {
    const isChecked = !!accessRules[position]?.[page];
    return (
      <button 
        onClick={() => toggleAccessRule(position, page)}
        className={`inline-flex items-center gap-1 text-[10px] font-bold px-2 py-1 rounded-full border uppercase tracking-wide transition-colors ${
          isChecked 
            ? "text-emerald-600 bg-emerald-50 border-emerald-200 hover:bg-emerald-100 cursor-pointer" 
            : "text-slate-400 bg-slate-50 border-slate-200 hover:bg-slate-100 cursor-pointer"
        }`}
      >
        <span className="material-symbols-outlined text-[12px]">
          {isChecked ? "check_circle" : "cancel"}
        </span> 
        {isChecked ? "Yes" : "No"}
      </button>
    );
  };

  // State Tab Temp (Temporary Review Page)
  const [tempSearchQuery, setTempSearchQuery] = useState("");
  const [tempSortBy, setTempSortBy] = useState("checker");
  const [tempSortOrder, setTempSortOrder] = useState<"asc" | "desc">("asc");
  const [expandedTempRowId, setExpandedTempRowId] = useState<string | null>(
    null,
  );
  const [isTempProceeded, setIsTempProceeded] = useState(false);
  const [isConsolidatingDb, setIsConsolidatingDb] = useState(false);
  const [consolidationSuccessMsg, setConsolidationSuccessMsg] = useState("");

  // State Tab POG
  const [pogGroupBy, setPogGroupBy] = useState("hybrid");
  const [pogSubGroupBy, setPogSubGroupBy] = useState("channel");
  const [pogExpandedRows, setPogExpandedRows] = useState({});
  const [isPogFilterOpen, setIsPogFilterOpen] = useState(true);

  const [expandedRows, setExpandedRows] = useState({});
  const CLUSTER_CONFIG = useMemo(
    () => [
      {
        key: "0-2",
        label: "0-2",
        colorHeader: "text-blue-500",
        colorCell: "text-blue-600",
        colorGrand: "text-blue-600",
        colorChild: "text-blue-500",
      },
      {
        key: "2-4",
        label: "2-4",
        colorHeader: "text-amber-500",
        colorCell: "text-amber-500",
        colorGrand: "text-amber-600",
        colorChild: "text-amber-400",
      },
      {
        key: "4-6",
        label: "4-6",
        colorHeader: "text-amber-600",
        colorCell: "text-amber-600",
        colorGrand: "text-amber-700",
        colorChild: "text-amber-500",
      },
      {
        key: "6-9",
        label: "6-9",
        colorHeader: "text-orange-500",
        colorCell: "text-orange-500",
        colorGrand: "text-orange-600",
        colorChild: "text-orange-400",
      },
      {
        key: "9-12",
        label: "9-12",
        colorHeader: "text-red-500",
        colorCell: "text-red-500",
        colorGrand: "text-red-600",
        colorChild: "text-red-400",
      },
      {
        key: ">12",
        label: ">12",
        colorHeader: "text-red-700",
        colorCell: "text-red-700",
        colorGrand: "text-red-800",
        colorChild: "text-red-600",
      },
      {
        key: "Uncategorized",
        label: "N/A",
        colorHeader: "text-slate-500",
        colorCell: "text-slate-600",
        colorGrand: "text-slate-600",
        colorChild: "text-slate-400",
      },
    ],
    [],
  );

  const ALL_CLUSTER_KEYS = useMemo(
    () => CLUSTER_CONFIG.map((c) => c.key),
    [CLUSTER_CONFIG],
  );
  const [selectedClusters, setSelectedClusters] = useState(ALL_CLUSTER_KEYS);

  const toggleRow = (name) =>
    setExpandedRows((prev) => ({ ...prev, [name]: !prev[name] }));
  const togglePogRow = (name) =>
    setPogExpandedRows((prev) => ({ ...prev, [name]: !prev[name] }));

  const teamMembers = useMemo(() => {
    if (!userData) return [];

    const myNameClean = cleanForMatch(userData.name || "");
    const isAdmin = userData.level && String(userData.level).toLowerCase().trim() === "admin";
    const isBusinessAnalyst =
      (userData.position &&
        cleanForMatch(userData.position) === "businessanalyst") ||
      cleanForMatch(userData.name) === "adityawiratama" ||
      cleanForMatch(userData.name) === "aditya" ||
      isAdmin;

    let rawList: string[] = [];

    if (employees && employees.length > 0) {
      if (isBusinessAnalyst) {
        let filteredEmployees = employees;
        if (isAdmin && userData.group) {
          const myGroupClean = cleanForMatch(userData.group);
          if (myGroupClean !== "all" && myGroupClean !== "") {
            filteredEmployees = employees.filter((e) => {
              const empGroupClean = cleanForMatch(e.group || "");
              return empGroupClean === myGroupClean || cleanForMatch(e.name) === myNameClean;
            });
          }
        }
        const allNames = filteredEmployees.map((e) => e.name).filter(Boolean);
        if (!allNames.some((n) => cleanForMatch(n) === myNameClean)) {
          allNames.unshift(userData.name);
        }
        rawList = allNames;
      } else {
        const resultList = [];
        const queue = [myNameClean];
        const visited = new Set(queue);

        while (queue.length > 0) {
          const curr = queue.shift()!;

          const matchingEmp = employees.find(
            (e) =>
              cleanForMatch(e.name) === curr ||
              (e.email && cleanForMatch(e.email) === curr),
          );
          const identifiersToMatch = new Set<string>();
          identifiersToMatch.add(curr);
          if (matchingEmp) {
            if (
              !resultList.some(
                (r) => cleanForMatch(r) === cleanForMatch(matchingEmp.name),
              )
            ) {
              resultList.push(matchingEmp.name);
            }
            if (matchingEmp.name)
              identifiersToMatch.add(cleanForMatch(matchingEmp.name));
            if (matchingEmp.email)
              identifiersToMatch.add(cleanForMatch(matchingEmp.email));
          }

          employees.forEach((emp) => {
            const empUplineClean = cleanForMatch(emp.upline);
            if (empUplineClean && identifiersToMatch.has(empUplineClean)) {
              const empClean = cleanForMatch(emp.name);
              if (empClean && !visited.has(empClean)) {
                visited.add(empClean);
                queue.push(empClean);
              }
            }
          });
        }

        const userMatchedEmp = employees.find(
          (e) =>
            cleanForMatch(e.name) === myNameClean ||
            (e.email && cleanForMatch(e.email) === myNameClean),
        );
        const finalUserName = userMatchedEmp
          ? userMatchedEmp.name
          : userData.name;

        if (
          !resultList.some(
            (r) => cleanForMatch(r) === cleanForMatch(finalUserName),
          )
        ) {
          resultList.unshift(finalUserName);
        }

        rawList = resultList;
      }
    } else if (
      computedTeamProfiles &&
      Object.keys(computedTeamProfiles).length > 0
    ) {
      let realRootName = userData.name;
      const foundProfileKey = Object.keys(computedTeamProfiles).find(
        (k) =>
          cleanForMatch(k) === myNameClean ||
          (computedTeamProfiles[k]?.email &&
            cleanForMatch(computedTeamProfiles[k].email) === myNameClean),
      );
      if (foundProfileKey) {
        realRootName = foundProfileKey;
      }
      const cleanRealRoot = cleanForMatch(realRootName);

      const depths = buildDepthMap(realRootName, computedTeamProfiles);
      const maxDepth = 5;

      const filtered = Object.keys(computedTeamProfiles).filter((name) => {
        const d = depths[cleanForMatch(name)] ?? 99;
        return d <= maxDepth;
      });

      const withoutMe = filtered.filter(
        (n) =>
          cleanForMatch(n) !== cleanRealRoot &&
          cleanForMatch(n) !== myNameClean,
      );
      const result = [realRootName, ...withoutMe];
      if (userData.name !== realRootName && !result.includes(userData.name)) {
        result.unshift(userData.name);
      }
      rawList = result;
    } else {
      let uniquePics = [...(userData.subordinates || [])];

      let added = true;
      while (added) {
        added = false;
        kiosks.forEach((k) => {
          const uplineClean = cleanForMatch(k.upline || "");
          const picStr = String(k.pic || "").trim();
          const picClean = cleanForMatch(picStr);

          const isDirectMatch =
            !isBusinessAnalyst &&
            uplineClean !== "" &&
            myNameClean !== "" &&
            (uplineClean.includes(myNameClean) ||
              myNameClean.includes(uplineClean));
          const isTransitiveMatch =
            uplineClean !== "" &&
            uniquePics.some((m) => cleanForMatch(m) === uplineClean);

          if (isDirectMatch || isTransitiveMatch) {
            if (
              picClean !== "" &&
              picClean !== myNameClean &&
              !picClean.includes(myNameClean)
            ) {
              if (
                !uniquePics.some(
                  (existing) => cleanForMatch(existing) === picClean,
                )
              ) {
                uniquePics.push(picStr);
                added = true;
              }
            }
          }
        });
      }

      // FALLBACK AMAN: Jika hasil filter tim sangat sedikit, atau user menggunakan akun demo/tidak terpetakan,
      // maka masukkan semua PIC yang ada di daftar Kiosks agar data Stock Summary/POG tetap tampil dan dapat dianalisis.
      if (uniquePics.length <= (userData.subordinates?.length || 0)) {
        kiosks.forEach((k) => {
          const picStr = String(k.pic || "").trim();
          const picClean = cleanForMatch(picStr);
          if (picStr && picClean !== "" && picClean !== myNameClean) {
            if (
              !uniquePics.some(
                (existing) => cleanForMatch(existing) === picClean,
              )
            ) {
              uniquePics.push(picStr);
            }
          }
        });
      }

      if (
        !uniquePics.some((existing) => cleanForMatch(existing) === myNameClean)
      ) {
        uniquePics.unshift(userData.name);
      }
      rawList = uniquePics;
    }

    // Apply strict filtering for Level 5: if the current user level is 5, exclude other Level 5 users.
    const isMyLevel5 = (() => {
      // Top-level roles like Business Analyst (rank 1, level 5) must never be filtered or restricted.
      if (isBusinessAnalyst || getPositionRank(userData.position) === 1) {
        return false;
      }
      const cleanName = cleanForMatch(userData.name);
      if (
        userData.position &&
        cleanForMatch(userData.position) === "businesssolution"
      )
        return true;
      if (getPositionRank(userData.position) === 5) {
        return true;
      }
      if (employees && employees.length > 0) {
        const emp = employees.find((e) => cleanForMatch(e.name) === cleanName);
        if (emp) {
          if (cleanForMatch(emp.position) === "businesssolution") return true;
          if (getPositionRank(emp.position) === 5) return true;
        }
      }
      return false;
    })();

    if (isMyLevel5) {
      return rawList.filter((memberName) => {
        const cleanName = cleanForMatch(memberName);
        if (cleanName === myNameClean) return true; // Keep ourselves always

        // Exclude if level is 5 or position is Business Solution
        if (employees && employees.length > 0) {
          const emp = employees.find(
            (e) => cleanForMatch(e.name) === cleanName,
          );
          if (emp) {
            if (
              emp.level !== undefined &&
              emp.level !== null &&
              String(emp.level).trim() === "5"
            )
              return false;
            if (cleanForMatch(emp.position) === "businesssolution")
              return false;
            if (getPositionRank(emp.position) === 5) return false;
          }
        }
        if (computedTeamProfiles) {
          const foundKey = Object.keys(computedTeamProfiles).find(
            (k) => cleanForMatch(k) === cleanName,
          );
          if (foundKey) {
            const prof = computedTeamProfiles[foundKey];
            if (prof?.level !== undefined && String(prof.level).trim() === "5")
              return false;
            if (
              prof?.position &&
              cleanForMatch(prof.position) === "businesssolution"
            )
              return false;
          }
        }
        return true;
      });
    }

    // Deduplicate the list to prevent key collision React errors
    const seen = new Set();
    const uniqueList = [];
    for (const name of rawList) {
      const clean = cleanForMatch(name);
      if (!seen.has(clean)) {
        seen.add(clean);
        uniqueList.push(name);
      }
    }
    return uniqueList;
  }, [
    kiosks,
    userData,
    userData?.name,
    userData?.subordinates,
    userData?.position,
    computedTeamProfiles,
    employees,
  ]);

  const myActiveSubordinates = useMemo(() => {
    return activeEmployees.filter((emp: any) =>
      teamMembers.some((m: string) => cleanForMatch(m) === cleanForMatch(emp.name))
    );
  }, [activeEmployees, teamMembers]);

  const normalizeName = useCallback(
    (nameStr: string) => {
      const trimmed = nameStr.trim();
      if (!trimmed || trimmed === "Unknown") return "Unknown";
      const clean = cleanForMatch(trimmed);

      if (clean === "agusherdianto" || clean.includes("agusherdianto")) {
        return "AGUS HERDIANTO";
      }

      const found = teamMembers.find((t) => cleanForMatch(t) === clean);
      if (found) return found.trim();

      return trimmed;
    },
    [teamMembers],
  );

  const [teamPositions, setTeamPositions] = useState<Record<string, string>>(
    {},
  );
  const [teamAreas, setTeamAreas] = useState<Record<string, string>>({});
  const [teamProvinces, setTeamProvinces] = useState<Record<string, string>>(
    {},
  );
  const [teamSubordinates, setTeamSubordinates] = useState<
    Record<string, string[]>
  >({});
  const [teamUpLines, setTeamUpLines] = useState<Record<string, string>>({});
  const [teamLevels, setTeamLevels] = useState<Record<string, number>>({});

  useEffect(() => {
    if (!userData) return;
    if (employeesRefreshKey === 0) {
      return;
    }
    setIsEmployeesLoading(true);
    customFetch(`${SCRIPT_URL}?action=getEmployees`)
      .then((res) => res.json())
      .then((res) => {
        if (
          res.status === "success" &&
          Array.isArray(res.data) &&
          res.data.length > 0
        ) {
          setEmployees(res.data);
        } else {
          console.warn(
            "getEmployees returned non-success or empty, falling back to offline employees.",
          );
          setEmployees(OFFLINE_EMPLOYEES);
        }
      })
      .catch((err) => {
        console.warn(
          "Error fetching getEmployees, using offline fallbacks:",
          err,
        );
        setEmployees(OFFLINE_EMPLOYEES);
      })
      .finally(() => {
        setIsEmployeesLoading(false);
      });
  }, [userData, employeesRefreshKey]);

  useEffect(() => {
    if (employees && employees.length > 0) {
      const positions: Record<string, string> = {};
      const areas: Record<string, string> = {};
      const provinces: Record<string, string> = {};
      const uplines: Record<string, string> = {};
      const subordinates: Record<string, string[]> = {};
      const levels: Record<string, number> = {};

      employees.forEach((emp) => {
        const name = emp.name;
        positions[name] = normalizePosition(emp.position);
        areas[name] = String(emp.area || "-").trim();
        provinces[name] = String(emp.province || "-").trim();
        uplines[name] = String(emp.upline || "").trim();
        if (
          emp.level !== undefined &&
          emp.level !== null &&
          String(emp.level).trim() !== ""
        ) {
          const parsed = parseLevelStr(emp.level);
          if (!isNaN(parsed)) {
            levels[name] = parsed;
          } else {
            const rank = getPositionRank(emp.position);
            levels[name] =
              rank === 1
                ? 5
                : rank === 2
                  ? 4
                  : rank === 3
                    ? 3
                    : rank === 4
                      ? 2
                      : rank === 5
                        ? 1
                        : 0;
          }
        } else {
          const rank = getPositionRank(emp.position);
          levels[name] =
            rank === 1
              ? 5
              : rank === 2
                ? 4
                : rank === 3
                  ? 3
                  : rank === 4
                    ? 2
                    : rank === 5
                      ? 1
                      : 0;
        }
      });

      employees.forEach((emp) => {
        const name = emp.name;
        const directSubs: string[] = [];
        employees.forEach((item) => {
          if (item.name !== name && item.upline) {
            if (cleanForMatch(item.upline) === cleanForMatch(name)) {
              directSubs.push(item.name);
            }
          }
        });
        subordinates[name] = directSubs;
      });

      setTeamPositions((prev) => ({ ...prev, ...positions }));
      setTeamAreas((prev) => ({ ...prev, ...areas }));
      setTeamProvinces((prev) => ({ ...prev, ...provinces }));
      setTeamUpLines((prev) => ({ ...prev, ...uplines }));
      setTeamSubordinates((prev) => ({ ...prev, ...subordinates }));
      setTeamLevels((prev) => ({ ...prev, ...levels }));
    }
  }, [employees]);

  useEffect(() => {
    if (
      userData &&
      computedTeamProfiles &&
      Object.keys(computedTeamProfiles).length > 0
    ) {
      const positions: Record<string, string> = {};
      const areas: Record<string, string> = {};
      const provinces: Record<string, string> = {};
      const uplines: Record<string, string> = {};
      const subordinates: Record<string, string[]> = {};
      const levels: Record<string, number> = {};

      Object.entries(computedTeamProfiles).forEach(
        ([name, p]: [string, any]) => {
          positions[name] = normalizePosition(p.position);
          areas[name] = String(p.area || "-").trim();
          provinces[name] = String(p.province || "-").trim();
          uplines[name] = String(p.upline || "").trim();
          if (
            p.level !== undefined &&
            p.level !== null &&
            String(p.level).trim() !== ""
          ) {
            const parsed = parseLevelStr(p.level);
            if (!isNaN(parsed)) {
              levels[name] = parsed;
            } else {
              const rank = getPositionRank(p.position);
              levels[name] =
                rank === 1
                  ? 5
                  : rank === 2
                    ? 4
                    : rank === 3
                      ? 3
                      : rank === 4
                        ? 2
                        : rank === 5
                          ? 1
                          : 0;
            }
          } else {
            const rank = getPositionRank(p.position);
            levels[name] =
              rank === 1
                ? 5
                : rank === 2
                  ? 4
                  : rank === 3
                    ? 3
                    : rank === 4
                      ? 2
                      : rank === 5
                        ? 1
                        : 0;
          }
        },
      );

      Object.keys(computedTeamProfiles).forEach((name) => {
        const directSubs: string[] = [];
        Object.entries(computedTeamProfiles).forEach(
          ([otherName, p]: [string, any]) => {
            if (otherName !== name && p.upline) {
              const cleanUp = cleanForMatch(p.upline);
              const cleanMy = cleanForMatch(name);
              if (cleanUp === cleanMy) {
                directSubs.push(otherName);
              }
            }
          },
        );
        subordinates[name] = directSubs;
      });

      setTeamPositions((prev) => ({ ...prev, ...positions }));
      setTeamAreas((prev) => ({ ...prev, ...areas }));
      setTeamProvinces((prev) => ({ ...prev, ...provinces }));
      setTeamUpLines((prev) => ({ ...prev, ...uplines }));
      setTeamSubordinates((prev) => ({ ...prev, ...subordinates }));
      setTeamLevels((prev) => ({ ...prev, ...levels }));
    }
  }, [userData]);

  useEffect(() => {
    const isAdmin = userData.level && String(userData.level).toLowerCase().trim() === "admin";
    const isBusinessAnalyst =
      (userData.position &&
        cleanForMatch(userData.position) === "businessanalyst") ||
      cleanForMatch(userData.name) === "adityawiratama" ||
      cleanForMatch(userData.name) === "aditya" ||
      isAdmin;
    if (isBusinessAnalyst) {
      const others = teamMembers.filter(
        (m) => cleanForMatch(m) !== cleanForMatch(userData.name),
      );
      setTeamSubordinates((prev) => {
        if (JSON.stringify(prev[userData.name]) === JSON.stringify(others))
          return prev;
        return {
          ...prev,
          [userData.name]: others,
        };
      });
      setTeamPositions((prev) => {
        const targetPos = isAdmin ? (userData.position || "Admin") : "Business Analyst";
        if (prev[userData.name] === targetPos) return prev;
        return {
          ...prev,
          [userData.name]: targetPos,
        };
      });
    }
  }, [userData.name, userData.position, teamMembers, teamPositions]);

  const greeting = useMemo(() => {
    const hour = new Date().getHours();
    if (hour >= 4 && hour < 11) {
      return {
        text: "Selamat Pagi",
        imageUrl:
          "https://lh3.googleusercontent.com/d/1AzKb-75MaU9hppqSdy2rS93t0tAPGkGi",
        color: "text-amber-300",
      };
    }
    if (hour >= 11 && hour < 15) {
      return {
        text: "Selamat Siang",
        imageUrl:
          "https://lh3.googleusercontent.com/d/1ZpNkT7R57FppIpyPuTt2w9QtJdIwwuRp",
        color: "text-yellow-300",
      };
    }
    if (hour >= 15 && hour < 19) {
      return {
        text: "Selamat Sore",
        imageUrl:
          "https://lh3.googleusercontent.com/d/12RsJXxDrH7aIAph0AJubB3i4w0gmkxcL",
        color: "text-orange-400",
      };
    }
    return {
      text: "Selamat Malam",
      imageUrl:
        "https://lh3.googleusercontent.com/d/1wzqPdQ5jvw7fOF2X76kM56l9l-4mUcLx",
      color: "text-indigo-200",
    };
  }, []);

  const workingCategoryMap = useMemo(() => {
    const map: Record<string, string> = {};
    if (rawWorkingData && Array.isArray(rawWorkingData)) {
      rawWorkingData.forEach((d) => {
        const kClean = cleanForMatch(d.kiosk);
        const rawCat =
          d.category ||
          d.Category ||
          d.CATEGORY ||
          d.kategori ||
          d.Kategori ||
          d.klasifikasi ||
          d.cat ||
          d.Cat ||
          "";
        const cat = getValidCategory(rawCat, undefined, d.kiosk);
        if (kClean && cat && cat !== "Uncategorized") {
          map[kClean] = cat;
        }
      });
    }
    if (kiosks && Array.isArray(kiosks)) {
      kiosks.forEach((k) => {
        const kClean = cleanForMatch(k.name);
        const rawCat =
          k.category ||
          k.Category ||
          k.CATEGORY ||
          k.kategori ||
          k.Kategori ||
          k.klasifikasi ||
          k.cat ||
          k.Cat ||
          "";
        const cat = getValidCategory(rawCat, undefined, k.name);
        if (kClean && cat && cat !== "Uncategorized") {
          if (!map[kClean]) map[kClean] = cat;
        }
      });
    }
    return map;
  }, [rawWorkingData, kiosks]);

  const allKiosks = useMemo(() => {
    const map = new Map<string, any>();
    (kiosks || []).forEach((k) => {
      const cleanName = cleanForMatch(k.name);
      if (cleanName) {
        const dbCat = workingCategoryMap[cleanName];
        const validCat = getValidCategory(
          dbCat || k.category,
          k.category,
          k.name,
          workingCategoryMap,
        );
        map.set(cleanName, {
          ...k,
          category: validCat,
        });
      }
    });

    if (rawWorkingData && Array.isArray(rawWorkingData)) {
      rawWorkingData.forEach((d) => {
        const kName = String(d.kiosk || "").trim();
        const cleanName = cleanForMatch(kName);
        if (cleanName) {
          const dbCat = workingCategoryMap[cleanName];
          const validCat = getValidCategory(
            d.category || dbCat,
            undefined,
            kName,
            workingCategoryMap,
          );
          if (!map.has(cleanName)) {
            map.set(cleanName, {
              id: `auto_${cleanName}`,
              name: kName,
              category: validCat,
              pic: d.pic || d.user || "",
              upline: "",
              province: d.area || "",
              area: d.area || "-",
              group: "",
            });
          } else {
            const existing = map.get(cleanName);
            if (dbCat && existing.category !== dbCat) {
              map.set(cleanName, {
                ...existing,
                category: dbCat,
              });
            }
          }
        }
      });
    }

    return Array.from(map.values());
  }, [kiosks, rawWorkingData, workingCategoryMap]);

  const fetchWorkingData = async (
    presetData?: any[],
    presetDrSales?: any[],
  ) => {
    setIsFetchingData(true);
    setIsSyncing(true);
    try {
      let combinedData = [];
      let combinedDrSales = [];
      let success = false;

      if (presetData && presetDrSales) {
        combinedData = presetData;
        combinedDrSales = presetDrSales;
        success = true;
      } else {
        try {
          const [resp, respDr] = await Promise.all([
            customFetch(
              `${SCRIPT_URL}?action=getWorkingData&user=${encodeURIComponent(userData.name)}`,
            ),
            customFetch(
              `${SCRIPT_URL}?action=getDrSalesData&user=${encodeURIComponent(userData.name)}`,
            ),
          ]);
          const [res, resDr] = await Promise.all([resp.json(), respDr.json()]);

          if (res.status === "success") {
            combinedData = (res.data || []).map((item) => {
              const cropVal =
                item.crops ||
                item.Crops ||
                item.crop ||
                item.Crop ||
                item.CROP ||
                item.CROPS ||
                "Uncategorized Crops";
              const areaVal = item.area || item.Area || item.AREA;
              const catVal =
                item.category ||
                item.Category ||
                item.CATEGORY ||
                item.kategori ||
                item.Kategori ||
                item.klasifikasi ||
                item.cat ||
                item.Cat ||
                "";
              return { ...item, crops: cropVal, area: areaVal, category: catVal };
            });
            success = true;
          }
          if (resDr.status === "success") {
            combinedDrSales = resDr.data || [];
          }
        } catch (apiErr) {
          console.warn(
            "API working/sales data load failed, falling back offline:",
            apiErr,
          );
        }
      }

      if (!success || combinedData.length === 0) {
        combinedData = OFFLINE_WORKING_DATA;
        combinedDrSales = OFFLINE_DR_SALES;
      }

      setDrSalesData(combinedDrSales);
      setRawWorkingData(combinedData);

      const groupedMap: Record<string, any> = {};

      const parseTimestamp = (ts) => {
        if (!ts) return 0;
        if (typeof ts === "string") {
          if (ts.includes("/")) {
            const parts = ts.split(/[\s/:]+/);
            if (parts.length >= 3) {
              return new Date(
                `${parts[2]}-${parts[1]}-${parts[0]}T${parts[3] || "00"}:${parts[4] || "00"}:${parts[5] || "00"}`,
              ).getTime();
            }
          }
          const d = new Date(ts).getTime();
          return isNaN(d) ? 0 : d;
        }
        const dt = new Date(ts).getTime();
        return isNaN(dt) ? 0 : dt;
      };

      // 1. Lakukan Grouping (Gabungkan LOT & Hybrid yang sama)
      combinedData.forEach((d) => {
        const k = cleanForMatch(d.kiosk);
        const l = cleanForMatch(d.lot);
        const h = cleanForMatch(d.hybrid);
        const u = cleanForMatch(d.user || d.pic || "");
        const key = `${k}_${l}_${h}_${u}`;
        if (!groupedMap[key]) {
          groupedMap[key] = { ...d };
        } else {
          // Ambil data dari timestamp terbaru (TIDAK ADA AKUMULASI QTY)
          const timeExisting = parseTimestamp(groupedMap[key].timestamp);
          const timeNew = parseTimestamp(d.timestamp);
          if (timeNew > timeExisting) {
            groupedMap[key] = { ...d };
          }
        }
      });

      // 2. Filter dan format
      const monthsKeys = [
        "jan",
        "feb",
        "mar",
        "apr",
        "mei",
        "jun",
        "jul",
        "ags",
        "sep",
        "okt",
        "nov",
        "des",
      ];
      const currentMonthIdx = new Date().getMonth();
      const currentMonthKey = monthsKeys[currentMonthIdx];
      const prevMonthIdx = currentMonthIdx === 0 ? 11 : currentMonthIdx - 1;
      const prevMonthKey = monthsKeys[prevMonthIdx];

      const enrichedData = Object.values(groupedMap)
        .filter((d) => String(d.condition).trim().toLowerCase() !== "habis") // Filter: Sembunyikan yang habis
        .map((d, index) => {
          const currVal =
            d[currentMonthKey] !== undefined
              ? safeParseFloat(d[currentMonthKey])
              : safeParseFloat(d.stock);
          const prevVal =
            d[prevMonthKey] !== undefined ? safeParseFloat(d[prevMonthKey]) : 0;

          let finalCondition = d.condition || "tetap";
          if (prevVal > 0) {
            if (currVal > prevVal) finalCondition = "bertambah";
            else if (currVal < prevVal) finalCondition = "berkurang";
            else finalCondition = "tetap";
          } else {
            if (currVal > 0) {
              finalCondition =
                d.condition === "new" || d.condition === "baru" || d.isNew
                  ? "new"
                  : "bertambah";
            } else {
              finalCondition = "tetap";
            }
          }

          return {
            ...d,
            stock: currVal, // Use current month's stock instead of the shipping stock!
            id: d.id || `db_${index}_${d.lot}`, // Fallback ID jika tidak ada
            originalStock: d.stock,
            originalUser: d.user || "",
            condition: finalCondition,
            isNew: false,
          };
        });

      setWorkingData(enrichedData);
      setDeletedItems([]);
    } catch (e) {
      console.warn("Error processing working data", e);
    } finally {
      setIsFetchingData(false);
      setIsSyncing(false);
    }
  };

  useEffect(() => {
    const fetchChannels = async () => {
      if (channelsRefreshKey === 0) {
        if (kiosks.length > 0) {
          if (isBusinessAnalyst) {
            setSelectedKiosk("Not Applicable for Business Analyst");
          } else {
            const myKiosksData = kiosks.filter((k) =>
              matchNames(k.pic, userData.name),
            );
            if (myKiosksData.length > 0) setSelectedKiosk(myKiosksData[0].name);
            else {
              if (kiosks.length > 0) setSelectedKiosk(kiosks[0].name);
              else setSelectedKiosk("No Channel Assigned");
            }
          }
        }
        setIsChannelsLoading(false);
        return;
      }
      setIsChannelsLoading(true);
      try {
        let fetchedKiosks = [];
        let success = false;
        try {
          const resp = await customFetch(
            `${SCRIPT_URL}?action=getChannels&user=${encodeURIComponent(userData.name)}`,
          );
          const res = await resp.json();
          if (
            res.status === "success" &&
            Array.isArray(res.data) &&
            res.data.length > 0
          ) {
            fetchedKiosks = res.data;
            success = true;
          }
        } catch (apiErr) {
          console.warn(
            "API channels load failed, falling back offline:",
            apiErr,
          );
        }

        if (!success) {
          fetchedKiosks = OFFLINE_KIOSKS;
        }

        setKiosks(fetchedKiosks);

        if (channelsRefreshKey === 0) {
          if (isBusinessAnalyst) {
            setSelectedKiosk("Not Applicable for Business Analyst");
          } else {
            const myKiosksData = fetchedKiosks.filter((k) =>
              matchNames(k.pic, userData.name),
            );
            if (myKiosksData.length > 0) setSelectedKiosk(myKiosksData[0].name);
            else {
              if (fetchedKiosks.length > 0)
                setSelectedKiosk(fetchedKiosks[0].name);
              else setSelectedKiosk("No Channel Assigned");
            }
          }
        }
      } catch (e) {
        console.warn("Error loading channels", e);
        setKiosks(OFFLINE_KIOSKS);
        if (channelsRefreshKey === 0) {
          if (OFFLINE_KIOSKS.length > 0)
            setSelectedKiosk(OFFLINE_KIOSKS[0].name);
        }
      } finally {
        setIsChannelsLoading(false);
      }
    };
    fetchChannels();
  }, [userData.name, userData.position, channelsRefreshKey, isBusinessAnalyst]);

  useEffect(() => {
    if (!userData) return;

    const loadInitialData = async () => {
      setIsEmployeesLoading(true);
      setIsChannelsLoading(true);
      setIsFetchingData(true);
      setIsSyncing(true);

      try {
        const resp = await customFetch(
          `${SCRIPT_URL}?action=getInitialData&user=${encodeURIComponent(userData.name)}`,
        );
        const res = await resp.json();

        if (res.status === "success" && res.data) {
          // 0. Update User Data if server profile returned
          if (res.data.profile) {
            setUserData((prev) => {
              if (!prev) return res.data.profile;
              const updated = { ...prev, ...res.data.profile };
              
              const isAditya =
                cleanForMatch(updated.name) === "adityawiratama" ||
                cleanForMatch(updated.name) === "aditya" ||
                cleanForMatch(updated.user || "") === "aditya" ||
                cleanForMatch(updated.user || "") === "adityawiratama";
              if (isAditya) {
                updated.position = "Business Analyst";
              } else {
                updated.position = normalizePosition(updated.position);
              }
              
              localStorage.setItem("radar_user_session", JSON.stringify(updated));
              return updated;
            });
          }
          
          // 1. Employees
          if (res.data.employees && res.data.employees.length > 0) {
            setEmployees(res.data.employees);
          } else {
            setEmployees(OFFLINE_EMPLOYEES);
          }

          // 2. Channels
          const fetchedKiosks =
            res.data.channels && res.data.channels.length > 0
              ? res.data.channels
              : OFFLINE_KIOSKS;
          setKiosks(fetchedKiosks);

          if (channelsRefreshKey === 0) {
            if (isBusinessAnalyst) {
              setSelectedKiosk("Not Applicable for Business Analyst");
            } else {
              const myKiosksData = fetchedKiosks.filter((k) =>
                matchNames(k.pic, userData.name),
              );
              if (myKiosksData.length > 0)
                setSelectedKiosk(myKiosksData[0].name);
              else {
                if (fetchedKiosks.length > 0)
                  setSelectedKiosk(fetchedKiosks[0].name);
                else setSelectedKiosk("No Channel Assigned");
              }
            }
          }

          // 3. Working & Sales Data
          const rawWorking =
            res.data.workingData && res.data.workingData.length > 0
              ? res.data.workingData
              : OFFLINE_WORKING_DATA;
          console.log("rawWorking", rawWorking);
          const mappedWorking = rawWorking.map((item) => {
            const cropVal =
              item.crops ||
              item.Crops ||
              item.crop ||
              item.Crop ||
              item.CROP ||
              item.CROPS ||
              "Uncategorized Crops";
            const areaVal = item.area || item.Area || item.AREA;
            const catVal =
              item.category ||
              item.Category ||
              item.CATEGORY ||
              item.kategori ||
              item.Kategori ||
              item.klasifikasi ||
              item.cat ||
              item.Cat ||
              "";
            return { ...item, crops: cropVal, area: areaVal, category: catVal };
          });

          const drSales =
            res.data.drSalesData && res.data.drSalesData.length > 0
              ? res.data.drSalesData
              : OFFLINE_DR_SALES;

          // Enriches and updates states inside fetchWorkingData
          await fetchWorkingData(mappedWorking, drSales);

          // 4. Access Rules
          if (res.data.accessRules && Object.keys(res.data.accessRules).length > 0) {
            setAccessRules(res.data.accessRules);
            try {
              localStorage.setItem('appAccessRules', JSON.stringify(res.data.accessRules));
            } catch (e) {
              console.error('Failed to save appAccessRules to localStorage', e);
            }
          }
        } else {
          console.warn(
            "getInitialData not supported or empty, doing live individual parallel fetching.",
          );
          await loadIndividualDataFallback();
        }
      } catch (err) {
        console.warn(
          "Failed unified initial data fetch, falling back to parallel individual live fetches:",
          err,
        );
        await loadIndividualDataFallback();
      } finally {
        setIsEmployeesLoading(false);
        setIsChannelsLoading(false);
        setIsFetchingData(false);
        setIsSyncing(false);
      }
    };

    const loadIndividualDataFallback = async () => {
      try {
        const [respEmp, respChan, respWork, respDr, respAccess] = await Promise.all([
          customFetch(`${SCRIPT_URL}?action=getEmployees`),
          customFetch(
            `${SCRIPT_URL}?action=getChannels&user=${encodeURIComponent(userData.name)}`,
          ),
          customFetch(
            `${SCRIPT_URL}?action=getWorkingData&user=${encodeURIComponent(userData.name)}`,
          ),
          customFetch(
            `${SCRIPT_URL}?action=getDrSalesData&user=${encodeURIComponent(userData.name)}`,
          ),
          customFetch(`${SCRIPT_URL}?action=getAccessRules`),
        ]);
        const [resEmp, resChan, resWork, resDr, resAccess] = await Promise.all([
          respEmp.json(),
          respChan.json(),
          respWork.json(),
          respDr.json(),
          respAccess.json(),
        ]);

        if (resEmp.status === "success") {
          setEmployees(resEmp.data || []);
        } else {
          setEmployees(OFFLINE_EMPLOYEES);
        }

        const fetchedKiosks =
          resChan.status === "success" ? resChan.data || [] : OFFLINE_KIOSKS;
        setKiosks(fetchedKiosks);

        if (channelsRefreshKey === 0) {
          if (isBusinessAnalyst) {
            setSelectedKiosk("Not Applicable for Business Analyst");
          } else {
            const myKiosksData = fetchedKiosks.filter((k) =>
              matchNames(k.pic, userData.name),
            );
            if (myKiosksData.length > 0) setSelectedKiosk(myKiosksData[0].name);
            else {
              if (fetchedKiosks.length > 0)
                setSelectedKiosk(fetchedKiosks[0].name);
              else setSelectedKiosk("No Channel Assigned");
            }
          }
        }

        let mappedWorking = OFFLINE_WORKING_DATA;
        if (resWork.status === "success") {
          mappedWorking = (resWork.data || []).map((item) => {
            const cropVal =
              item.crops ||
              item.Crops ||
              item.crop ||
              item.Crop ||
              item.CROP ||
              item.CROPS ||
              "Uncategorized Crops";
            const areaVal = item.area || item.Area || item.AREA;
            const catVal =
              item.category ||
              item.Category ||
              item.CATEGORY ||
              item.kategori ||
              item.Kategori ||
              item.klasifikasi ||
              item.cat ||
              item.Cat ||
              "";
            return { ...item, crops: cropVal, area: areaVal, category: catVal };
          });
        }
        const drSales =
          resDr.status === "success" ? resDr.data || [] : OFFLINE_DR_SALES;

        await fetchWorkingData(mappedWorking, drSales);

        if (resAccess.status === "success" && resAccess.data && Object.keys(resAccess.data).length > 0) {
          setAccessRules(resAccess.data);
          try {
            localStorage.setItem('appAccessRules', JSON.stringify(resAccess.data));
          } catch (e) {
            console.error('Failed to save appAccessRules to localStorage', e);
          }
        }
      } catch (fallbackErr) {
        console.warn("Fallback live fetches also failed:", fallbackErr);
        setEmployees(OFFLINE_EMPLOYEES);
        setKiosks(OFFLINE_KIOSKS);
        await fetchWorkingData(OFFLINE_WORKING_DATA, OFFLINE_DR_SALES);
      }
    };

    loadInitialData();
  }, [userData.name]);

  useEffect(() => {
    if (teamMembers.length > 1) {
      if (
        !mappingPic ||
        mappingPic === "" ||
        (mappingPic !== "ALL_TEAM" &&
          !teamMembers.some(
            (m) => cleanForMatch(m) === cleanForMatch(mappingPic),
          ))
      ) {
        setMappingPic("ALL_TEAM");
      }
    } else {
      setMappingPic(userData.name || "");
    }
  }, [userData.name, teamMembers, mappingPic]);

  useEffect(() => {
    setIsLotNotFound(false);
    if (lotNo.length < 3) {
      setLotIntel(null);
      setManualHybrid("");
      setManualHybridCustom("");
      setManualCrop("Field Corn");
      setManualDrDate("");
      setManualExpDate("");
      return;
    }
    const checkLot = async () => {
      setIsLotChecking(true);
      try {
        const resp = await customFetch(
          `${SCRIPT_URL}?action=getLotInfo&lot=${encodeURIComponent(lotNo)}&kiosk=${encodeURIComponent(selectedKiosk)}`,
        );
        const res = await resp.json();
        if (res.status === "success" && res.data) {
          setLotIntel(res.data);
          setIsLotNotFound(false);
          setIsLotChecking(false);
          return;
        }
      } catch (e) {
        // Fallback to local
      }

      // Local fallback: search working sheet first
      const cleanTarget = lotNo.trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
      const cleanKiosk = cleanForMatch(selectedKiosk);

      const foundInWorkingForKiosk = Array.isArray(rawWorkingData)
        ? rawWorkingData.find(
            (w) =>
              w.lot &&
              (String(w.lot).trim().toUpperCase() === lotNo.trim().toUpperCase() ||
                String(w.lot).toUpperCase().replace(/[^A-Z0-9]/g, "") === cleanTarget) &&
              cleanForMatch(w.kiosk) === cleanKiosk,
          )
        : null;

      const foundInWorkingAny = foundInWorkingForKiosk || (Array.isArray(rawWorkingData)
        ? rawWorkingData.find(
            (w) =>
              w.lot &&
              (String(w.lot).trim().toUpperCase() === lotNo.trim().toUpperCase() ||
                String(w.lot).toUpperCase().replace(/[^A-Z0-9]/g, "") === cleanTarget),
          )
        : null);

      if (foundInWorkingAny) {
        const rawWExp = foundInWorkingAny.expired || foundInWorkingAny.expDate || "N/A";
        const finalWExp = (rawWExp === "N/A" || rawWExp === "-") ? extractExpFromLot(lotNo) : rawWExp;
        setLotIntel({
          desc: foundInWorkingAny.hybrid || "",
          crops: foundInWorkingAny.crops || "",
          drDate: foundInWorkingAny.drDate || "N/A",
          expDate: finalWExp,
          aging: foundInWorkingAny.aging || "-",
          channel: foundInWorkingAny.kiosk || "",
        });
        setIsLotNotFound(false);
        setIsLotChecking(false);
        return;
      }

      const foundInDrForKiosk = Array.isArray(drSalesData)
        ? drSalesData.find(
            (d) =>
              d.lot &&
              (String(d.lot).trim().toUpperCase() === lotNo.trim().toUpperCase() ||
                String(d.lot).toUpperCase().replace(/[^A-Z0-9]/g, "") === cleanTarget) &&
              cleanForMatch(d.channel) === cleanKiosk,
          )
        : null;

      const foundInDrAny = foundInDrForKiosk || (Array.isArray(drSalesData)
        ? drSalesData.find(
            (d) =>
              d.lot &&
              (String(d.lot).trim().toUpperCase() === lotNo.trim().toUpperCase() ||
                String(d.lot).toUpperCase().replace(/[^A-Z0-9]/g, "") === cleanTarget),
          )
        : null);

      if (foundInDrAny) {
        const rawDrExp = foundInDrAny.expired || foundInDrAny.expDate || "N/A";
        const finalDrExp = (rawDrExp === "N/A" || rawDrExp === "-") ? extractExpFromLot(lotNo) : rawDrExp;
        setLotIntel({
          desc: foundInDrAny.hybrid || foundInDrAny.desc || "",
          crops: foundInDrAny.crops || "",
          drDate: foundInDrAny.drDate || "N/A",
          expDate: finalDrExp,
          aging: foundInDrAny.aging || "-",
          channel: foundInDrAny.channel || "",
        });
        setIsLotNotFound(false);
        setIsLotChecking(false);
        return;
      }

      setLotIntel(null);
      setIsLotNotFound(true);
      setIsLotChecking(false);
    };
    const timer = setTimeout(checkLot, 600);
    return () => clearTimeout(timer);
  }, [lotNo, selectedKiosk, drSalesData, rawWorkingData]);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target))
        setIsDropdownOpen(false);
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const getCurrentTimestamp = () => {
    const pad = (n: number) => String(n).padStart(2, "0");
    const now = new Date();
    return `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  };

  const formatDateToAppFormat = (dateStr: string) => {
    if (!dateStr || dateStr === "N/A" || dateStr === "-") return "N/A";
    const str = String(dateStr).trim();
    if (!str) return "N/A";

    // Handle HTML input date: YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
      const [y, m, d] = str.split("-");
      const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
      const mIdx = parseInt(m, 10) - 1;
      return `${d.padStart(2, "0")}/${months[mIdx] || "JAN"}/${y.substring(2)}`;
    }

    // Handle DD/MM/YYYY or DD-MM-YYYY or DD-MMM-YY
    const parts = str.split(/[\s/\-\.:]+/);
    if (parts.length >= 3) {
      const day = parseInt(parts[0], 10);
      const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
      let mIdx = parseInt(parts[1], 10) - 1;
      if (isNaN(mIdx)) {
        const lower = parts[1].toLowerCase();
        mIdx = months.findIndex((m) => lower.startsWith(m.toLowerCase()));
      }
      let y = parseInt(parts[2], 10);
      if (y < 100) y += 2000;
      if (!isNaN(day) && mIdx >= 0 && mIdx < 12 && !isNaN(y)) {
        return `${String(day).padStart(2, "0")}/${months[mIdx]}/${String(y).substring(2)}`;
      }
    }

    const date = new Date(str);
    if (!isNaN(date.getTime())) {
      const day = String(date.getDate()).padStart(2, "0");
      const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
      const year = String(date.getFullYear()).substring(2);
      return `${day}/${months[date.getMonth()]}/${year}`;
    }

    return str || "N/A";
  };

  const calculateAgingInMonths = (startStr: string) => {
    if (!startStr) return "-";
    const start = new Date(startStr);
    const end = new Date();
    if (isNaN(start.getTime())) return "-";
    return String(Math.round((end.getTime() - start.getTime()) / (1000 * 3600 * 24) / 30.416));
  };

  const handleAddLocal = () => {
    if (!lotNo || !qty) return;
    const cleanLot = lotNo.trim().toUpperCase();
    
    let cleanHybrid = "Unknown";
    let cleanCrops = "";
    let cleanDrDate = "";
    let cleanExpired = "N/A";
    let cleanAging = "-";

    if (lotIntel) {
      cleanHybrid = lotIntel.desc || "Unknown";
      cleanCrops = lotIntel.crops || "";
      cleanDrDate = lotIntel.drDate || "";
      const intelExp = lotIntel.expDate;
      cleanExpired = (intelExp && intelExp !== "N/A" && intelExp !== "-") ? intelExp : extractExpFromLot(cleanLot);
      cleanAging = lotIntel.aging || "-";
    } else if (isLotNotFound) {
      const selectedH = manualHybrid === "CUSTOM" ? manualHybridCustom.trim() : manualHybrid.trim();
      cleanHybrid = selectedH || "Unknown";
      cleanCrops = manualCrop;
      cleanDrDate = manualDrDate ? formatDateToAppFormat(manualDrDate) : "N/A";
      cleanExpired = manualExpDate ? formatDateToAppFormat(manualExpDate) : extractExpFromLot(cleanLot);
      cleanAging = manualDrDate ? calculateAgingInMonths(manualDrDate) : "-";
    }

    const existingItemIndex = workingData.findIndex(
      (item) =>
        cleanForMatch(item.kiosk) === cleanForMatch(selectedKiosk) &&
        cleanForMatch(item.lot) === cleanForMatch(cleanLot) &&
        cleanForMatch(item.hybrid) === cleanForMatch(cleanHybrid),
    );

    if (existingItemIndex !== -1) {
      setWorkingData((prev) =>
        prev.map((item, index) => {
          if (index === existingItemIndex) {
            const nQty = safeParseFloat(qty);
            const monthsKeys = [
              "jan",
              "feb",
              "mar",
              "apr",
              "mei",
              "jun",
              "jul",
              "ags",
              "sep",
              "okt",
              "nov",
              "des",
            ];
            const currentMonthIdx = new Date().getMonth();
            const currentMonthKey = monthsKeys[currentMonthIdx];
            const prevMonthIdx =
              currentMonthIdx === 0 ? 11 : currentMonthIdx - 1;
            const prevMonthKey = monthsKeys[prevMonthIdx];
            const prevVal =
              item[prevMonthKey] !== undefined ? safeParseFloat(item[prevMonthKey]) : 0;

            let cond = "tetap";
            if (prevVal === 0) {
              cond = item.isNew ? "new" : nQty > 0 ? "bertambah" : "tetap";
            } else {
              if (nQty > prevVal) cond = "bertambah";
              else if (nQty < prevVal) cond = "berkurang";
              else cond = "tetap";
            }

            return {
              ...item,
              stock: nQty,
              condition: cond,
              user: userData.name,
              timestamp: getCurrentTimestamp(),
              [currentMonthKey]: nQty,
            };
          }
          return item;
        }),
      );
    } else {
      const monthsKeys = [
        "jan",
        "feb",
        "mar",
        "apr",
        "mei",
        "jun",
        "jul",
        "ags",
        "sep",
        "okt",
        "nov",
        "des",
      ];
      const currentMonthKey = monthsKeys[new Date().getMonth()];
      const resolvedCategory = getValidCategory(
        undefined,
        undefined,
        selectedKiosk,
        workingCategoryMap,
      );
      const newItem = {
        id: "local_" + Date.now(),
        lot: cleanLot,
        hybrid: cleanHybrid,
        crops: cleanCrops,
        drDate: cleanDrDate,
        stock: safeParseFloat(qty),
        aging: cleanAging,
        expired: cleanExpired,
        kiosk: selectedKiosk,
        condition: "new", // BARU
        isNew: true,
        originalStock: safeParseFloat(qty),
        user: userData.name,
        timestamp: getCurrentTimestamp(),
        category: resolvedCategory,
        [currentMonthKey]: safeParseFloat(qty),
      };
      setWorkingData((prev) => [newItem, ...prev]);
    }

    // Reset Form
    setLotNo("");
    setQty("");
    setLotIntel(null);
    setIsLotNotFound(false);
    setManualHybrid("");
    setManualHybridCustom("");
    setManualCrop("Field Corn");
    setManualDrDate("");
    setManualExpDate("");
  };

  const handleEditLocal = (id, newQty, updatedFields = {}) => {
    const monthsKeys = [
      "jan",
      "feb",
      "mar",
      "apr",
      "mei",
      "jun",
      "jul",
      "ags",
      "sep",
      "okt",
      "nov",
      "des",
    ];
    const currentMonthIdx = new Date().getMonth();
    const currentMonthKey = monthsKeys[currentMonthIdx];
    const prevMonthIdx = currentMonthIdx === 0 ? 11 : currentMonthIdx - 1;
    const prevMonthKey = monthsKeys[prevMonthIdx];

    setWorkingData((prev) =>
      prev.map((item) => {
        if (item.id === id) {
          const nQty = safeParseFloat(newQty);
          const prevVal =
            item[prevMonthKey] !== undefined ? safeParseFloat(item[prevMonthKey]) : 0;

          let cond = "tetap";
          if (prevVal === 0) {
            cond = item.isNew ? "new" : nQty > 0 ? "bertambah" : "tetap";
          } else {
            if (nQty > prevVal) cond = "bertambah";
            else if (nQty < prevVal) cond = "berkurang";
            else cond = "tetap";
          }
          const resolvedCategory = getValidCategory(
            item.category || (updatedFields as any).category,
            undefined,
            item.kiosk || (updatedFields as any).kiosk,
            workingCategoryMap,
          );
          return {
            ...item,
            ...updatedFields,
            category: resolvedCategory,
            stock: nQty,
            condition: cond,
            user: userData.name,
            timestamp: getCurrentTimestamp(),
            [currentMonthKey]: nQty,
          };
        }
        return item;
      }),
    );
    setEditModal({ isOpen: false, item: null });
  };

  const handleDeleteLocal = () => {
    const id = deleteModal.item.id;
    const isNew = deleteModal.item.isNew;
    if (!isNew) {
      const monthsKeys = [
        "jan",
        "feb",
        "mar",
        "apr",
        "mei",
        "jun",
        "jul",
        "ags",
        "sep",
        "okt",
        "nov",
        "des",
      ];
      const currentMonthKey = monthsKeys[new Date().getMonth()];
      // Jangan dihapus dari list, ubah saja condition-nya jadi 'habis'
      setWorkingData((prev) =>
        prev.map((i) =>
          i.id === id
            ? {
                ...i,
                stock: 0,
                condition: "habis",
                user: userData.name,
                timestamp: getCurrentTimestamp(),
                [currentMonthKey]: 0,
              }
            : i,
        ),
      );
    } else {
      setWorkingData((prev) => prev.filter((i) => i.id !== id));
    }
    setDeleteModal({ isOpen: false, item: null });
  };

  const hasChanges =
    workingData.some(
      (i) =>
        i.isNew ||
        i.condition !== "tetap" ||
        String(i.user || "")
          .trim()
          .toLowerCase() !==
          String(i.originalUser || "")
            .trim()
            .toLowerCase(),
    ) || deletedItems.length > 0;

  const handleUploadActivity = async () => {
    // Ambil hanya list yang tampil (berdasarkan Kiosk yang terpilih)
    const currentKioskItems = workingData.filter(
      (item) =>
        cleanForMatch(item.kiosk) === cleanForMatch(selectedKiosk) ||
        matchNames(item.kiosk, selectedKiosk),
    );
    if (currentKioskItems.length === 0) return;

    setIsFetchingData(true);
    setIsSyncing(true);
    try {
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "batchActivity",
          user: userData.name,
          kiosk: selectedKiosk,
          area: userData.area,
          items: currentKioskItems, // Kirim semua data yang terlihat di layer (dan yg status 'habis')
        }),
      });
      const res = await resp.json();
      if (res.status === "error") {
        console.error("Batch activity failed:", res.message);
        alert("Upload activity gagal: " + res.message);
      } else {
        alert("Upload activity berhasil!");
      }
      fetchWorkingData();
    } catch (e) {
      console.warn(
        "Activity sync failed, using offline fallback capabilities:",
        e,
      );
      setIsFetchingData(false);
      setIsSyncing(false);
    }
  };

  const handleConsolidateDatabase = async () => {
    setIsConsolidatingDb(true);
    try {
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "consolidateDatabase",
          user: userData.name,
        }),
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        throw new Error("Respon server tidak valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        setConsolidationSuccessMsg("Konsolidasi database berhasil disimpan!");
        setTimeout(() => setConsolidationSuccessMsg(""), 5000);
        await fetchWorkingData(); // Refresh data lists
      } else {
        alert(
          "Gagal melakukan konsolidasi: " + (res.message || "Unknown error"),
        );
      }
    } catch (e) {
      console.warn("Consolidation fetch failed:", e);
      alert("Error menghubungi server untuk konsolidasi.");
    } finally {
      setIsConsolidatingDb(false);
    }
  };

  const handleEditPartnerSave = async (
    id: any,
    newPic: any,
    additionalData: any = {},
  ) => {
    setIsActionLoading(true);
    try {
      const isAdd = !id || additionalData.isAdd;
      const payload = {
        action: isAdd ? "addPartner" : "updatePartner",
        id: id || "partner_" + Date.now(),
        pic: newPic,
        name: additionalData.name || "",
        originalName: additionalData.originalName || "",
        originalPic: additionalData.originalPic || "",
        originalProvince: additionalData.originalProvince || "",
        originalGroup: additionalData.originalGroup || "",
        originalCategory: additionalData.originalCategory || "",
        category: additionalData.category || "",
        user: userData.name,
        group: additionalData.group || userData?.group || "",
        province: additionalData.province || userData?.province || "",
      };

      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify(payload),
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        throw new Error("Respon server tidak valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        if (isAdd) {
          const newPartner = {
            id: res.id || payload.id,
            name: payload.name,
            category: payload.category,
            pic: payload.pic,
            upline: "",
            province: payload.province,
            area: payload.province,
            group: payload.group,
          };
          setKiosks((prev) => [...prev, newPartner]);
        } else {
          setKiosks((prev) =>
            prev.map((k) => {
              const matchesId = String(k.id) === String(id);
              if (matchesId) {
                return {
                  ...k,
                  name: payload.name || k.name,
                  category: payload.category || k.category,
                  pic: payload.pic,
                  province: payload.province || k.province || "",
                  area: payload.province || k.area || "",
                  group: payload.group || k.group || "",
                };
              }
              return k;
            }),
          );
        }
        setPartnerEditModal({ isOpen: false, item: null });
        setChannelsRefreshKey((prev) => prev + 1);
        alert(res.message || "Partner berhasil disimpan");
      } else {
        alert("Gagal simpan partner: " + (res.message || "Unknown error"));
      }
    } catch (e: any) {
      console.warn("Gagal update data partner", e);
      alert("Terjadi kesalahan saat menyimpan data partner: " + e.message);
    } finally {
      setIsActionLoading(false);
    }
  };

  const handleDeletePartnerConfirm = async () => {
    if (!partnerDeleteModal.item?.name) {
      alert("Data partner tidak valid untuk dihapus (Nama kosong)");
      return;
    }
    const targetName = partnerDeleteModal.item.name;
    const targetPic = partnerDeleteModal.item.pic || "";
    const targetId = partnerDeleteModal.item.id;
    const targetProvince = partnerDeleteModal.item.province || partnerDeleteModal.item.area || "";
    const targetGroup = partnerDeleteModal.item.group || "";
    const targetCategory = partnerDeleteModal.item.category || "";

    setIsActionLoading(true);
    try {
      console.log("[Delete] Sending request to:", SCRIPT_URL);
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "deletePartner",
          id: targetId,
          name: targetName,
          pic: targetPic,
          originalProvince: targetProvince,
          originalGroup: targetGroup,
          originalCategory: targetCategory,
          user: userData.name,
        }),
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        const text = await resp.text();
        console.error("Non-JSON response from server:", text);
        throw new Error("Server tidak memberikan respon JSON yang valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        setKiosks((prev) =>
          prev.filter((k) => {
            if (targetId && k.id) {
              return String(k.id) !== String(targetId);
            }
            // Fallback match using name, pic, group, and province
            const matchName = String(k.name).trim().toLowerCase() === String(targetName).trim().toLowerCase();
            const matchPic = String(k.pic || "").trim().toLowerCase() === String(targetPic || "").trim().toLowerCase();
            const matchProvince = String(k.province || k.area || "").trim().toLowerCase() === String(targetProvince).trim().toLowerCase();
            const matchGroup = String(k.group || "").trim().toLowerCase() === String(targetGroup).trim().toLowerCase();
            return !(matchName && matchPic && matchProvince && matchGroup);
          })
        );
        setPartnerDeleteModal({ isOpen: false, item: null });
        setChannelsRefreshKey((prev) => prev + 1);
        alert(res.message || "Partner berhasil dihapus");
      } else {
        alert("Gagal hapus partner: " + (res.message || "Unknown error"));
      }
    } catch (e) {
      console.warn("Gagal hapus data partner", e);
      alert("Terjadi kesalahan saat menghapus partner. Silakan coba lagi.");
    } finally {
      setIsActionLoading(false);
    }
  };

  const handleEditEmployeeSave = async (
    originalName: string,
    updatedFields: any,
    isAdd = false,
  ) => {
    if (!isAdd && !originalName) return;
    setIsActionLoading(true);
    try {
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "updateEmployee",
          originalName: isAdd ? "" : originalName,
          ...updatedFields,
        }),
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        throw new Error("Respon server tidak valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        if (isAdd) {
          setEmployees((prev) => [...prev, { ...updatedFields }]);
        } else {
          if (
            userData &&
            cleanForMatch(userData.name) === cleanForMatch(originalName)
          ) {
            setUserData((prev) =>
              prev ? { ...prev, ...updatedFields } : null,
            );
          }
          setEmployees((prev) =>
            prev.map((emp) => {
              if (cleanForMatch(emp.name) === cleanForMatch(originalName)) {
                return { ...emp, ...updatedFields };
              }
              return emp;
            }),
          );
        }
        setEmployeeEditModal({ isOpen: false, item: null });
        setEmployeesRefreshKey((prev) => prev + 1);
      }
    } catch (e) {
      console.warn("Gagal update data karyawan", e);
    } finally {
      setIsActionLoading(false);
    }
  };

  const handleDeleteEmployeeConfirm = async () => {
    if (!employeeDeleteModal.item?.name) return;
    setIsActionLoading(true);
    try {
      const resp = await customFetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          action: "deleteEmployee",
          name: employeeDeleteModal.item.name,
        }),
      });

      const contentType = resp.headers.get("content-type");
      if (!resp.ok || !contentType || !contentType.includes("application/json")) {
        throw new Error("Respon server tidak valid. Pastikan backend sudah terkonfigurasi.");
      }

      const res = await resp.json();
      if (res.status === "success") {
        setEmployees((prev) =>
          prev.filter(
            (emp) =>
              cleanForMatch(emp.name) !==
              cleanForMatch(employeeDeleteModal.item?.name),
          ),
        );
        setEmployeeDeleteModal({ isOpen: false, item: null });
        setEmployeesRefreshKey((prev) => prev + 1);
        alert(res.message || "Employee berhasil dihapus");
      } else {
        alert("Gagal hapus employee: " + (res.message || "Unknown error"));
      }
    } catch (e) {
      console.warn("Gagal hapus data karyawan", e);
    } finally {
      setIsActionLoading(false);
    }
  };

  const teamStats = useMemo(() => {
    const memberChannelsMap: Record<string, typeof kiosks> = {};
    const kioskCategoryMap: Record<string, string> = {};

    teamMembers.forEach((memberName) => {
      memberChannelsMap[cleanForMatch(memberName)] = [];
    });

    kiosks.forEach((k) => {
      const resolvedPic = getDdaOfUser(
        k.pic || "",
        userData?.name,
        computedTeamProfiles,
      );
      const cleanPic = cleanForMatch(resolvedPic);
      const cleanKiosk = cleanForMatch(k.name);

      kioskCategoryMap[cleanKiosk] = String(
        k.category || "Uncategorized",
      ).trim();

      if (memberChannelsMap[cleanPic]) {
        memberChannelsMap[cleanPic].push(k);
      } else {
        const match = teamMembers.find((m) => cleanForMatch(m) === cleanPic);
        if (match) {
          const mClean = cleanForMatch(match);
          if (!memberChannelsMap[mClean]) memberChannelsMap[mClean] = [];
          memberChannelsMap[mClean].push(k);
        }
      }
    });

    rawWorkingData.forEach((d) => {
      if (d.kiosk && d.category) {
        const cleanK = cleanForMatch(d.kiosk);
        if (cleanK) {
          kioskCategoryMap[cleanK] = String(d.category).trim();
        }
      }
    });

    const monthKeys = [
      "jan",
      "feb",
      "mar",
      "apr",
      "mei",
      "jun",
      "jul",
      "ags",
      "sep",
      "okt",
      "nov",
      "des",
    ];
    let currentMonthIdx = new Date().getMonth();
    if (filterBelowMonth && filterBelowMonth !== "All") {
      const parts = filterBelowMonth.split(" ");
      const mIdx = INDO_MONTHS.indexOf(parts[0]);
      if (mIdx !== -1) {
        currentMonthIdx = mIdx;
      }
    }
    const currentMonthKey = monthKeys[currentMonthIdx];
    const updColName = `upd_${currentMonthKey}`;
    const visitedKiosksByUser: Record<string, Set<string>> = {};

    rawWorkingData.forEach((item) => {
      if (!item.kiosk) return;
      const updVal = String(item[updColName] || "")
        .trim()
        .toLowerCase();
      if (updVal === "sales") {
        const cleanKiosk = cleanForMatch(item.kiosk);
        const resolvedUser = getDdaOfUser(
          String(item.user || ""),
          userData?.name,
          computedTeamProfiles,
        );
        const cleanUser = cleanForMatch(resolvedUser);

        let finalUser = cleanUser;
        if (!teamMembers.some((m) => cleanForMatch(m) === cleanUser)) {
          const match = teamMembers.find((m) => cleanForMatch(m) === cleanUser);
          if (match) finalUser = cleanForMatch(match);
        }

        if (!visitedKiosksByUser[finalUser]) {
          visitedKiosksByUser[finalUser] = new Set();
        }
        visitedKiosksByUser[finalUser].add(cleanKiosk);
      }
    });

    return teamMembers
      .filter((memberName) => {
        const p = teamPositions[memberName];
        const pos = normalizePosition(p);
        return pos !== "Unknown";
      })
      .map((memberName) => {
        const cleanMember = cleanForMatch(memberName);
        const memberChannels = memberChannelsMap[cleanMember] || [];
        const visitedKiosks = visitedKiosksByUser[cleanMember] || new Set();

        const counts: Record<string, any> = {};

        memberChannels.forEach((c) => {
          const cat = String(c.category || "Uncategorized").trim();
          if (!counts[cat]) counts[cat] = { total: 0, visited: 0 };
          counts[cat].total += 1;
        });

        let totalVisited = 0;
        visitedKiosks.forEach((kClean) => {
          const cat = kioskCategoryMap[kClean] || "Uncategorized";
          if (!counts[cat]) counts[cat] = { total: 0, visited: 0 };
          counts[cat].visited += 1;
          totalVisited += 1;
        });

        return {
          name: memberName,
          total: memberChannels.length,
          totalVisited: totalVisited,
          counts: counts,
        };
      })
      .sort((a, b) =>
        compareMembersByLevel(
          a.name,
          b.name,
          teamLevels,
          teamPositions,
          userData,
        ),
      );
  }, [
    kiosks,
    teamMembers,
    rawWorkingData,
    teamPositions,
    userData,
    filterBelowMonth,
  ]);

  const getStatsForPic = (name: string) => {
    const cleanName = cleanForMatch(name);
    const stat = teamStats.find((s) => cleanForMatch(s.name) === cleanName);

    if (!stat) {
      return { visited: 0, total: 0, percentage: 0 };
    }
    const percentage =
      stat.total > 0 ? Math.round((stat.totalVisited / stat.total) * 100) : 0;
    return { visited: stat.totalVisited, total: stat.total, percentage };
  };

  const isKioskVisited = (kioskName: string) => {
    const cleanK = cleanForMatch(kioskName);
    const monthKeys = [
      "jan",
      "feb",
      "mar",
      "apr",
      "mei",
      "jun",
      "jul",
      "ags",
      "sep",
      "okt",
      "nov",
      "des",
    ];
    let currentMonthIdx = new Date().getMonth();
    if (filterBelowMonth && filterBelowMonth !== "All") {
      const parts = filterBelowMonth.split(" ");
      const mIdx = INDO_MONTHS.indexOf(parts[0]);
      if (mIdx !== -1) {
        currentMonthIdx = mIdx;
      }
    }
    const currentMonthKey = monthKeys[currentMonthIdx];
    const updColName = `upd_${currentMonthKey}`;
    return rawWorkingData.some((item) => {
      if (!item.kiosk) return false;
      if (cleanForMatch(item.kiosk) !== cleanK) return false;
      const updVal = String(item[updColName] || "")
        .trim()
        .toLowerCase();
      return updVal === "sales";
    });
  };

  const isRowVisitZero = (rowName: string) => {
    const cleanName = cleanForMatch(rowName);
    const stat = teamStats.find((s) => cleanForMatch(s.name) === cleanName);
    if (stat) {
      return stat.totalVisited === 0;
    }
    // Check if it matches a kiosk name
    const isKiosk = kiosks.some((k) => cleanForMatch(k.name) === cleanName);
    if (isKiosk) {
      return !isKioskVisited(rowName);
    }
    return false;
  };

  const mappedChannelsByPic = useMemo(() => {
    const enrichedKiosks = allKiosks.map((k) => {
      const resolvedPic = getDdaOfUser(
        k.pic || "",
        userData?.name,
        computedTeamProfiles,
      );
      const cat = getValidCategory(
        k.category,
        undefined,
        k.name,
        workingCategoryMap,
      );
      return { ...k, pic: resolvedPic, category: cat };
    });

    const isAll =
      !mappingPic ||
      cleanForMatch(mappingPic) === "allteam" ||
      cleanForMatch(mappingPic) === "all_team";
    if (isAll) {
      return enrichedKiosks.filter((k) => {
        return teamMembers.some((m) => matchNames(k.pic, m));
      });
    }
    return enrichedKiosks.filter((k) => matchNames(k.pic, mappingPic));
  }, [allKiosks, mappingPic, teamMembers, userData, workingCategoryMap]);

  const mappingCategories = useMemo(() => {
    const rawCats = [
      ...new Set<string>(
        mappedChannelsByPic.map((k) =>
          String(k.category || "Uncategorized").trim(),
        ),
      ),
    ];
    const order = ["Distributor", "R1", "R2"];
    return rawCats.sort((a, b) => {
      const idxA = order.indexOf(a);
      const idxB = order.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a.localeCompare(b);
    });
  }, [mappedChannelsByPic]);

  const allCategories = useMemo(() => {
    const rawCats = [
      ...new Set<string>(
        allKiosks.map((k) => String(k.category || "Uncategorized").trim()),
      ),
    ];
    const order = ["Distributor", "R1", "R2"];
    return rawCats.sort((a, b) => {
      const idxA = order.indexOf(a);
      const idxB = order.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a.localeCompare(b);
    });
  }, [allKiosks]);

  useEffect(() => {
    if (
      mappingCategories.length > 0 &&
      !mappingCategories.includes(mappingCategory)
    ) {
      setMappingCategory(mappingCategories[0]);
    } else if (mappingCategories.length === 0) {
      setMappingCategory("");
    }
  }, [mappingCategories, mappingCategory]);

  const displayedPartnerChannels = useMemo(() => {
    return mappedChannelsByPic.filter(
      (k) => String(k.category || "Uncategorized").trim() === mappingCategory,
    );
  }, [mappedChannelsByPic, mappingCategory]);

  const myKiosks = useMemo(() => {
    if (isBusinessAnalyst) {
      return allKiosks;
    }
    const userPicName = userData?.name || "";
    const direct = allKiosks.filter(
      (k) =>
        matchNames(k.pic, userPicName) ||
        (computedTeamProfiles &&
          getDdaOfUser(k.pic, userPicName, computedTeamProfiles).toLowerCase() ===
            userPicName.toLowerCase()),
    );
    if (direct.length > 0) return direct;
    return allKiosks;
  }, [allKiosks, userData, isBusinessAnalyst, computedTeamProfiles]);

  const filteredKiosks = useMemo(
    () =>
      myKiosks.filter((k) =>
        String(k.name || "")
          .toLowerCase()
          .includes(searchTerm.toLowerCase()),
      ),
    [myKiosks, searchTerm],
  );

  useEffect(() => {
    if (myKiosks.length > 0) {
      if (
        selectedKiosk === "Loading Kiosk..." ||
        selectedKiosk === "No Channel Assigned" ||
        !myKiosks.some(
          (k) => cleanForMatch(k.name) === cleanForMatch(selectedKiosk),
        )
      ) {
        setSelectedKiosk(myKiosks[0].name);
      }
    } else if (kiosks.length > 0) {
      setSelectedKiosk(kiosks[0].name);
    } else {
      setSelectedKiosk("No Channel Assigned");
    }
  }, [myKiosks, kiosks, selectedKiosk]);

  // Fungsi Parser Tanggal POG
  const parseDateForPog = (timestamp) => {
    if (!timestamp) return new Date(0);
    if (timestamp instanceof Date) return timestamp;
    let d = new Date(timestamp);
    if (!isNaN(d.getTime())) return d;

    if (typeof timestamp === "string" && timestamp.includes("/")) {
      const parts = timestamp.split(/[\s/:]+/);
      if (parts.length >= 3) {
        let year = parts[2];
        if (year.length === 2 && !isNaN(Number(year))) {
          year = "20" + year;
        }
        const dStr = `${year}-${parts[1]}-${parts[0]}T${parts[3] || "00"}:${parts[4] || "00"}:${parts[5] || "00"}`;
        d = new Date(dStr);
      }
    }
    return isNaN(d.getTime()) ? new Date(0) : d;
  };

  // Kalkulasi & Filter Data POG
  const pogDataProcessed = useMemo(() => {
    const now = new Date();
    let currentMonth = now.getMonth();
    let currentYear = now.getFullYear();

    if (filterBelowMonth && filterBelowMonth !== "All") {
      const parts = filterBelowMonth.split(" ");
      const mIdx = INDO_MONTHS.indexOf(parts[0]);
      if (mIdx !== -1) {
        currentMonth = mIdx;
        currentYear = parseInt(parts[1], 10);
      }
    }

    const startOfCurrentMonthTime = new Date(
      currentYear,
      currentMonth,
      1,
    ).getTime();

    const picToUplineMap: Record<string, string> = {};
    allKiosks.forEach((k) => {
      const p = String(k.pic || "").trim();
      const u = String(k.upline || "").trim();
      if (p && u) {
        picToUplineMap[p.toLowerCase()] = normalizeName(u);
      }
    });
    picToUplineMap["listianto"] = "AGUS HERDIANTO";

    // Pre-compute maps to avoid nested array scans (O(N*M) -> O(N+M))
    const kiosksMapByCleanName: Record<string, any> = {};
    allKiosks.forEach((k) => {
      kiosksMapByCleanName[cleanForMatch(k.name)] = k;
    });

    const teamMembersMapByCleanName: Record<string, string> = {};
    teamMembers.forEach((m) => {
      teamMembersMapByCleanName[cleanForMatch(m)] = m;
    });
    const getTeamMemberMatch = (name: string): string | undefined => {
      const clean = cleanForMatch(name);
      if (teamMembersMapByCleanName[clean])
        return teamMembersMapByCleanName[clean];
      return teamMembers.find((m) => matchNames(m, name));
    };

    const lotMap: Record<string, any> = {};

    const resolveHierarchy = (
      kioskName: string,
      itemUser?: string,
      itemArea?: string,
      itemCategory?: string,
      itemLot?: string,
      itemPic?: string,
    ) => {
      const cleanKName = cleanForMatch(kioskName);
      const kInfo =
        kiosksMapByCleanName[cleanKName] ||
        kiosks.find((k) => matchNames(k.name, kioskName)) ||
        {};
      const rawPic = normalizeName(getResolvedPic(kioskName, kInfo.pic, itemUser, itemLot, itemPic));
      const pic = getDdaOfUser(rawPic, userData?.name, computedTeamProfiles);
      let upline = normalizeName(String(kInfo.upline || ""));
      if (pic.toLowerCase() === "listianto") {
        upline = "AGUS HERDIANTO";
      } else {
        const matchedMember = getTeamMemberMatch(pic);
        const foundUp = getFromRecord<string>(
          teamUpLines,
          matchedMember || pic,
        );
        if (foundUp) {
          upline = normalizeName(foundUp);
        } else if (!upline && pic !== "Unknown") {
          const foundUpline = picToUplineMap[pic.toLowerCase()];
          if (foundUpline) upline = normalizeName(foundUpline);
        }
      }

      let area = "-";
      const cleanPic = cleanForMatch(pic);
      const matchedMember = getTeamMemberMatch(pic);
      const foundArea = getFromRecord<string>(teamAreas, matchedMember || pic);
      if (foundArea && foundArea !== "-") {
        area = foundArea;
      } else if (itemArea && String(itemArea).trim() !== "") {
        area = String(itemArea).trim();
      } else if (cleanPic === cleanForMatch(userData?.name)) {
        area = userData?.area || "-";
      }

      const category = getValidCategory(
        itemCategory,
        kInfo.category,
        kioskName,
        workingCategoryMap,
      );
      return { pic, upline, area, category, kInfo };
    };

    const monthCols = [
      "jan",
      "feb",
      "mar",
      "apr",
      "mei",
      "jun",
      "jul",
      "ags",
      "sep",
      "okt",
      "nov",
      "des",
    ];
    const prevMonthCol =
      currentMonth === 3 // April
        ? null
        : currentMonth === 0 // January
        ? "des"
        : monthCols[currentMonth - 1];
    const targetMonthCol = monthCols[currentMonth];

    rawWorkingData.forEach((d) => {
      const h = resolveHierarchy(d.kiosk, d.user, d.area, d.category, d.lot, d.pic);
      const crops =
        d.crops && String(d.crops).trim() !== ""
          ? d.crops
          : "Uncategorized Crops";

      const dateObj = parseDateForPog(d.timestamp);
      const t = dateObj.getTime();
      const key = `${d.kiosk}_${String(d.lot).trim().toUpperCase()}_${d.hybrid}_${cleanForMatch(h.pic || "")}`;
      if (!lotMap[key]) {
        lotMap[key] = {
          kiosk: d.kiosk,
          lot: String(d.lot).toUpperCase(),
          hybrid: d.hybrid,
          pic: h.pic,
          upline: h.upline,
          area: h.area,
          category: h.category,
          crops,
          lastQty: 0,
          currentQty: 0,
          latestTime: 0,
          latestRow: null,
          sellIn: 0, // representing Stock in
          sellOut: 0, // representing Stock out
          idleStock: 0, // representing idle stock
          pogAccumulated: 0,
        };
      }

      if (t > lotMap[key].latestTime || lotMap[key].latestRow === null) {
        lotMap[key].latestTime = t;
        lotMap[key].latestRow = d;
      }

      const pogVal = safeParseFloat(d.pog);
      if (pogVal < 0) {
        lotMap[key].sellIn += Math.abs(pogVal);
      } else if (pogVal > 0) {
        lotMap[key].sellOut += pogVal;
      }

      const condClean = String(d.condition || "").trim().toLowerCase();
      if (condClean === "berkurang" && pogVal > 0) {
        lotMap[key].pogAccumulated += pogVal;
      }
      if (condClean === "tetap") {
        lotMap[key].idleStock += safeParseFloat(d.stock);
      }
    });

    // Extract opening (lastQty) and ending (currentQty) directly from monthly columns of the latest row
    Object.values(lotMap).forEach((item: any) => {
      if (item.latestRow) {
        const row = item.latestRow;
        item.lastQty =
          prevMonthCol &&
          row[prevMonthCol] !== undefined &&
          safeParseFloat(row[prevMonthCol]) > 0
            ? safeParseFloat(row[prevMonthCol])
            : 0;

        const isUpdated = row["upd_" + targetMonthCol] && row["upd_" + targetMonthCol] !== "";
        item.currentQty = isUpdated
          ? safeParseFloat(row[targetMonthCol])
          : safeParseFloat(row.stock);
      }
    });

    return Object.values(lotMap)
      .map((item: any) => {
        return {
          ...item,
          pog: (item.lastQty || 0) + (item.sellIn || 0) - (item.currentQty || 0),
        };
      })
      .filter(
        (item) =>
          item.lastQty > 0 ||
          item.currentQty > 0 ||
          item.sellIn > 0 ||
          item.sellOut > 0 ||
          item.idleStock > 0,
      );
  }, [
    rawWorkingData,
    drSalesData,
    kiosks,
    userData,
    teamAreas,
    teamMembers,
    teamUpLines,
    filterBelowMonth,
  ]);

  // Dynamic filter options based on raw data
  const filterOptions = useMemo(() => {
    const rawData =
      rawWorkingData && rawWorkingData.length > 0 ? rawWorkingData : [];

    // 1. Months (April - Maret, matching the database columns / crop year)
    const months = [
      "April 2026",
      "Mei 2026",
      "Juni 2026",
      "Juli 2026",
      "Agustus 2026",
      "September 2026",
      "Oktober 2026",
      "November 2026",
      "Desember 2026",
      "Januari 2027",
      "Februari 2027",
      "Maret 2027",
    ];

    // 2. Channel (Category)
    const channelsSet = new Set<string>();
    kiosks.forEach((k) => {
      const cat = getValidCategory(k.category, undefined, k.name, workingCategoryMap);
      if (cat && cat !== "Uncategorized") channelsSet.add(cat);
    });
    rawData.forEach((d) => {
      const cat = getValidCategory(d.category, undefined, d.kiosk, workingCategoryMap);
      if (cat && cat !== "Uncategorized") channelsSet.add(cat);
    });
    const channels = Array.from(channelsSet).filter(Boolean).sort();

    // 3. Material (Hybrid)
    const materialsSet = new Set<string>();
    rawData.forEach((d) => {
      if (d.hybrid) materialsSet.add(String(d.hybrid).trim());
    });
    const materials = Array.from(materialsSet).filter(Boolean).sort();

    // 4. Team (PIC)
    const teamsSet = new Set<string>();
    rawData.forEach((d) => {
      const cleanKName = cleanForMatch(d.kiosk);
      const kInfo =
        kiosks.find((k) => cleanForMatch(k.name) === cleanKName) || {};
      const rawPic = normalizeName(getResolvedPic(d.kiosk, kInfo.pic, d.user, d.lot, d.pic));
      const pic = getDdaOfUser(rawPic, userData?.name, computedTeamProfiles);
      if (pic && pic !== "Unknown") teamsSet.add(String(pic).trim());
    });
    const teams = Array.from(teamsSet).filter(Boolean).sort();

    // 5. Area
    const areasSet = new Set<string>();
    if (employees && employees.length > 0) {
      employees.forEach((emp) => {
        const area = String(emp.area || "").trim();
        if (area && area !== "-") {
          areasSet.add(area);
        }
      });
    } else {
      rawData.forEach((d) => {
        const cleanKName = cleanForMatch(d.kiosk);
        const kInfo =
          kiosks.find((k) => cleanForMatch(k.name) === cleanKName) || {};
        const rawPic = normalizeName(getResolvedPic(d.kiosk, kInfo.pic, d.user, d.lot, d.pic));
        const pic = getDdaOfUser(rawPic, userData?.name, computedTeamProfiles);
        const matchedMember = teamMembers.find((m) => matchNames(m, pic));
        const area =
          getFromRecord<string>(teamAreas, matchedMember || pic) ||
          kInfo.area ||
          d.area;
        if (area && area !== "-") areasSet.add(String(area).trim());
      });
    }
    const areas = Array.from(areasSet).filter(Boolean).sort();

    return {
      months,
      channels,
      materials,
      teams,
      areas,
    };
  }, [
    rawWorkingData,
    kiosks,
    userData,
    teamMembers,
    teamAreas,
    employees,
    workingCategoryMap,
  ]);

  // Recalculate processed POG data base on selected month filter
  const pogDataProcessedForOverview = useMemo(() => {
    const now = new Date();
    let targetMonth = now.getMonth();
    let targetYear = now.getFullYear();
    const isFilteredMonth = filterBelowMonth && filterBelowMonth !== "All";

    if (isFilteredMonth) {
      const parts = filterBelowMonth.split(" ");
      const mIdx = INDO_MONTHS.indexOf(parts[0]);
      if (mIdx !== -1) {
        targetMonth = mIdx;
        targetYear = parseInt(parts[1], 10);
      }
    }

    const startOfTargetMonthTime = new Date(
      targetYear,
      targetMonth,
      1,
    ).getTime();
    const endOfTargetMonthTime = new Date(
      targetYear,
      targetMonth + 1,
      1,
    ).getTime();

    const monthCols = [
      "jan",
      "feb",
      "mar",
      "apr",
      "mei",
      "jun",
      "jul",
      "ags",
      "sep",
      "okt",
      "nov",
      "des",
    ];
    const prevMonthCol =
      targetMonth === 3 // April
        ? null
        : targetMonth === 0 // January
        ? "des"
        : monthCols[targetMonth - 1];
    const targetMonthCol = monthCols[targetMonth];

    const picToUplineMap: Record<string, string> = {};
    allKiosks.forEach((k) => {
      const p = String(k.pic || "").trim();
      const u = String(k.upline || "").trim();
      if (p && u) {
        picToUplineMap[p.toLowerCase()] = normalizeName(u);
      }
    });
    picToUplineMap["listianto"] = "AGUS HERDIANTO";

    const kiosksMapByCleanName: Record<string, any> = {};
    allKiosks.forEach((k) => {
      kiosksMapByCleanName[cleanForMatch(k.name)] = k;
    });

    const teamMembersMapByCleanName: Record<string, string> = {};
    teamMembers.forEach((m) => {
      teamMembersMapByCleanName[cleanForMatch(m)] = m;
    });
    const getTeamMemberMatch = (name: string): string | undefined => {
      const clean = cleanForMatch(name);
      if (teamMembersMapByCleanName[clean])
        return teamMembersMapByCleanName[clean];
      return teamMembers.find((m) => matchNames(m, name));
    };

    const lotMap: Record<string, any> = {};

    const resolveHierarchy = (
      kioskName: string,
      itemUser?: string,
      itemArea?: string,
      itemCategory?: string,
      itemLot?: string,
      itemPic?: string,
    ) => {
      const cleanKName = cleanForMatch(kioskName);
      const kInfo =
        kiosksMapByCleanName[cleanKName] ||
        kiosks.find((k) => matchNames(k.name, kioskName)) ||
        {};
      const rawPic = normalizeName(getResolvedPic(kioskName, kInfo.pic, itemUser, itemLot, itemPic));
      const pic = getDdaOfUser(rawPic, userData?.name, computedTeamProfiles);
      let upline = normalizeName(String(kInfo.upline || ""));
      if (pic.toLowerCase() === "listianto") {
        upline = "AGUS HERDIANTO";
      }

      let area = "-";
      const cleanPic = cleanForMatch(pic);
      const matchedMember = getTeamMemberMatch(pic);
      const foundArea = getFromRecord<string>(teamAreas, matchedMember || pic);
      if (foundArea && foundArea !== "-") {
        area = foundArea;
      } else if (itemArea && String(itemArea).trim() !== "") {
        area = String(itemArea).trim();
      } else if (cleanPic === cleanForMatch(userData?.name)) {
        area = userData?.area || "-";
      }

      const category = getValidCategory(
        itemCategory,
        kInfo.category,
        kioskName,
        workingCategoryMap,
      );
      return { pic, upline, area, category, kInfo };
    };

    const rawData =
      rawWorkingData && rawWorkingData.length > 0 ? rawWorkingData : [];

    rawData.forEach((d) => {
      const dateObj = parseDateForPog(d.timestamp);
      const t = dateObj.getTime();

      // If filtering by specific month, skip transactions that happened after the end of this target month
      if (isFilteredMonth && t >= endOfTargetMonthTime) {
        return;
      }

      const h = resolveHierarchy(d.kiosk, d.user, d.area, d.category, d.lot, d.pic);
      const crops =
        d.crops && String(d.crops).trim() !== ""
          ? d.crops
          : "Uncategorized Crops";

      const key = `${d.kiosk}_${String(d.lot).trim().toUpperCase()}_${d.hybrid}_${cleanForMatch(h.pic || "")}`;
      if (!lotMap[key]) {
        lotMap[key] = {
          kiosk: d.kiosk,
          lot: String(d.lot).toUpperCase(),
          hybrid: d.hybrid,
          pic: h.pic,
          upline: h.upline,
          area: h.area,
          category: h.category,
          crops,
          lastQty: 0,
          currentQty: 0,
          latestTime: 0,
          latestRow: null,
          sellIn: 0,
          sellOut: 0,
          idleStock: 0,
          pogAccumulated: 0,
        };
      }

      if (t > lotMap[key].latestTime || lotMap[key].latestRow === null) {
        lotMap[key].latestTime = t;
        lotMap[key].latestRow = d;
      }

      const pogVal = safeParseFloat(d.pog);
      if (pogVal < 0) {
        lotMap[key].sellIn += Math.abs(pogVal);
      } else if (pogVal > 0) {
        lotMap[key].sellOut += pogVal;
      }

      const condClean = String(d.condition || "").trim().toLowerCase();
      if (condClean === "berkurang" && pogVal > 0) {
        lotMap[key].pogAccumulated += pogVal;
      }
      if (condClean === "tetap") {
        lotMap[key].idleStock += safeParseFloat(d.stock);
      }
    });

    // Extract opening (lastQty) and ending (currentQty) directly from monthly columns of the latest row
    Object.values(lotMap).forEach((item: any) => {
      if (item.latestRow) {
        const row = item.latestRow;
        item.lastQty =
          prevMonthCol &&
          row[prevMonthCol] !== undefined &&
          safeParseFloat(row[prevMonthCol]) > 0
            ? safeParseFloat(row[prevMonthCol])
            : 0;

        const isUpdated = row["upd_" + targetMonthCol] && row["upd_" + targetMonthCol] !== "";
        item.currentQty = isUpdated
          ? safeParseFloat(row[targetMonthCol])
          : safeParseFloat(row.stock);
      }
    });

    const teamMembersCleanSet = new Set(
      teamMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );

    let result = Object.values(lotMap)
      .map((item: any) => {
        return {
          ...item,
          pog: (item.lastQty || 0) + (item.sellIn || 0) - (item.currentQty || 0),
        };
      })
      .filter(
        (item) =>
          item.lastQty > 0 ||
          item.currentQty > 0 ||
          item.sellIn > 0 ||
          item.sellOut > 0 ||
          item.idleStock > 0,
      )
      .filter((item) => {
        if (isBusinessAnalyst) return true;
        const picClean = cleanForMatch(item.pic);
        let isTeamMember = teamMembersCleanSet.has(picClean);
        if (!isTeamMember && computedTeamProfiles) {
          const emp = computedTeamProfiles[picClean];
          if (emp && emp.name) {
            const empNameClean = cleanForMatch(emp.name);
            if (teamMembersCleanSet.has(empNameClean)) {
              isTeamMember = true;
            }
          }
        }
        return isTeamMember;
      });

    return result;
  }, [
    rawWorkingData,
    kiosks,
    userData,
    teamAreas,
    teamMembers,
    teamUpLines,
    filterBelowMonth,
    isBusinessAnalyst,
  ]);

  // Apply other category filters (channel, material, team, area) to computed list
  const pogDataOverviewFiltered = useMemo(() => {
    let result = pogDataProcessedForOverview;

    // Filter 2: channel (Category)
    if (filterBelowChannel && filterBelowChannel !== "All") {
      result = result.filter(
        (item) =>
          cleanForMatch(item.category) === cleanForMatch(filterBelowChannel),
      );
    }

    // Filter 2: material (Hybrid)
    if (filterBelowMaterial && filterBelowMaterial !== "All") {
      result = result.filter(
        (item) =>
          cleanForMatch(item.hybrid) === cleanForMatch(filterBelowMaterial),
      );
    }

    // Filter 2: team (PIC)
    if (filterBelowTeam && filterBelowTeam !== "All") {
      result = result.filter(
        (item) => cleanForMatch(item.pic) === cleanForMatch(filterBelowTeam),
      );
    }

    // Filter 2: area
    if (filterBelowArea && filterBelowArea !== "All") {
      result = result.filter(
        (item) => cleanForMatch(item.area) === cleanForMatch(filterBelowArea),
      );
    }

    // Filter 2: crop (Crops)
    if (filterBelowCrop && filterBelowCrop !== "All") {
      result = result.filter((item) =>
        checkCropMatch(item.crops, filterBelowCrop),
      );
    }

    return result;
  }, [
    pogDataProcessedForOverview,
    filterBelowChannel,
    filterBelowMaterial,
    filterBelowTeam,
    filterBelowArea,
    filterBelowCrop,
  ]);

  const overviewHistoryData = useMemo(() => {
    // Generate historical trends grouped by Month from April 2026 to March 2027
    const monthsSequence = [
      { key: "2026-04", label: "Apr 26", prop: "apr", monthIdx: 3, year: 2026 },
      { key: "2026-05", label: "Mei 26", prop: "mei", monthIdx: 4, year: 2026 },
      { key: "2026-06", label: "Jun 26", prop: "jun", monthIdx: 5, year: 2026 },
      { key: "2026-07", label: "Jul 26", prop: "jul", monthIdx: 6, year: 2026 },
      { key: "2026-08", label: "Ags 26", prop: "ags", monthIdx: 7, year: 2026 },
      { key: "2026-09", label: "Sep 26", prop: "sep", monthIdx: 8, year: 2026 },
      { key: "2026-10", label: "Okt 26", prop: "okt", monthIdx: 9, year: 2026 },
      {
        key: "2026-11",
        label: "Nov 26",
        prop: "nov",
        monthIdx: 10,
        year: 2026,
      },
      {
        key: "2026-12",
        label: "Des 26",
        prop: "des",
        monthIdx: 11,
        year: 2026,
      },
      { key: "2027-01", label: "Jan 27", prop: "jan", monthIdx: 0, year: 2027 },
      { key: "2027-02", label: "Feb 27", prop: "feb", monthIdx: 1, year: 2027 },
      { key: "2027-03", label: "Mar 27", prop: "mar", monthIdx: 2, year: 2027 },
    ];

    const monthlyMap: Record<
      string,
      {
        monthKey: string;
        monthLabel: string;
        opening: number;
        ending: number;
        stockIn: number;
        idle: number;
        pog: number;
      }
    > = {};
    monthsSequence.forEach((item) => {
      monthlyMap[item.key] = {
        monthKey: item.key,
        monthLabel: item.label,
        opening: 0,
        ending: 0,
        stockIn: 0,
        idle: 0,
        pog: 0,
      };
    });

    rawWorkingData.forEach((d) => {
      // Dynamic filters matching Category, Material, Team, Area
      if (filterBelowChannel && filterBelowChannel !== "All") {
        const kName = cleanForMatch(d.kiosk);
        const kInfo = allKiosks.find((k) => cleanForMatch(k.name) === kName);
        const effectiveCat = kInfo?.category || d.category;
        if (
          cleanForMatch(effectiveCat) !== cleanForMatch(filterBelowChannel)
        ) {
          return;
        }
      }
      if (filterBelowMaterial && filterBelowMaterial !== "All") {
        if (cleanForMatch(d.hybrid) !== cleanForMatch(filterBelowMaterial)) {
          return;
        }
      }
      if (filterBelowTeam && filterBelowTeam !== "All") {
        const kName = cleanForMatch(d.kiosk);
        const kInfo = kiosks.find((k) => cleanForMatch(k.name) === kName) || {};
        const rawPic = normalizeName(getResolvedPic(d.kiosk, kInfo.pic, d.user, d.lot, d.pic));
        const pic = getDdaOfUser(
          rawPic,
          userData?.name,
          computedTeamProfiles,
        );
        if (cleanForMatch(pic) !== cleanForMatch(filterBelowTeam)) {
          return;
        }
      }
      if (filterBelowArea && filterBelowArea !== "All") {
        const kName = cleanForMatch(d.kiosk);
        const kInfo = kiosks.find((k) => cleanForMatch(k.name) === kName) || {};
        const rawPic = normalizeName(getResolvedPic(d.kiosk, kInfo.pic, d.user, d.lot, d.pic));
        const pic = getDdaOfUser(
          rawPic,
          userData?.name,
          computedTeamProfiles,
        );
        const matchedMember = teamMembers.find((m) => matchNames(m, pic));
        const area =
          getFromRecord<string>(teamAreas, matchedMember || pic) ||
          kInfo.area ||
          d.area ||
          "-";
        if (cleanForMatch(area) !== cleanForMatch(filterBelowArea)) {
          return;
        }
      }
      if (filterBelowCrop && filterBelowCrop !== "All") {
        const itemCrop =
          d.crops || d.Crops || d.crop || d.Crop || d.CROP || d.CROPS || "";
        if (!checkCropMatch(itemCrop, filterBelowCrop)) {
          return;
        }
      }

      // Find max POG across months to estimate capacity
      let maxPogForKiosk = 0;
      monthsSequence.forEach((cfg) => {
        const val = safeParseFloat((d as any)[cfg.prop]);
        if (val > maxPogForKiosk) maxPogForKiosk = val;
      });
      if (maxPogForKiosk === 0) maxPogForKiosk = 100; // default baseline

      // Initialize sequential inventory simulation for this specific channel partner
      let currentOpening = Math.round(maxPogForKiosk * 1.5 + 40);

      monthsSequence.forEach((cfg) => {
        const pogVal = safeParseFloat((d as any)[cfg.prop]);
        const stockInVal = Math.round(pogVal * 1.08 + (pogVal > 0 ? 12 : 3));
        const endingVal = Math.max(0, currentOpening + stockInVal - pogVal);
        const idleVal = Math.min(
          endingVal,
          Math.round(currentOpening * 0.15 + 8),
        );

        monthlyMap[cfg.key].opening += currentOpening;
        monthlyMap[cfg.key].ending += endingVal;
        monthlyMap[cfg.key].stockIn += stockInVal;
        monthlyMap[cfg.key].idle += idleVal;
        monthlyMap[cfg.key].pog += pogVal;

        currentOpening = endingVal;
      });
    });

    return monthsSequence.map((cfg) => monthlyMap[cfg.key]);
  }, [
    rawWorkingData,
    kiosks,
    filterBelowChannel,
    filterBelowMaterial,
    filterBelowTeam,
    filterBelowArea,
    filterBelowCrop,
    userData,
    teamMembers,
    teamAreas,
  ]);

  const overviewStats = useMemo(() => {
    let activeKiosks = (allKiosks || []).filter((k) => {
      if (isBusinessAnalyst) return true;
      const resolvedPic = getDdaOfUser(
        k.pic || "",
        userData?.name,
        computedTeamProfiles,
      );
      return teamMembers.some((m) => matchNames(m, resolvedPic));
    });

    if (filterBelowChannel && filterBelowChannel !== "All") {
      activeKiosks = activeKiosks.filter(
        (k) => cleanForMatch(k.category) === cleanForMatch(filterBelowChannel),
      );
    }
    if (filterBelowTeam && filterBelowTeam !== "All") {
      activeKiosks = activeKiosks.filter((k) => {
        const resolvedPic = getDdaOfUser(
          k.pic || "",
          userData?.name,
          computedTeamProfiles,
        );
        return cleanForMatch(resolvedPic) === cleanForMatch(filterBelowTeam);
      });
    }
    if (filterBelowArea && filterBelowArea !== "All") {
      activeKiosks = activeKiosks.filter((k) => {
        const resolvedPic = getDdaOfUser(
          k.pic || "",
          userData?.name,
          computedTeamProfiles,
        );
        const matchedMember = teamMembers.find((m) =>
          matchNames(m, resolvedPic),
        );
        const area =
          getFromRecord<string>(teamAreas, matchedMember || resolvedPic) ||
          k.area ||
          "-";
        return cleanForMatch(area) === cleanForMatch(filterBelowArea);
      });
    }
    if (filterBelowCrop && filterBelowCrop !== "All") {
      const activeKioskNames = new Set(
        pogDataOverviewFiltered.map((item) => cleanForMatch(item.kiosk)),
      );
      activeKiosks = activeKiosks.filter((k) =>
        activeKioskNames.has(cleanForMatch(k.name)),
      );
    }

    let totalKiosks = activeKiosks.length;
    let totalOpeningStock = 0;
    let totalSellIn = 0;
    let totalSellOut = 0;
    let totalCurrentStock = 0;
    let totalIdleStock = 0;

    pogDataOverviewFiltered.forEach((item) => {
      const isDistributor = cleanForMatch(item.category) === "distributor";
      totalOpeningStock += safeParseFloat(item.lastQty);
      if (isDistributor) totalSellIn += safeParseFloat(item.sellIn);
      totalSellOut += safeParseFloat(item.sellOut);
      totalCurrentStock += safeParseFloat(item.currentQty);
      totalIdleStock += safeParseFloat(item.idleStock);
    });

    let totalPog = totalOpeningStock + totalSellIn - totalCurrentStock;

    // Grouping dimension: group by area, province, sales_agronomist, business_solution, material, or distributor
    const groupMap: Record<
      string,
      {
        name: string;
        pog: number;
        stock: number;
        sellIn: number;
        sellOut: number;
        opening: number;
        idle: number;
      }
    > = {};

    pogDataOverviewFiltered.forEach((item) => {
      let gVal = "";
      const itemPic = item.pic;
      const rawPos = getFromRecord<string>(teamPositions, itemPic) || "";
      const pos = normalizePosition(rawPos);

      if (overviewGroupDimension === "area") {
        gVal = item.area || "Tanpa Area";
      } else if (overviewGroupDimension === "province") {
        gVal =
          getFromRecord<string>(teamProvinces, itemPic) ||
          item.province ||
          "Tanpa Provinsi";
      } else if (overviewGroupDimension === "sales_agronomist") {
        if (pos === "Sales Agronomist") {
          gVal = itemPic || "Unknown";
        } else {
          return;
        }
      } else if (overviewGroupDimension === "business_solution") {
        if (pos === "Business Solution") {
          gVal = itemPic || "Unknown";
        } else {
          return;
        }
      } else if (overviewGroupDimension === "material") {
        gVal = item.hybrid || "Lainnya";
      } else if ((overviewGroupDimension as string) === "distributor") {
        if (cleanForMatch(item.category) === "distributor") {
          gVal = item.kiosk || "Tanpa Distributor";
        } else {
          return;
        }
      } else {
        gVal = item.area || "Tanpa Area";
      }

      if (!groupMap[gVal]) {
        groupMap[gVal] = {
          name: gVal,
          pog: 0,
          stock: 0,
          sellIn: 0,
          sellOut: 0,
          opening: 0,
          idle: 0,
        };
      }
      groupMap[gVal].pog += safeParseFloat(item.pog);
      groupMap[gVal].stock += safeParseFloat(item.currentQty);
      groupMap[gVal].sellIn += safeParseFloat(item.sellIn);
      groupMap[gVal].sellOut += safeParseFloat(item.sellOut);
      groupMap[gVal].opening += safeParseFloat(item.lastQty);
      groupMap[gVal].idle += safeParseFloat(item.idleStock);
    });

    const areaChartData = Object.values(groupMap).sort((a, b) => {
      if (overviewMetricFilter === "movement") {
        return b.sellIn + b.pog - (a.sellIn + a.pog);
      } else if (overviewMetricFilter === "idle") {
        return b.idle - a.idle;
      } else if (overviewMetricFilter === "total_stock") {
        return b.stock - a.stock;
      } else if (overviewMetricFilter === "Opening") {
        return b.opening - a.opening;
      } else {
        return b.pog - a.pog;
      }
    });

    // Group by Category of Kiosk (Channel)
    const catMap: Record<string, { name: string; value: number }> = {};
    activeKiosks.forEach((k) => {
      const cat = getValidCategory(
        k.category,
        undefined,
        k.name,
        workingCategoryMap,
      );
      if (!catMap[cat]) {
        catMap[cat] = { name: cat, value: 0 };
      }
      catMap[cat].value += 1;
    });
    const categoryChartData = Object.values(catMap);

    // Group by Crops
    const cropMap: Record<
      string,
      { name: string; stock: number; pog: number }
    > = {};
    pogDataOverviewFiltered.forEach((item) => {
      const crop = item.crops || "Lainnya";
      if (!cropMap[crop]) {
        cropMap[crop] = { name: crop, stock: 0, pog: 0 };
      }
      cropMap[crop].stock += safeParseFloat(item.currentQty);
      cropMap[crop].pog += safeParseFloat(item.pog);
    });
    const cropsChartData = Object.values(cropMap).sort(
      (a, b) => b.stock - a.stock,
    );

    return {
      totalKiosks,
      totalOpeningStock,
      totalSellIn,
      totalSellOut,
      totalCurrentStock,
      totalIdleStock,
      totalPog,
      areaChartData,
      categoryChartData,
      cropsChartData,
    };
  }, [
    allKiosks,
    pogDataOverviewFiltered,
    overviewMetricFilter,
    overviewGroupDimension,
    teamPositions,
    teamProvinces,
    filterBelowChannel,
    filterBelowTeam,
    filterBelowArea,
    filterBelowCrop,
    userData,
    teamMembers,
    teamAreas,
  ]);

  const topKiosksData = useMemo(() => {
    const kioskSales: Record<
      string,
      {
        kiosk: string;
        category: string;
        pic: string;
        area: string;
        pog: number;
        currentStock: number;
      }
    > = {};

    pogDataOverviewFiltered.forEach((item) => {
      const kName = item.kiosk;
      if (!kioskSales[kName]) {
        kioskSales[kName] = {
          kiosk: kName,
          category: getValidCategory(
            item.category,
            undefined,
            item.kiosk,
            workingCategoryMap,
          ),
          pic: item.pic || "Unknown",
          area: item.area || "-",
          pog: 0,
          currentStock: 0,
        };
      }
      kioskSales[kName].pog += Number(item.pog || 0);
      kioskSales[kName].currentStock += Number(item.currentQty || 0);
    });

    return Object.values(kioskSales)
      .sort((a, b) => b.pog - a.pog)
      .slice(0, 5);
  }, [pogDataOverviewFiltered]);

  const employeePerformanceData = useMemo(() => {
    const empSales: Record<
      string,
      { employee: string; area: string; pog: number; currentStock: number }
    > = {};

    teamMembers.forEach((member) => {
      const mClean = cleanForMatch(member);

      // Get area for this member
      let area = "-";
      const foundArea = getFromRecord<string>(teamAreas, member);
      if (foundArea) {
        area = foundArea;
      } else if (mClean === cleanForMatch(userData?.name)) {
        area = userData?.area || "-";
      }

      // Filter by Area if chosen
      if (
        filterBelowArea &&
        filterBelowArea !== "All" &&
        cleanForMatch(area) !== cleanForMatch(filterBelowArea)
      ) {
        return;
      }

      // Filter by PIC if chosen
      if (
        filterBelowTeam &&
        filterBelowTeam !== "All" &&
        mClean !== cleanForMatch(filterBelowTeam)
      ) {
        return;
      }

      empSales[mClean] = {
        employee: member,
        area: area,
        pog: 0,
        currentStock: 0,
      };
    });

    // Accumulate sales and stock from pogDataOverviewFiltered
    pogDataOverviewFiltered.forEach((item) => {
      const picClean = cleanForMatch(item.pic);
      // Try to find matching employee
      let matchedEmpClean = picClean;
      if (!empSales[picClean]) {
        const found = teamMembers.find(
          (m) => cleanForMatch(m) === picClean || matchNames(m, item.pic),
        );
        if (found) {
          matchedEmpClean = cleanForMatch(found);
        }
      }

      if (empSales[matchedEmpClean]) {
        empSales[matchedEmpClean].pog += Number(item.pog || 0);
        empSales[matchedEmpClean].currentStock += Number(item.currentQty || 0);
      }
    });

    const empSalesList = Object.values(empSales);

    // Sort for Highest (Top 5)
    const highest = [...empSalesList].sort((a, b) => b.pog - a.pog).slice(0, 5);

    // Sort for Lowest (Bottom 5)
    const lowest = [...empSalesList].sort((a, b) => a.pog - b.pog).slice(0, 5);

    return { highest, lowest };
  }, [
    teamMembers,
    pogDataOverviewFiltered,
    teamAreas,
    userData,
    filterBelowArea,
    filterBelowTeam,
  ]);

  // Filter active POG team members based on selected crop filter
  const activePogMembers = useMemo(() => {
    if (!filterBelowCrop || filterBelowCrop === "All") return teamMembers;

    const kiosksMapByCleanName: Record<string, any> = {};
    allKiosks.forEach((k) => {
      kiosksMapByCleanName[cleanForMatch(k.name)] = k;
    });
    const teamMembersMapByCleanName: Record<string, string> = {};
    teamMembers.forEach((m) => {
      teamMembersMapByCleanName[cleanForMatch(m)] = m;
    });
    const getTeamMemberMatch = (name: string): string | undefined => {
      const clean = cleanForMatch(name);
      if (teamMembersMapByCleanName[clean])
        return teamMembersMapByCleanName[clean];
      return teamMembers.find((m) => matchNames(m, name));
    };

    const matchedMembersWithCrop = new Set<string>();
    pogDataProcessed.forEach((item) => {
      const itemCrop = String(item.crops || "")
        .trim()
        .toLowerCase();
      if (checkCropMatch(itemCrop, filterBelowCrop)) {
        const matchedMember = getTeamMemberMatch(item.pic);
        if (matchedMember) {
          matchedMembersWithCrop.add(cleanForMatch(matchedMember));
        }
      }
    });

    const activeSet = new Set<string>();
    activeSet.add(cleanForMatch(userData?.name));

    teamMembers.forEach((m) => {
      const mClean = cleanForMatch(m);
      if (matchedMembersWithCrop.has(mClean)) {
        activeSet.add(mClean);
        let current = m;
        for (let i = 0; i < 15; i++) {
          const upRaw = getFromRecord<string>(teamUpLines, current);
          if (!upRaw) break;

          const up = getTeamMemberMatch(upRaw as string) || (upRaw as string);

          const upClean = cleanForMatch(up);
          if (activeSet.has(upClean)) {
            break;
          }
          activeSet.add(upClean);
          current = up;
        }
      }
    });

    return teamMembers.filter((m) => activeSet.has(cleanForMatch(m)));
  }, [
    teamMembers,
    filterBelowCrop,
    pogDataProcessed,
    kiosks,
    userData,
    teamUpLines,
  ]);

  const aggregatedPogData = useMemo(() => {
    const teamMembersCleanSet = new Set(
      activePogMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );
    

    const filteredData = pogDataProcessed.filter((item) => {
      const isTeamMember = isBusinessAnalyst || teamMembersCleanSet.has(cleanForMatch(item.pic));
      const isCropMatch = checkCropMatch(item.crops, filterBelowCrop);
      return isTeamMember && isCropMatch;
    });

    // Filter active POG team members based on selected crop filter
    const hasGroupInfo = employees.some((e) => {
      const g = String(
        e.group || e.Group || e["group"] || e["Group"] || "",
      ).trim();
      return g.length > 0;
    });

    if (pogGroupBy === "subordinate") {
      // Pre-compute uplines for activePogMembers to avoid O(M) recursive scans inside buildNode loops
      const pMembersUplines: Record<string, string | null> = {};
      const pDirectSubsMap: Record<string, string[]> = {};
      activePogMembers.forEach((m) => {
        pMembersUplines[m] = getUplineInTeam(m, activePogMembers, teamUpLines);
      });

      const rootMembers = activePogMembers.filter((m) => {
        return pMembersUplines[m] === null;
      });
      rootMembers.sort((a, b) =>
        compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
      );

      activePogMembers.forEach((m) => {
        const upl = pMembersUplines[m];
        if (upl !== null) {
          const uplClean = cleanForMatch(upl);
          if (!pDirectSubsMap[uplClean]) {
            pDirectSubsMap[uplClean] = [];
          }
          pDirectSubsMap[uplClean].push(m);
        }
      });

      Object.keys(pDirectSubsMap).forEach((key) => {
        pDirectSubsMap[key].sort((a, b) =>
          compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
        );
      });

      const buildNode = (name: string): any => {
        const nameClean = cleanForMatch(name);
        const directSubs = pDirectSubsMap[nameClean] || [];
        const myItems = filteredData.filter(
          (item) => cleanForMatch(item.pic) === nameClean,
        );
        const childrenNodes = directSubs.map((sub) => buildNode(sub));

        const ownPog = {
          lastQty: 0,
          currentQty: 0,
          pog: 0,
          sellIn: 0,
          sellOut: 0,
          totalInv: 0,
          idleStock: 0,
        };
        myItems.forEach((item) => {
          const isDistributor = cleanForMatch(item.category) === "distributor";
          ownPog.lastQty += Number(item.lastQty) || 0;
          ownPog.currentQty += Number(item.currentQty) || 0;
          if (isDistributor) ownPog.sellIn += Number(item.sellIn) || 0;
          ownPog.sellOut += Number(item.sellOut) || 0;
          ownPog.totalInv += Number(item.totalInv) || 0;
          ownPog.idleStock += Number(item.idleStock) || 0;
          ownPog.pog += Number(item.pog) || 0;
        });

        const transPog = { ...ownPog };
        childrenNodes.forEach((child) => {
          transPog.lastQty += child.lastQty || 0;
          transPog.currentQty += child.currentQty || 0;
          transPog.sellIn += child.sellIn || 0;
          transPog.sellOut += child.sellOut || 0;
          transPog.totalInv += child.totalInv || 0;
          transPog.pog += child.pog || 0;
          transPog.idleStock += child.idleStock || 0;
        });

        let finalChildren = [];

        if (pogSubGroupBy === "subordinate") {
          finalChildren = [...childrenNodes];
        } else {
          const leafGroups: Record<string, any> = {};

          // Accumulate own items
          myItems.forEach((item) => {
            const isDistributor = cleanForMatch(item.category) === "distributor";
            let leafKey = "Unknown";
            if (pogSubGroupBy === "channel" || pogSubGroupBy === "kiosk")
              leafKey = item.kiosk || "Unknown Channel";
            else if (pogSubGroupBy === "hybrid")
              leafKey = item.hybrid || "Unknown";
            else if (pogSubGroupBy === "area") leafKey = item.area || "Unknown";
            else if (pogSubGroupBy === "category")
              leafKey = item.category || "Unknown";
            else if (pogSubGroupBy === "crops")
              leafKey = item.crops || "Unknown";

            if (!leafGroups[leafKey]) {
              leafGroups[leafKey] = {
                lastQty: 0,
                currentQty: 0,
                pog: 0,
                sellIn: 0,
                sellOut: 0,
                totalInv: 0,
                idleStock: 0,
                category: item.category,
                pic: item.pic,
              };
            }
            leafGroups[leafKey].lastQty += Number(item.lastQty) || 0;
            leafGroups[leafKey].currentQty += Number(item.currentQty) || 0;
            leafGroups[leafKey].sellIn += Number(item.sellIn) || 0;
            leafGroups[leafKey].sellOut += Number(item.sellOut) || 0;
            leafGroups[leafKey].totalInv += Number(item.totalInv) || 0;
            leafGroups[leafKey].pog += Number(item.pog) || 0;
            leafGroups[leafKey].idleStock += Number(item.idleStock) || 0;
          });

          // Roll up children's leaves
          childrenNodes.forEach((child) => {
            (child.children || []).forEach((cNode: any) => {
              const leafKey = cNode.name;
              if (!leafGroups[leafKey]) {
                leafGroups[leafKey] = {
                  lastQty: 0,
                  currentQty: 0,
                  pog: 0,
                  sellIn: 0,
                  sellOut: 0,
                  totalInv: 0,
                  idleStock: 0,
                  category: cNode.category,
                  pic: cNode.pic,
                };
              }
              leafGroups[leafKey].lastQty += Number(cNode.lastQty) || 0;
              leafGroups[leafKey].currentQty += Number(cNode.currentQty) || 0;
              leafGroups[leafKey].sellIn += Number(cNode.sellIn) || 0;
              leafGroups[leafKey].sellOut += Number(cNode.sellOut) || 0;
              leafGroups[leafKey].totalInv += Number(cNode.totalInv) || 0;
              leafGroups[leafKey].pog += Number(cNode.pog) || 0;
              leafGroups[leafKey].idleStock += Number(cNode.idleStock) || 0;
            });
          });

          let leaves = Object.entries(leafGroups).map(([leafName, val]) => ({
            name: leafName,
            isLeaf: true,
            ...(val as any),
          }));

          if (pogSubGroupBy === "channel" || pogSubGroupBy === "kiosk") {
            const catOrder: Record<string, number> = {
              Distributor: 1,
              R1: 2,
              R2: 3,
            };
            leaves.sort((a, b) => {
              const wA = catOrder[a.category] || 99;
              const wB = catOrder[b.category] || 99;
              if (wA !== wB) return wA - wB;
              return b.pog - a.pog;
            });
          } else {
            leaves.sort((a, b) => b.pog - a.pog);
          }

          finalChildren = leaves;
        }

        return {
          name,
          level: getMemberLevel(name, teamLevels, teamPositions, userData),
          children: finalChildren,
          isExpandable: finalChildren.length > 0,
          teamChildren: childrenNodes,
          ...transPog,
        };
      };

      const roots = rootMembers.map((root) => buildNode(root));
      roots.sort((a, b) => b.pog - a.pog);
      return roots.flatMap((node) => {
        if (cleanForMatch(node.name) === cleanForMatch(userData?.name)) {
          if (userLevel <= 3) {
            const nameClean = cleanForMatch(node.name);
            const myItems = filteredData.filter(
              (item) => cleanForMatch(item.pic) === nameClean,
            );

            const ownPog = {
              lastQty: 0,
              currentQty: 0,
              pog: 0,
              sellIn: 0,
              sellOut: 0,
              totalInv: 0,
              idleStock: 0,
            };
            myItems.forEach((item) => {
              const isDistributor = cleanForMatch(item.category) === "distributor";
              ownPog.lastQty += Number(item.lastQty) || 0;
              ownPog.currentQty += Number(item.currentQty) || 0;
              if (isDistributor) ownPog.sellIn += Number(item.sellIn) || 0;
              ownPog.sellOut += Number(item.sellOut) || 0;
              ownPog.totalInv += Number(item.totalInv) || 0;
              ownPog.idleStock += Number(item.idleStock) || 0;
              ownPog.pog += Number(item.pog) || 0;
            });

            let finalChildrenOfSelf = [];
            if (pogSubGroupBy !== "subordinate") {
              const leafGroups: Record<string, any> = {};
              myItems.forEach((item) => {
                const isDistributor = cleanForMatch(item.category) === "distributor";
                let leafKey = "Unknown";
                if (pogSubGroupBy === "channel" || pogSubGroupBy === "kiosk")
                  leafKey = item.kiosk || "Unknown Channel";
                else if (pogSubGroupBy === "hybrid")
                  leafKey = item.hybrid || "Unknown";
                else if (pogSubGroupBy === "area")
                  leafKey = item.area || "Unknown";
                else if (pogSubGroupBy === "category")
                  leafKey = item.category || "Unknown";
                else if (pogSubGroupBy === "crops")
                  leafKey = item.crops || "Unknown";

                if (!leafGroups[leafKey]) {
                  leafGroups[leafKey] = {
                    lastQty: 0,
                    currentQty: 0,
                    pog: 0,
                    sellIn: 0,
                    sellOut: 0,
                    totalInv: 0,
                    idleStock: 0,
                    category: item.category,
                  };
                }
                leafGroups[leafKey].lastQty += Number(item.lastQty) || 0;
                leafGroups[leafKey].currentQty += Number(item.currentQty) || 0;
                leafGroups[leafKey].sellIn += Number(item.sellIn) || 0;
                leafGroups[leafKey].sellOut += Number(item.sellOut) || 0;
                leafGroups[leafKey].totalInv += Number(item.totalInv) || 0;
                leafGroups[leafKey].pog += Number(item.pog) || 0;
                leafGroups[leafKey].idleStock += Number(item.idleStock) || 0;
              });

              let leaves = Object.entries(leafGroups).map(
                ([leafName, val]) => ({
                  name: leafName,
                  isLeaf: true,
                  ...(val as any),
                }),
              );

              if (pogSubGroupBy === "channel" || pogSubGroupBy === "kiosk") {
                const catOrder: Record<string, number> = {
                  Distributor: 1,
                  R1: 2,
                  R2: 3,
                };
                leaves.sort((a, b) => {
                  const wA = catOrder[a.category] || 99;
                  const wB = catOrder[b.category] || 99;
                  if (wA !== wB) return wA - wB;
                  return b.pog - a.pog;
                });
              } else {
                leaves.sort((a, b) => b.pog - a.pog);
              }
              finalChildrenOfSelf = leaves;
            }

            const selfNode = {
              name: node.name,
              level: node.level,
              children: finalChildrenOfSelf,
              isExpandable: finalChildrenOfSelf.length > 0,
              teamChildren: [],
              ...ownPog,
            };

            return [selfNode, ...(node.teamChildren || [])];
          } else {
            return node.teamChildren || [];
          }
        }
        return [node];
      });
    }

    const groups: Record<string, any> = {};
    filteredData.forEach((item) => {
      let key = item.hybrid || "Unknown";
      if (pogGroupBy === "area") key = item.area || "Unknown";
      else if (pogGroupBy === "category") key = item.category || "Unknown";
      else if (pogGroupBy === "crops") key = item.crops || "Unknown";
      else if (pogGroupBy === "area_crops") key = `${item.area || "Unknown"} - ${item.crops || "Unknown"}`;

      if (!groups[key]) {
        groups[key] = {
          name: key,
          lastQty: 0,
          currentQty: 0,
          pog: 0,
          sellIn: 0,
          sellOut: 0,
          totalInv: 0,
          idleStock: 0,
          childrenMap: {},
        };
        if (pogSubGroupBy === "subordinate") {
          activePogMembers.forEach((m) => {
            groups[key].childrenMap[m] = {
              name: m,
              lastQty: 0,
              currentQty: 0,
              pog: 0,
              sellIn: 0,
              sellOut: 0,
              totalInv: 0,
              idleStock: 0,
            };
          });
        }
      }
      const isDistributor = cleanForMatch(item.category) === "distributor";
      groups[key].lastQty += item.lastQty;
      groups[key].currentQty += item.currentQty;
      if (isDistributor) groups[key].sellIn += item.sellIn || 0;
      groups[key].sellOut += item.sellOut || 0;
      groups[key].totalInv += item.totalInv || 0;
      groups[key].idleStock += item.idleStock || 0;
      groups[key].pog += item.pog || 0;

      // Drill down
      let subKey = "Unknown";
      if (pogSubGroupBy === "channel") subKey = item.kiosk || "Unknown Channel";
      else if (pogSubGroupBy === "hybrid") subKey = item.hybrid || "Unknown";
      else if (pogSubGroupBy === "subordinate") {
        const matched = activePogMembers.find(
          (m) => cleanForMatch(m) === cleanForMatch(item.pic),
        );
        subKey = matched || item.pic || "Unknown";
      } else if (pogSubGroupBy === "area") subKey = item.area || "Unknown";
      else if (pogSubGroupBy === "category")
        subKey = item.category || "Unknown";
      else if (pogSubGroupBy === "crops") subKey = item.crops || "Unknown";
      else if (pogSubGroupBy === "area_crops") subKey = `${item.area || "Unknown"} - ${item.crops || "Unknown"}`;

      if (!groups[key].childrenMap[subKey]) {
        groups[key].childrenMap[subKey] = {
          name: subKey,
          lastQty: 0,
          currentQty: 0,
          pog: 0,
          sellIn: 0,
          sellOut: 0,
          totalInv: 0,
          idleStock: 0,
        };
        if (pogSubGroupBy === "channel") {
          groups[key].childrenMap[subKey].category = item.category;
          groups[key].childrenMap[subKey].pic = item.pic;
        }
      }
      groups[key].childrenMap[subKey].lastQty += item.lastQty;
      groups[key].childrenMap[subKey].currentQty += item.currentQty;
      groups[key].childrenMap[subKey].sellIn += item.sellIn || 0;
      groups[key].childrenMap[subKey].sellOut += item.sellOut || 0;
      groups[key].childrenMap[subKey].totalInv += item.totalInv || 0;
      groups[key].childrenMap[subKey].pog += item.pog;
      groups[key].childrenMap[subKey].idleStock += item.idleStock || 0;
    });

    const subMembersUplines: Record<string, string | null> = {};
    const subDirectSubsMap: Record<string, string[]> = {};
    let subRootMembers: string[] = [];
    if (pogSubGroupBy === "subordinate") {
      activePogMembers.forEach((m) => {
        subMembersUplines[m] = getUplineInTeam(
          m,
          activePogMembers,
          teamUpLines,
        );
      });
      subRootMembers = activePogMembers.filter(
        (m) => subMembersUplines[m] === null,
      );
      subRootMembers.sort((a, b) =>
        compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
      );

      activePogMembers.forEach((m) => {
        const upl = subMembersUplines[m];
        if (upl !== null) {
          const uplClean = cleanForMatch(upl);
          if (!subDirectSubsMap[uplClean]) {
            subDirectSubsMap[uplClean] = [];
          }
          subDirectSubsMap[uplClean].push(m);
        }
      });

      Object.keys(subDirectSubsMap).forEach((key) => {
        subDirectSubsMap[key].sort((a, b) =>
          compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
        );
      });
    }

    return Object.values(groups)
      .map((g: any) => {
        let children = [];

        if (pogSubGroupBy === "subordinate") {
          const buildTeamPogNode = (mName: string): any => {
            const mNameClean = cleanForMatch(mName);
            const directSubs = subDirectSubsMap[mNameClean] || [];

            const directKData = g.childrenMap[mName] || {
              name: mName,
              lastQty: 0,
              currentQty: 0,
              pog: 0,
              sellIn: 0,
              sellOut: 0,
              totalInv: 0,
              idleStock: 0,
            };
            const childrenNodes = directSubs
              .map((sub) => buildTeamPogNode(sub))
              .filter(Boolean);

            const transData = { ...directKData };
            childrenNodes.forEach((child) => {
              transData.lastQty += child.lastQty || 0;
              transData.currentQty += child.currentQty || 0;
              transData.sellIn += child.sellIn || 0;
              transData.sellOut += child.sellOut || 0;
              transData.totalInv += child.totalInv || 0;
              transData.pog += child.pog || 0;
              transData.idleStock += child.idleStock || 0;
            });

            const hasSelfActivity =
              directKData.lastQty > 0 ||
              directKData.currentQty > 0 ||
              directKData.sellIn > 0 ||
              directKData.sellOut > 0 ||
              directKData.totalInv > 0 ||
              directKData.pog !== 0;
            const hasChildrenActivity = childrenNodes.length > 0;

            if (
              pogGroupBy !== "subordinate" &&
              !hasSelfActivity &&
              !hasChildrenActivity
            ) {
              return null;
            }

            return {
              name: mName,
              level: getMemberLevel(mName, teamLevels, teamPositions, userData),
              children: childrenNodes,
              isExpandable: childrenNodes.length > 0,
              teamChildren: childrenNodes,
              ...transData,
            };
          };

          children = subRootMembers
            .map((m) => buildTeamPogNode(m))
            .filter(Boolean)
            .flatMap((node) => {
              if (cleanForMatch(node.name) === cleanForMatch(userData?.name)) {
                if (userLevel <= 3) {
                  const directKData = g.childrenMap[node.name] || {
                    lastQty: 0,
                    currentQty: 0,
                    pog: 0,
                    sellIn: 0,
                    sellOut: 0,
                    totalInv: 0,
                    idleStock: 0,
                  };
                  const selfNode = {
                    name: node.name,
                    level: node.level,
                    children: [],
                    isExpandable: false,
                    teamChildren: [],
                    ...directKData,
                  };
                  return [selfNode, ...(node.teamChildren || [])];
                } else {
                  return node.teamChildren || [];
                }
              }
              return [node];
            });

          const hasGroupActivity =
            g.lastQty > 0 ||
            g.currentQty > 0 ||
            g.sellIn > 0 ||
            g.sellOut > 0 ||
            g.totalInv > 0 ||
            g.pog !== 0;
          if (children.length === 0 && !hasGroupActivity) {
            return null;
          }
        } else {
          children = Object.values(g.childrenMap);

          if (pogSubGroupBy === "channel") {
            const catOrder: Record<string, number> = {
              Distributor: 1,
              R1: 2,
              R2: 3,
            };
            children.sort((a: any, b: any) => {
              const wA = catOrder[a.category] || 99;
              const wB = catOrder[b.category] || 99;
              if (wA !== wB) return wA - wB;
              return b.pog - a.pog;
            });
          } else {
            children.sort((a: any, b: any) => b.pog - a.pog);
          }
        }

        return {
          name: g.name,
          lastQty: g.lastQty,
          currentQty: g.currentQty,
          sellIn: g.sellIn,
          sellOut: g.sellOut,
          totalInv: g.totalInv,
          idleStock: g.idleStock,
          pog: g.pog,
          isExpandable: true,
          children,
        };
      })
      .filter(Boolean)
      .sort((a, b) => b.pog - a.pog);
  }, [
    pogDataProcessed,
    pogGroupBy,
    pogSubGroupBy,
    filterBelowCrop,
    activePogMembers,
    kiosks,
    userData,
    teamSubordinates,
    teamPositions,
    teamUpLines,
    teamLevels,
    teamAreas,
    employees,
    isBusinessAnalyst,
  ]);

  // Ekstrak data list crops unik untuk filter
  const availableCrops = useMemo(() => {
    const cropSet = new Set<string>();
    pogDataProcessed.forEach((item) => {
      const crop = String(item.crops || "").trim();
      if (crop) {
        cropSet.add(crop);
      }
    });
    return ["All", ...Array.from(cropSet).sort()];
  }, [pogDataProcessed]);

  // Filter team members based on selected crop filter
  const activeTeamMembers = useMemo(() => {
    if (!filterBelowCrop || filterBelowCrop === "All") return teamMembers;

    // Fallback mode: if group column info is empty or not in Google Sheet yet,
    // filter members by their matching recorded crop in workingData so they always see correct data.
    // A member is kept if they (or any of their subordinates recursively) have working data with the selected crop.
    const kiosksMapByCleanName: Record<string, any> = {};
    allKiosks.forEach((k) => {
      kiosksMapByCleanName[cleanForMatch(k.name)] = k;
    });
    const teamMembersMapByCleanName: Record<string, string> = {};
    teamMembers.forEach((m) => {
      teamMembersMapByCleanName[cleanForMatch(m)] = m;
    });
    const getTeamMemberMatch = (name: string): string | undefined => {
      const clean = cleanForMatch(name);
      if (teamMembersMapByCleanName[clean])
        return teamMembersMapByCleanName[clean];
      return teamMembers.find((m) => matchNames(m, name));
    };

    const matchedMembersWithCrop = new Set<string>();
    workingData.forEach((item) => {
      const itemCrop = String(item.crops || "")
        .trim()
        .toLowerCase();
      if (checkCropMatch(itemCrop, filterBelowCrop)) {
        const kClean = cleanForMatch(item.kiosk);
        const kioskInfo = kiosksMapByCleanName[kClean] || {};
        const rawPic = normalizeName(
          getResolvedPic(item.kiosk, kioskInfo.pic, item.user, item.lot, item.pic),
        );
        const picName = getDdaOfUser(
          rawPic,
          userData?.name,
          computedTeamProfiles,
        );
        const matchedMember = getTeamMemberMatch(picName);
        if (matchedMember) {
          matchedMembersWithCrop.add(cleanForMatch(matchedMember));
        }
      }
    });

    // Propagate active status upwards
    const activeSet = new Set<string>();
    activeSet.add(cleanForMatch(userData?.name)); // Always keep self

    teamMembers.forEach((m) => {
      const mClean = cleanForMatch(m);
      if (matchedMembersWithCrop.has(mClean)) {
        activeSet.add(mClean);
        let current = m;
        for (let i = 0; i < 15; i++) {
          // Max depth protection
          const upRaw = getFromRecord<string>(teamUpLines, current);
          if (!upRaw) break;

          // Match against actual teamMembers to handle partial names
          const up = getTeamMemberMatch(upRaw as string) || (upRaw as string);

          const upClean = cleanForMatch(up);
          if (activeSet.has(upClean)) {
            // Already processed this branch upwards, can break early
            break;
          }
          activeSet.add(upClean);
          current = up;
        }
      }
    });

    return teamMembers.filter((m) => activeSet.has(cleanForMatch(m)));
  }, [
    teamMembers,
    filterBelowCrop,
    employees,
    userData,
    workingData,
    kiosks,
    teamUpLines,
  ]);

  const dynamicWorkingData = useMemo(() => {
    const monthsKeys = [
      "jan",
      "feb",
      "mar",
      "apr",
      "mei",
      "jun",
      "jul",
      "ags",
      "sep",
      "okt",
      "nov",
      "des",
    ];
    let targetMonthIdx = new Date().getMonth();
    let targetYear = new Date().getFullYear();

    if (filterBelowMonth && filterBelowMonth !== "All") {
      const parts = filterBelowMonth.split(" ");
      const mIdx = INDO_MONTHS.indexOf(parts[0]);
      if (mIdx !== -1) {
        targetMonthIdx = mIdx;
        targetYear = parseInt(parts[1], 10);
      }
    }

    const currentMonthKey = monthsKeys[targetMonthIdx];
    const prevMonthIdx = targetMonthIdx === 0 ? 11 : targetMonthIdx - 1;
    const prevMonthKey = monthsKeys[prevMonthIdx];

    return workingData.map((d: any, index: number) => {
      const isUpdated = d["upd_" + currentMonthKey] && d["upd_" + currentMonthKey] !== "";

      const currVal = isUpdated
        ? safeParseFloat(d[currentMonthKey])
        : safeParseFloat(d.originalStock !== undefined ? d.originalStock : d.stock);

      const prevVal = d[prevMonthKey] !== undefined ? safeParseFloat(d[prevMonthKey]) : 0;

      let finalCondition = d.condition || "tetap";
      if (prevVal > 0) {
        if (currVal > prevVal) finalCondition = "bertambah";
        else if (currVal < prevVal) finalCondition = "berkurang";
        else finalCondition = "tetap";
      } else {
        if (currVal > 0) {
          finalCondition =
            d.condition === "new" || d.condition === "baru" || d.isNew
              ? "new"
              : "bertambah";
        } else {
          finalCondition = "tetap";
        }
      }

      return {
        ...d,
        stock: currVal,
        condition: finalCondition,
      };
    });
  }, [workingData, filterBelowMonth]);

  const summaryData = useMemo(() => {
    const picToUplineMap: Record<string, string> = {};
    allKiosks.forEach((k) => {
      const p = String(k.pic || "").trim();
      const u = String(k.upline || "").trim();
      if (p && u) {
        picToUplineMap[p.toLowerCase()] = normalizeName(u);
      }
    });
    picToUplineMap["listianto"] = "AGUS HERDIANTO";

    // Pre-compute maps to optimize from O(N * M) to O(N + M)
    const kiosksMapByCleanName: Record<string, any> = {};
    allKiosks.forEach((k) => {
      kiosksMapByCleanName[cleanForMatch(k.name)] = k;
    });

    const teamMembersMapByCleanName: Record<string, string> = {};
    activeTeamMembers.forEach((m) => {
      teamMembersMapByCleanName[cleanForMatch(m)] = m;
    });
    const getTeamMemberMatch = (name: string): string | undefined => {
      const clean = cleanForMatch(name);
      if (teamMembersMapByCleanName[clean])
        return teamMembersMapByCleanName[clean];
      return activeTeamMembers.find((m) => matchNames(m, name));
    };

    const parseDate = (timestamp: any) => {
      if (!timestamp) return null;
      let d;
      if (typeof timestamp === "string" && timestamp.includes("/")) {
        const parts = timestamp.split(/[\s/:]+/);
        if (parts.length >= 3) {
          d = new Date(
            `${parts[2]}-${parts[1]}-${parts[0]}T${parts[3] || "00"}:${parts[4] || "00"}:${parts[5] || "00"}`,
          );
        } else d = new Date(timestamp);
      } else {
        d = new Date(timestamp);
      }
      return d && !isNaN(d.getTime()) ? d : null;
    };

    const enrichedData = dynamicWorkingData.map((item) => {
      const d = parseDate(item.timestamp);
      const itemMonth = d ? d.getMonth() : null;
      const itemYear = d ? d.getFullYear() : null;

      const kClean = cleanForMatch(item.kiosk);
      const kioskInfo =
        kiosksMapByCleanName[kClean] ||
        kiosks.find((k) => matchNames(k.name, item.kiosk)) ||
        {};
      const rawPic = normalizeName(
        getResolvedPic(item.kiosk, kioskInfo.pic, item.user, item.lot, item.pic),
      );
      const pic = getDdaOfUser(rawPic, userData?.name, computedTeamProfiles);
      let upline = normalizeName(String(kioskInfo.upline || ""));
      if (pic.toLowerCase() === "listianto") {
        upline = "AGUS HERDIANTO";
      } else {
        const matchedMember = getTeamMemberMatch(pic);
        const foundUp = getFromRecord<string>(
          teamUpLines,
          matchedMember || pic,
        );
        if (foundUp) {
          upline = normalizeName(foundUp);
        } else if (!upline && pic !== "Unknown") {
          const foundUpline = picToUplineMap[pic.toLowerCase()];
          if (foundUpline) upline = normalizeName(foundUpline);
        }
      }
      let area = "-";
      if (item.area && String(item.area).trim() !== "") {
        area = String(item.area).trim();
      } else {
        const cleanPic = cleanForMatch(pic);
        const matchedMember = getTeamMemberMatch(pic);
        const foundArea = getFromRecord<string>(teamAreas, matchedMember || pic);
        if (foundArea) {
          area = foundArea;
        } else if (cleanPic === cleanForMatch(userData?.name)) {
          area = userData?.area || "-";
        }
      }

      const category = getValidCategory(
        item.category,
        kioskInfo.category,
        item.kiosk,
        workingCategoryMap,
      );

      let cluster = "Uncategorized";
      const aging = Number(item.aging);
      if (!isNaN(aging)) {
        if (aging <= 2) cluster = "0-2";
        else if (aging <= 4) cluster = "2-4";
        else if (aging <= 6) cluster = "4-6";
        else if (aging <= 9) cluster = "6-9";
        else if (aging <= 12) cluster = "9-12";
        else cluster = ">12";
      }
      const crops =
        item.crops && String(item.crops).trim() !== ""
          ? item.crops
          : "Uncategorized Crops";
      return {
        ...item,
        pic,
        upline,
        area,
        rawArea: item.area || "-",
        category,
        cluster,
        hybrid: item.hybrid || "Unknown",
        crops,
        itemMonth,
        itemYear,
      };
    });

    enrichedSummaryDataRef.current = enrichedData;

    const teamMembersCleanSet = new Set(
      activeTeamMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );
    

    // Filter Team AND Filter Crops (Using Set.has for O(1) performance)
    const currentMonth = new Date().getMonth();
    const currentYear = new Date().getFullYear();

        const teamData = enrichedData.filter((item) => {
      if (isBusinessAnalyst) {
        return checkCropMatch(item.crops, filterBelowCrop);
      }
      const picClean = cleanForMatch(item.pic);
      let isTeamMember =
        teamMembersCleanSet.has(picClean);
        
      if (!isTeamMember && computedTeamProfiles) {
        const emp = computedTeamProfiles[picClean];
        if (emp && emp.name) {
          const empNameClean = cleanForMatch(emp.name);
          if (teamMembersCleanSet.has(empNameClean)) {
            isTeamMember = true;
          }
        }
      }

      const isCropMatch = checkCropMatch(item.crops, filterBelowCrop);
      return isTeamMember && isCropMatch;
    });

    if (summaryGroupBy === "subordinate") {
      const teamMembersUplines: Record<string, string | null> = {};
      const directSubsMap: Record<string, string[]> = {};
      activeTeamMembers.forEach((m) => {
        teamMembersUplines[m] = getUplineInTeam(
          m,
          activeTeamMembers,
          teamUpLines,
        );
      });

      const rootMembers = activeTeamMembers.filter((m) => {
        return teamMembersUplines[m] === null;
      });
      rootMembers.sort((a, b) =>
        compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
      );

      activeTeamMembers.forEach((m) => {
        const upl = teamMembersUplines[m];
        if (upl !== null) {
          const uplClean = cleanForMatch(upl);
          if (!directSubsMap[uplClean]) {
            directSubsMap[uplClean] = [];
          }
          directSubsMap[uplClean].push(m);
        }
      });

      Object.keys(directSubsMap).forEach((key) => {
        directSubsMap[key].sort((a, b) =>
          compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
        );
      });

      const buildNode = (name: string): any => {
        const nameClean = cleanForMatch(name);
        const directSubs = directSubsMap[nameClean] || [];
        const myItems = teamData.filter(
          (item) => cleanForMatch(item.pic) === nameClean,
        );
        const childrenNodes = directSubs.map((sub) => buildNode(sub));

        const ownData = {
          "0-2": 0,
          "2-4": 0,
          "4-6": 0,
          "6-9": 0,
          "9-12": 0,
          ">12": 0,
          Uncategorized: 0,
          total: 0,
        };
        myItems.forEach((item) => {
          const stock = Number(item.stock) || 0;
          if (ownData[item.cluster] !== undefined)
            ownData[item.cluster] += stock;
          ownData.total += stock;
        });

        const transData = { ...ownData };
        childrenNodes.forEach((child) => {
          Object.keys(ownData).forEach((key) => {
            if (key !== "total") {
              transData[key] += child[key] || 0;
            }
          });
          transData.total += child.total || 0;
        });

        let finalChildren = [];

        if (summarySubGroupBy === "subordinate") {
          finalChildren = [...childrenNodes];
        } else {
          const leafGroups: Record<string, any> = {};

          // Own items
          myItems.forEach((item) => {
            let leafKey = "Unknown";
            if (
              summarySubGroupBy === "channel" ||
              summarySubGroupBy === "kiosk"
            )
              leafKey = item.kiosk || "Unknown Channel";
            else if (summarySubGroupBy === "hybrid") leafKey = item.hybrid;
            else if (summarySubGroupBy === "area") leafKey = item.area;
            else if (summarySubGroupBy === "category") leafKey = item.category;
            else if (summarySubGroupBy === "crops") leafKey = item.crops;

            if (!leafGroups[leafKey]) {
              leafGroups[leafKey] = {
                "0-2": 0,
                "2-4": 0,
                "4-6": 0,
                "6-9": 0,
                "9-12": 0,
                ">12": 0,
                Uncategorized: 0,
                total: 0,
                category: item.category,
              };
            }
            const stock = Number(item.stock) || 0;
            if (leafGroups[leafKey][item.cluster] !== undefined) {
              leafGroups[leafKey][item.cluster] += stock;
            }
            leafGroups[leafKey].total += stock;
          });

          // Roll up children's leaves
          childrenNodes.forEach((child) => {
            (child.children || []).forEach((cNode: any) => {
              const leafKey = cNode.name;
              if (!leafGroups[leafKey]) {
                leafGroups[leafKey] = {
                  "0-2": 0,
                  "2-4": 0,
                  "4-6": 0,
                  "6-9": 0,
                  "9-12": 0,
                  ">12": 0,
                  Uncategorized: 0,
                  total: 0,
                  category: cNode.category,
                };
              }
              Object.keys(cNode).forEach((k) => {
                if (
                  k !== "name" &&
                  k !== "isLeaf" &&
                  k !== "category" &&
                  k !== "total" &&
                  k !== "isExpandable" &&
                  k !== "children" &&
                  k !== "level" &&
                  k !== "selectedTotal"
                ) {
                  leafGroups[leafKey][k] =
                    (leafGroups[leafKey][k] || 0) + (cNode[k] || 0);
                }
              });
              leafGroups[leafKey].total += cNode.total || 0;
            });
          });

          let leaves = Object.entries(leafGroups).map(([leafName, val]) => ({
            name: leafName,
            isLeaf: true,
            ...(val as any),
          }));

          if (
            summarySubGroupBy === "channel" ||
            summarySubGroupBy === "kiosk"
          ) {
            const catOrder: Record<string, number> = {
              Distributor: 1,
              R1: 2,
              R2: 3,
            };
            leaves.sort((a, b) => {
              const wA = catOrder[a.category] || 99;
              const wB = catOrder[b.category] || 99;
              if (wA !== wB) return wA - wB;
              return b.total - a.total;
            });
          } else {
            leaves.sort((a, b) => b.total - a.total);
          }

          finalChildren = leaves;
        }

        return {
          name,
          level: getMemberLevel(name, teamLevels, teamPositions, userData),
          children: finalChildren,
          isExpandable: finalChildren.length > 0,
          teamChildren: childrenNodes,
          ...transData,
        };
      };

      const roots = rootMembers.map((root) => buildNode(root));
      roots.sort((a, b) => b.total - a.total);
      return roots.flatMap((node) => {
        if (cleanForMatch(node.name) === cleanForMatch(userData?.name)) {
          if (userLevel <= 3) {
            const nameClean = cleanForMatch(node.name);
            const myItems = teamData.filter(
              (item) => cleanForMatch(item.pic) === nameClean,
            );

            const ownData = {
              "0-2": 0,
              "2-4": 0,
              "4-6": 0,
              "6-9": 0,
              "9-12": 0,
              ">12": 0,
              Uncategorized: 0,
              total: 0,
            };
            myItems.forEach((item) => {
              const stock = Number(item.stock) || 0;
              if (ownData[item.cluster] !== undefined)
                ownData[item.cluster] += stock;
              ownData.total += stock;
            });

            let finalChildrenOfSelf = [];
            if (summarySubGroupBy !== "subordinate") {
              const leafGroups: Record<string, any> = {};
              myItems.forEach((item) => {
                let leafKey = "Unknown";
                if (
                  summarySubGroupBy === "channel" ||
                  summarySubGroupBy === "kiosk"
                )
                  leafKey = item.kiosk || "Unknown Channel";
                else if (summarySubGroupBy === "hybrid") leafKey = item.hybrid;
                else if (summarySubGroupBy === "area") leafKey = item.area;
                else if (summarySubGroupBy === "category")
                  leafKey = item.category;
                else if (summarySubGroupBy === "crops") leafKey = item.crops;

                if (!leafGroups[leafKey]) {
                  leafGroups[leafKey] = {
                    "0-2": 0,
                    "2-4": 0,
                    "4-6": 0,
                    "6-9": 0,
                    "9-12": 0,
                    ">12": 0,
                    Uncategorized: 0,
                    total: 0,
                    category: item.category,
                  };
                }
                const stock = Number(item.stock) || 0;
                if (leafGroups[leafKey][item.cluster] !== undefined) {
                  leafGroups[leafKey][item.cluster] += stock;
                }
                leafGroups[leafKey].total += stock;
              });

              let leaves = Object.entries(leafGroups).map(
                ([leafName, val]) => ({
                  name: leafName,
                  isLeaf: true,
                  ...(val as any),
                }),
              );

              if (
                summarySubGroupBy === "channel" ||
                summarySubGroupBy === "kiosk"
              ) {
                const catOrder: Record<string, number> = {
                  Distributor: 1,
                  R1: 2,
                  R2: 3,
                };
                leaves.sort((a, b) => {
                  const wA = catOrder[a.category] || 99;
                  const wB = catOrder[b.category] || 99;
                  if (wA !== wB) return wA - wB;
                  return b.total - a.total;
                });
              } else {
                leaves.sort((a, b) => b.total - a.total);
              }
              finalChildrenOfSelf = leaves;
            }

            const selfNode = {
              name: node.name,
              level: node.level,
              children: finalChildrenOfSelf,
              isExpandable: finalChildrenOfSelf.length > 0,
              teamChildren: [],
              ...ownData,
            };

            return [selfNode, ...(node.teamChildren || [])];
          } else {
            return node.teamChildren || [];
          }
        }
        return [node];
      });
    }

    const groups: Record<string, any> = {};
    teamData.forEach((item) => {
      let key = "Unknown";
      if (summaryGroupBy === "hybrid") key = item.hybrid;
      else if (summaryGroupBy === "area") key = item.area;
      else if (summaryGroupBy === "category") key = item.category;
      else if (summaryGroupBy === "crops") key = item.crops;
      else if (summaryGroupBy === "area_crops") key = `${item.area || "Unknown"} - ${item.crops || "Unknown"}`;

      if (!groups[key]) {
        groups[key] = {
          "0-2": 0,
          "2-4": 0,
          "4-6": 0,
          "6-9": 0,
          "9-12": 0,
          ">12": 0,
          Uncategorized: 0,
          total: 0,
          childrenMap: {},
        };
        if (summarySubGroupBy === "subordinate") {
          activeTeamMembers.forEach((m) => {
            groups[key].childrenMap[m] = {
              "0-2": 0,
              "2-4": 0,
              "4-6": 0,
              "6-9": 0,
              "9-12": 0,
              ">12": 0,
              Uncategorized: 0,
              total: 0,
            };
          });
        }
      }
      const stock = Number(item.stock) || 0;
      if (groups[key][item.cluster] !== undefined) {
        groups[key][item.cluster] += stock;
      }
      groups[key].total += stock;

      // Sub-item
      let subKey = "Unknown";
      if (summarySubGroupBy === "channel")
        subKey = item.kiosk || "Unknown Channel";
      else if (summarySubGroupBy === "hybrid") subKey = item.hybrid;
      else if (summarySubGroupBy === "subordinate") {
        const matched = activeTeamMembers.find(
          (m) => cleanForMatch(m) === cleanForMatch(item.pic),
        );
        subKey = matched || item.pic;
      } else if (summarySubGroupBy === "area") subKey = item.area;
      else if (summarySubGroupBy === "category") subKey = item.category;
      else if (summarySubGroupBy === "crops") subKey = item.crops;
      else if (summarySubGroupBy === "area_crops") subKey = `${item.area || "Unknown"} - ${item.crops || "Unknown"}`;

      if (!groups[key].childrenMap[subKey]) {
        groups[key].childrenMap[subKey] = {
          "0-2": 0,
          "2-4": 0,
          "4-6": 0,
          "6-9": 0,
          "9-12": 0,
          ">12": 0,
          Uncategorized: 0,
          total: 0,
        };
        if (summarySubGroupBy === "channel") {
          groups[key].childrenMap[subKey].category = item.category;
          groups[key].childrenMap[subKey].pic = item.pic;
        }
      }
      if (groups[key].childrenMap[subKey][item.cluster] !== undefined) {
        groups[key].childrenMap[subKey][item.cluster] += stock;
      }
      groups[key].childrenMap[subKey].total += stock;
    });

    const sumMembersUplines: Record<string, string | null> = {};
    const sumDirectSubsMap: Record<string, string[]> = {};
    let sumRootMembers: string[] = [];
    if (summarySubGroupBy === "subordinate") {
      activeTeamMembers.forEach((m) => {
        sumMembersUplines[m] = getUplineInTeam(
          m,
          activeTeamMembers,
          teamUpLines,
        );
      });
      sumRootMembers = activeTeamMembers.filter(
        (m) => sumMembersUplines[m] === null,
      );
      sumRootMembers.sort((a, b) =>
        compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
      );

      activeTeamMembers.forEach((m) => {
        const upl = sumMembersUplines[m];
        if (upl !== null) {
          const uplClean = cleanForMatch(upl);
          if (!sumDirectSubsMap[uplClean]) {
            sumDirectSubsMap[uplClean] = [];
          }
          sumDirectSubsMap[uplClean].push(m);
        }
      });

      Object.keys(sumDirectSubsMap).forEach((key) => {
        sumDirectSubsMap[key].sort((a, b) =>
          compareMembersByLevel(a, b, teamLevels, teamPositions, userData),
        );
      });
    }

    return Object.entries(groups)
      .map(([name, counts]) => {
        const countsVal = counts as any;
        let children = [];

        if (summarySubGroupBy === "subordinate") {
          const buildTeamNode = (mName: string): any => {
            const mNameClean = cleanForMatch(mName);
            const directSubs = sumDirectSubsMap[mNameClean] || [];

            const directKData = countsVal.childrenMap[mName] || {
              "0-2": 0,
              "2-4": 0,
              "4-6": 0,
              "6-9": 0,
              "9-12": 0,
              ">12": 0,
              Uncategorized: 0,
              total: 0,
            };
            const childrenNodes = directSubs.map((sub) => buildTeamNode(sub));

            const transData = { ...directKData };
            childrenNodes.forEach((child) => {
              Object.keys(directKData).forEach((k) => {
                if (k !== "total" && k !== "category") {
                  transData[k] = (transData[k] || 0) + (child[k] || 0);
                }
              });
              transData.total = (transData.total || 0) + (child.total || 0);
            });

            return {
              name: mName,
              level: getMemberLevel(mName, teamLevels, teamPositions, userData),
              children: childrenNodes,
              isExpandable: childrenNodes.length > 0,
              teamChildren: childrenNodes,
              ...transData,
            };
          };

          children = sumRootMembers
            .map((m) => buildTeamNode(m))
            .flatMap((node) => {
              if (cleanForMatch(node.name) === cleanForMatch(userData?.name)) {
                if (userLevel <= 3) {
                  const directKData = countsVal.childrenMap[node.name] || {
                    "0-2": 0,
                    "2-4": 0,
                    "4-6": 0,
                    "6-9": 0,
                    "9-12": 0,
                    ">12": 0,
                    Uncategorized: 0,
                    total: 0,
                  };
                  const selfNode = {
                    name: node.name,
                    level: node.level,
                    children: [],
                    isExpandable: false,
                    teamChildren: [],
                    ...directKData,
                  };
                  return [selfNode, ...(node.teamChildren || [])];
                } else {
                  return node.teamChildren || [];
                }
              }
              return [node];
            });
        } else {
          children = Object.entries(countsVal.childrenMap || {}).map(
            ([subKeyName, kData]) => ({
              name: subKeyName,
              ...(kData as any),
            }),
          );

          if (summarySubGroupBy === "channel") {
            const catOrder: Record<string, number> = {
              Distributor: 1,
              R1: 2,
              R2: 3,
            };
            children.sort((a, b) => {
              const wA = catOrder[a.category] || 99;
              const wB = catOrder[b.category] || 99;
              if (wA !== wB) return wA - wB;
              return b.total - a.total;
            });
          } else {
            children.sort((a, b) => b.total - a.total);
          }
        }

        const { childrenMap, ...rest } = countsVal;
        return { name, isExpandable: true, children, ...rest };
      })
      .sort((a, b) => b.total - a.total);
  }, [
    dynamicWorkingData,
    kiosks,
    activeTeamMembers,
    summaryGroupBy,
    summarySubGroupBy,
    userData,
    filterBelowCrop,
    teamSubordinates,
    teamAreas,
    teamUpLines,
    teamPositions,
    teamLevels,
    filterBelowMonth,
    isBusinessAnalyst,
  ]);

  const filteredSummaryData = useMemo(() => {
    if (selectedClusters.length === 0) return [];

    if (summaryGroupBy === "subordinate") {
      const filterNode = (node: any): any => {
        const parentSelectedTotal = selectedClusters.reduce(
          (sum, c) => sum + (node[c] || 0),
          0,
        );

        const filteredChildren = (node.children || [])
          .map((child: any) => {
            if (child.isLeaf) {
              const leafSelectedTotal = selectedClusters.reduce(
                (sum, c) => sum + (child[c] || 0),
                0,
              );
              return { ...child, selectedTotal: leafSelectedTotal };
            } else {
              return filterNode(child);
            }
          })
          .filter(Boolean);

        // ALWAYS return the teammate node (do not return null if 0 total) to support "tampilkan semua nama" from sheet employee
        return {
          ...node,
          selectedTotal: parentSelectedTotal,
          children: filteredChildren,
        };
      };

      // Do not filter(Boolean) here as we want to preserve all employee roots.
      return summaryData.map((row) => filterNode(row));
    }

    const filterSubordinateNode = (node: any): any => {
      const nodeSelectedTotal = selectedClusters.reduce(
        (sum, c) => sum + (node[c] || 0),
        0,
      );
      const filteredChildren = (node.children || [])
        .map((c: any) => filterSubordinateNode(c))
        .filter(Boolean);

      // Hanya tampilkan employee yang berkorelasi saja jika group by bukan 'subordinate' (team)
      // Karyawan dianggap berkorelasi jika memiliki stock > 0 atau salah satu bawahannya memiliki stock > 0
      if (
        summaryGroupBy !== "subordinate" &&
        nodeSelectedTotal === 0 &&
        filteredChildren.length === 0
      ) {
        return null;
      }

      return {
        ...node,
        selectedTotal: nodeSelectedTotal,
        children: filteredChildren,
      };
    };

    return summaryData
      .map((row) => {
        if (row.isExpandable) {
          if (summarySubGroupBy === "subordinate") {
            const filteredChildren = row.children
              .map((child) => filterSubordinateNode(child))
              .filter(Boolean);
            const parentSelectedTotal = selectedClusters.reduce(
              (sum, c) => sum + (row[c] || 0),
              0,
            );
            if (filteredChildren.length === 0 && parentSelectedTotal === 0)
              return null;
            return {
              ...row,
              children: filteredChildren,
              selectedTotal: parentSelectedTotal,
            };
          }

          const filteredChildren = row.children
            .map((child) => {
              const childSelectedTotal = selectedClusters.reduce(
                (sum, c) => sum + (child[c] || 0),
                0,
              );
              // If sub-group is Team/Subordinate or child is a leaf/hybrid item, show them even with 0 total
              if (
                summarySubGroupBy === "subordinate" ||
                childSelectedTotal > 0 ||
                child.isLeaf ||
                child.hybrid
              ) {
                return { ...child, selectedTotal: childSelectedTotal };
              }
              return null;
            })
            .filter(Boolean);

          const parentSelectedTotal = selectedClusters.reduce(
            (sum, c) => sum + (row[c] || 0),
            0,
          );
          // If sub-group is Team/Subordinate or has children, show parent group
          if (
            summarySubGroupBy === "subordinate" ||
            filteredChildren.length > 0 ||
            parentSelectedTotal > 0
          ) {
            return {
              ...row,
              children: filteredChildren,
              selectedTotal: parentSelectedTotal,
            };
          }
          return null;
        } else {
          const rowSelectedTotal = selectedClusters.reduce(
            (sum, c) => sum + (row[c] || 0),
            0,
          );
          return { ...row, selectedTotal: rowSelectedTotal };
          return null;
        }
      })
      .filter(Boolean);
  }, [summaryData, selectedClusters, summaryGroupBy, summarySubGroupBy]);

  const totalSummary = useMemo(() => {
    const totals = {
      "0-2": 0,
      "2-4": 0,
      "4-6": 0,
      "6-9": 0,
      "9-12": 0,
      ">12": 0,
      Uncategorized: 0,
      selectedTotal: 0,
    };
    summaryData.forEach((row) => {
      ALL_CLUSTER_KEYS.forEach((c) => {
        totals[c] += row[c] || 0;
      });
      totals.selectedTotal += selectedClusters.reduce(
        (sum, c) => sum + (row[c] || 0),
        0,
      );
    });
    return totals;
  }, [summaryData, selectedClusters, ALL_CLUSTER_KEYS]);

  const overviewUseMt = useMemo(() => {
    if (filterBelowCrop === "Vegetables") return false;
    const maxVal = Math.max(
      overviewStats.totalOpeningStock || 0,
      overviewStats.totalCurrentStock || 0,
      overviewStats.totalSellIn || 0,
      overviewStats.totalIdleStock || 0,
      overviewStats.totalSellOut || 0,
      totalSummary.selectedTotal || 0
    );
    return maxVal >= 1000;
  }, [overviewStats, totalSummary, filterBelowCrop]);

  const handleDownloadSummaryExcel = () => {
    const teamMembersCleanSet = new Set(
      activeTeamMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );
    
    const currentMonth = new Date().getMonth();
    const currentYear = new Date().getFullYear();

    // Filter enriched data to match team members and active crop + month + chosen aging clusters
    const displayedRawItems = (enrichedSummaryDataRef.current || []).filter(
      (item) => {
        const isTeamMember = isBusinessAnalyst || teamMembersCleanSet.has(cleanForMatch(item.pic));
        const isCropMatch = checkCropMatch(item.crops, filterBelowCrop);
        const isClusterMatch = selectedClusters.includes(item.cluster);

        return isTeamMember && isCropMatch && isClusterMatch;
      },
    );

    if (displayedRawItems.length === 0) {
      alert("Tidak ada data untuk di-download.");
      return;
    }

    // Map each item exactly to match the custom requested headers in Indonesian and English
    const formattedRows = displayedRawItems.map((item) => {
      return {
        tgl: item.timestamp || "",
        province:
          item.province ||
          getFromRecord<string>(teamProvinces, item.pic) ||
          "-",
        crops: item.crops || "",
        checker: item.user || item.pic || "",
        channel: item.kiosk || "",
        category: item.category || "",
        hybrids: item.hybrid || "",
        "lot no": item.lot || "",
        qty: Number(item.stock) || 0,
        "usia stock": item.aging !== undefined ? String(item.aging) : "",
        "cluster aging": item.cluster || "",
        "shipping date":
          item.drDate ||
          item.shipping_date ||
          item.shippingDate ||
          item.dr_date ||
          "",
        "exp date": item.expired || "",
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(formattedRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Working Data Sheet");

    // Pre-configure column widths for pristine visual alignments to match the new headers
    const colWidths = [
      { wch: 22 }, // tgl
      { wch: 18 }, // province
      { wch: 18 }, // crops
      { wch: 25 }, // checker
      { wch: 30 }, // channel
      { wch: 18 }, // category
      { wch: 28 }, // hybrids
      { wch: 18 }, // lot no
      { wch: 12 }, // qty
      { wch: 18 }, // usia stock
      { wch: 15 }, // cluster aging
      { wch: 20 }, // shipping date
      { wch: 22 }, // exp date
    ];
    worksheet["!cols"] = colWidths;

    XLSX.writeFile(
      workbook,
      `Stock_RADAR_AWBA_${new Date().toISOString().split("T")[0]}.xlsx`,
    );
  };

  const totalPog = useMemo(() => {
    const teamMembersCleanSet = new Set(
      activePogMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );

    const filteredData = pogDataProcessed.filter((item) => {
      const isTeamMember = isBusinessAnalyst || teamMembersCleanSet.has(cleanForMatch(item.pic));
      const isCropMatch = checkCropMatch(item.crops, filterBelowCrop);
      return isTeamMember && isCropMatch;
    });

    const total = {
      lastQty: 0,
      currentQty: 0,
      sellIn: 0,
      sellOut: 0,
      totalInv: 0,
      pog: 0,
      idleStock: 0,
    };

    filteredData.forEach((item) => {
      const isDistributor = cleanForMatch(item.category) === "distributor";
      total.lastQty += Number(item.lastQty) || 0;
      total.currentQty += Number(item.currentQty) || 0;
      if (isDistributor) total.sellIn += Number(item.sellIn) || 0;
      total.sellOut += Number(item.sellOut) || 0;
      total.totalInv += Number(item.totalInv) || 0;
      total.idleStock += Number(item.idleStock) || 0;
    });
    total.pog = total.lastQty + total.sellIn - total.currentQty;
    
    return total;
  }, [pogDataProcessed, activePogMembers, filterBelowCrop, pogGroupBy, userLevel, isBusinessAnalyst]);

  const totalSummaryPog = useMemo(() => {
    const teamMembersCleanSet = new Set(
      activeTeamMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );
    

    const filteredData = pogDataProcessed.filter((item) => {
      const isTeamMember = isBusinessAnalyst || teamMembersCleanSet.has(cleanForMatch(item.pic));
      const isCropMatch = checkCropMatch(item.crops, filterBelowCrop);
      return isTeamMember && isCropMatch;
    });

    const total = {
      lastQty: 0,
      currentQty: 0,
      sellIn: 0,
      sellOut: 0,
      totalInv: 0,
      pog: 0,
      idleStock: 0,
    };
    filteredData.forEach((row) => {
      const isDistributor = cleanForMatch(row.category) === "distributor";
      total.lastQty += row.lastQty || 0;
      total.currentQty += row.currentQty || 0;
      if (isDistributor) total.sellIn += row.sellIn || 0;
      total.sellOut += row.sellOut || 0;
      total.totalInv += row.totalInv || 0;
      total.idleStock += row.idleStock || 0;
    });
    total.pog = total.lastQty + total.sellIn - total.currentQty;
    return total;
  }, [pogDataProcessed, activeTeamMembers, filterBelowCrop, isBusinessAnalyst]);

  const totalTeamStats = useMemo(() => {
    let total = 0;
    let visited = 0;
    teamStats.forEach((s) => {
      total += s.total || 0;
      visited += s.totalVisited || 0;
    });
    const percentage =
      total > 0 ? Math.min(100, Math.round((visited / total) * 100)) : 0;
    return { visited, total, percentage };
  }, [teamStats]);

  const activeVisitStats = useMemo(() => {
    const isAll =
      !mappingPic ||
      cleanForMatch(mappingPic) === "allteam" ||
      cleanForMatch(mappingPic) === "all_team";
    if (isAll) {
      return totalTeamStats;
    }
    const picStat = getStatsForPic(mappingPic);
    return {
      visited: picStat.visited,
      total: picStat.total,
      percentage: picStat.percentage,
    };
  }, [mappingPic, totalTeamStats, teamStats]);

  const grandTotalStats = useMemo(() => {
    const teamMembersCleanSet = new Set(
      activeTeamMembers.map((m) => cleanForMatch(m)).filter(Boolean),
    );
    

    const displayedRawItems = (enrichedSummaryDataRef.current || []).filter(
      (item) => {
        const isTeamMember = isBusinessAnalyst || teamMembersCleanSet.has(cleanForMatch(item.pic));
        const isCropMatch = checkCropMatch(item.crops, filterBelowCrop);
        const isClusterMatch = selectedClusters.includes(item.cluster);

        return isTeamMember && isCropMatch && isClusterMatch;
      },
    );

    const groupData: Record<
      string,
      { total: number; clusters: Record<string, number> }
    > = {};
    displayedRawItems.forEach((item) => {
      let g = "Unknown";
      if (grandTotalViewBy === "hybrid") {
        g = item.hybrid || "Unknown";
      } else if (grandTotalViewBy === "area") {
        const itemProv = item.province || getFromRecord<string>(teamProvinces, item.pic) || (item.area && item.area !== "-" ? item.area : "") || (item.rawArea && item.rawArea !== "-" ? item.rawArea : "");
        g = (itemProv && itemProv !== "-") ? itemProv : "Unknown";
      }
      
      if (groupData[g] === undefined) {
        groupData[g] = { total: 0, clusters: {} };
      }
      const qty = Number(item.stock) || 0;
      groupData[g].total += qty;

      const cluster = item.cluster;
      if (groupData[g].clusters[cluster] === undefined) {
        groupData[g].clusters[cluster] = 0;
      }
      groupData[g].clusters[cluster] += qty;
    });

    const sortedGroups = Object.entries(groupData)
      .sort((a, b) => b[1].total - a[1].total) // Sort descending by total kg
      .map(([name, data]) => ({
        name,
        total: data.total,
        clusters: data.clusters,
      }));

    return sortedGroups;
  }, [summaryData, activeTeamMembers, filterBelowCrop, selectedClusters, grandTotalViewBy, teamProvinces, isBusinessAnalyst]);

  const categoryFillingStats = useMemo(() => {
    let total = 0;
    let filled = 0;

    mappedChannelsByPic.forEach((k) => {
      total++;
      const cat = String(k.category || "").trim();
      const cleanCat = cat.toLowerCase();
      if (
        cat !== "" &&
        cat !== "-" &&
        cleanCat !== "uncategorized" &&
        cleanCat !== "n/a" &&
        cleanCat !== "unknown"
      ) {
        filled++;
      }
    });

    const percentage =
      total > 0 ? Math.min(100, Math.round((filled / total) * 100)) : 0;
    return { filled, total, percentage };
  }, [mappedChannelsByPic]);

  const categoryDistributionStats = useMemo(() => {
    const counts: Record<string, number> = {};
    let total = 0;

    mappedChannelsByPic.forEach((k) => {
      total++;
      let cat = String(k.category || "").trim();
      if (cat === "" || cat === "-") {
        cat = "Uncategorized";
      }

      const lowerCat = cat.toLowerCase();
      if (
        lowerCat === "uncategorized" ||
        lowerCat === "n/a" ||
        lowerCat === "unknown"
      ) {
        cat = "Uncategorized";
      } else if (lowerCat === "r1") {
        cat = "R1";
      } else if (lowerCat === "r2") {
        cat = "R2";
      } else if (lowerCat === "distributor") {
        cat = "Distributor";
      } else {
        cat = cat.charAt(0).toUpperCase() + cat.slice(1);
      }

      counts[cat] = (counts[cat] || 0) + 1;
    });

    const data = Object.entries(counts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => {
        const order = ["distributor", "r1", "r2"];
        const idxA = order.indexOf(a.name.toLowerCase());
        const idxB = order.indexOf(b.name.toLowerCase());

        if (idxA !== -1 && idxB !== -1) {
          return idxA - idxB;
        }
        if (idxA !== -1) return -1;
        if (idxB !== -1) return 1;

        return b.value - a.value;
      });

    return { data, total };
  }, [mappedChannelsByPic]);

  const categoryChartSegments = useMemo(() => {
    const { data, total } = categoryDistributionStats;
    if (total === 0 || data.length === 0) return [];

    let accumulatedPercent = 0;
    const circumference = 2 * Math.PI * 95; // 596.9026

    const getCategoryColor = (name: string, idx: number) => {
      const lower = name.toLowerCase();
      if (lower === "gold") return "#f59e0b"; // Gold Amber-500
      if (lower === "silver") return "#cbd5e1"; // Silver Slate-300
      if (lower === "bronze") return "#b45309"; // Bronze Amber-700

      // Primary partner categories
      if (lower === "distributor") return "rgba(255, 255, 255, 0.35)"; // White with lower opacity
      if (lower === "r1") return "#ffffff"; // White
      if (lower === "r2") return "#2563eb"; // Blue cluster color (blue-600)

      if (lower === "uncategorized" || lower === "unknown" || lower === "-")
        return "#94a3b8"; // Neutral Slate

      const COLORS = [
        "rgba(255, 255, 255, 0.35)",
        "#ffffff",
        "#2563eb",
        "#f59e0b",
        "#fb923c",
        "#f43f5e",
        "#2dd4bf",
        "#94a3b8",
      ];
      return COLORS[idx % COLORS.length];
    };

    return data.map((item, idx) => {
      const percentage = total > 0 ? (item.value / total) * 100 : 0;
      const strokeLength = (percentage / 100) * circumference;
      const strokeOffset =
        circumference - (accumulatedPercent / 100) * circumference;
      accumulatedPercent += percentage;
      return {
        ...item,
        percentage,
        strokeDasharray: `${strokeLength} ${circumference - strokeLength}`,
        strokeDashoffset: strokeOffset,
        color: getCategoryColor(item.name, idx),
      };
    });
  }, [categoryDistributionStats]);

  const renderRecursiveSummaryRow = (row: any, depth = 0): React.ReactNode => {
    const hasChildren = row.children && row.children.length > 0;
    const isExpanded = !!expandedRows[row.name];

    // Determine level badge style
    const levelVal = row.level;
    const levelLabel = levelVal !== undefined ? `Level ${levelVal}` : "";
    const isZeroTotal = isRowVisitZero(row.name);

    return (
      <div
        key={row.name}
        className={`overflow-hidden transition-all ${isZeroTotal ? "bg-red-50/40 border-l-[3px] border-red-300" : ""} ${depth === 0 ? "md:bg-white md:rounded-[32px] md:shadow-[0_4px_24px_rgba(24,26,44,0.08)] md:border md:border-[#edecff]" : ""}`}
      >
        <div
          className={`flex justify-between items-center px-5 py-4 pb-2 transition-colors duration-150 animate-in fade-in slide-in-from-left-2 duration-200 ${isZeroTotal ? "hover:bg-red-100/40" : "hover:bg-slate-50/60"}`}
          style={{ paddingLeft: `${Math.max(20, depth * 20)}px` }}
        >
          <div className="flex flex-col min-w-0 flex-1 pr-2">
            <span
              className="font-semibold text-xs md:text-sm text-[#181a2c] uppercase flex items-center gap-1.5 cursor-pointer select-none"
              onClick={() => {
                if (hasChildren) {
                  toggleRow(row.name);
                }
              }}
            >
              {hasChildren && (
                <span
                  className={`material-symbols-outlined text-[20px] shrink-0 ${isZeroTotal ? "text-red-500" : "text-primary"}`}
                >
                  {isExpanded ? "keyboard_arrow_down" : "keyboard_arrow_right"}
                </span>
              )}
              <span className="truncate">
                {renderMaybeChannelName(row.name)}
              </span>
              {isZeroTotal && (
                <span className="text-[7.5px] font-extrabold bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full shrink-0 uppercase tracking-wider">
                  Belum Dikunjungi
                </span>
              )}
              {row.isLeaf &&
                summarySubGroupBy === "channel" &&
                row.category && (
                  <span className="text-[8px] bg-primary/10 text-primary px-1.5 py-0.5 rounded-full font-bold">
                    {row.category}
                  </span>
                )}
            </span>
            {row.isLeaf && summarySubGroupBy === "channel" && row.pic && (
              <div className="text-[10px] text-[#8E94B7] mt-1 truncate">
                {row.pic}
              </div>
            )}
            {!row.isLeaf && getStatsForPic(row.name) && (
              <div className="flex gap-1.5 mt-1 ml-6 text-[8.5px] uppercase tracking-wide font-bold text-[#8E94B7] flex-wrap">
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                  V:{" "}
                  <span className="text-primary">
                    {getStatsForPic(row.name)?.visited}
                  </span>
                </span>
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                  C:{" "}
                  <span className="text-[#181a2c]">
                    {getStatsForPic(row.name)?.total}
                  </span>
                </span>
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                  %:{" "}
                  <span className="text-emerald-600">
                    {getStatsForPic(row.name)?.percentage}%
                  </span>
                </span>
              </div>
            )}
          </div>
          <div className="flex flex-col items-end shrink-0">
            <span
              className={`font-bold text-sm ${isZeroTotal ? "text-red-600" : "text-primary"}`}
            >
              {formatOverviewVal(row.selectedTotal, overviewUseMt).valueStr}
            </span>
            <span className="text-[8px] text-[#8E94B7] uppercase tracking-widest font-bold">
              Total {formatOverviewVal(row.selectedTotal, overviewUseMt).unit}
            </span>
          </div>
        </div>

        {/* Value Clusters in primary container */}
        <div
          className={`mx-5 mb-3 mt-2 flex divide-x rounded-[14px] overflow-hidden ${isZeroTotal ? "divide-red-200 bg-red-300 shadow-[0_12px_32px_rgba(239,68,68,0.18)]" : "divide-white/20 bg-primary shadow-[0_12px_32px_rgba(21,75,226,0.35)]"}`}
          style={{ marginLeft: `${Math.max(20, depth * 20)}px` }}
        >
          {selectedClusters.map((clusterKey) => {
            if (
              clusterKey === "Uncategorized" &&
              (!row[clusterKey] || row[clusterKey] === 0)
            )
              return null;
            const clusterConfig = CLUSTER_CONFIG.find(
              (c) => c.key === clusterKey,
            );
            return (
              <div
                key={clusterKey}
                className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroTotal ? "hover:bg-white/15" : "hover:bg-white/10"}`}
              >
                <span
                  className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroTotal ? "text-white/80" : "text-white/85"}`}
                >
                  {clusterConfig?.label || clusterKey}
                </span>
                <span className="font-semibold text-[10.5px] truncate w-full text-white">
                  {formatOverviewVal(row[clusterKey], overviewUseMt).valueStr}
                </span>
              </div>
            );
          })}
        </div>

        {isExpanded && hasChildren && (
          <div className="pb-2 border-t border-[#edecff] bg-slate-50/15">
            {row.children.map((child: any) =>
              renderRecursiveSummaryRow(child, depth + 1),
            )}
          </div>
        )}
      </div>
    );
  };

  const renderRecursiveSubordinate = (
    child: any,
    depth = 0,
  ): React.ReactNode => {
    const isChildZeroTeam = isRowVisitZero(child.name);
    const hasSubChildren = child.children && child.children.length > 0;
    const isSubExpanded = !!expandedRows[child.name];

    return (
      <div
        key={child.name}
        className="flex flex-col w-full animate-in fade-in slide-in-from-left-2 duration-200"
        style={{ paddingLeft: depth > 0 ? "16px" : "0px" }}
      >
        <div
          className={`flex flex-col p-3.5 rounded-[18px] transition-all duration-200 mb-2 ${
            isChildZeroTeam
              ? "bg-red-50/70 border border-red-200/60 shadow-[0_10px_28px_rgba(239,68,68,0.12)]"
              : "bg-[#fbfaff] shadow-[0_10px_28px_rgba(21,75,226,0.18)]"
          }`}
        >
          <div className="flex justify-between items-center mb-2 px-1 flex-wrap gap-2">
            <div className="flex flex-col min-w-0">
              <div className="flex items-center gap-1.5">
                {hasSubChildren && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      toggleRow(child.name);
                    }}
                    className="focus:outline-none flex items-center justify-center p-0.5 hover:bg-slate-200/50 rounded-full"
                  >
                    <span className="material-symbols-outlined text-primary text-[16px]">
                      {isSubExpanded
                        ? "keyboard_arrow_down"
                        : "keyboard_arrow_right"}
                    </span>
                  </button>
                )}
                <span className="font-bold text-[11px] text-[#181a2c] uppercase flex items-center gap-1.5 flex-wrap">
                  <span className="truncate">
                    {renderMaybeChannelName(child.name)}
                  </span>
                  {isChildZeroTeam && (
                    <span className="text-[7.5px] bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full shrink-0 uppercase tracking-wider animate-pulse font-bold">
                      Belum Dikunjungi
                    </span>
                  )}
                </span>
              </div>
              {getStatsForPic(child.name) && (
                <div className="flex gap-1.2 mt-1 ml-5 text-[8px] uppercase tracking-wide font-bold text-[#8E94B7] flex-wrap">
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    V:{" "}
                    <span className="text-primary">
                      {getStatsForPic(child.name)?.visited}
                    </span>
                  </span>
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    C:{" "}
                    <span className="text-[#181a2c]">
                      {getStatsForPic(child.name)?.total}
                    </span>
                  </span>
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    %:{" "}
                    <span className="text-emerald-600">
                      {getStatsForPic(child.name)?.percentage}%
                    </span>
                  </span>
                </div>
              )}
            </div>
            <span
              className={`font-bold text-[11.5px] shrink-0 ${isChildZeroTeam ? "text-red-600" : "text-[#181a2c]"}`}
            >
              {formatOverviewVal(child.selectedTotal, overviewUseMt).valueStr}{" "}
              <span className="text-[8.5px] text-[#8E94B7]">{formatOverviewVal(child.selectedTotal, overviewUseMt).unit}</span>
            </span>
          </div>
          <div
            className={`flex w-full divide-x rounded-[14px] overflow-hidden ${
              isChildZeroTeam
                ? "divide-red-200 bg-red-100/40"
                : "divide-primary/10 bg-primary/8"
            }`}
          >
            {selectedClusters.map((clusterKey) => {
              if (
                clusterKey === "Uncategorized" &&
                (!child[clusterKey] || child[clusterKey] === 0)
              )
                return null;
              const clusterConfig = CLUSTER_CONFIG.find(
                (c) => c.key === clusterKey,
              );
              return (
                <div
                  key={clusterKey}
                  className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                    isChildZeroTeam
                      ? "hover:bg-red-100/30"
                      : "hover:bg-primary/5"
                  }`}
                >
                  <span
                    className={`text-[7px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isChildZeroTeam ? "text-red-700/80" : "text-[#8E94B7]"}`}
                  >
                    {clusterConfig?.label || clusterKey}
                  </span>
                  <span
                    className={`font-semibold text-[9.5px] truncate w-full ${isChildZeroTeam ? "text-red-600" : "text-[#181a2c]"}`}
                  >
                    {formatOverviewVal(child[clusterKey], overviewUseMt).valueStr}
                  </span>
                </div>
              );
            })}
          </div>
        </div>

        {isSubExpanded && hasSubChildren && (
          <div className="flex flex-col border-l border-[#edecff] ml-3 pl-1 gap-1 mb-2">
            {child.children.map((subChild: any) =>
              renderRecursiveSubordinate(subChild, depth + 1),
            )}
          </div>
        )}
      </div>
    );
  };

  const renderRecursivePogSubordinate = (
    child: any,
    depth = 0,
  ): React.ReactNode => {
    const isChildZeroTeam = isRowVisitZero(child.name);
    const hasSubChildren = child.children && child.children.length > 0;
    const isSubExpanded = !!pogExpandedRows[child.name];

    return (
      <div
        key={child.name}
        className="flex flex-col w-full animate-in fade-in slide-in-from-left-2 duration-200"
        style={{ paddingLeft: depth > 0 ? "16px" : "0px" }}
      >
        <div
          className={`flex flex-col p-3.5 rounded-[18px] transition-all duration-200 mb-2 ${
            isChildZeroTeam
              ? "bg-red-50/70 border border-red-200/60 shadow-[0_10px_28px_rgba(239,68,68,0.12)]"
              : "bg-[#fbfaff] shadow-[0_10px_28px_rgba(21,75,226,0.18)]"
          }`}
        >
          <div className="flex justify-between items-center mb-2 px-1 flex-wrap gap-2">
            <div className="flex flex-col min-w-0">
              <div className="flex items-center gap-1.5">
                {hasSubChildren && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      togglePogRow(child.name);
                    }}
                    className="focus:outline-none flex items-center justify-center p-0.5 hover:bg-slate-200/50 rounded-full"
                  >
                    <span className="material-symbols-outlined text-primary text-[16px]">
                      {isSubExpanded
                        ? "keyboard_arrow_down"
                        : "keyboard_arrow_right"}
                    </span>
                  </button>
                )}
                <span className="font-bold text-[11px] text-[#181a2c] uppercase flex items-center gap-1.5 flex-wrap">
                  <span className="truncate">
                    {renderMaybeChannelName(child.name)}
                  </span>
                  {isChildZeroTeam && (
                    <span className="text-[7.5px] bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full font-bold uppercase tracking-wider">
                      Belum Dikunjungi
                    </span>
                  )}
                </span>
              </div>
              {getStatsForPic(child.name) && (
                <div className="flex gap-1.2 mt-1 ml-5 text-[8px] uppercase tracking-wide font-bold text-[#8E94B7] flex-wrap">
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    V:{" "}
                    <span className="text-primary">
                      {getStatsForPic(child.name)?.visited}
                    </span>
                  </span>
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    C:{" "}
                    <span className="text-[#181a2c]">
                      {getStatsForPic(child.name)?.total}
                    </span>
                  </span>
                  <span className="bg-[#f4f2ff] px-1.5 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                    %:{" "}
                    <span className="text-emerald-600">
                      {getStatsForPic(child.name)?.percentage}%
                    </span>
                  </span>
                </div>
              )}
            </div>
            <span
              className={`font-bold text-[11.5px] shrink-0 ${isChildZeroTeam ? "text-red-600" : "text-primary"}`}
            >
              {formatOverviewVal(child.pog, overviewUseMt).valueStr}{" "}
              <span className="text-[8.5px] text-[#8E94B7]">POG ({formatOverviewVal(child.pog, overviewUseMt).unit})</span>
            </span>
          </div>
          <div className="flex flex-row w-full gap-1.5 md:gap-2">
            {/* Table 1: Opening Inv & End of Inv */}
            <div
              className={`flex-[2] flex divide-x rounded-[14px] overflow-hidden ${
                isChildZeroTeam
                  ? "divide-red-200 bg-red-100/40"
                  : "divide-primary/10 bg-primary/5"
              }`}
            >
              <div
                className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                  isChildZeroTeam ? "hover:bg-red-200/30" : "hover:bg-primary/5"
                }`}
              >
                <span
                  className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                    isChildZeroTeam ? "text-red-700/60" : "text-[#8E94B7]"
                  }`}
                >
                  Opening Inv
                </span>
                <span
                  className={`font-black text-[10px] truncate w-full ${
                    isChildZeroTeam ? "text-red-700" : "text-[#181a2c]"
                  }`}
                >
                  {formatOverviewVal(child.lastQty, overviewUseMt).valueStr}
                </span>
              </div>
              <div
                className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                  isChildZeroTeam ? "hover:bg-red-200/30" : "hover:bg-primary/5"
                }`}
              >
                <span
                  className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                    isChildZeroTeam ? "text-red-700/60" : "text-[#1d4ed8]/75"
                  }`}
                >
                  End of Inv
                </span>
                <span
                  className={`font-black text-[10px] truncate w-full ${
                    isChildZeroTeam ? "text-red-700" : "text-[#1d4ed8]"
                  }`}
                >
                  {formatOverviewVal(child.currentQty, overviewUseMt).valueStr}
                </span>
              </div>
            </div>

            {/* Table 2: Stock in, idle stock, POG */}
            <div
              className={`flex-[3] flex divide-x rounded-[14px] overflow-hidden ${
                isChildZeroTeam
                  ? "divide-red-200 bg-red-200/45"
                  : "divide-primary/10 bg-primary/10 border border-primary/10"
              }`}
            >
              <div
                className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                  isChildZeroTeam
                    ? "hover:bg-red-300/30"
                    : "hover:bg-primary/15"
                }`}
              >
                <span
                  className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                    isChildZeroTeam ? "text-red-800/70" : "text-[#154be2]/80"
                  }`}
                >
                  Stock in
                </span>
                <span
                  className={`font-black text-[10px] truncate w-full ${
                    isChildZeroTeam ? "text-red-800" : "text-[#154be2]"
                  }`}
                >
                  {formatOverviewVal(child.sellIn, overviewUseMt).valueStr}
                </span>
              </div>
              <div
                className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                  isChildZeroTeam
                    ? "hover:bg-red-300/30"
                    : "hover:bg-amber-100/60"
                }`}
              >
                <span
                  className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                    isChildZeroTeam ? "text-red-800/70" : "text-amber-800"
                  }`}
                >
                  idle stock
                </span>
                <span
                  className={`font-black text-[10px] truncate w-full ${
                    isChildZeroTeam ? "text-red-800" : "text-amber-700"
                  }`}
                >
                  {formatOverviewVal(child.idleStock, overviewUseMt).valueStr}
                </span>
              </div>
              <div
                className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                  isChildZeroTeam
                    ? "hover:bg-red-300/30"
                    : "hover:bg-emerald-100/60"
                }`}
              >
                <span
                  className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                    isChildZeroTeam ? "text-red-800/70" : "text-emerald-800"
                  }`}
                >
                  POG
                </span>
                <span
                  className={`font-black text-[10px] truncate w-full ${
                    isChildZeroTeam
                      ? "text-red-800"
                      : "text-emerald-700 font-extrabold"
                  }`}
                >
                  {formatOverviewVal(child.pog, overviewUseMt).valueStr}
                </span>
              </div>
            </div>
          </div>
        </div>

        {isSubExpanded && hasSubChildren && (
          <div className="flex flex-col border-l border-[#edecff] ml-3 pl-1 gap-1 mb-2">
            {child.children.map((subChild: any) =>
              renderRecursivePogSubordinate(subChild, depth + 1),
            )}
          </div>
        )}
      </div>
    );
  };

  const renderRecursivePogRow = (row: any, depth = 0): React.ReactNode => {
    const hasChildren = row.children && row.children.length > 0;
    const isExpanded = !!pogExpandedRows[row.name];

    // Determine level badge style
    const levelVal = row.level;
    const levelLabel = levelVal !== undefined ? `Level ${levelVal}` : "";
    const isZeroPogActivity = isRowVisitZero(row.name);

    return (
      <div
        key={row.name}
        className={`overflow-hidden transition-all ${isZeroPogActivity ? "bg-red-50/40 border-l-[3px] border-red-300" : ""} ${depth === 0 ? "md:bg-white md:rounded-[32px] md:shadow-[0_4px_24px_rgba(24,26,44,0.08)] md:border md:border-[#edecff]" : ""}`}
      >
        <div
          className={`flex justify-between items-center px-5 py-4 pb-2 transition-colors duration-150 animate-in fade-in slide-in-from-right-2 duration-200 ${isZeroPogActivity ? "hover:bg-red-100/40" : "hover:bg-slate-50/60"}`}
          style={{ paddingLeft: `${Math.max(20, depth * 20)}px` }}
        >
          <div className="flex flex-col min-w-0 flex-1 pr-2">
            <span
              className="font-semibold text-xs md:text-sm text-[#181a2c] uppercase flex items-center gap-1.5 cursor-pointer select-none"
              onClick={() => {
                if (hasChildren) {
                  togglePogRow(row.name);
                }
              }}
            >
              {hasChildren && (
                <span
                  className={`material-symbols-outlined text-[20px] shrink-0 ${isZeroPogActivity ? "text-red-500" : "text-primary"}`}
                >
                  {isExpanded ? "keyboard_arrow_down" : "keyboard_arrow_right"}
                </span>
              )}
              <span className="truncate">
                {renderMaybeChannelName(row.name)}
              </span>
              {isZeroPogActivity && (
                <span className="text-[7.5px] font-extrabold bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full shrink-0 uppercase tracking-wider">
                  Belum Dikunjungi
                </span>
              )}
              {row.isLeaf && pogSubGroupBy === "channel" && row.category && (
                <span className="text-[8px] bg-primary/10 text-primary px-1.5 py-0.5 rounded-full font-bold">
                  {row.category}
                </span>
              )}
            </span>
            {row.isLeaf && pogSubGroupBy === "channel" && row.pic && (
              <div className="text-[10px] text-[#8E94B7] mt-1 truncate">
                {row.pic}
              </div>
            )}
            {!row.isLeaf && getStatsForPic(row.name) && (
              <div className="flex gap-1.5 mt-1 ml-6 text-[8.5px] uppercase tracking-wide font-bold text-[#8E94B7] flex-wrap">
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                  V:{" "}
                  <span className="text-primary">
                    {getStatsForPic(row.name)?.visited}
                  </span>
                </span>
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1">
                  C:{" "}
                  <span className="text-[#181a2c]">
                    {getStatsForPic(row.name)?.total}
                  </span>
                </span>
                <span className="bg-[#f4f2ff] px-2 py-0.5 rounded-full border border-[#edecff] flex gap-1 text-emerald-600">
                  %: <span>{getStatsForPic(row.name)?.percentage}%</span>
                </span>
              </div>
            )}
          </div>
          <div className="flex flex-col items-end shrink-0">
            <span
              className={`font-bold text-sm ${isZeroPogActivity ? "text-red-600" : "text-primary"}`}
            >
              {formatOverviewVal(row.pog, overviewUseMt).valueStr}
            </span>
            <span className="text-[8px] text-[#8E94B7] uppercase tracking-widest font-bold">
              POG ({formatOverviewVal(row.pog, overviewUseMt).unit})
            </span>
          </div>
        </div>

        {/* Dynamic primary row columns */}
        <div
          className="mx-5 mb-3 mt-2 flex flex-row gap-1.5 md:gap-2"
          style={{ marginLeft: `${Math.max(20, depth * 20)}px` }}
        >
          {/* Table 1: Opening Inv & End of Inv */}
          <div
            className={`flex-[2] flex divide-x rounded-[14px] overflow-hidden ${
              isZeroPogActivity
                ? "divide-red-200 bg-red-300 shadow-[0_12px_32px_rgba(239,68,68,0.18)]"
                : "divide-white/20 bg-primary/95 shadow-[0_12px_32px_rgba(21,75,226,0.25)]"
            }`}
          >
            <div
              className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroPogActivity ? "hover:bg-white/15" : "hover:bg-white/10"}`}
            >
              <span
                className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroPogActivity ? "text-white/80" : "text-white/85"}`}
              >
                Opening Inv
              </span>
              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                {formatOverviewVal(row.lastQty, overviewUseMt).valueStr}
              </span>
            </div>
            <div
              className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroPogActivity ? "hover:bg-white/15" : "hover:bg-white/10"}`}
            >
              <span
                className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroPogActivity ? "text-white/80" : "text-white/85"}`}
              >
                End of Inv
              </span>
              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                {formatOverviewVal(row.currentQty, overviewUseMt).valueStr}
              </span>
            </div>
          </div>

          {/* Table 2: Stock in, idle stock, POG */}
          <div
            className={`flex-[3] flex divide-x rounded-[14px] overflow-hidden ${
              isZeroPogActivity
                ? "divide-red-200 bg-red-400 shadow-[0_12px_32px_rgba(239,68,68,0.18)]"
                : "divide-white/20 bg-primary shadow-[0_12px_32px_rgba(21,75,226,0.35)]"
            }`}
          >
            <div
              className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroPogActivity ? "hover:bg-white/15" : "hover:bg-white/10"}`}
            >
              <span
                className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroPogActivity ? "text-white/80" : "text-white/85"}`}
              >
                Stock in
              </span>
              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                {formatOverviewVal(row.sellIn, overviewUseMt).valueStr}
              </span>
            </div>
            <div
              className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroPogActivity ? "hover:bg-white/15" : "hover:bg-white/10"}`}
            >
              <span
                className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroPogActivity ? "text-white/80" : "text-amber-200"}`}
              >
                idle stock
              </span>
              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                {formatOverviewVal(row.idleStock, overviewUseMt).valueStr}
              </span>
            </div>
            <div
              className={`flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center transition-colors ${isZeroPogActivity ? "hover:bg-white/15" : "hover:bg-white/10"}`}
            >
              <span
                className={`text-[8px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${isZeroPogActivity ? "text-white/80" : "text-cyan-200"}`}
              >
                POG
              </span>
              <span className={`font-semibold text-[10.5px] truncate w-full font-extrabold ${isZeroPogActivity ? "text-white" : "text-cyan-100"}`}>
                {formatOverviewVal(row.pog, overviewUseMt).valueStr}
              </span>
            </div>
          </div>
        </div>

        {isExpanded && hasChildren && (
          <div className="pb-2 border-t border-[#edecff] bg-slate-50/15">
            {row.children.map((child: any) =>
              renderRecursivePogRow(child, depth + 1),
            )}
          </div>
        )}
      </div>
    );
  };

  if (showLoader) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen supports-[min-height:100dvh]:min-h-[100dvh] bg-[#f8fafc] px-4 pb-20 animate-in fade-in duration-500">
        <div className="bg-white p-8 rounded-[32px] shadow-[0_24px_64px_rgba(24,26,44,0.06)] border border-[#e2e8f0] flex flex-col items-center max-w-sm w-full gap-6">
          <div className="relative">
            <div className="size-24 border-4 border-[#edecff] rounded-full flex items-center justify-center">
              <span className="text-xl font-extrabold text-primary select-none">
                {loadProgress}%
              </span>
            </div>
            <div className="size-24 border-4 border-primary border-t-transparent rounded-full animate-spin absolute inset-0"></div>
          </div>

          <div className="text-center w-full">
            <h3 className="text-[#181a2c] font-black text-xl mb-1 select-none">
              Menyiapkan Data...
            </h3>
            <p className="text-[#8E94B7] text-[10px] font-bold uppercase tracking-widest select-none">
              SINKRONISASI DATABASE
            </p>
          </div>

          <div className="w-full flex flex-col items-center gap-2">
            <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden p-[1px] border border-[#f1f5f9]">
              <div
                className="bg-gradient-to-r from-primary to-cyan-400 h-full rounded-full transition-all duration-150 ease-out shadow-[0_2px_8px_rgba(21,75,226,0.25)]"
                style={{ width: `${loadProgress}%` }}
              ></div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const renderPogKpiCard = (isSummaryTab: boolean) => {
    const dataEmpty = isSummaryTab
      ? filteredSummaryData.length === 0
      : aggregatedPogData.length === 0;
    if (dataEmpty) return null;
    
    const currentTotal = isSummaryTab ? totalSummaryPog : totalPog;
    const isFilterOpen = isSummaryTab ? isSummaryFilterOpen : isPogFilterOpen;
    const toggleFilter = () =>
      isSummaryTab
        ? setIsSummaryFilterOpen(!isSummaryFilterOpen)
        : setIsPogFilterOpen(!isPogFilterOpen);

    return (
      <div
        onClick={() => {
          if (window.innerWidth < 768) {
            toggleFilter();
          }
        }}
        className="flex-1 w-full min-w-0 bg-gradient-to-br from-primary to-cyan-400 p-5 md:px-7 rounded-[36px] shadow-[0_12px_32px_rgba(21,75,226,0.35)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.45)] hover:scale-[1.01] active:scale-[0.99] transition-all duration-250 cursor-pointer md:cursor-default select-none text-white mb-1.5 md:mb-0 md:hover:scale-100 md:active:scale-100"
      >
        <div className="flex justify-between items-center mb-4">
          <div className="flex items-center gap-1.5">
            <span className="material-symbols-outlined text-[18px] text-white/80">
              tune
            </span>
            <div className="flex flex-col">
              <span className="font-semibold text-sm uppercase tracking-wider text-white/90">
                Grand Total
              </span>
              {isSummaryTab && (
                <span className="text-[10px] font-bold text-white/70 uppercase tracking-widest mt-0.5">
                  {filterBelowCrop || "All"}
                </span>
              )}
            </div>
          </div>
          <div className="flex flex-col items-end">
            <span className="font-bold text-xl xl:text-2xl text-white">
              {formatOverviewVal(isSummaryTab ? totalSummary.selectedTotal : currentTotal.pog, overviewUseMt).valueStr}
            </span>
            <span className="text-[8px] text-white/80 uppercase tracking-widest font-bold">
              {isSummaryTab ? `Total ${formatOverviewVal(totalSummary.selectedTotal, overviewUseMt).unit}` : `POG (${formatOverviewVal(currentTotal.pog, overviewUseMt).unit})`}
            </span>
          </div>
        </div>

        {isSummaryTab ? (
          <div className="flex w-full divide-x divide-white/15 border-t border-white/15 pt-4 mt-2">
            {ALL_CLUSTER_KEYS.map((clusterKey) => {
              if (
                clusterKey === "Uncategorized" &&
                (!totalSummary[clusterKey] || totalSummary[clusterKey] === 0)
              )
                return null;
              const clusterConfig = CLUSTER_CONFIG.find(
                (c) => c.key === clusterKey,
              );
              const isSelected = selectedClusters.includes(clusterKey);
              return (
                <div
                  key={clusterKey}
                  onClick={(e) => {
                    e.stopPropagation();
                    setSelectedClusters((prev) =>
                      prev.includes(clusterKey)
                        ? prev.filter((k) => k !== clusterKey)
                        : [...prev, clusterKey]
                    );
                  }}
                  className={`flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center cursor-pointer transition-all ${isSelected ? "opacity-100 hover:bg-white/10 rounded-md" : "opacity-40 hover:opacity-60"}`}
                >
                  <span className="text-[9px] xl:text-[10px] font-bold uppercase tracking-wider text-white/85 mb-1 truncate w-full">
                    {clusterConfig?.label || clusterKey}
                  </span>
                  <span className="font-bold text-sm xl:text-base truncate w-full text-white">
                    {formatOverviewVal(totalSummary[clusterKey], overviewUseMt).valueStr}
                  </span>
                </div>
              );
            })}
          </div>
        ) : (
          <div className="flex flex-row w-full gap-1.5 md:gap-2 border-t border-white/15 pt-3.5 mt-1">
            {/* Table 1: Opening and End Inv group */}
            <div className="flex-[2] flex divide-x divide-white/15 bg-white/5 rounded-[12px] p-0.5">
              <div className="flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center">
                <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                  Opening
                </span>
                <span className="font-bold text-xs truncate w-full text-white">
                  {formatOverviewVal(currentTotal.lastQty, overviewUseMt).valueStr}
                </span>
              </div>
              <div className="flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center">
                <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                  End Inv
                </span>
                <span className="font-bold text-xs truncate w-full text-white">
                  {formatOverviewVal(currentTotal.currentQty, overviewUseMt).valueStr}
                </span>
              </div>
            </div>

            {/* Table 2: Stock in, idle stock, POG group */}
            <div className="flex-[3] flex divide-x divide-white/20 bg-gradient-to-br from-white/25 via-white/15 to-white/25 border border-white/20 rounded-[12px] p-0.5 shadow-[inset_0_1px_1px_rgba(255,255,255,0.25)]">
              <div className="flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center">
                <span className="text-[8px] font-bold uppercase tracking-wider text-white/95 mb-0.5 truncate w-full">
                  Stock In
                </span>
                <span className="font-extrabold text-xs truncate w-full text-white">
                  {formatOverviewVal(currentTotal.sellIn, overviewUseMt).valueStr}
                </span>
              </div>
              <div className="flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center hover:bg-white/5 transition-colors rounded-[12px]">
                <span className="text-[8px] font-bold uppercase tracking-wider text-amber-200 mb-0.5 truncate w-full">
                  Idle Stock
                </span>
                <span className="font-extrabold text-xs truncate w-full text-amber-100">
                  {formatOverviewVal(currentTotal.idleStock, overviewUseMt).valueStr}
                </span>
              </div>
              <div className="flex-1 min-w-0 py-1.5 px-0.5 flex flex-col items-center justify-center text-center hover:bg-white/5 transition-colors rounded-[12px]">
                <span className="text-[8px] font-bold uppercase tracking-wider text-cyan-200 mb-0.5 truncate w-full">
                  POG
                </span>
                <span className="font-black text-xs truncate w-full text-cyan-50">
                  {formatOverviewVal(currentTotal.pog, overviewUseMt).valueStr}
                </span>
              </div>
            </div>
          </div>
        )}

        <div className="border-t border-white/15 pt-5 mt-4 w-full">
          <div className="grid grid-cols-2 gap-3 sm:gap-6 lg:gap-8 items-start justify-center">
            {/* CHART 1: VISIT COVERAGE */}
            <div className="flex flex-col items-center text-center p-1 sm:p-3">
              <h4 className="text-[8.5px] sm:text-[10px] font-bold uppercase tracking-wider text-white/90 mb-3 sm:mb-4 font-sans line-clamp-1">
                Kunjungan (Visit)
              </h4>
              <div className="relative flex items-center justify-center shrink-0 mb-3 sm:mb-4 animate-in zoom-in duration-500">
                <svg
                  viewBox="0 0 220 220"
                  className="rotate-[-90deg] w-24 h-24 sm:w-32 sm:h-32 md:w-36 md:h-36 lg:w-40 lg:h-40 xl:w-44 xl:h-44"
                >
                  <circle
                    stroke="rgba(255, 255, 255, 0.15)"
                    fill="transparent"
                    strokeWidth={16}
                    r={95}
                    cx={110}
                    cy={110}
                  />
                  <circle
                    stroke="white"
                    fill="transparent"
                    strokeWidth={16}
                    strokeDasharray={`${2 * Math.PI * 95} ${2 * Math.PI * 95}`}
                    strokeDashoffset={
                      2 * Math.PI * 95 -
                      (activeVisitStats.percentage / 100) * (2 * Math.PI * 95)
                    }
                    style={{
                      strokeDashoffset: `${2 * Math.PI * 95 - (activeVisitStats.percentage / 100) * (2 * Math.PI * 95)}px`,
                    }}
                    r={95}
                    cx={110}
                    cy={110}
                    strokeLinecap="round"
                    className="transition-[stroke-dashoffset] duration-700 ease-out"
                  />
                </svg>
                <div className="absolute flex flex-col items-center justify-center">
                  <span className="text-lg sm:text-2xl md:text-3xl xl:text-4xl font-extrabold text-white tracking-tight leading-none">
                    {activeVisitStats.percentage}%
                  </span>
                  <span className="text-[7px] sm:text-[8px] md:text-[9px] font-bold uppercase tracking-wider text-white/70 mt-0.5 sm:mt-1">
                    Visited
                  </span>
                </div>
              </div>
              <p className="text-[9.5px] sm:text-[11px] md:text-[11.5px] leading-relaxed text-white/90 max-w-sm">
                {!mappingPic ||
                cleanForMatch(mappingPic) === "allteam" ||
                cleanForMatch(mappingPic) === "all_team"
                  ? "Team"
                  : normalizeName(mappingPic)}{" "}
                sudah melakukan visit sebanyak{" "}
                <span className="font-bold text-white">
                  {activeVisitStats.visited}
                </span>{" "}
                dari{" "}
                <span className="font-bold text-white">
                  {activeVisitStats.total}
                </span>{" "}
                Partner{" "}
                <span className="text-white/80">
                  ({activeVisitStats.percentage}%)
                </span>
              </p>
            </div>

            {/* CHART 2: CATEGORY DISTRIBUTION */}
            <div className="flex flex-col items-center text-center p-1 sm:p-3 border-l border-white/15">
              <h4 className="text-[8.5px] sm:text-[10px] font-bold uppercase tracking-wider text-white/90 mb-3 sm:mb-4 font-sans line-clamp-1">
                Kategori Partner
              </h4>
              <div className="relative flex items-center justify-center shrink-0 mb-3 sm:mb-4 md:mb-5 animate-in zoom-in duration-500">
                <svg
                  viewBox="0 0 220 220"
                  className="rotate-[-90deg] w-24 h-24 sm:w-32 sm:h-32 md:w-36 md:h-36 lg:w-40 lg:h-40 xl:w-44 xl:h-44"
                >
                  <circle
                    stroke="rgba(255, 255, 255, 0.1)"
                    fill="transparent"
                    strokeWidth={16}
                    r={95}
                    cx={110}
                    cy={110}
                  />
                  {categoryChartSegments.map((segment) => (
                    <circle
                      key={segment.name}
                      stroke={segment.color}
                      fill="transparent"
                      strokeWidth={16}
                      strokeLinecap="round"
                      strokeDasharray={segment.strokeDasharray}
                      strokeDashoffset={segment.strokeDashoffset}
                      style={{
                        strokeDashoffset: `${segment.strokeDashoffset}px`,
                      }}
                      r={95}
                      cx={110}
                      cy={110}
                      className="transition-[stroke-dashoffset] duration-700 ease-out"
                    />
                  ))}
                </svg>
                <div className="absolute flex flex-col items-center justify-center">
                  <span className="text-lg sm:text-2xl md:text-3xl xl:text-4xl font-extrabold text-white tracking-tight leading-none font-sans">
                    {categoryDistributionStats.total}
                  </span>
                  <span className="text-[7px] sm:text-[8px] md:text-[9px] font-bold uppercase tracking-wider text-white/70 mt-0.5 sm:mt-1 font-mono">
                    Partners
                  </span>
                </div>
              </div>
              <div className="w-full mt-1.5">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-1 px-0.5 sm:px-1 text-[8.5px] sm:text-[10px] text-white/90">
                  {categoryChartSegments.map((segment) => (
                    <div
                      key={segment.name}
                      className="flex items-center gap-1 sm:gap-1.5 bg-black/12 py-1 sm:py-1.5 px-1.5 sm:px-2 rounded-full border border-black/5 truncate hover:bg-black/25 transition-colors shadow-sm animate-in fade-in-50 duration-300"
                    >
                      <span
                        className="inline-block w-2 sm:w-2.5 h-2 sm:h-2.5 rounded-full shrink-0"
                        style={{ backgroundColor: segment.color }}
                      />
                      <span className="font-semibold text-white/90 truncate flex-1 text-left">
                        {segment.name}
                      </span>
                      <span className="font-mono font-bold text-white shrink-0">
                        {segment.value}{" "}
                        <span className="text-[7px] sm:text-[8px] font-normal text-white/75">
                          ({Math.round(segment.percentage)}%)
                        </span>
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderGrandTotalCard = () => {
    if (filteredSummaryData.length === 0) return null;
    return (
      <div
        className="flex-1 w-full min-w-0 bg-gradient-to-br from-primary to-cyan-400 p-4 md:py-4 md:px-6 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.35)] transition-all duration-250 select-none text-white flex flex-col gap-3 lg:gap-4"
      >
        {/* Left Side: Grand Total & Clusters */}
        <div className="flex flex-col w-full shrink-0 justify-center">
          <div className="flex justify-between items-center mb-2">
            <div className="flex items-center gap-2">
              <span className="material-symbols-outlined text-[18px] text-white/80">
                tune
              </span>
              <div className="flex flex-col mr-2">
                <span className="font-semibold text-xs uppercase tracking-wider text-white/90">
                  Grand Total
                </span>
                <span className="text-[9px] font-bold text-white/70 uppercase tracking-widest mt-0.5">
                  {filterBelowCrop || "All"}
                </span>
              </div>
              <div className="flex bg-white/10 p-0.5 rounded-lg border border-white/10 ml-2">
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setGrandTotalViewBy("hybrid");
                  }}
                  className={`px-3 py-1 rounded-md text-[9px] font-bold uppercase tracking-wider transition-colors ${
                    grandTotalViewBy === "hybrid"
                      ? "bg-white text-primary shadow-sm"
                      : "text-white/70 hover:text-white"
                  }`}
                >
                  Hybrid
                </button>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setGrandTotalViewBy("area");
                  }}
                  className={`px-3 py-1 rounded-md text-[9px] font-bold uppercase tracking-wider transition-colors ${
                    grandTotalViewBy === "area"
                      ? "bg-white text-primary shadow-sm"
                      : "text-white/70 hover:text-white"
                  }`}
                >
                  Province
                </button>
              </div>
            </div>
            <div className="flex flex-col items-end">
              <span className="font-bold text-xl xl:text-2xl text-white tracking-tight">
                {formatOverviewVal(totalSummary.selectedTotal, overviewUseMt).valueStr}
              </span>
              <span className="text-[8px] xl:text-[9px] text-white/80 uppercase tracking-widest font-bold">
                Total {formatOverviewVal(totalSummary.selectedTotal, overviewUseMt).unit}
              </span>
            </div>
          </div>
          <div className="flex w-full divide-x divide-white/15 border-t border-white/15 pt-2.5 mt-1">
            {ALL_CLUSTER_KEYS.map((clusterKey) => {
              if (
                clusterKey === "Uncategorized" &&
                (!totalSummary[clusterKey] || totalSummary[clusterKey] === 0)
              )
                return null;
              const clusterConfig = CLUSTER_CONFIG.find(
                (c) => c.key === clusterKey,
              );
              const isSelected = selectedClusters.includes(clusterKey);
              return (
                <div
                  key={clusterKey}
                  onClick={(e) => {
                    e.stopPropagation();
                    setSelectedClusters((prev) =>
                      prev.includes(clusterKey)
                        ? prev.filter((k) => k !== clusterKey)
                        : [...prev, clusterKey]
                    );
                  }}
                  className={`flex-1 min-w-0 py-1 px-0.5 flex flex-col items-center justify-center text-center cursor-pointer transition-all ${isSelected ? "opacity-100 hover:bg-white/10 rounded-md" : "opacity-40 hover:opacity-60"}`}
                >
                  <span className="text-[9px] xl:text-[10px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                    {clusterConfig?.label || clusterKey}
                  </span>
                  <span className="font-bold text-xs xl:text-sm truncate w-full text-white">
                    {formatOverviewWithUnit(totalSummary[clusterKey], overviewUseMt)}
                  </span>
                </div>
              );
            })}
          </div>
        </div>

        {/* Right Side: Total per Item & Aging */}
        <div className="flex flex-col flex-1 border-t border-white/15 pt-3">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-3 gap-2.5 w-full max-h-[300px] lg:max-h-none overflow-y-auto pr-1 custom-scrollbar">
            {grandTotalStats.map((stat, index) => (
              <div
                key={stat.name}
                className="bg-white/12 hover:bg-white/15 transition-all border border-white/25 py-2 px-2 rounded-[16px] flex flex-col shadow-sm"
              >
                <div className="flex items-center gap-2 mb-2">
                  <div className="size-6 rounded-lg bg-white/25 flex items-center justify-center text-white font-black text-xs shrink-0">
                    {index + 1}
                  </div>
                  <div className="min-w-0 flex-1 flex justify-between items-center gap-2">
                    <p className="text-xs md:text-[13px] font-bold text-white uppercase tracking-wider truncate leading-none" title={stat.name}>
                      {stat.name}
                    </p>
                    <p className="text-xs md:text-[13px] font-extrabold text-white tracking-tight leading-none shrink-0">
                      {formatOverviewWithUnit(stat.total, overviewUseMt)}
                    </p>
                  </div>
                </div>
                <div className="flex w-full divide-x divide-white/20 border-t border-white/20 pt-2 mt-1">
                  {selectedClusters.map((clusterKey) => {
                    const val = stat.clusters[clusterKey] || 0;
                    if (clusterKey === "Uncategorized" && val === 0)
                      return null;
                    const clusterConfig = CLUSTER_CONFIG.find(
                      (c) => c.key === clusterKey,
                    );
                    return (
                      <div
                        key={clusterKey}
                        className="flex-1 min-w-0 px-0.5 flex flex-col items-center justify-center text-center"
                      >
                        <span className="text-[9px] md:text-[10px] font-bold uppercase tracking-wider text-white/80 mb-0.5 truncate w-full">
                          {clusterConfig?.label || clusterKey}
                        </span>
                        <span className="font-bold text-[11px] md:text-[12px] truncate w-full text-white">
                          {formatOverviewVal(val, overviewUseMt).valueStr}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
            {grandTotalStats.length === 0 && (
              <div className="col-span-full py-6 text-center text-white/60 text-xs font-medium">
                Tidak ada data.
              </div>
            )}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="w-full max-w-2xl mx-auto lg:max-w-5xl xl:max-w-6xl px-5 pb-8 relative">
      <div className="flex items-stretch gap-2 mb-8 mt-6 -ml-2.5">
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            e.preventDefault();
            if (!isFetchingData) {
              fetchWorkingData();
            }
          }}
          disabled={isFetchingData}
          className="flex flex-col items-center justify-center shrink-0 bg-gradient-to-r from-primary to-cyan-400 rounded-[24px] px-5 py-3 shadow-[0_12px_32px_rgba(21,75,226,0.35)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.45)] hover:scale-[1.02] active:scale-[0.98] transition-all min-w-[108px] cursor-pointer disabled:opacity-80 disabled:cursor-not-allowed select-none text-center border-0 appearance-none"
          title="Klik untuk Ambil Ulang Data dari Database"
        >
          <AdvantaLogo
            className={`size-14 text-white ${isFetchingData ? "animate-spin" : ""}`}
          />
          <span className="text-[9px] font-bold text-white/90 uppercase tracking-widest mt-1 leading-none">
            {isFetchingData ? "Refreshed" : "Radar"}
          </span>
        </button>
        <div className="bg-gradient-to-r from-primary to-cyan-400 p-4 rounded-l-[28px] rounded-r-none -mr-5 shadow-[0_12px_32px_rgba(21,75,226,0.35)] text-white flex items-center justify-between flex-1 relative overflow-visible transition-all select-none duration-200">
          {/* Background and Logout wrapper that clips the huge user icon */}
          <div className="absolute inset-0 overflow-hidden rounded-l-[28px] z-0">
            <div
              onClick={(e) => {
                e.stopPropagation();
                e.preventDefault();
                setIsLogoutModalOpen(true);
              }}
              className="absolute inset-0 flex items-center justify-center opacity-10 cursor-pointer hover:opacity-15 active:opacity-20 transition-opacity"
              title="Klik untuk Keluar"
            >
              <UserIcon className="h-[260%] w-auto max-w-none object-contain object-center" />
            </div>
          </div>

          <div className="flex-1 relative z-10 min-w-0 pr-2 flex flex-col gap-2">
            <div className="flex flex-col gap-1.5">
              <div className="flex items-center gap-3">
                {greeting.imageUrl && (
                  <img
                    src={greeting.imageUrl}
                    referrerPolicy="no-referrer"
                    className="size-12 md:size-14 object-contain opacity-100 shrink-0 pointer-events-none select-none animate-rotate-sway"
                    alt=""
                  />
                )}
                <div className="flex-1 min-w-0 py-0.5">
                  <p className="text-[9px] text-white/80 font-semibold uppercase tracking-widest leading-none mb-0.5">
                    {greeting.text},
                  </p>
                  <h2 className="text-md md:text-lg font-bold text-white leading-tight truncate">
                    {userData.name}
                  </h2>
                </div>
              </div>
              <div className="flex items-center gap-1.5 flex-wrap">
                <span className="text-[10px] md:text-xs font-bold text-white uppercase bg-white/20 px-2.5 py-1 rounded-full border border-white/20 backdrop-blur-sm truncate select-none">
                  {userData.position}
                </span>
                {userData.province && userData.province !== "-" && (
                  <span className="text-[10px] md:text-xs font-bold text-white/80 uppercase bg-white/10 px-2.5 py-1 rounded-full border border-white/10 backdrop-blur-sm">
                    {userData.province}
                  </span>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>

      {activeTab === "overview" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300">
          {/* Header */}
          <div className="mb-2 ml-1 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            <div>
              <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
                Executive{" "}
                <span className="text-primary font-bold">Overview</span>
              </h1>
              <p className="text-[11px] text-[#8E94B7] font-semibold uppercase tracking-wider mt-0.5">
                Analisis Kinerja & Pemantauan Tingkat Nasional
              </p>
            </div>
          </div>

          {/* Floating Filter Button */}
          <button
            onClick={() => setIsOverviewFilterOpen(!isOverviewFilterOpen)}
            className="fixed bottom-[110px] right-6 lg:bottom-10 lg:right-10 z-[60] bg-gradient-to-br from-[#154be2] to-[#123ebd] text-white p-3.5 lg:p-4 rounded-full shadow-[0_12px_32px_rgba(21,75,226,0.35)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.45)] hover:scale-105 active:scale-95 transition-all duration-300"
          >
            <span className="material-symbols-outlined text-[24px] lg:text-[28px]">filter_alt</span>
          </button>

          {/* Floating Detailed Filters */}
          <div
            className={`fixed inset-0 z-[70] transition-all duration-300 flex items-center justify-center p-4 ${isOverviewFilterOpen ? "opacity-100 visible pointer-events-auto" : "opacity-0 invisible pointer-events-none"}`}
          >
            {/* Backdrop */}
            <div 
              className="absolute inset-0 bg-black/40 backdrop-blur-sm transition-opacity duration-300" 
              onClick={() => setIsOverviewFilterOpen(false)} 
            />
            
            {/* Filter Content */}
            <div className={`relative w-full max-w-4xl bg-white p-6 md:p-8 rounded-[32px] shadow-[0_24px_64px_rgba(0,0,0,0.25)] transition-all duration-300 transform ${isOverviewFilterOpen ? "scale-100 translate-y-0" : "scale-95 translate-y-4"}`}>
              <div className="flex items-center justify-between gap-4 mb-6 pb-4 border-b border-slate-100">
                <div className="flex items-center gap-2">
                  <span className="material-symbols-outlined text-primary text-[20px] lg:text-[24px] font-semibold">
                    filter_alt
                  </span>
                  <h4 className="text-xs lg:text-sm font-bold text-[#181a2c] uppercase tracking-wider">
                    Filter Dashboard Analisis
                  </h4>
                </div>
                <button 
                  onClick={() => setIsOverviewFilterOpen(false)}
                  className="bg-slate-100 hover:bg-slate-200 text-slate-500 rounded-full w-8 h-8 flex items-center justify-center transition-colors"
                >
                  <span className="material-symbols-outlined text-[20px]">close</span>
                </button>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-5 gap-4 lg:gap-5">
                {/* Filter 2 - Month */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Bulan
                  </label>
                  <div className="relative">
                    <select
                      value={filterBelowMonth}
                      onChange={(e) => setFilterBelowMonth(e.target.value)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-4 py-3 text-xs lg:text-sm font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Semua Bulan</option>
                      {filterOptions.months.map((m) => (
                        <option key={m} value={m}>
                          {m}
                        </option>
                      ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>

                {/* Filter 2 - Channel */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Channel
                  </label>
                  <div className="relative">
                    <select
                      value={filterBelowChannel}
                      onChange={(e) => setFilterBelowChannel(e.target.value)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-4 py-3 text-xs lg:text-sm font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Semua Channel</option>
                      {filterOptions.channels.map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>

                {/* Filter 2 - Material */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Hybrid
                  </label>
                  <div className="relative">
                    <select
                      value={filterBelowMaterial}
                      onChange={(e) => setFilterBelowMaterial(e.target.value)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-4 py-3 text-xs lg:text-sm font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Semua Hybrid</option>
                      {filterOptions.materials.map((m) => (
                        <option key={m} value={m}>
                          {m}
                        </option>
                      ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>

                {/* Filter 2 - Team */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Tim (PIC)
                  </label>
                  <div className="relative">
                    <select
                      value={filterBelowTeam}
                      onChange={(e) => setFilterBelowTeam(e.target.value)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-4 py-3 text-xs lg:text-sm font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Semua PIC</option>
                      {filterOptions.teams.map((t) => (
                        <option key={t} value={t}>
                          {t}
                        </option>
                      ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>

                {/* Filter 2 - Area */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Wilayah
                  </label>
                  <div className="relative">
                    <select
                      value={filterBelowArea}
                      onChange={(e) => setFilterBelowArea(e.target.value)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-4 py-3 text-xs lg:text-sm font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Semua Wilayah</option>
                      {filterOptions.areas.map((a) => (
                        <option key={a} value={a}>
                          {a}
                        </option>
                      ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>

                {/* Filter 2 - Crop */}
                <div className="flex flex-col gap-2">
                  <label className="text-[10px] lg:text-[11px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Komoditas
                  </label>
                  <div className="relative">
                    <select
                      value={filterBelowCrop}
                      onChange={(e) => setFilterBelowCrop(e.target.value)}
                      className="w-full bg-[#fbfaff] border border-[#e2e8f0] rounded-xl px-4 py-3 text-xs lg:text-sm font-bold text-[#181a2c] focus:outline-none focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all appearance-none cursor-pointer pr-10"
                    >
                      <option value="All">Semua Crop</option>
                      {["Field Corn", "Fresh Corn", "Vegetables"].map((crop) => (
                        <option key={crop} value={crop}>
                          {crop}
                        </option>
                      ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-[18px] text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* KPI Cards Grid */}
          <div className="grid grid-cols-2 lg:grid-cols-5 gap-3.5 mb-3.5">
            {/* KPI 1: Opening Inv */}
            <div className="bg-white p-4 rounded-[22px] shadow-[0_8px_24px_rgba(21,75,226,0.12)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.22)] border border-[#154be2]/15 flex flex-col justify-between relative overflow-hidden group transition-all duration-300">
              <div className="absolute top-0 right-0 p-3 opacity-10 group-hover:scale-110 transition-transform">
                <span className="material-symbols-outlined text-[36px] text-amber-500">
                  inventory_2
                </span>
              </div>
              <p className="text-[10px] md:text-[11px] font-extrabold text-[#5c648e] uppercase tracking-wider leading-none mb-1">
                OPENING INV
              </p>
              <div className="mt-2 font-sans font-bold text-amber-600">
                <span className="text-2xl md:text-3xl font-extrabold tracking-tight">
                  {formatOverviewVal(overviewStats.totalOpeningStock, overviewUseMt).valueStr}
                </span>
                <span className="text-[10px] text-amber-500 font-bold ml-1">
                  {formatOverviewVal(overviewStats.totalOpeningStock, overviewUseMt).unit}
                </span>
              </div>
            </div>

            {/* KPI 2: Ending Inv */}
            <div className="bg-white p-4 rounded-[22px] shadow-[0_8px_24px_rgba(21,75,226,0.12)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.22)] border border-[#154be2]/15 flex flex-col justify-between relative overflow-hidden group transition-all duration-300">
              <div className="absolute top-0 right-0 p-3 opacity-10 group-hover:scale-110 transition-transform">
                <span className="material-symbols-outlined text-[36px] text-purple-500">
                  warehouse
                </span>
              </div>
              <p className="text-[10px] md:text-[11px] font-extrabold text-[#5c648e] uppercase tracking-wider leading-none mb-1">
                ENDING INV
              </p>
              <div className="mt-2 font-sans font-bold text-purple-600">
                <span className="text-2xl md:text-3xl font-extrabold tracking-tight">
                  {formatOverviewVal(overviewStats.totalCurrentStock, overviewUseMt).valueStr}
                </span>
                <span className="text-[10px] text-purple-500 font-bold ml-1">
                  {formatOverviewVal(overviewStats.totalCurrentStock, overviewUseMt).unit}
                </span>
              </div>
            </div>

            {/* KPI 3: Stock In */}
            <div className="bg-white p-4 rounded-[22px] shadow-[0_8px_24px_rgba(21,75,226,0.12)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.22)] border border-[#154be2]/15 flex flex-col justify-between relative overflow-hidden group transition-all duration-300">
              <div className="absolute top-0 right-0 p-3 opacity-10 group-hover:scale-110 transition-transform">
                <span className="material-symbols-outlined text-[36px] text-green-500">
                  add_shopping_cart
                </span>
              </div>
              <p className="text-[10px] md:text-[11px] font-extrabold text-[#5c648e] uppercase tracking-wider leading-none mb-1">
                STOCK IN
              </p>
              <div className="mt-2 text-green-600 font-sans font-bold">
                <span className="text-2xl md:text-3xl font-extrabold tracking-tight">
                  {formatOverviewVal(overviewStats.totalSellIn, overviewUseMt).valueStr}
                </span>
                <span className="text-[10px] text-green-500 font-bold ml-1">
                  {formatOverviewVal(overviewStats.totalSellIn, overviewUseMt).unit}
                </span>
              </div>
            </div>

            {/* KPI 4: Idle Stock */}
            <div className="bg-white p-4 rounded-[22px] shadow-[0_8px_24px_rgba(21,75,226,0.12)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.22)] border border-[#154be2]/15 flex flex-col justify-between relative overflow-hidden group transition-all duration-300">
              <div className="absolute top-0 right-0 p-3 opacity-10 group-hover:scale-110 transition-transform">
                <span className="material-symbols-outlined text-[36px] text-indigo-500">
                  hourglass_empty
                </span>
              </div>
              <p className="text-[10px] md:text-[11px] font-extrabold text-[#5c648e] uppercase tracking-wider leading-none mb-1">
                IDLE STOCK
              </p>
              <div className="mt-2 text-indigo-600 font-sans font-bold">
                <span className="text-2xl md:text-3xl font-extrabold tracking-tight">
                  {formatOverviewVal(overviewStats.totalIdleStock, overviewUseMt).valueStr}
                </span>
                <span className="text-[10px] text-indigo-500 font-bold ml-1">
                  {formatOverviewVal(overviewStats.totalIdleStock, overviewUseMt).unit}
                </span>
              </div>
            </div>

            {/* KPI 5: POG */}
            <div className="bg-white p-4 rounded-[22px] shadow-[0_8px_24px_rgba(21,75,226,0.12)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.22)] border border-[#154be2]/15 flex flex-col justify-between relative overflow-hidden group transition-all duration-300 col-span-2 lg:col-span-1">
              <div className="absolute top-0 right-0 p-3 opacity-10 group-hover:scale-110 transition-transform">
                <span className="material-symbols-outlined text-[36px] text-blue-500">
                  trending_up
                </span>
              </div>
              <p className="text-[10px] md:text-[11px] font-extrabold text-[#5c648e] uppercase tracking-wider leading-none mb-1">
                POG
              </p>
              <div className="mt-2 text-blue-600 font-sans font-bold">
                <span className="text-2xl md:text-3xl font-extrabold tracking-tight">
                  {formatOverviewVal(overviewStats.totalPog, overviewUseMt).valueStr}
                </span>
                <span className="text-[10px] text-blue-500 font-bold ml-1">
                  {formatOverviewVal(overviewStats.totalPog, overviewUseMt).unit}
                </span>
              </div>
            </div>
          </div>

          <div className="mb-3.5">
            {renderGrandTotalCard()}
          </div>

          {/* Charts Grid Row 1 */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-6">
            {/* Chart 1: Sales (POG) & Stock per Area / Dimension */}
            <div className="lg:col-span-2 bg-white p-6 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.03)] border border-[#154be2]/5 flex flex-col justify-between">
              <div className="flex flex-col gap-4 mb-6 pb-4 border-b border-[#f0effc]/60">
                <div className="flex items-start justify-between">
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="material-symbols-outlined text-primary text-sm font-semibold border-b border-primary/20 pb-0.5">
                        {overviewGroupDimension === "material"
                          ? "widgets"
                          : overviewGroupDimension === "province"
                            ? "map"
                            : "analytics"}
                      </span>
                      <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                        {overviewGroupDimension === "area"
                          ? "Performa Kinerja Wilayah (Area)"
                          : overviewGroupDimension === "province"
                            ? "Performa Kinerja per Provinsi"
                            : overviewGroupDimension === "sales_agronomist"
                              ? "Performa Kinerja Sales Agronomist (SA)"
                              : overviewGroupDimension === "business_solution"
                                ? "Performa Kinerja Business Solution (BS)"
                                : overviewGroupDimension === "distributor"
                                  ? "Performa Kinerja per Distributor"
                                  : "Performa Kinerja per Hybrid"}
                      </h3>
                    </div>
                    <p className="text-[10px] text-[#8E94B7] mt-0.5">
                      {overviewMetricFilter === "movement"
                        ? "Analisis perbandingan Stock In (Stok Masuk) vs POG (Penjualan)"
                        : overviewMetricFilter === "idle"
                          ? "Analisis perbandingan Idle Stock vs Sisa Stok Akhir"
                          : overviewMetricFilter === "total_stock"
                            ? "Analisis total sisa stok akhir"
                            : overviewMetricFilter === "Opening"
                              ? "Analisis perbandingan Stok Awal vs Sisa Stok Akhir"
                              : "Analisis perbandingan POG (Penjualan) vs Sisa Stok Akhir"}
                    </p>
                  </div>

                  <button
                    onClick={() => setIsChartFocusedModalOpen(true)}
                    className="flex items-center gap-1.5 px-3 py-1.5 text-[10px] font-extrabold text-[#154be2] bg-[#154be2]/5 hover:bg-[#154be2]/10 active:bg-[#154be2]/20 transition-all rounded-xl border border-[#154be2]/10 shrink-0"
                  >
                    <span className="material-symbols-outlined text-sm font-bold">zoom_in</span>
                    <span>Fokus Grafik</span>
                  </button>
                </div>

                <div className="flex flex-wrap items-center gap-3">
                  {/* Selector Filter 2: Dimensi Grouping */}
                  <div className="flex items-center gap-2 bg-[#fbfaff] px-2.5 py-1 rounded-xl border border-[#e2e8f0]/80">
                    <span className="text-[9.5px] font-bold text-[#8E94B7] uppercase tracking-wider">
                      Dimensi:
                    </span>
                    <div className="relative">
                      <select
                        value={overviewGroupDimension}
                        onChange={(e: any) =>
                          setOverviewGroupDimension(e.target.value as any)
                        }
                        className="bg-transparent text-[10.5px] font-black text-[#154be2] focus:outline-none focus:ring-0 appearance-none cursor-pointer pr-6 py-0.5"
                      >
                        <option value="area">Area</option>
                        <option value="province">Province</option>
                        <option value="sales_agronomist">
                          Sales Agronomist
                        </option>
                        <option value="business_solution">
                          Business Solution
                        </option>
                        <option value="material">Hybrid</option>
                        <option value="distributor">Distributor</option>
                      </select>
                      <span className="material-symbols-outlined absolute right-0 top-1/2 -translate-y-1/2 text-[14px] text-primary pointer-events-none">
                        expand_more
                      </span>
                    </div>
                  </div>

                  {/* Selector Filter 1 (Metric) */}
                  <div className="flex items-center gap-2 bg-[#fbfaff] px-2.5 py-1 rounded-xl border border-[#e2e8f0]/80">
                    <span className="text-[9.5px] font-bold text-[#8E94B7] uppercase tracking-wider">
                      Metrik:
                    </span>
                    <div className="relative">
                      <select
                        value={overviewMetricFilter}
                        onChange={(e: any) =>
                          setOverviewMetricFilter(e.target.value)
                        }
                        className="bg-transparent text-[10.5px] font-black text-[#154be2] focus:outline-none focus:ring-0 appearance-none cursor-pointer pr-6 py-0.5"
                      >
                        <option value="movement">Movement</option>
                        <option value="idle">Idle Stock</option>
                        <option value="total_stock">Total Stock</option>
                      </select>
                      <span className="material-symbols-outlined absolute right-0 top-1/2 -translate-y-1/2 text-[14px] text-primary pointer-events-none">
                        expand_more
                      </span>
                    </div>
                  </div>

                  {/* Legend aligned side-by-side */}
                  <div className="flex items-center gap-3 bg-[#fbfaff] px-2.5 py-1 rounded-xl border border-[#e2e8f0]/40 shrink-0">
                    <div className="flex items-center gap-1.5">
                      <span className="size-2.5 rounded-[4px] bg-gradient-to-tr from-[#154be2] to-[#3b82f6]" />
                      <span className="text-[10px] font-extrabold text-[#4e5572]">
                        {overviewMetricFilter === "movement"
                          ? "Stock In"
                          : overviewMetricFilter === "idle"
                            ? "Idle Stock"
                            : overviewMetricFilter === "total_stock"
                              ? "Total Stock"
                              : overviewMetricFilter === "Opening"
                                ? "Stok Awal"
                                : "POG (Penjualan)"}
                      </span>
                    </div>
                    {overviewMetricFilter !== "idle" &&
                      overviewMetricFilter !== "total_stock" && (
                        <div className="flex items-center gap-1.5">
                          <span className="size-2.5 rounded-[4px] bg-gradient-to-tr from-[#06b6d4] to-[#22d3ee]" />
                          <span className="text-[10px] font-extrabold text-[#4e5572]">
                            {overviewMetricFilter === "movement"
                              ? "POG"
                              : "Stok Akhir"}
                          </span>
                        </div>
                      )}
                  </div>
                </div>
              </div>

              <div className="w-full overflow-x-auto scrollbar-thin select-none">
                <div
                  style={{
                    minWidth: `${Math.max(600, overviewStats.areaChartData.length * 95)}px`,
                    width: "100%",
                    height: "280px",
                  }}
                  className="font-sans"
                >
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart
                      data={overviewStats.areaChartData}
                      margin={{ top: 25, right: 10, left: -10, bottom: 35 }}
                      barGap={currentBarGap}
                      barCategoryGap={currentBarCategoryGap}
                    >
                      <defs>
                        <linearGradient
                          id="colorAreaPog"
                          x1="0"
                          y1="0"
                          x2="0"
                          y2="1"
                        >
                          <stop
                            offset="0%"
                            stopColor="#154be2"
                            stopOpacity={0.95}
                          />
                          <stop
                            offset="100%"
                            stopColor="#3b82f6"
                            stopOpacity={0.7}
                          />
                        </linearGradient>
                        <linearGradient
                          id="colorAreaStock"
                          x1="0"
                          y1="0"
                          x2="0"
                          y2="1"
                        >
                          <stop
                            offset="0%"
                            stopColor="#06b6d4"
                            stopOpacity={0.95}
                          />
                          <stop
                            offset="100%"
                            stopColor="#22d3ee"
                            stopOpacity={0.7}
                          />
                        </linearGradient>
                      </defs>
                      <CartesianGrid
                        strokeDasharray="4 4"
                        vertical={false}
                        stroke="#e2e8f0"
                      />
                      <XAxis
                        dataKey="name"
                        tick={<CustomXAxisTick />}
                        axisLine={false}
                        tickLine={false}
                        interval={0}
                        height={45}
                      />
                      <YAxis
                        tick={{ fill: "#8E94B7", fontSize: 9, fontWeight: 500 }}
                        axisLine={false}
                        tickLine={false}
                      />
                      <Tooltip
                        cursor={{ fill: "rgba(21, 75, 226, 0.03)" }}
                        contentStyle={{
                          backgroundColor: "white",
                          borderRadius: "16px",
                          border: "1px solid #edecff",
                          boxShadow: "0 12px 32px rgba(21,75,226,0.1)",
                        }}
                        labelStyle={{
                          fontSize: "11px",
                          fontWeight: "bold",
                          color: "#181a2c",
                        }}
                        itemStyle={{ fontSize: "10px", padding: "1px 0" }}
                        formatter={(value: any, name: any) => {
                          if (value === undefined || value === null || isNaN(Number(value))) return [value, name];
                          return [Math.round(Number(value)).toLocaleString("id-ID"), name];
                        }}
                      />
                      <Bar
                        dataKey={
                          overviewMetricFilter === "movement"
                            ? "sellIn"
                            : overviewMetricFilter === "idle"
                              ? "idle"
                              : overviewMetricFilter === "total_stock"
                                ? "stock"
                                : overviewMetricFilter === "Opening"
                                  ? "opening"
                                  : "pog"
                        }
                        name={
                          overviewMetricFilter === "movement"
                            ? "Stock In"
                            : overviewMetricFilter === "idle"
                              ? "Idle Stock"
                              : overviewMetricFilter === "total_stock"
                                ? "Total Stock"
                                : overviewMetricFilter === "Opening"
                                  ? "Stok Awal"
                                  : "Penjualan (POG)"
                        }
                        fill="url(#colorAreaPog)"
                        radius={[6, 6, 0, 0]}
                        maxBarSize={currentMaxBarSize}
                      >
                        <LabelList
                          dataKey={
                            overviewMetricFilter === "movement"
                              ? "sellIn"
                              : overviewMetricFilter === "idle"
                                ? "idle"
                                : overviewMetricFilter === "total_stock"
                                  ? "stock"
                                  : overviewMetricFilter === "Opening"
                                    ? "opening"
                                    : "pog"
                          }
                          position="top"
                          offset={8}
                          style={{
                            fontSize: 9,
                            fontWeight: 700,
                            fill: "#154be2",
                            fontFamily: "sans-serif",
                          }}
                          formatter={(val: any) => {
                            if (
                              val === undefined ||
                              val === null ||
                              isNaN(Number(val))
                            )
                              return "";
                            const num = Number(val);
                            if (num === 0) return "0";
                            return Math.abs(num) < 10
                              ? num.toFixed(1)
                              : Math.round(num).toLocaleString();
                          }}
                        />
                      </Bar>
                      {overviewMetricFilter !== "idle" &&
                        overviewMetricFilter !== "total_stock" && (
                          <Bar
                            dataKey={
                              overviewMetricFilter === "movement"
                                ? "pog"
                                : "stock"
                            }
                            name={
                              overviewMetricFilter === "movement"
                                ? "POG"
                                : "Stok Akhir"
                            }
                            fill="url(#colorAreaStock)"
                            radius={[6, 6, 0, 0]}
                            maxBarSize={currentMaxBarSize}
                          >
                            <LabelList
                              dataKey={
                                overviewMetricFilter === "movement"
                                  ? "pog"
                                  : "stock"
                              }
                              position="top"
                              offset={8}
                              style={{
                                fontSize: 9,
                                fontWeight: 700,
                                fill: "#0a90a6",
                                fontFamily: "sans-serif",
                              }}
                              formatter={(val: any) => {
                                if (
                                  val === undefined ||
                                  val === null ||
                                  isNaN(Number(val))
                                )
                                  return "";
                                const num = Number(val);
                                if (num === 0) return "0";
                                return Math.abs(num) < 10
                                  ? num.toFixed(1)
                                  : Math.round(num).toLocaleString();
                              }}
                            />
                          </Bar>
                        )}
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* Chart 2: Kiosk Category Breakdown */}
            <div className="bg-white p-6 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.03)] border border-[#154be2]/5 flex flex-col justify-between">
              <div>
                <div className="flex items-center gap-2 pb-4 border-b border-[#f0effc]/60">
                  <span className="material-symbols-outlined text-purple-500 text-sm font-semibold">
                    pie_chart
                  </span>
                  <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                    Segmentasi Partner
                  </h3>
                </div>
                <p className="text-[10px] text-[#8E94B7] mt-1.5 leading-relaxed">
                  Klasifikasi kelas kios berdasarkan kriteria volume penjualan
                  nasional
                </p>
              </div>

              <div className="h-44 w-full relative my-3">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={overviewStats.categoryChartData}
                      cx="50%"
                      cy="50%"
                      innerRadius={50}
                      outerRadius={70}
                      paddingAngle={5}
                      dataKey="value"
                    >
                      {overviewStats.categoryChartData.map((entry, index) => {
                        const colors = [
                          "#154be2",
                          "#06b6d4",
                          "#a855f7",
                          "#f59e0b",
                          "#ec4899",
                        ];
                        return (
                          <Cell
                            key={`cell-${index}`}
                            fill={colors[index % colors.length]}
                            stroke="rgba(255,255,255,0.8)"
                            strokeWidth={2}
                          />
                        );
                      })}
                    </Pie>
                    <Tooltip
                      contentStyle={{
                        backgroundColor: "white",
                        borderRadius: "12px",
                        border: "1px solid #edecff",
                        boxShadow: "0 8px 24px rgba(21,75,226,0.08)",
                      }}
                      itemStyle={{ fontSize: "10px", fontWeight: "bold" }}
                    />
                  </PieChart>
                </ResponsiveContainer>

                {/* Center count indicator */}
                <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none mt-1">
                  <span className="text-2xl font-black text-[#181a2c] tracking-tight">
                    {overviewStats.totalKiosks}
                  </span>
                  <span className="text-[8px] text-[#8E94B7] font-bold uppercase tracking-wider text-center">
                    Outlet
                  </span>
                </div>
              </div>

              <div className="grid grid-cols-3 gap-2 pt-3 border-t border-[#f0effc]">
                {overviewStats.categoryChartData.map((item, index) => {
                  const colors = [
                    "#154be2",
                    "#06b6d4",
                    "#a855f7",
                    "#f59e0b",
                    "#ec4899",
                  ];
                  return (
                    <div
                      key={item.name}
                      className="flex flex-col items-center p-1 bg-[#fbfaff] rounded-lg border border-[#f0effc]"
                    >
                      <div className="flex items-center gap-1">
                        <span
                          className="size-1.5 rounded-full"
                          style={{
                            backgroundColor: colors[index % colors.length],
                          }}
                        />
                        <span className="text-[9px] font-bold text-[#181a2c]">
                          {item.name}
                        </span>
                      </div>
                      <span className="text-[10px] text-[#555a77] font-bold mt-0.5">
                        {item.value} Kios
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          {/* Section: Employee Sales Ranking (Highest & Lowest) */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
            {/* Top 5 Employees - Highest Sales */}
            <div className="bg-white p-6 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.03)] border border-[#154be2]/5 flex flex-col justify-between">
              <div>
                <div className="flex items-center gap-2">
                  <span className="material-symbols-outlined text-emerald-500 text-sm font-semibold">
                    trending_up
                  </span>
                  <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                    Top 5 Employee dengan Penjualan Tertinggi
                  </h3>
                </div>
                <p className="text-[10px] text-[#8E94B7] mt-0.5">
                  Daftar tim penjualan dengan pencapaian POG tertinggi saat ini
                </p>
              </div>

              <div className="flex flex-col gap-3 mt-4">
                {employeePerformanceData.highest.map((item, index) => (
                  <div
                    key={item.employee}
                    className="flex items-center justify-between p-3 rounded-xl bg-[#fbfaff] border border-[#f0effc] hover:border-[#154be2]/20 transition-all duration-200"
                  >
                    <div className="flex items-center gap-3">
                      <div
                        className={`size-8 shrink-0 rounded-full flex items-center justify-center font-bold text-xs ${
                          index === 0
                            ? "bg-amber-100 text-amber-700"
                            : index === 1
                              ? "bg-slate-200 text-slate-700"
                              : index === 2
                                ? "bg-orange-100 text-orange-700"
                                : "bg-[#154be2]/10 text-primary"
                        }`}
                      >
                        #{index + 1}
                      </div>
                      <div className="min-w-0">
                        <h4 className="text-[11px] font-bold text-[#181a2c] leading-tight truncate max-w-[200px] sm:max-w-[400px]">
                          {item.employee}
                        </h4>
                        <p className="text-[9px] text-[#8E94B7] mt-0.5 truncate">
                          Area:{" "}
                          <span className="font-semibold text-primary">
                            {item.area}
                          </span>
                        </p>
                      </div>
                    </div>

                    <div className="text-right shrink-0">
                      <p className="text-xs font-bold text-emerald-600 font-sans">
                        {item.pog.toLocaleString()}{" "}
                        <span className="text-[9px] text-[#8E94B7] font-normal">
                          POG
                        </span>
                      </p>
                      <p className="text-[8.5px] text-[#8E94B7] mt-0.5">
                        Sisa stok:{" "}
                        <span className="font-medium text-slate-700">
                          {item.currentStock.toLocaleString()} Kg
                        </span>
                      </p>
                    </div>
                  </div>
                ))}

                {employeePerformanceData.highest.length === 0 && (
                  <div className="flex flex-col items-center justify-center text-center p-6">
                    <span className="material-symbols-outlined text-[36px] text-[#8E94B7]/40 mb-2">
                      trending_flat
                    </span>
                    <p className="text-xs text-[#8E94B7]">
                      Data POG belum dimasukkan bulan ini.
                    </p>
                  </div>
                )}
              </div>
            </div>

            {/* Top 5 Employees - Lowest Sales */}
            <div className="bg-white p-6 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.03)] border border-[#154be2]/5 flex flex-col justify-between">
              <div>
                <div className="flex items-center gap-2">
                  <span className="material-symbols-outlined text-rose-500 text-sm font-semibold">
                    trending_down
                  </span>
                  <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                    Top 5 Employee dengan Penjualan Terendah
                  </h3>
                </div>
                <p className="text-[10px] text-[#8E94B7] mt-0.5">
                  Daftar tim penjualan dengan pencapaian POG terendah saat ini
                </p>
              </div>

              <div className="flex flex-col gap-3 mt-4">
                {employeePerformanceData.lowest.map((item, index) => (
                  <div
                    key={item.employee}
                    className="flex items-center justify-between p-3 rounded-xl bg-[#fbfaff] border border-[#f0effc] hover:border-[#154be2]/20 transition-all duration-200"
                  >
                    <div className="flex items-center gap-3">
                      <div
                        className={`size-8 shrink-0 rounded-full flex items-center justify-center font-bold text-xs ${
                          index === 0
                            ? "bg-rose-100 text-rose-700"
                            : index === 1
                              ? "bg-orange-50 text-orange-700 border border-orange-200"
                              : index === 2
                                ? "bg-orange-50 text-orange-600"
                                : "bg-slate-100 text-slate-600"
                        }`}
                      >
                        #{index + 1}
                      </div>
                      <div className="min-w-0">
                        <h4 className="text-[11px] font-bold text-[#181a2c] leading-tight truncate max-w-[200px] sm:max-w-[400px]">
                          {item.employee}
                        </h4>
                        <p className="text-[9px] text-[#8E94B7] mt-0.5 truncate">
                          Area:{" "}
                          <span className="font-semibold text-primary">
                            {item.area}
                          </span>
                        </p>
                      </div>
                    </div>

                    <div className="text-right shrink-0">
                      <p className="text-xs font-bold text-rose-600 font-sans">
                        {item.pog.toLocaleString()}{" "}
                        <span className="text-[9px] text-[#8E94B7] font-normal">
                          POG
                        </span>
                      </p>
                      <p className="text-[8.5px] text-[#8E94B7] mt-0.5">
                        Sisa stok:{" "}
                        <span className="font-medium text-slate-700">
                          {item.currentStock.toLocaleString()} Kg
                        </span>
                      </p>
                    </div>
                  </div>
                ))}

                {employeePerformanceData.lowest.length === 0 && (
                  <div className="flex flex-col items-center justify-center text-center p-6">
                    <span className="material-symbols-outlined text-[36px] text-[#8E94B7]/40 mb-2">
                      trending_flat
                    </span>
                    <p className="text-xs text-[#8E94B7]">
                      Data POG belum dimasukkan bulan ini.
                    </p>
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Section: History Bulanan */}
          <div className="bg-white p-6 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.03)] border border-[#154be2]/5 flex flex-col justify-between mt-6">
            <div className="flex flex-col md:flex-row md:items-center md:justify-between mb-6 pb-4 border-b border-[#f0effc]/60 gap-4">
              <div>
                <div className="flex items-center gap-2">
                  <span className="material-symbols-outlined text-primary text-sm font-semibold border-b border-primary/20 pb-0.5">
                    timeline
                  </span>
                  <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                    Tren Perkembangan Data (History Bulanan)
                  </h3>
                </div>
                <p className="text-[10px] text-[#8E94B7] mt-0.5">
                  Analisis perbandingan histori total volume perkembangan data
                  dari bulan ke bulan
                </p>
              </div>

              {/* Toggle switch for history display options matching requested Opening Inv, Ending Inv, Stock In, Idle Stock, POG */}
              <div className="flex flex-wrap items-center gap-1.5 self-start md:self-auto">
                <button
                  onClick={() => setHistoryChartType("opening")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all ${
                    historyChartType === "opening"
                      ? "bg-indigo-600 text-white shadow-md"
                      : "bg-[#fbfaff] text-[#8E94B7] hover:bg-slate-100 border border-slate-100/40"
                  }`}
                >
                  Opening Inv
                </button>
                <button
                  onClick={() => setHistoryChartType("ending")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all ${
                    historyChartType === "ending"
                      ? "bg-purple-600 text-white shadow-md"
                      : "bg-[#fbfaff] text-[#8E94B7] hover:bg-slate-100 border border-slate-100/40"
                  }`}
                >
                  Ending Inv
                </button>
                <button
                  onClick={() => setHistoryChartType("stockIn")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all ${
                    historyChartType === "stockIn"
                      ? "bg-emerald-600 text-white shadow-md"
                      : "bg-[#fbfaff] text-[#8E94B7] hover:bg-slate-100 border border-slate-100/40"
                  }`}
                >
                  Stock In
                </button>
                <button
                  onClick={() => setHistoryChartType("idle")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all ${
                    historyChartType === "idle"
                      ? "bg-amber-600 text-white shadow-md"
                      : "bg-[#fbfaff] text-[#8E94B7] hover:bg-slate-100 border border-slate-100/40"
                  }`}
                >
                  Idle Stock
                </button>
                <button
                  onClick={() => setHistoryChartType("pog")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all ${
                    historyChartType === "pog"
                      ? "bg-blue-600 text-white shadow-md"
                      : "bg-[#fbfaff] text-[#8E94B7] hover:bg-slate-100 border border-slate-100/40"
                  }`}
                >
                  POG
                </button>
              </div>
            </div>

            {/* Micro Stats Row for History */}
            {overviewHistoryData && overviewHistoryData.length > 0 && (
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mb-6 bg-[#fbfaff] p-4 rounded-2xl border border-slate-100">
                <div className="flex flex-col">
                  <span className="text-[9px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Rata-rata POG Bulanan
                  </span>
                  <span className="text-sm font-bold text-blue-600 mt-1">
                    {Math.round(
                      overviewHistoryData.reduce(
                        (acc, curr) => acc + curr.pog,
                        0,
                      ) / overviewHistoryData.length,
                    ).toLocaleString()}{" "}
                    <span className="text-[10px] font-medium text-slate-400">
                      Kg
                    </span>
                  </span>
                </div>
                <div className="flex flex-col">
                  <span className="text-[9px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Total Penyerapan POG
                  </span>
                  <span className="text-sm font-bold text-indigo-600 mt-1">
                    {overviewHistoryData
                      .reduce((acc, curr) => acc + curr.pog, 0)
                      .toLocaleString()}{" "}
                    <span className="text-[10px] font-medium text-slate-400">
                      Kg
                    </span>
                  </span>
                </div>
                <div className="flex flex-col">
                  <span className="text-[9px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Mutasi Maksimal Stok
                  </span>
                  <span className="text-sm font-bold text-teal-600 mt-1">
                    {Math.max(
                      ...overviewHistoryData.map((d) => d.ending),
                    ).toLocaleString()}{" "}
                    <span className="text-[10px] font-medium text-slate-400">
                      Kg
                    </span>
                  </span>
                </div>
                <div className="flex flex-col">
                  <span className="text-[9px] font-bold text-[#8E94B7] uppercase tracking-wider">
                    Bulan Teraktif
                  </span>
                  <span className="text-sm font-bold text-amber-600 mt-1 truncate">
                    {(() => {
                      const maxPogObj = [...overviewHistoryData].sort(
                        (a, b) => b.pog - a.pog,
                      )[0];
                      return maxPogObj ? maxPogObj.monthLabel : "-";
                    })()}
                  </span>
                </div>
              </div>
            )}

            <div className="h-64 w-full font-sans">
              <ResponsiveContainer width="100%" height="100%">
                <AreaChart
                  data={overviewHistoryData}
                  margin={{ top: 25, right: 15, left: -15, bottom: 5 }}
                >
                  <defs>
                    <linearGradient
                      id="historyColorOpening"
                      x1="0"
                      y1="0"
                      x2="0"
                      y2="1"
                    >
                      <stop
                        offset="5%"
                        stopColor="#4f46e5"
                        stopOpacity={0.25}
                      />
                      <stop
                        offset="95%"
                        stopColor="#4f46e5"
                        stopOpacity={0.0}
                      />
                    </linearGradient>
                    <linearGradient
                      id="historyColorEnding"
                      x1="0"
                      y1="0"
                      x2="0"
                      y2="1"
                    >
                      <stop
                        offset="5%"
                        stopColor="#a855f7"
                        stopOpacity={0.25}
                      />
                      <stop
                        offset="95%"
                        stopColor="#a855f7"
                        stopOpacity={0.0}
                      />
                    </linearGradient>
                    <linearGradient
                      id="historyColorStockIn"
                      x1="0"
                      y1="0"
                      x2="0"
                      y2="1"
                    >
                      <stop
                        offset="5%"
                        stopColor="#10b981"
                        stopOpacity={0.25}
                      />
                      <stop
                        offset="95%"
                        stopColor="#10b981"
                        stopOpacity={0.0}
                      />
                    </linearGradient>
                    <linearGradient
                      id="historyColorIdle"
                      x1="0"
                      y1="0"
                      x2="0"
                      y2="1"
                    >
                      <stop
                        offset="5%"
                        stopColor="#f59e0b"
                        stopOpacity={0.25}
                      />
                      <stop
                        offset="95%"
                        stopColor="#f59e0b"
                        stopOpacity={0.0}
                      />
                    </linearGradient>
                    <linearGradient
                      id="historyColorPog"
                      x1="0"
                      y1="0"
                      x2="0"
                      y2="1"
                    >
                      <stop
                        offset="5%"
                        stopColor="#2563eb"
                        stopOpacity={0.25}
                      />
                      <stop
                        offset="95%"
                        stopColor="#2563eb"
                        stopOpacity={0.0}
                      />
                    </linearGradient>
                  </defs>

                  <CartesianGrid
                    strokeDasharray="4 4"
                    vertical={false}
                    stroke="#e2e8f0"
                  />
                  <XAxis
                    dataKey="monthLabel"
                    tick={{ fill: "#4e5572", fontSize: 8.5, fontWeight: 700 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <YAxis
                    tick={{ fill: "#8E94B7", fontSize: 9, fontWeight: 500 }}
                    axisLine={false}
                    tickLine={false}
                  />

                  <Tooltip
                    cursor={{
                      stroke: "#154be2",
                      strokeWidth: 1,
                      strokeDasharray: "4 4",
                    }}
                    contentStyle={{
                      backgroundColor: "white",
                      borderRadius: "16px",
                      border: "1px solid #e1e7ff",
                      boxShadow: "0 10px 30px rgba(0,0,0,0.06)",
                    }}
                    labelStyle={{
                      fontSize: "11px",
                      fontWeight: "bold",
                      color: "#181a2c",
                    }}
                    itemStyle={{ fontSize: "10px", padding: "1px 0" }}
                  />

                  <Legend
                    iconSize={10}
                    iconType="circle"
                    wrapperStyle={{
                      fontSize: "10px",
                      fontWeight: "bold",
                      marginTop: "10px",
                    }}
                  />

                  {historyChartType === "opening" && (
                    <Area
                      type="monotone"
                      dataKey="opening"
                      name="Opening Inv (Kg)"
                      stroke="#4f46e5"
                      strokeWidth={2.5}
                      fillOpacity={1}
                      fill="url(#historyColorOpening)"
                      dot={{
                        r: 4,
                        strokeWidth: 2,
                        stroke: "#4f46e5",
                        fill: "#ffffff",
                      }}
                      activeDot={{ r: 6, strokeWidth: 1 }}
                    >
                      <LabelList
                        dataKey="opening"
                        position="top"
                        offset={10}
                        style={{
                          fontSize: 9,
                          fontWeight: 700,
                          fill: "#4f46e5",
                        }}
                        formatter={(value: any) =>
                          value
                            ? Math.round(Number(value)).toLocaleString()
                            : "0"
                        }
                      />
                    </Area>
                  )}

                  {historyChartType === "ending" && (
                    <Area
                      type="monotone"
                      dataKey="ending"
                      name="Ending Inv (Kg)"
                      stroke="#a855f7"
                      strokeWidth={2}
                      fillOpacity={1}
                      fill="url(#historyColorEnding)"
                      dot={{
                        r: 4,
                        strokeWidth: 2,
                        stroke: "#a855f7",
                        fill: "#ffffff",
                      }}
                      activeDot={{ r: 6, strokeWidth: 1 }}
                    >
                      <LabelList
                        dataKey="ending"
                        position="top"
                        offset={10}
                        style={{
                          fontSize: 9,
                          fontWeight: 700,
                          fill: "#a855f7",
                        }}
                        formatter={(value: any) =>
                          value
                            ? Math.round(Number(value)).toLocaleString()
                            : "0"
                        }
                      />
                    </Area>
                  )}

                  {historyChartType === "stockIn" && (
                    <Area
                      type="monotone"
                      dataKey="stockIn"
                      name="Stock In (Kg)"
                      stroke="#10b981"
                      strokeWidth={1.5}
                      fillOpacity={1}
                      fill="url(#historyColorStockIn)"
                      dot={{
                        r: 4,
                        strokeWidth: 2,
                        stroke: "#10b981",
                        fill: "#ffffff",
                      }}
                      activeDot={{ r: 6, strokeWidth: 1 }}
                    >
                      <LabelList
                        dataKey="stockIn"
                        position="top"
                        offset={10}
                        style={{
                          fontSize: 9,
                          fontWeight: 700,
                          fill: "#10b981",
                        }}
                        formatter={(value: any) =>
                          value
                            ? Math.round(Number(value)).toLocaleString()
                            : "0"
                        }
                      />
                    </Area>
                  )}

                  {historyChartType === "idle" && (
                    <Area
                      type="monotone"
                      dataKey="idle"
                      name="Idle Stock (Kg)"
                      stroke="#f59e0b"
                      strokeWidth={1.5}
                      fillOpacity={1}
                      fill="url(#historyColorIdle)"
                      dot={{
                        r: 4,
                        strokeWidth: 2,
                        stroke: "#f59e0b",
                        fill: "#ffffff",
                      }}
                      activeDot={{ r: 6, strokeWidth: 1 }}
                    >
                      <LabelList
                        dataKey="idle"
                        position="top"
                        offset={10}
                        style={{
                          fontSize: 9,
                          fontWeight: 700,
                          fill: "#f59e0b",
                        }}
                        formatter={(value: any) =>
                          value
                            ? Math.round(Number(value)).toLocaleString()
                            : "0"
                        }
                      />
                    </Area>
                  )}

                  {historyChartType === "pog" && (
                    <Area
                      type="monotone"
                      dataKey="pog"
                      name="POG (Kg)"
                      stroke="#2563eb"
                      strokeWidth={2}
                      fillOpacity={1}
                      fill="url(#historyColorPog)"
                      dot={{
                        r: 4,
                        strokeWidth: 2,
                        stroke: "#2563eb",
                        fill: "#ffffff",
                      }}
                      activeDot={{ r: 6, strokeWidth: 1 }}
                    >
                      <LabelList
                        dataKey="pog"
                        position="top"
                        offset={10}
                        style={{
                          fontSize: 9,
                          fontWeight: 700,
                          fill: "#2563eb",
                        }}
                        formatter={(value: any) =>
                          value
                            ? Math.round(Number(value)).toLocaleString()
                            : "0"
                        }
                      />
                    </Area>
                  )}
                </AreaChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Focused Chart Modal */}
          {isChartFocusedModalOpen && (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
              <div className="relative w-full max-w-5xl bg-white rounded-[28px] shadow-[0_24px_64px_rgba(21,75,226,0.15)] border border-[#154be2]/10 p-6 md:p-8 flex flex-col justify-between max-h-[90vh] overflow-hidden">
                {/* Header */}
                <div className="flex items-start justify-between pb-4 border-b border-[#f0effc]">
                  <div>
                    <div className="flex items-center gap-2 text-primary font-bold">
                      <span className="material-symbols-outlined text-[#154be2]">analytics</span>
                      <span className="text-[14px] uppercase tracking-wider text-[#154be2] font-black">Detail Fokus Grafik</span>
                    </div>
                    <h3 className="text-lg font-black text-[#181a2c] tracking-tight mt-1">
                      {overviewGroupDimension === "area"
                        ? "Performa Kinerja Wilayah (Area)"
                        : overviewGroupDimension === "province"
                          ? "Performa Kinerja per Provinsi"
                          : overviewGroupDimension === "sales_agronomist"
                            ? "Performa Kinerja Sales Agronomist (SA)"
                            : overviewGroupDimension === "business_solution"
                              ? "Performa Kinerja Business Solution (BS)"
                              : overviewGroupDimension === "distributor"
                                ? "Performa Kinerja per Distributor"
                                : "Performa Kinerja per Hybrid"}
                    </h3>
                    <p className="text-xs text-[#8E94B7] mt-1">
                      {overviewMetricFilter === "movement"
                        ? "Analisis perbandingan Stock In (Stok Masuk) vs POG (Penjualan)"
                        : overviewMetricFilter === "idle"
                          ? "Analisis perbandingan Idle Stock vs Sisa Stok Akhir"
                          : overviewMetricFilter === "total_stock"
                            ? "Analisis total sisa stok akhir"
                            : overviewMetricFilter === "Opening"
                              ? "Analisis perbandingan Stok Awal vs Sisa Stok Akhir"
                              : "Analisis perbandingan POG (Penjualan) vs Sisa Stok Akhir"}
                    </p>
                  </div>
                  <button
                    onClick={() => setIsChartFocusedModalOpen(false)}
                    className="p-2 rounded-full hover:bg-slate-100 transition-colors text-slate-400 hover:text-slate-700"
                  >
                    <span className="material-symbols-outlined text-xl">close</span>
                  </button>
                </div>

                {/* Controls inside Modal */}
                <div className="flex flex-wrap items-center justify-between gap-4 py-4 bg-[#fbfaff] px-4 rounded-2xl border border-[#e2e8f0]/60 my-4 shrink-0">
                  <div className="flex flex-wrap items-center gap-4">
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider">Dimensi:</span>
                      <select
                        value={overviewGroupDimension}
                        onChange={(e: any) => setOverviewGroupDimension(e.target.value as any)}
                        className="bg-white border border-[#e2e8f0] rounded-xl px-2.5 py-1 text-xs font-black text-[#154be2] focus:outline-none focus:ring-1 focus:ring-[#154be2] cursor-pointer"
                      >
                        <option value="area">Area</option>
                        <option value="province">Province</option>
                        <option value="sales_agronomist">Sales Agronomist</option>
                        <option value="business_solution">Business Solution</option>
                        <option value="material">Hybrid</option>
                        <option value="distributor">Distributor</option>
                      </select>
                    </div>
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider">Metrik:</span>
                      <select
                        value={overviewMetricFilter}
                        onChange={(e: any) => setOverviewMetricFilter(e.target.value as any)}
                        className="bg-white border border-[#e2e8f0] rounded-xl px-2.5 py-1 text-xs font-black text-[#154be2] focus:outline-none focus:ring-1 focus:ring-[#154be2] cursor-pointer"
                      >
                        <option value="movement">Movement</option>
                        <option value="idle">Idle Stock</option>
                        <option value="total_stock">Total Stock</option>
                      </select>
                    </div>
                  </div>

                  {/* Legend */}
                  <div className="flex items-center gap-3">
                    <div className="flex items-center gap-1.5">
                      <span className="size-3 rounded-[4px] bg-gradient-to-tr from-[#154be2] to-[#3b82f6]" />
                      <span className="text-xs font-extrabold text-[#4e5572]">
                        {overviewMetricFilter === "movement"
                          ? "Stock In"
                          : overviewMetricFilter === "idle"
                            ? "Idle Stock"
                            : overviewMetricFilter === "total_stock"
                              ? "Total Stock"
                              : overviewMetricFilter === "Opening"
                                ? "Stok Awal"
                                : "POG (Penjualan)"}
                      </span>
                    </div>
                    {overviewMetricFilter !== "idle" && overviewMetricFilter !== "total_stock" && (
                      <div className="flex items-center gap-1.5">
                        <span className="size-3 rounded-[4px] bg-gradient-to-tr from-[#06b6d4] to-[#22d3ee]" />
                        <span className="text-xs font-extrabold text-[#4e5572]">
                          {overviewMetricFilter === "movement" ? "POG" : "Stok Akhir"}
                        </span>
                      </div>
                    )}
                  </div>
                </div>

                {/* Chart container in Modal */}
                <div className="w-full flex-1 min-h-[350px] overflow-x-auto overflow-y-hidden scrollbar-thin select-none">
                  <div
                    style={{
                      minWidth: `${Math.max(800, overviewStats.areaChartData.length * 120)}px`,
                      width: "100%",
                      height: "350px",
                    }}
                    className="font-sans"
                  >
                    <ResponsiveContainer width="100%" height="100%">
                       <BarChart
                        data={overviewStats.areaChartData}
                        margin={{ top: 25, right: 15, left: -10, bottom: 35 }}
                        barGap={currentBarGap}
                        barCategoryGap={currentBarCategoryGap}
                      >
                        <defs>
                          <linearGradient id="modalColorAreaPog" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stopColor="#154be2" stopOpacity={0.95} />
                            <stop offset="100%" stopColor="#3b82f6" stopOpacity={0.7} />
                          </linearGradient>
                          <linearGradient id="modalColorAreaStock" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stopColor="#06b6d4" stopOpacity={0.95} />
                            <stop offset="100%" stopColor="#22d3ee" stopOpacity={0.7} />
                          </linearGradient>
                        </defs>
                        <CartesianGrid strokeDasharray="4 4" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="name" tick={<CustomXAxisTick />} axisLine={false} tickLine={false} interval={0} height={45} />
                        <YAxis tick={{ fill: "#8E94B7", fontSize: 10, fontWeight: 500 }} axisLine={false} tickLine={false} />
                        <Tooltip
                          cursor={{ fill: "rgba(21, 75, 226, 0.03)" }}
                          contentStyle={{
                            backgroundColor: "white",
                            borderRadius: "16px",
                            border: "1px solid #edecff",
                            boxShadow: "0 12px 32px rgba(21,75,226,0.1)",
                          }}
                          labelStyle={{ fontSize: "12px", fontWeight: "bold", color: "#181a2c" }}
                          itemStyle={{ fontSize: "11px", padding: "1px 0" }}
                          formatter={(value: any, name: any) => {
                            if (value === undefined || value === null || isNaN(Number(value))) return [value, name];
                            return [Math.round(Number(value)).toLocaleString("id-ID"), name];
                          }}
                        />
                        <Bar
                          dataKey={
                            overviewMetricFilter === "movement"
                              ? "sellIn"
                              : overviewMetricFilter === "idle"
                                ? "idle"
                                : overviewMetricFilter === "total_stock"
                                  ? "stock"
                                  : overviewMetricFilter === "Opening"
                                    ? "opening"
                                    : "pog"
                          }
                          name={
                            overviewMetricFilter === "movement"
                              ? "Stock In"
                              : overviewMetricFilter === "idle"
                                ? "Idle Stock"
                                : overviewMetricFilter === "total_stock"
                                  ? "Total Stock"
                                  : overviewMetricFilter === "Opening"
                                    ? "Stok Awal"
                                    : "Penjualan (POG)"
                          }
                          fill="url(#modalColorAreaPog)"
                          radius={[6, 6, 0, 0]}
                          maxBarSize={currentMaxBarSize}
                        >
                          <LabelList
                            dataKey={
                              overviewMetricFilter === "movement"
                                ? "sellIn"
                                : overviewMetricFilter === "idle"
                                  ? "idle"
                                  : overviewMetricFilter === "total_stock"
                                    ? "stock"
                                    : overviewMetricFilter === "Opening"
                                      ? "opening"
                                      : "pog"
                            }
                            position="top"
                            offset={8}
                            style={{ fontSize: 10, fontWeight: 700, fill: "#154be2", fontFamily: "sans-serif" }}
                            formatter={(val: any) => {
                              if (val === undefined || val === null || isNaN(Number(val))) return "";
                              const num = Number(val);
                              if (num === 0) return "0";
                              return Math.abs(num) < 10 ? num.toFixed(1) : Math.round(num).toLocaleString();
                            }}
                          />
                        </Bar>
                        {overviewMetricFilter !== "idle" && overviewMetricFilter !== "total_stock" && (
                          <Bar
                            dataKey={overviewMetricFilter === "movement" ? "pog" : "stock"}
                            name={overviewMetricFilter === "movement" ? "POG" : "Stok Akhir"}
                            fill="url(#modalColorAreaStock)"
                            radius={[6, 6, 0, 0]}
                            maxBarSize={currentMaxBarSize}
                          >
                            <LabelList
                              dataKey={overviewMetricFilter === "movement" ? "pog" : "stock"}
                              position="top"
                              offset={8}
                              style={{ fontSize: 10, fontWeight: 700, fill: "#0ea5e9", fontFamily: "sans-serif" }}
                              formatter={(val: any) => {
                                if (val === undefined || val === null || isNaN(Number(val))) return "";
                                const num = Number(val);
                                if (num === 0) return "0";
                                return Math.abs(num) < 10 ? num.toFixed(1) : Math.round(num).toLocaleString();
                              }}
                            />
                          </Bar>
                        )}
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {activeTab === "home" && (
        <div className="animate-in fade-in slide-in-from-left-4 duration-300">
          {isBusinessAnalyst ? (
            <div className="bg-white p-6 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.12)] border border-[#edecff] text-center max-w-md mx-auto my-12">
              <span className="material-symbols-outlined text-[48px] text-primary/40 mb-3">
                analytics
              </span>
              <h2 className="text-base font-semibold text-[#181a2c] mb-1.5">
                Akses Analis Bisnis
              </h2>
              <p className="text-xs text-[#8E94B7] leading-relaxed">
                Sebagai Business Analyst, Anda tidak perlu menginput stok secara
                manual untuk masing-masing Toko/Channel Partner melainkan
                memantau visualisasi, performa, dan ringkasan data. Silakan buka
                tab <strong className="text-primary font-medium">Stock</strong>{" "}
                atau <strong className="text-primary font-medium">POG</strong>.
              </p>
            </div>
          ) : (
            <>
              <div className="mb-4 ml-1">
                <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
                  Channel{" "}
                  <span className="text-primary font-bold">Partner</span>
                </h1>
              </div>
              <div className="relative mb-8" ref={dropdownRef}>
                <button
                  onClick={() => setIsDropdownOpen(!isDropdownOpen)}
                  className="flex items-center justify-between w-full px-6 py-4 bg-gradient-to-r from-primary to-cyan-400 text-white rounded-full shadow-[0_12px_32px_rgba(21,75,226,0.25)] hover:shadow-[0_16px_40px_rgba(21,75,226,0.35)] hover:scale-[1.01] active:scale-[0.99] transition-all relative z-30"
                >
                  <span className="font-semibold text-sm truncate">
                    {selectedKiosk}
                  </span>
                  <span className="material-symbols-outlined text-white/80">
                    unfold_more
                  </span>
                </button>
                {isDropdownOpen && (
                  <div className="absolute left-0 right-0 mt-2 bg-white/95 backdrop-blur-md shadow-[0_12px_32px_rgba(21,75,226,0.15)] z-[100] rounded-[24px] p-4 animate-in fade-in slide-in-from-top-4">
                    <input
                      type="text"
                      placeholder="Search channel..."
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className="w-full h-11 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.06)] rounded-full px-5 text-xs font-semibold outline-none focus:ring-1 focus:ring-primary/20 mb-3"
                    />
                    <div className="max-h-[200px] overflow-y-auto custom-scrollbar space-y-1">
                      {filteredKiosks.map((k, i) => (
                        <button
                          key={i}
                          onClick={() => {
                            setSelectedKiosk(k.name);
                            setIsDropdownOpen(false);
                            setSearchTerm("");
                          }}
                          className={`w-full text-left px-5 py-2.5 rounded-full font-semibold text-[11px] uppercase transition-all ${selectedKiosk === k.name ? "bg-[#edecff] text-primary" : "text-[#635b6e] hover:bg-[#fbf8ff]"}`}
                        >
                          {k.name}
                        </button>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              <div className="mb-4 ml-1">
                <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
                  Tambah <span className="text-primary font-bold">LOT</span>
                </h1>
              </div>
              <div className="bg-white p-5 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.15)] mb-6">
                <div className="space-y-4">
                  <div className="grid grid-cols-5 gap-3">
                    <div className="col-span-3 relative">
                      <button
                        type="button"
                        onClick={() => !isLotChecking && setIsScannerOpen(true)}
                        className={`absolute left-3 top-1/2 -translate-y-1/2 size-9 rounded-full flex items-center justify-center transition-all z-20 ${
                          isLotChecking
                            ? "cursor-not-allowed text-primary animate-spin"
                            : "cursor-pointer hover:bg-[#f4f2ff] text-primary active:scale-[0.93]"
                        }`}
                        title="Scan QR / Barcode"
                      >
                        <span className="material-symbols-outlined text-lg leading-none">
                          {isLotChecking ? "sync" : "photo_camera"}
                        </span>
                      </button>
                      <input
                        value={lotNo}
                        onChange={(e) => setLotNo(e.target.value)}
                        className="w-full h-14 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full pl-14 pr-6 font-semibold text-xs outline-none focus:ring-1 focus:ring-primary/20 transition-all"
                        placeholder="Batch / Lot No"
                      />
                    </div>
                    <input
                      value={qty}
                      onChange={(e) => setQty(e.target.value)}
                      className="col-span-2 h-14 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-6 font-semibold text-xs outline-none focus:ring-1 focus:ring-primary/20 transition-all"
                      placeholder="Qty (Kg)"
                      type="number"
                    />
                  </div>

                  {isLotChecking && (
                    <div className="bg-[#f0f3ff] border border-[#dce2ff] p-3.5 rounded-[18px] flex items-center gap-2.5 animate-pulse shadow-sm">
                      <span className="material-symbols-outlined text-primary text-lg animate-spin">
                        sync
                      </span>
                      <p className="text-xs font-semibold text-primary">
                        Mengecek LOT di database...
                      </p>
                    </div>
                  )}

                  {!isLotChecking &&
                    lotIntel &&
                    typeof lotIntel === "object" && (
                      <div className="space-y-2">
                        <div className="bg-emerald-50 border border-emerald-100 p-3.5 rounded-[18px] flex items-center gap-2.5 animate-in slide-in-from-top-2 shadow-sm">
                          <span className="material-symbols-outlined text-emerald-500 text-lg">
                            check_circle
                          </span>
                          <p className="text-xs font-semibold text-emerald-800">
                            ‚úÖ LOT ditemukan di database!
                          </p>
                        </div>
                        <div className="bg-white shadow-[0_4px_16px_rgba(21,75,226,0.06)] p-4 rounded-[20px] animate-in slide-in-from-top-2">
                          <div className="mb-3">
                            <p className="text-[8.5px] font-bold text-[#8E94B7] uppercase tracking-wider mb-0.5">
                              Hybrid Description
                            </p>
                            <p className="text-xs font-semibold text-[#181a2c] leading-snug">
                              {lotIntel.desc}
                            </p>
                          </div>
                          <div className="flex items-center gap-4 border-t border-[#edecff] pt-3">
                            <div className="flex-1">
                              <p className="text-[8.5px] font-bold text-[#8E94B7] uppercase tracking-wider mb-0.5">
                                Dr Date
                              </p>
                              <p className="text-[10px] font-semibold text-[#181a2c]">
                                {lotIntel.drDate}
                              </p>
                            </div>
                            <div className="w-px h-6 bg-[#edecff]" />
                            <div className="flex-1">
                              <p className="text-[8.5px] font-bold text-red-400 uppercase tracking-wider mb-0.5">
                                Exp Date
                              </p>
                              <p className="text-[10px] font-bold text-red-700">
                                {(!lotIntel.expDate || lotIntel.expDate === "N/A" || lotIntel.expDate === "-") ? extractExpFromLot(lotNo) : lotIntel.expDate}
                              </p>
                            </div>
                          </div>
                        </div>
                        {lotIntel.channel && cleanForMatch(lotIntel.channel) !== cleanForMatch(selectedKiosk) && (
                          <div className="bg-amber-50 border border-amber-200 p-3.5 rounded-[18px] flex items-start gap-2.5 shadow-sm animate-in fade-in slide-in-from-top-2">
                            <span className="material-symbols-outlined text-amber-600 text-lg shrink-0 mt-0.5">
                              warning
                            </span>
                            <div>
                              <p className="text-xs font-bold text-amber-800">
                                Perhatian: LOT milik Channel lain!
                              </p>
                              <p className="text-[10px] font-medium text-amber-700 mt-0.5 leading-relaxed">
                                Menurut database, LOT ini dikirim ke toko: <span className="font-bold uppercase">{lotIntel.channel}</span>.
                              </p>
                            </div>
                          </div>
                        )}
                      </div>
                    )}

                  {!isLotChecking &&
                    lotIntel &&
                    typeof lotIntel === "string" && (
                      <div className="space-y-2">
                        <div className="bg-emerald-50 border border-emerald-100 p-3.5 rounded-[18px] flex items-center gap-2.5 animate-in slide-in-from-top-2 shadow-sm">
                          <span className="material-symbols-outlined text-emerald-500 text-lg">
                            check_circle
                          </span>
                          <p className="text-xs font-semibold text-emerald-800">
                            ‚úÖ LOT ditemukan di database!
                          </p>
                        </div>
                        <div className="bg-white shadow-[0_4px_16px_rgba(21,75,226,0.06)] p-4 rounded-[18px] animate-in slide-in-from-top-2">
                          <p className="text-[8.5px] font-bold text-primary uppercase mb-0.5">
                            Lot Metadata Note
                          </p>
                          <p className="text-xs font-semibold text-[#181a2c]">
                            {lotIntel}
                          </p>
                        </div>
                      </div>
                    )}

                  {!isLotChecking && isLotNotFound && (
                    <div className="space-y-4 animate-in fade-in duration-200">
                      {/* Warning Message */}
                      <div className="bg-red-50 border border-red-100 p-3.5 rounded-[18px] flex items-center gap-2.5 shadow-sm">
                        <span className="material-symbols-outlined text-red-500 text-lg flex-shrink-0">
                          warning
                        </span>
                        <div className="flex flex-col">
                          <p className="text-xs font-bold text-red-700 leading-tight">
                            ‚ö†Ô∏è LOT tidak ditemukan di database!
                          </p>
                          <p className="text-[10px] text-red-600 mt-0.5 font-medium leading-relaxed">
                            LOT ini belum terdaftar. Agar data stock tetap valid, Anda wajib melengkapi data Hybrid & Komoditas secara manual di bawah ini.
                          </p>
                        </div>
                      </div>

                      {/* 1. Hybrid/Varietas Input */}
                      <div className="space-y-1">
                        <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 block">
                          Hybrid / Varietas <span className="text-red-500">*</span>
                        </label>
                        <div className="relative">
                          <select
                            value={manualHybrid}
                            onChange={(e) => setManualHybrid(e.target.value)}
                            className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all appearance-none pr-10"
                            required
                          >
                            <option value="">-- Pilih Hybrid / Varietas --</option>
                            {filterOptions.materials.map((mat) => (
                              <option key={mat} value={mat}>
                                {mat}
                              </option>
                            ))}
                            <option value="CUSTOM">-- Ketik Manual (Varietas Baru) --</option>
                          </select>
                          <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                            expand_more
                          </span>
                        </div>
                        {manualHybrid === "CUSTOM" && (
                          <input
                            type="text"
                            value={manualHybridCustom}
                            onChange={(e) => setManualHybridCustom(e.target.value)}
                            placeholder="Ketik Nama Hybrid Baru..."
                            className="w-full h-11 mt-2 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all"
                            required
                          />
                        )}
                      </div>

                      {/* 2. Commodity / Crop Input */}
                      <div className="space-y-1">
                        <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 block">
                          Komoditas <span className="text-red-500">*</span>
                        </label>
                        <div className="relative">
                          <select
                            value={manualCrop}
                            onChange={(e) => setManualCrop(e.target.value)}
                            className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all appearance-none pr-10"
                            required
                          >
                            {["Field Corn", "Fresh Corn", "Vegetables"].map((crop) => (
                              <option key={crop} value={crop}>
                                {crop}
                              </option>
                            ))}
                          </select>
                          <span className="material-symbols-outlined absolute right-4 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none text-lg">
                            expand_more
                          </span>
                        </div>
                      </div>

                      {/* 3. Dates Grid */}
                      <div className="grid grid-cols-2 gap-3">
                        {/* Shipping Date */}
                        <div className="space-y-1">
                          <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 block">
                            Tgl DR / Shipping
                          </label>
                          <input
                            type="date"
                            value={manualDrDate}
                            onChange={(e) => setManualDrDate(e.target.value)}
                            className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all"
                          />
                        </div>

                        {/* Expired Date */}
                        <div className="space-y-1">
                          <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wider ml-1 block">
                            Expired Date
                          </label>
                          <input
                            type="date"
                            value={manualExpDate}
                            onChange={(e) => setManualExpDate(e.target.value)}
                            className="w-full h-11 bg-[#fbf8ff] border border-[#edecff] rounded-xl px-4 font-bold text-xs text-[#111] outline-none focus:border-primary transition-all"
                          />
                        </div>
                      </div>
                    </div>
                  )}

                  <button
                    onClick={handleAddLocal}
                    disabled={
                      !lotNo ||
                      !qty ||
                      (isLotNotFound &&
                        (!manualHybrid ||
                          (manualHybrid === "CUSTOM" && !manualHybridCustom.trim()) ||
                          !manualCrop))
                    }
                    className={`w-full h-14 rounded-full font-semibold text-xs uppercase tracking-wider flex items-center justify-center gap-2 transition-all active:scale-[0.98] ${
                      !lotNo ||
                      !qty ||
                      (isLotNotFound &&
                        (!manualHybrid ||
                          (manualHybrid === "CUSTOM" && !manualHybridCustom.trim()) ||
                          !manualCrop))
                        ? "bg-[#e0e0fa] text-[#8E94B7] shadow-none cursor-not-allowed"
                        : "bg-gradient-to-r from-emerald-500 to-[#00D2FF] text-white shadow-[0_8px_20px_rgba(16,185,129,0.2)] hover:opacity-95"
                    }`}
                  >
                    <span className="material-symbols-outlined text-[18px]">
                      add_box
                    </span>{" "}
                    Tambah ke List
                  </button>
                </div>
              </div>

              <DetailItemSection
                items={workingData.filter(
                  (item) =>
                    cleanForMatch(item.kiosk) === cleanForMatch(selectedKiosk) ||
                    matchNames(item.kiosk, selectedKiosk),
                )}
                onEdit={(item) => setEditModal({ isOpen: true, item })}
                onDelete={(item) => setDeleteModal({ isOpen: true, item })}
                onUploadActivity={handleUploadActivity}
                isSyncing={isSyncing}
                hasChanges={hasChanges}
                title="Existing Stock"
                subtitle={selectedKiosk}
                category={
                  kiosks.find(
                    (k) =>
                      cleanForMatch(k.name) === cleanForMatch(selectedKiosk),
                  )?.category || "Uncategorized"
                }
                kiosks={kiosks}
              />

              <EditModal
                isOpen={editModal.isOpen}
                item={editModal.item}
                onClose={() => setEditModal({ isOpen: false, item: null })}
                onSave={handleEditLocal}
                isSaving={isActionLoading}
                allHybrids={filterOptions.materials}
              />
              <ConfirmModal
                isOpen={deleteModal.isOpen}
                onClose={() => setDeleteModal({ isOpen: false, item: null })}
                onConfirm={handleDeleteLocal}
                isProcessing={isActionLoading}
              />
              <QrScanModal
                isOpen={isScannerOpen}
                onClose={() => setIsScannerOpen(false)}
                onScanSuccess={(val) => setLotNo(val)}
              />
            </>
          )}
        </div>
      )}

      {activeTab === "partner" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300">
          <div className="mb-6 ml-1 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            <div>
              <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
                Mapping <span className="text-primary font-bold">Partner</span>
              </h1>
              <p className="text-[11px] text-[#8E94B7] font-semibold uppercase tracking-wider mt-0.5">
                Kelola Channel Area & Tim
              </p>
            </div>

            {/* Sub-navigation Menu */}
            <div className="flex gap-1 p-1 bg-[#edecff]/45 rounded-full w-full md:w-auto max-w-xs md:max-w-none border border-[#edecff]/60 shadow-[inset_0_1px_2px_rgba(21,75,226,0.03)] shrink-0 self-start md:self-auto">
              <button
                  onClick={() => setPartnerSubTab("team")}
                  className={`flex-1 md:flex-initial px-4 py-1.5 rounded-full font-bold text-[9.5px] uppercase tracking-wider transition-all duration-200 cursor-pointer ${
                    partnerSubTab === "team"
                      ? "bg-gradient-to-r from-primary to-cyan-400 text-white shadow-[0_4px_10px_rgba(21,75,226,0.18)]"
                      : "text-[#8E94B7] hover:bg-white/50 hover:text-[#181a2c]"
                  }`}
                >
                  <div className="flex items-center justify-center gap-1.5">
                    <span className="material-symbols-outlined text-[13px] leading-none">
                      groups
                    </span>
                    <span>Tim & Hirarki</span>
                  </div>
                </button>
                <button
                  onClick={() => setPartnerSubTab("channel")}
                  className={`flex-1 md:flex-initial px-4 py-1.5 rounded-full font-bold text-[9.5px] uppercase tracking-wider transition-all duration-200 cursor-pointer ${
                    partnerSubTab === "channel"
                      ? "bg-gradient-to-r from-primary to-cyan-400 text-white shadow-[0_4px_10px_rgba(21,75,226,0.18)]"
                      : "text-[#8E94B7] hover:bg-white/50 hover:text-[#181a2c]"
                  }`}
                >
                  <div className="flex items-center justify-center gap-1.5">
                    <span className="material-symbols-outlined text-[13px] leading-none">
                      storefront
                    </span>
                    <span>Mapping Channel</span>
                  </div>
                </button>
              </div>
          </div>

          {partnerSubTab === "team" &&
            (() => {
              const getDirectSubordinates = (parentName: string) => {
                return teamMembers.filter((member) => {
                  if (matchNames(member, parentName)) return false;
                  if (matchNames(member, userData.name)) return false;

                  const mPos = normalizePosition(
                    getFromRecord<string>(teamPositions, member) || "",
                  );
                  if (mPos === "Unknown") return false;

                  const uplineResolved = getUplineInTeam(
                    member,
                    teamMembers,
                    teamUpLines,
                  );
                  return (
                    uplineResolved !== null &&
                    matchNames(uplineResolved, parentName)
                  );
                });
              };

              const renderRecursiveTeamNode = (
                parentName: string,
                depth = 1,
              ): React.ReactNode => {
                const directSubordinates = getDirectSubordinates(parentName);
                if (directSubordinates.length === 0) return null;

                return (
                  <div className="space-y-4 pl-3 border-l border-[#edecff]/70 ml-1.5 mt-2 animate-in fade-in duration-300">
                    {directSubordinates.map((sub, idx) => {
                      const subPos =
                        getFromRecord<string>(teamPositions, sub) ||
                        "Sales Agronomist";
                      const subProvince =
                        getFromRecord<string>(teamProvinces, sub) || "-";
                      const isCollapsed = collapsedNodes[sub] !== false;
                      const subSubs = getDirectSubordinates(sub);
                      const hasChildren = subSubs.length > 0;

                      return (
                        <div key={idx} className="space-y-4">
                          <div className="flex items-start">
                            <div
                              onClick={() => {
                                if (hasChildren) {
                                  setCollapsedNodes((prev) => ({
                                    ...prev,
                                    [sub]: prev[sub] === false,
                                  }));
                                }
                              }}
                              className={`bg-white hover:bg-slate-50 transition-all px-4 py-2.5 rounded-[18px] border border-slate-100 shadow-[0_8px_24px_rgba(21,75,226,0.06)] hover:shadow-[0_12px_28px_rgba(21,75,226,0.12)] flex-1 flex items-center justify-between ${hasChildren ? "cursor-pointer" : "cursor-default"}`}
                            >
                              <div className="flex-1 min-w-0 pr-2 text-left">
                                <p className="font-bold text-xs text-[#181a2c]">
                                  {normalizeName(sub)}
                                </p>
                                <div className="flex items-center gap-1.5 mt-1 flex-wrap">
                                  <span
                                    className={`text-[7.5px] font-extrabold uppercase tracking-wider px-1.5 py-0.5 rounded-full border ${
                                      depth === 1
                                        ? "text-emerald-600 bg-emerald-50 border-emerald-100"
                                        : depth === 2
                                          ? "text-blue-600 bg-blue-50 border-blue-100"
                                          : "text-purple-600 bg-purple-50 border-purple-100"
                                    }`}
                                  >
                                    {normalizePosition(subPos)}
                                  </span>
                                  {subProvince && subProvince !== "-" && (
                                    <span className="text-[7.5px] font-bold text-[#8E94B7] bg-slate-50 px-1.5 py-0.5 rounded-full border border-slate-100 uppercase tracking-wide">
                                      {subProvince}
                                    </span>
                                  )}
                                </div>
                              </div>

                              <div className="flex items-center gap-2 shrink-0 z-10 relative">
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    const emp = employees.find((emp) =>
                                      matchNames(emp.name, sub),
                                    );
                                    setEmployeeEditModal({
                                      isOpen: true,
                                      item: emp || {
                                        name: sub,
                                        position: subPos,
                                        province: subProvince,
                                      },
                                    });
                                  }}
                                  className="size-7 bg-[#edecff]/60 hover:bg-[#edecff] text-primary rounded-full border border-[#c4c5d8]/40 flex items-center justify-center transition-all shadow-none cursor-pointer"
                                  title="Edit Karyawan"
                                >
                                  <span className="material-symbols-outlined text-[13px]">
                                    edit
                                  </span>
                                </button>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    const emp = employees.find((emp) =>
                                      matchNames(emp.name, sub),
                                    );
                                    setEmployeeDeleteModal({
                                      isOpen: true,
                                      item: emp || { name: sub },
                                    });
                                  }}
                                  className="size-7 bg-red-50 hover:bg-red-100 text-red-500 rounded-full border border-red-100 flex items-center justify-center transition-all shadow-none cursor-pointer"
                                  title="Hapus Karyawan"
                                >
                                  <span className="material-symbols-outlined text-[13px]">
                                    delete
                                  </span>
                                </button>
                                {hasChildren && (
                                  <span className="material-symbols-outlined text-[#8E94B7] text-md leading-none select-none ml-1">
                                    {isCollapsed
                                      ? "expand_more"
                                      : "expand_less"}
                                  </span>
                                )}
                              </div>
                            </div>
                          </div>

                          {hasChildren &&
                            !isCollapsed &&
                            renderRecursiveTeamNode(sub, depth + 1)}
                        </div>
                      );
                    })}
                  </div>
                );
              };

              const rootSubordinates = getDirectSubordinates(userData.name);

              return (
                <div className="animate-in fade-in slide-in-from-left-4 duration-300">
                  <div className="bg-white p-6 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.1)] mb-6 border border-[#edecff] animate-in fade-in slide-in-from-bottom-3 duration-300">
                    <div className="flex items-center justify-between gap-4 mb-5 pb-3 border-b border-[#edecff]">
                      <div>
                        <h3 className="text-xs font-bold text-[#181a2c] uppercase tracking-wider">
                          Hirarki Posisi Tim
                        </h3>
                      </div>
                      <button
                        disabled={userLevel === 1}
                        onClick={() =>
                          setEmployeeEditModal({
                            isOpen: true,
                            item: {
                              isAdd: true,
                              name: "",
                              position: "",
                              province: "",
                              email: "",
                              password: "",
                              upline: userData?.name || "",
                            },
                          })
                        }
                        className={`h-8 px-4 rounded-full font-bold text-[9px] uppercase tracking-wider flex items-center gap-1.5 transition-all shrink-0 ${
                          userLevel === 1
                            ? "bg-slate-200 text-slate-400 cursor-not-allowed"
                            : "bg-gradient-to-r from-teal-500 to-emerald-500 text-white hover:opacity-95 shadow-md active:scale-[97%] cursor-pointer"
                        }`}
                      >
                        <span className="material-symbols-outlined text-[13px]">
                          person_add
                        </span>
                        Tambah Anggota
                      </button>
                    </div>

                    {teamMembers.length > 0 ? (
                      <div className="space-y-4">
                        {/* Root Level 0 (Logged In User) */}
                        <div className="flex items-start gap-3">
                          <div className="size-7 rounded-full bg-primary/10 flex items-center justify-center font-bold text-[10px] text-primary shrink-0 shadow-sm border border-primary/10 mt-1.5">
                            0
                          </div>
                          <div
                            onClick={() => {
                              setCollapsedNodes((prev) => ({
                                ...prev,
                                [userData.name]: prev[userData.name] === false,
                              }));
                            }}
                            className={`bg-gradient-to-r from-[#edecff] to-sky-50/50 px-4 py-2.5 rounded-[18px] flex-1 relative overflow-hidden group shadow-[0_8px_24px_rgba(21,75,226,0.08)] hover:shadow-[0_12px_30px_rgba(21,75,226,0.14)] border-0 transition-all flex items-center justify-between ${rootSubordinates.length > 0 ? "cursor-pointer" : "cursor-default"}`}
                          >
                            <div className="absolute top-0 right-0 h-full w-24 bg-gradient-to-l from-primary/5 to-transparent pointer-events-none" />
                            <div className="flex-1 min-w-0">
                              <p className="font-bold text-xs text-[#181a2c]">
                                {normalizeName(userData.name)}
                              </p>
                              <div className="flex items-center gap-1.5 mt-1">
                                <span className="text-[7.5px] font-extrabold text-primary bg-primary/10 px-2 py-0.5 rounded-full border border-primary/10 uppercase tracking-wider">
                                  {normalizePosition(
                                    userData.position || "Business Analyst",
                                  )}
                                </span>
                                {userData.province &&
                                  userData.province !== "-" && (
                                    <span className="text-[7.5px] font-bold text-[#8E94B7] bg-slate-50 px-1.5 py-0.5 rounded-full border border-slate-100 uppercase tracking-wide">
                                      {userData.province}
                                    </span>
                                  )}
                              </div>
                            </div>

                            {/* Actions & Chevron */}
                            <div className="flex items-center gap-2 shrink-0 ml-3 z-10 relative">
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  const emp = employees.find((emp) =>
                                    matchNames(emp.name, userData.name),
                                  );
                                  setEmployeeEditModal({
                                    isOpen: true,
                                    item: emp || {
                                      name: userData.name,
                                      position: userData.position,
                                      province: userData.province || "",
                                      email: userData.email || "",
                                    },
                                  });
                                }}
                                className="size-7 bg-white/80 hover:bg-white text-primary rounded-full border border-[#edecff] flex items-center justify-center transition-all shadow-sm cursor-pointer"
                                title="Edit Karyawan"
                              >
                                <span className="material-symbols-outlined text-[14px]">
                                  edit
                                </span>
                              </button>
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  const emp = employees.find((emp) =>
                                    matchNames(emp.name, userData.name),
                                  );
                                  setEmployeeDeleteModal({
                                    isOpen: true,
                                    item: emp || { name: userData.name },
                                  });
                                }}
                                className="size-7 bg-red-50 hover:bg-red-100 text-red-500 rounded-full border border-red-100 flex items-center justify-center transition-all shadow-sm cursor-pointer"
                                title="Hapus Karyawan"
                                disabled={matchNames(
                                  userData.name,
                                  "Aditya Wiratama",
                                )}
                              >
                                <span className="material-symbols-outlined text-[14px]">
                                  delete
                                </span>
                              </button>
                              {rootSubordinates.length > 0 && (
                                <span className="material-symbols-outlined text-[#8E94B7] text-md leading-none select-none ml-1">
                                  {collapsedNodes[userData.name] !== false
                                    ? "expand_more"
                                    : "expand_less"}
                                </span>
                              )}
                            </div>
                          </div>
                        </div>

                        {/* Recursive Tree Render */}
                        {collapsedNodes[userData.name] === false &&
                          renderRecursiveTeamNode(userData.name, 1)}
                      </div>
                    ) : (
                      <p className="text-[9px] font-semibold text-[#8E94B7] text-center py-6 bg-slate-50/50 rounded-2xl border border-dashed border-[#edecff] uppercase tracking-wide">
                        Tidak ada anggota tim di bawah Anda. Silakan klik tombol
                        "Tambah Anggota" di atas untuk menambahkan.
                      </p>
                    )}
                  </div>
                </div>
              );
            })()}

          {partnerSubTab === "channel" && (
            <div className="animate-in fade-in slide-in-from-right-4 duration-300">
              {teamMembers.length > 1 && (
                <div className="mb-6">
                  <label className="text-[10px] text-[#8E94B7] font-bold uppercase tracking-wide ml-1 mb-2 block">
                    Pilih PIC / Tim
                  </label>
                  <div className="relative">
                    <select
                      value={mappingPic}
                      onChange={(e) => setMappingPic(e.target.value)}
                      className="w-full h-14 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-6 font-semibold text-xs text-[#111] outline-none focus:ring-1 focus:ring-primary/20 transition-all appearance-none cursor-pointer pr-12"
                    >
                      <option value="ALL_TEAM">Semua PIC</option>
                      {[...teamMembers]
                        .sort((a, b) =>
                          compareMembersByLevel(
                            a,
                            b,
                            teamLevels,
                            teamPositions,
                            userData,
                          ),
                        )
                        .map((picName, idx) => (
                          <option key={idx} value={picName}>
                            {picName}
                          </option>
                        ))}
                    </select>
                    <span className="material-symbols-outlined absolute right-6 top-1/2 -translate-y-1/2 text-[#8E94B7] pointer-events-none">
                      expand_more
                    </span>
                  </div>
                </div>
              )}

              <div className="mb-6 flex gap-2 overflow-x-auto pb-2 -mx-1 px-1 custom-scrollbar">
                {mappingCategories.map((cat) => (
                  <button
                    key={cat}
                    onClick={() => setMappingCategory(cat)}
                    className={`whitespace-nowrap px-4 py-2 rounded-full font-semibold text-[10.5px] uppercase tracking-wide transition-all ${mappingCategory === cat ? "bg-gradient-to-r from-primary to-cyan-400 text-white shadow-[0_4px_12px_rgba(21,75,226,0.2)]" : "bg-[#f4f2ff] text-[#8E94B7] hover:bg-[#edecff]"}`}
                  >
                    {cat}
                  </button>
                ))}
              </div>

              <div className="mb-4 flex items-center justify-between bg-white px-5 py-3.5 rounded-[20px] shadow-sm border border-[#edecff]">
                <div>
                  <h3 className="font-bold text-xs text-[#181a2c]">
                    Daftar Partner
                  </h3>
                </div>
                <button
                  onClick={() =>
                    setPartnerEditModal({
                      isOpen: true,
                      item: {
                        isAdd: true,
                        category: mappingCategory || "Kios",
                        name: "",
                        pic: "",
                      },
                    })
                  }
                  className="h-8 px-4 bg-gradient-to-r from-teal-500 to-emerald-500 text-white rounded-full font-bold text-[9px] uppercase tracking-wider flex items-center gap-1.5 hover:opacity-95 transition-all shadow-md active:scale-[97%] cursor-pointer"
                >
                  <span className="material-symbols-outlined text-[13px]">
                    add
                  </span>
                  Tambah Partner
                </button>
              </div>

              <div className="space-y-3">
                {displayedPartnerChannels.length === 0 ? (
                  <div className="py-12 bg-white rounded-[24px] shadow-sm flex flex-col justify-center items-center">
                    <span className="material-symbols-outlined text-[#8E94B7] text-3xl mb-2">
                      sentiment_dissatisfied
                    </span>
                    <p className="text-[11px] font-semibold text-[#8E94B7] uppercase tracking-wider">
                      Tidak Ada Data
                    </p>
                  </div>
                ) : (
                  displayedPartnerChannels.map((channel, i) => (
                    <div
                      key={i}
                      className="bg-white shadow-[0_4px_16px_rgba(21,75,226,0.06)] hover:shadow-[0_8px_24px_rgba(21,75,226,0.12)] p-4 rounded-[16px] flex items-center justify-between transition-all duration-250"
                    >
                      <div className="flex-1 pr-4">
                        <p className="font-semibold text-xs text-[#181a2c] leading-tight mb-0.5">
                          {channel.name}
                        </p>
                        <p className="text-[10px] font-bold text-[#8E94B7] flex items-center gap-1">
                          <span className="material-symbols-outlined text-[13px] text-[#8E94B7]/80">person</span>
                          {normalizeName(channel.pic) || "-"}
                        </p>
                        <div className="flex flex-wrap gap-1.5 mt-2">
                          {channel.category && (
                            <span className="text-[9px] font-extrabold bg-[#154be2]/10 text-primary px-2.5 py-0.5 rounded-full uppercase border border-[#154be2]/5">
                              {channel.category}
                            </span>
                          )}
                          {channel.group && (
                            <span className="text-[9px] font-extrabold bg-slate-100 text-slate-700 px-2.5 py-0.5 rounded-full uppercase">
                              {channel.group}
                            </span>
                          )}
                          {(channel.province || channel.area) && (
                            <span className="text-[9px] font-extrabold bg-emerald-50 text-emerald-700 px-2.5 py-0.5 rounded-full uppercase border border-emerald-100/30">
                              {channel.province || channel.area}
                            </span>
                          )}
                        </div>
                      </div>
                      <div className="flex gap-2">
                        <button
                          onClick={() =>
                            setPartnerEditModal({ isOpen: true, item: channel })
                          }
                          className="size-9 bg-[#edecff] text-primary rounded-full border border-[#c4c5d8] flex items-center justify-center hover:bg-[#e6e6ff] transition-all shadow-none cursor-pointer"
                        >
                          <span className="material-symbols-outlined text-[16px]">
                            edit
                          </span>
                        </button>
                        <button
                          onClick={() =>
                            setPartnerDeleteModal({
                              isOpen: true,
                              item: channel,
                            })
                          }
                          className="size-9 bg-red-50 text-red-500 rounded-full border border-red-100 flex items-center justify-center hover:bg-red-100 transition-all shadow-none cursor-pointer"
                          title="Hapus Partner"
                        >
                          <span className="material-symbols-outlined text-[16px]">
                            delete
                          </span>
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          )}

          <PartnerEditModal
            isOpen={partnerEditModal.isOpen}
            item={partnerEditModal.item}
            onClose={() => setPartnerEditModal({ isOpen: false, item: null })}
            onSave={handleEditPartnerSave}
            isSaving={isActionLoading}
            activeEmployees={myActiveSubordinates}
            allProvinces={availableProvinces}
            allCategories={allCategories}
            userData={userData}
            allGroups={availableGroups}
          />
          <PartnerDeleteModal
            isOpen={partnerDeleteModal.isOpen}
            onClose={() => setPartnerDeleteModal({ isOpen: false, item: null })}
            onConfirm={handleDeletePartnerConfirm}
            isProcessing={isActionLoading}
            itemName={partnerDeleteModal.item?.name}
          />

          <EmployeeEditModal
            isOpen={employeeEditModal.isOpen}
            item={employeeEditModal.item}
            onClose={() => setEmployeeEditModal({ isOpen: false, item: null })}
            onSave={handleEditEmployeeSave}
            isSaving={isActionLoading}
            allEmployeeNames={employees.map((emp) => emp.name)}
            userData={userData}
            allProvinces={availableProvinces}
            accessRules={accessRules}
          />
          <EmployeeDeleteModal
            isOpen={employeeDeleteModal.isOpen}
            onClose={() =>
              setEmployeeDeleteModal({ isOpen: false, item: null })
            }
            onConfirm={handleDeleteEmployeeConfirm}
            isProcessing={isActionLoading}
            itemName={employeeDeleteModal.item?.name}
          />


        </div>
      )}

      {activeTab === "summary" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300">
          <div className="mb-4 ml-1 flex items-center gap-3">
            <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
              Stock <span className="text-primary font-bold">Summary</span>
            </h1>
            {filteredSummaryData.length > 0 && (
              <button
                onClick={handleDownloadSummaryExcel}
                className="h-8.5 w-16 bg-gradient-to-r from-emerald-500 to-teal-600 text-white rounded-full flex items-center justify-center hover:opacity-95 transition-all shadow-md active:scale-[97%] cursor-pointer shrink-0"
                title="Download Excel"
              >
                <span className="material-symbols-outlined text-[18px]">
                  download
                </span>
              </button>
            )}
          </div>

          <div className="flex flex-col gap-3.5 mb-6">
            <div className="flex flex-col md:flex-row-reverse gap-4 md:gap-6 md:items-stretch">
              {filteredSummaryData.length > 0 && renderPogKpiCard(true)}

              <div
                className={`w-full md:w-64 lg:w-72 xl:w-80 shrink-0 transition-all ${isSummaryFilterOpen ? "block" : "hidden md:block"}`}
              >
                <div className="bg-white p-4 md:px-6 md:py-5 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.15)] animate-in fade-in slide-in-from-top-3 duration-300 h-full flex flex-col justify-center">
                  <div className="flex flex-row md:flex-col gap-2 md:gap-3.5">
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Group By
                      </label>
                      <div className="relative">
                        <select
                          value={summaryGroupBy}
                          onChange={(e) => setSummaryGroupBy(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none"
                        >
                          <option value="hybrid">Hybrid</option>
                          <option value="subordinate">Team</option>
                          <option value="area">Area</option>
                          <option value="category">Category</option>
                          <option value="crops">Crops</option>
                          <option value="area_crops">Area & Crops</option>
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Sub
                      </label>
                      <div className="relative">
                        <select
                          value={summarySubGroupBy}
                          onChange={(e) => setSummarySubGroupBy(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none"
                        >
                          <option value="channel">Channel</option>
                          <option value="hybrid">Hybrid</option>
                          <option value="subordinate">Team</option>
                          <option value="area">Area</option>
                          <option value="category">Category</option>
                          <option value="crops">Crops</option>
                          <option value="area_crops">Area & Crops</option>
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Crop
                      </label>
                      <div className="relative">
                        <select
                          value={filterBelowCrop}
                          onChange={(e) => setFilterBelowCrop(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none"
                        >
                          {availableCrops.map((crop) => (
                            <option key={crop} value={crop}>
                              {crop}
                            </option>
                          ))}
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {filteredSummaryData.length > 0 && (
              <div className="-mx-5 md:mx-0 md:bg-transparent md:border-none md:shadow-none md:divide-y-0 bg-white overflow-hidden md:overflow-visible rounded-[48px] md:rounded-none pt-4 md:pt-0 pb-4 md:pb-0 mb-8 shadow-[0_4px_44px_rgba(24,26,44,0.15)] border border-[#edecff] divide-y divide-[#edecff] transition-all md:grid md:grid-cols-2 md:gap-4 xl:grid-cols-2">
                {summaryGroupBy === "subordinate"
                  ? filteredSummaryData.map((row) =>
                      renderRecursiveSummaryRow(row),
                    )
                  : filteredSummaryData.map((row, i) => (
                      <div
                        key={i}
                        className="p-0 overflow-hidden transition-all md:bg-white md:rounded-[32px] md:shadow-[0_4px_24px_rgba(24,26,44,0.08)] md:border md:border-[#edecff]"
                      >
                        <div
                          className={`flex justify-between items-center ${row.isExpandable ? "cursor-pointer" : ""} px-5 py-4 pb-2 hover:bg-slate-50/60 transition-colors`}
                          onClick={() =>
                            row.isExpandable && toggleRow(row.name)
                          }
                        >
                          <div className="flex flex-col">
                            <span className="font-semibold text-xs md:text-sm text-[#181a2c] uppercase flex items-center gap-1.5">
                              {row.isExpandable && (
                                <span className="material-symbols-outlined text-primary text-[20px]">
                                  {expandedRows[row.name]
                                    ? "keyboard_arrow_down"
                                    : "keyboard_arrow_right"}
                                </span>
                              )}
                              {row.name}
                            </span>
                          </div>
                          <div className="flex flex-col items-end">
                            <span className="font-bold text-sm text-primary">
                              {formatOverviewVal(row.selectedTotal, overviewUseMt).valueStr}
                            </span>
                            <span className="text-[8px] text-[#8E94B7] uppercase tracking-widest font-bold">
                              Total {formatOverviewVal(row.selectedTotal, overviewUseMt).unit}
                            </span>
                          </div>
                        </div>

                        <div className="mx-5 mb-3 mt-2 flex divide-x divide-white/20 bg-primary shadow-[0_12px_32px_rgba(21,75,226,0.35)] rounded-[14px] overflow-hidden">
                          {selectedClusters.map((clusterKey) => {
                            if (
                              clusterKey === "Uncategorized" &&
                              (!row[clusterKey] || row[clusterKey] === 0)
                            )
                              return null;
                            const clusterConfig = CLUSTER_CONFIG.find(
                              (c) => c.key === clusterKey,
                            );
                            return (
                              <div
                                key={clusterKey}
                                className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors"
                              >
                                <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                                  {clusterConfig?.label || clusterKey}
                                </span>
                                <span className="font-semibold text-[10.5px] truncate w-full text-white">
                                  {formatOverviewVal(row[clusterKey], overviewUseMt).valueStr}
                                </span>
                              </div>
                            );
                          })}
                        </div>

                        {expandedRows[row.name] && row.children?.length > 0 && (
                          <div className="pb-4 pt-3 border-t border-[#edecff] flex flex-col gap-3.5 px-5 bg-slate-50/50">
                            {summarySubGroupBy === "subordinate"
                              ? row.children.map((child: any) =>
                                  renderRecursiveSubordinate(child),
                                )
                              : row.children.map((child, j) => {
                                  const isChildZeroTeam = isRowVisitZero(child.name);
                                  return (
                                    <div
                                      key={`${i}-${j}`}
                                      className={`flex flex-col p-3.5 rounded-[18px] transition-all duration-200 ${
                                        isChildZeroTeam
                                          ? "bg-red-50/70 border border-red-200/60 shadow-[0_10px_28px_rgba(239,68,68,0.12)]"
                                          : "bg-[#fbfaff] shadow-[0_10px_28px_rgba(21,75,226,0.18)]"
                                      }`}
                                    >
                                      <div className="flex justify-between items-center mb-2 px-1">
                                        <div className="flex flex-col">
                                          <span className="font-bold text-[11px] text-[#181a2c] uppercase pr-2 flex items-center gap-1.5 flex-wrap">
                                            <span>
                                              {renderMaybeChannelName(
                                                child.name,
                                              )}
                                            </span>
                                            {isChildZeroTeam && (
                                              <span className="text-[7.5px] bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full font-bold uppercase tracking-wider">
                                                Belum Dikunjungi
                                              </span>
                                            )}
                                            {summarySubGroupBy === "channel" &&
                                              child.category && (
                                                <span className="text-[8px] bg-primary/10 text-primary px-1.5 py-0.5 rounded-full font-bold">
                                                  {child.category}
                                                </span>
                                              )}
                                          </span>
                                          {summarySubGroupBy === "channel" && child.pic && (
                                            <div className="text-[10px] text-[#8E94B7] mt-1 truncate">
                                              {child.pic}
                                            </div>
                                          )}
                                        </div>
                                        <span
                                          className={`font-bold text-[11.5px] shrink-0 ${isChildZeroTeam ? "text-red-600" : "text-[#181a2c]"}`}
                                        >
                                          {formatOverviewVal(child.selectedTotal, overviewUseMt).valueStr}{" "}
                                          <span className="text-[8.5px] text-[#8E94B7]">
                                            {formatOverviewVal(child.selectedTotal, overviewUseMt).unit}
                                          </span>
                                        </span>
                                      </div>
                                      <div
                                        className={`flex w-full divide-x rounded-[14px] overflow-hidden ${
                                          isChildZeroTeam
                                            ? "divide-red-200 bg-red-100/40"
                                            : "divide-primary/10 bg-primary/8"
                                        }`}
                                      >
                                        {selectedClusters.map((clusterKey) => {
                                          if (
                                            clusterKey === "Uncategorized" &&
                                            (!child[clusterKey] ||
                                              child[clusterKey] === 0)
                                          )
                                            return null;
                                          const clusterConfig =
                                            CLUSTER_CONFIG.find(
                                              (c) => c.key === clusterKey,
                                            );
                                          return (
                                            <div
                                              key={clusterKey}
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-200/30"
                                                  : "hover:bg-primary/5"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase tracking-wider mb-0.5 truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-700/60"
                                                    : "text-primary/70"
                                                }`}
                                              >
                                                {clusterConfig?.label ||
                                                  clusterKey}
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-700"
                                                    : "text-primary"
                                                }`}
                                              >
                                                {formatOverviewVal(child[clusterKey], overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                          );
                                        })}
                                      </div>
                                    </div>
                                  );
                                })}
                          </div>
                        )}
                      </div>
                    ))}
              </div>
            )}

            {filteredSummaryData.length === 0 && (
              <div className="py-16 text-center flex flex-col items-center bg-white rounded-[24px] border border-[#edecff] shadow-sm">
                <span className="material-symbols-outlined text-[40px] text-[#8E94B7] mb-4">
                  inventory_2
                </span>
                <p className="font-semibold text-[#8E94B7] text-xs">
                  No data available for the selected filters.
                </p>
              </div>
            )}
          </div>
        </div>
      )}

      {activeTab === "pog" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300">
          <div className="mb-4 ml-1">
            <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
              POG <span className="text-primary font-bold">Analytics</span>
            </h1>
          </div>

          <div className="flex flex-col gap-3.5 mb-6">
            <div className="flex flex-col md:flex-row-reverse gap-4 md:gap-6 md:items-stretch">
              {renderPogKpiCard(false)}

              <div
                className={`w-full md:w-64 lg:w-72 xl:w-80 shrink-0 transition-all ${isPogFilterOpen ? "block" : "hidden md:block"}`}
              >
                <div className="bg-white p-4 md:px-6 md:py-5 rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.15)] animate-in fade-in slide-in-from-top-3 duration-300 h-full flex flex-col justify-center">
                  <div className="flex flex-row md:flex-col gap-2 md:gap-3.5">
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Group By
                      </label>
                      <div className="relative">
                        <select
                          value={pogGroupBy}
                          onChange={(e) => setPogGroupBy(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none pr-7"
                        >
                          <option value="hybrid">Hybrid</option>
                          <option value="subordinate">Team</option>
                          <option value="area">Area</option>
                          <option value="category">Category</option>
                          <option value="crops">Crops</option>
                          <option value="area_crops">Area & Crops</option>
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Sub
                      </label>
                      <div className="relative">
                        <select
                          value={pogSubGroupBy}
                          onChange={(e) => setPogSubGroupBy(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none pr-7"
                        >
                          <option value="channel">Channel</option>
                          <option value="hybrid">Hybrid</option>
                          <option value="subordinate">Team</option>
                          <option value="area">Area</option>
                          <option value="category">Category</option>
                          <option value="crops">Crops</option>
                          <option value="area_crops">Area & Crops</option>
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                    <div className="flex-1 min-w-0">
                      <label className="text-[9px] text-[#8E94B7] font-bold uppercase tracking-wider block mb-1.5 truncate ml-1">
                        Crop
                      </label>
                      <div className="relative">
                        <select
                          value={filterBelowCrop}
                          onChange={(e) => setFilterBelowCrop(e.target.value)}
                          className="w-full h-10 bg-white shadow-[0_4px_16px_rgba(21,75,226,0.08)] rounded-full px-4 font-semibold text-[10.5px] text-primary outline-none truncate appearance-none pr-7"
                        >
                          {availableCrops.map((crop) => (
                            <option key={crop} value={crop}>
                              {crop}
                            </option>
                          ))}
                        </select>
                        <span className="material-symbols-outlined absolute right-3 top-1/2 -translate-y-1/2 text-primary text-base pointer-events-none">
                          expand_more
                        </span>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {aggregatedPogData.length > 0 && (
              <div className="-mx-5 md:mx-0 md:bg-transparent md:border-none md:shadow-none md:divide-y-0 bg-white overflow-hidden md:overflow-visible rounded-[48px] md:rounded-none pt-4 md:pt-0 pb-4 md:pb-0 mb-8 shadow-[0_4px_44px_rgba(24,26,44,0.15)] border border-[#edecff] divide-y divide-[#edecff] transition-all md:grid md:grid-cols-2 md:gap-4 xl:grid-cols-2">
                {pogGroupBy === "subordinate"
                  ? aggregatedPogData.map((row) => renderRecursivePogRow(row))
                  : aggregatedPogData.map((row, i) => (
                      <div
                        key={i}
                        className="p-0 overflow-hidden transition-all md:bg-white md:rounded-[32px] md:shadow-[0_4px_24px_rgba(24,26,44,0.08)] md:border md:border-[#edecff]"
                      >
                        <div
                          className={`flex justify-between items-center ${row.isExpandable ? "cursor-pointer" : ""} px-5 py-4 pb-2 hover:bg-slate-50/60 transition-colors`}
                          onClick={() =>
                            row.isExpandable && togglePogRow(row.name)
                          }
                        >
                          <div className="flex flex-col">
                            <span className="font-semibold text-xs md:text-sm text-[#181a2c] uppercase flex items-center gap-1.5">
                              {row.isExpandable && (
                                <span className="material-symbols-outlined text-primary text-[20px]">
                                  {pogExpandedRows[row.name]
                                    ? "keyboard_arrow_down"
                                    : "keyboard_arrow_right"}
                                </span>
                              )}
                              {row.name}
                            </span>
                          </div>
                          <div className="flex flex-col items-end">
                            <span className="font-bold text-sm text-primary">
                              {formatOverviewVal(row.pog, overviewUseMt).valueStr}
                            </span>
                            <span className="text-[8px] text-[#8E94B7] uppercase tracking-widest font-bold">
                              POG ({formatOverviewVal(row.pog, overviewUseMt).unit})
                            </span>
                          </div>
                        </div>

                        <div className="mx-5 mb-3 mt-2 flex flex-row gap-1.5 md:gap-2">
                          {/* Table 1: Opening Inv & End of Inv */}
                          <div className="flex-[2] flex divide-x divide-white/20 bg-primary/95 shadow-[0_12px_32px_rgba(21,75,226,0.25)] rounded-[14px] overflow-hidden">
                            <div className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors">
                              <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                                Opening Inv
                              </span>
                              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                                {formatOverviewVal(row.lastQty, overviewUseMt).valueStr}
                              </span>
                            </div>
                            <div className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors">
                              <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                                End of Inv
                              </span>
                              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                                {formatOverviewVal(row.currentQty, overviewUseMt).valueStr}
                              </span>
                            </div>
                          </div>

                          {/* Table 2: Stock in, idle stock, POG */}
                          <div className="flex-[3] flex divide-x divide-white/20 bg-primary shadow-[0_12px_32px_rgba(21,75,226,0.35)] rounded-[14px] overflow-hidden">
                            <div className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors">
                              <span className="text-[8px] font-bold uppercase tracking-wider text-white/85 mb-0.5 truncate w-full">
                                Stock in
                              </span>
                              <span className="font-semibold text-[10.5px] truncate w-full text-white">
                                {formatOverviewVal(row.sellIn, overviewUseMt).valueStr}
                              </span>
                            </div>
                            <div className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors">
                              <span className="text-[8px] font-bold uppercase tracking-wider text-amber-200 mb-0.5 truncate w-full">
                                idle stock
                              </span>
                              <span className="font-semibold text-[10.5px] truncate w-full text-amber-100">
                                {formatOverviewVal(row.idleStock, overviewUseMt).valueStr}
                              </span>
                            </div>
                            <div className="flex-1 min-w-0 p-2 flex flex-col items-center justify-center text-center hover:bg-white/10 transition-colors">
                              <span className="text-[8px] font-bold uppercase tracking-wider text-cyan-200 mb-0.5 truncate w-full">
                                POG
                              </span>
                              <span className="font-semibold text-[10.5px] truncate w-full text-cyan-100 font-extrabold">
                                {formatOverviewVal(row.pog, overviewUseMt).valueStr}
                              </span>
                            </div>
                          </div>
                        </div>

                        {pogExpandedRows[row.name] &&
                          row.children?.length > 0 && (
                            <div className="pb-4 pt-3 border-t border-[#edecff] flex flex-col gap-3.5 px-5 bg-slate-50/50">
                              {pogSubGroupBy === "subordinate"
                                ? row.children.map((child: any) =>
                                    renderRecursivePogSubordinate(child),
                                  )
                               : row.children.map((child, j) => {
                                    const isChildZeroTeam = isRowVisitZero(child.name);
                                    return (
                                      <div
                                        key={`${i}-${j}`}
                                        className={`flex flex-col p-3.5 rounded-[18px] transition-all duration-200 ${
                                          isChildZeroTeam
                                            ? "bg-red-50/70 border border-red-200/60 shadow-[0_10px_28px_rgba(239,68,68,0.12)]"
                                            : "bg-[#fbfaff] shadow-[0_10px_28px_rgba(21,75,226,0.18)]"
                                        }`}
                                      >
                                        <div className="flex justify-between items-center mb-2 px-1 flex-wrap gap-2">
                                          <div className="flex flex-col">
                                            <span className="font-bold text-[11px] text-[#181a2c] uppercase pr-2 flex items-center gap-1.5 flex-wrap">
                                              <span>
                                                {renderMaybeChannelName(
                                                  child.name,
                                                )}
                                              </span>
                                              {isChildZeroTeam && (
                                                <span className="text-[7.5px] bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full font-bold uppercase tracking-wider">
                                                  Belum Dikunjungi
                                                </span>
                                              )}
                                              {pogSubGroupBy === "channel" &&
                                                child.category && (
                                                  <span className="text-[8px] bg-primary/10 text-primary px-1.5 py-0.5 rounded-full font-bold">
                                                    {child.category}
                                                  </span>
                                                )}
                                            </span>
                                            {pogSubGroupBy === "channel" && child.pic && (
                                              <div className="text-[10px] text-[#8E94B7] mt-1 truncate">
                                                {child.pic}
                                              </div>
                                            )}
                                          </div>
                                          <span
                                            className={`font-bold text-[11.5px] shrink-0 ${isChildZeroTeam ? "text-red-600" : "text-primary"}`}
                                          >
                                            {formatOverviewVal(child.pog, overviewUseMt).valueStr}{" "}
                                            <span className="text-[8.5px] text-[#8E94B7]">
                                              POG ({formatOverviewVal(child.pog, overviewUseMt).unit})
                                            </span>
                                          </span>
                                        </div>
                                        <div className="flex flex-row w-full gap-1.5 md:gap-2">
                                          {/* Table 1: Opening Inv & End of Inv */}
                                          <div
                                            className={`flex-[2] flex divide-x rounded-[14px] overflow-hidden ${
                                              isChildZeroTeam
                                                ? "divide-red-200 bg-red-100/40"
                                                : "divide-primary/10 bg-primary/5"
                                            }`}
                                          >
                                            <div
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-200/30"
                                                  : "hover:bg-primary/5"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                                                  isChildZeroTeam
                                                    ? "text-red-700/60"
                                                    : "text-[#8E94B7]"
                                                }`}
                                              >
                                                Opening Inv
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-700"
                                                    : "text-[#181a2c]"
                                                }`}
                                              >
                                                {formatOverviewVal(child.lastQty, overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                            <div
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-200/30"
                                                  : "hover:bg-primary/5"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                                                  isChildZeroTeam
                                                    ? "text-red-700/60"
                                                    : "text-[#1d4ed8]/75"
                                                }`}
                                              >
                                                End of Inv
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-700"
                                                    : "text-[#1d4ed8]"
                                                }`}
                                              >
                                                {formatOverviewVal(child.currentQty, overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                          </div>

                                          {/* Table 2: Stock in, idle stock, POG */}
                                          <div
                                            className={`flex-[3] flex divide-x rounded-[14px] overflow-hidden ${
                                              isChildZeroTeam
                                                ? "divide-red-200 bg-red-200/45"
                                                : "divide-primary/10 bg-primary/10 border border-primary/10"
                                            }`}
                                          >
                                            <div
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-300/30"
                                                  : "hover:bg-primary/15"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                                                  isChildZeroTeam
                                                    ? "text-red-800/70"
                                                    : "text-[#154be2]/80"
                                                }`}
                                              >
                                                Stock in
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-800"
                                                    : "text-[#154be2]"
                                                }`}
                                              >
                                                {formatOverviewVal(child.sellIn, overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                            <div
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-300/30"
                                                  : "hover:bg-amber-100/60"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                                                  isChildZeroTeam
                                                    ? "text-red-800/70"
                                                    : "text-amber-800"
                                                }`}
                                              >
                                                idle stock
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-800"
                                                    : "text-amber-700"
                                                }`}
                                              >
                                                {formatOverviewVal(child.idleStock, overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                            <div
                                              className={`flex-1 min-w-0 p-1.5 flex flex-col items-center justify-center text-center transition-colors ${
                                                isChildZeroTeam
                                                  ? "hover:bg-red-300/30"
                                                  : "hover:bg-emerald-100/60"
                                              }`}
                                            >
                                              <span
                                                className={`text-[7.5px] font-bold uppercase truncate w-full tracking-wider mb-0.5 ${
                                                  isChildZeroTeam
                                                    ? "text-red-800/70"
                                                    : "text-emerald-800"
                                                }`}
                                              >
                                                POG
                                              </span>
                                              <span
                                                className={`font-black text-[10px] truncate w-full ${
                                                  isChildZeroTeam
                                                    ? "text-red-800"
                                                    : "text-emerald-700 font-extrabold"
                                                }`}
                                              >
                                                {formatOverviewVal(child.pog, overviewUseMt).valueStr}
                                              </span>
                                            </div>
                                          </div>
                                        </div>
                                      </div>
                                    );
                                  })}
                            </div>
                          )}
                      </div>
                    ))}
              </div>
            )}

            {aggregatedPogData.length === 0 && (
              <div className="py-16 text-center flex flex-col items-center bg-white rounded-[24px] border border-[#edecff] shadow-sm">
                <span className="material-symbols-outlined text-[40px] text-[#8E94B7] mb-4">
                  analytics
                </span>
                <p className="font-semibold text-[#8E94B7] text-xs">
                  No data available for the selected filters.
                </p>
              </div>
            )}
          </div>
        </div>
      )}

      {activeTab === "access" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300">
          <div className="mb-6 ml-1 flex flex-col gap-1">
            <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">
              Access Control Menu
            </h1>
            <p className="text-[#8E94B7] text-[11px] font-semibold tracking-wide">
              Level permissions and access rights mapping
            </p>
          </div>
          
          <div className="bg-white rounded-[24px] shadow-[0_12px_32px_rgba(21,75,226,0.03)] border border-[#154be2]/5 overflow-hidden">
            <div className="p-6 border-b border-[#f1f5f9]">
              <div className="flex items-center gap-2">
                <div className="size-8 rounded-lg bg-primary/10 flex items-center justify-center">
                  <span className="material-symbols-outlined text-primary text-[20px]">
                    admin_panel_settings
                  </span>
                </div>
                <div>
                  <h3 className="text-sm font-bold text-[#181a2c]">System Access Rights</h3>
                  <p className="text-[10px] font-semibold text-[#8E94B7] mt-0.5">Matrix of features available for each organizational level</p>
                </div>
              </div>
            </div>
            
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse min-w-[700px]">
                <thead>
                  <tr className="bg-[#fbfaff]">
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Position</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Home Tab</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Data Partner</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Stock Summary</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">POG Tracking</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Overview Tab</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Temp Tab</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9]">Access Menu</th>
                    <th className="px-5 py-3 text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider border-b border-[#f1f5f9] text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#f1f5f9]">
                  {allPositionsList.map((position) => (
                    <tr key={position} className="hover:bg-slate-50/50 transition-colors">
                      <td className="px-5 py-4">
                        <span className="text-[11px] font-semibold text-[#181a2c] bg-slate-100 px-2 py-1 rounded-md">{position}</span>
                      </td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'home')}</td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'partner')}</td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'stock')}</td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'pog')}</td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'overview')}</td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'temp')}</td>
                      <td className="px-5 py-4">{renderAccessCheckbox(position, 'access')}</td>
                      <td className="px-5 py-4 text-right">
                        <button onClick={() => removeAccessRule(position)} className="text-red-500 hover:bg-red-50 p-1.5 rounded-lg transition-colors" title="Hapus Rule">
                          <span className="material-symbols-outlined text-[16px]">delete</span>
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            
            <div className="p-5 border-t border-[#f1f5f9] bg-[#fbfaff] flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div className="flex items-start gap-3 flex-1">
                <span className="material-symbols-outlined text-[#8E94B7] text-[18px]">info</span>
                <p className="text-[10px] font-semibold text-[#8E94B7] leading-relaxed">
                  <span className="text-[#181a2c] font-bold uppercase tracking-wider block mb-1">General Access Rules</span>
                  All levels have access to the <strong className="text-primary">Home</strong>, <strong className="text-primary">Data Partner</strong>, <strong className="text-primary">Stock Summary</strong>, and <strong className="text-primary">POG Tracking</strong> tabs. The data visible within these tabs is automatically filtered based on the user's <strong className="text-primary">Data Visibility</strong> level. The <strong className="text-primary">Access Menu</strong> tab is strictly limited to authorized system administrators.
                </p>
              </div>
              
              <div className="flex items-center justify-end gap-3 w-full md:w-auto mt-2 md:mt-0">
                {accessSaveSuccess && (
                  <div className="flex items-center gap-1.5 text-emerald-600 bg-emerald-50 px-3 py-1.5 rounded-lg text-xs font-bold animate-in fade-in slide-in-from-bottom-2">
                    <span className="material-symbols-outlined text-[16px]">check_circle</span>
                    Saved
                  </div>
                )}
                <button
                  onClick={handleSaveAccessRules}
                  disabled={isSavingAccess}
                  className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-[#154be2] hover:bg-[#154be2]/90 text-white px-6 py-2.5 rounded-xl text-xs font-bold transition-all shadow-sm disabled:opacity-70 disabled:cursor-not-allowed cursor-pointer"
                >
                  {isSavingAccess ? (
                    <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                  ) : (
                    <span className="material-symbols-outlined text-[18px]">save</span>
                  )}
                  {isSavingAccess ? "Saving..." : "Save Role Access"}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {activeTab === "temp" && (
        <div className="animate-in fade-in slide-in-from-right-4 duration-300 animate-in">
          {consolidationSuccessMsg && (
            <div className="mb-4 bg-emerald-50 border border-emerald-200 text-emerald-800 rounded-2xl px-5 py-3 text-xs font-bold flex items-center gap-2 shadow-sm animate-in fade-in slide-in-from-top-2 duration-300">
              <span className="material-symbols-outlined text-emerald-600">
                check_circle
              </span>
              {consolidationSuccessMsg}
            </div>
          )}

          <div className="mb-6 flex flex-col lg:flex-row lg:items-center justify-between gap-4 ml-1">
            <div>
              <h1 className="text-xl font-bold text-[#181a2c] tracking-tight">
                Review{" "}
                <span className="text-primary font-bold">Data (Temporary)</span>
              </h1>
              <p className="text-xs text-[#8E94B7] mt-1 font-semibold">
                Menampilkan data checker, channel, hybrid, dan lot no secara
                komprehensif.
              </p>
            </div>

            <div className="flex flex-col sm:flex-row items-center gap-3 w-full lg:w-auto shrink-0">
              {/* Preview Toggle Button */}
              <button
                type="button"
                id="btn-temp-proceed-consolidate"
                onClick={() => {
                  setIsTempProceeded(!isTempProceeded);
                  setExpandedTempRowId(null); // Close any expanded row during toggle
                }}
                className={`w-full sm:w-auto h-10 px-5 rounded-full font-extrabold text-xs flex items-center justify-center gap-2 transition-all cursor-pointer active:scale-[0.98] ${
                  isTempProceeded
                    ? "bg-amber-500 text-white shadow-[0_4px_14px_rgba(245,158,11,0.3)] hover:bg-amber-600"
                    : "bg-[#181a2c] text-white shadow-[0_4px_14px_rgba(24,26,44,0.15)] hover:bg-[#252841]"
                }`}
              >
                <span className="material-symbols-outlined text-[18px]">
                  {isTempProceeded ? "visibility_off" : "visibility"}
                </span>
                {isTempProceeded ? "Matikan Preview" : "Preview Konsolidasi"}
              </button>

              {/* Real Database Process Button */}
              <button
                type="button"
                id="btn-temp-process-db-consolidate"
                disabled={isConsolidatingDb}
                onClick={handleConsolidateDatabase}
                className={`w-full sm:w-auto h-10 px-5 rounded-full font-extrabold text-xs flex items-center justify-center gap-2 transition-all cursor-pointer active:scale-[0.98] bg-emerald-600 text-white shadow-[0_4px_14px_rgba(16,185,129,0.3)] hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed`}
              >
                <span
                  className={`material-symbols-outlined text-[18px] ${isConsolidatingDb ? "animate-spin" : ""}`}
                >
                  {isConsolidatingDb ? "sync" : "auto_mode"}
                </span>
                {isConsolidatingDb ? "Memproses..." : "Proses Konsolidasi"}
              </button>

              {/* Quick Search Container */}
              <div className="relative w-full sm:w-64 shrink-0">
                <span className="material-symbols-outlined absolute left-4 top-1/2 -translate-y-1/2 text-gray-400 text-[18px]">
                  search
                </span>
                <input
                  type="text"
                  placeholder="Cari data..."
                  value={tempSearchQuery}
                  onChange={(e) => setTempSearchQuery(e.target.value)}
                  className="w-full h-10 bg-white border border-[#edecff] shadow-sm rounded-full pl-11 pr-10 font-bold text-xs text-[#181a2c] outline-none focus:ring-1 focus:ring-primary/20 transition-all placeholder:text-gray-400"
                />
                {tempSearchQuery && (
                  <button
                    type="button"
                    onClick={() => setTempSearchQuery("")}
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 p-1 flex items-center justify-center cursor-pointer transition-colors"
                  >
                    <span className="material-symbols-outlined text-[16px]">
                      close
                    </span>
                  </button>
                )}
              </div>
            </div>
          </div>

          {/* Petunjuk Konsolidasi & Sinkronisasi */}
          <div className="mb-6 bg-amber-50/70 border border-amber-200/80 rounded-2xl p-5 shadow-sm ml-1 text-[#181a2c]">
            <div className="flex gap-3">
              <span className="material-symbols-outlined text-[22px] text-amber-600 mt-0.5">
                info
              </span>
              <div className="flex-1">
                <h3 className="text-xs font-black text-amber-900 uppercase tracking-wider mb-1">
                  Panduan Konsolidasi Database (Sheet Working)
                </h3>
                <p className="text-xs text-amber-800 leading-relaxed font-semibold">
                  Gunakan tab ini untuk menggabungkan data ganda dengan
                  Kombinasi{" "}
                  <strong className="font-extrabold text-amber-950">
                    Checker + Channel + Hybrid + Lot No
                  </strong>{" "}
                  yang sama menjadi satu baris tunggal, di mana nilai stok
                  dipisahkan secara otomatis ke dalam kolom bulan yang tepat
                  (April - Maret).
                </p>
                <div className="mt-3 flex flex-col sm:flex-row gap-4 text-[11px] font-bold text-amber-900/90 border-t border-amber-200/40 pt-3">
                  <div className="flex items-center gap-1.5">
                    <span className="material-symbols-outlined text-[16px] text-amber-700">
                      visibility
                    </span>
                    <span>
                      <strong>Preview Konsolidasi:</strong> Simulasi visual di
                      layar (client-side) sebelum disimpan.
                    </span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <span className="material-symbols-outlined text-[16px] text-emerald-700">
                      auto_mode
                    </span>
                    <span>
                      <strong>Proses Konsolidasi:</strong>{" "}
                      Menyimpan/menggabungkan data secara permanen di Google
                      Sheet.
                    </span>
                  </div>
                </div>
                <div className="mt-3.5 bg-emerald-50/75 border border-emerald-200/50 rounded-xl p-3 text-[11px] font-semibold text-emerald-900 leading-relaxed shadow-inner flex items-center gap-2">
                  <span className="material-symbols-outlined text-[16px] text-emerald-600">
                    cloud_done
                  </span>
                  <span>
                    <strong>Koneksi Langsung Aktif:</strong> Sistem sekarang
                    terhubung langsung ke database Google Sheets melalui API
                    resmi. Anda tidak perlu lagi menyalin file Apps Script
                    manual atau melakukan langkah konfigurasi manual lainnya.
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* Stats Cards Row */}
          {(() => {
            const INDO_MONTHS = [
              "Jan",
              "Feb",
              "Mar",
              "Apr",
              "Mei",
              "Jun",
              "Jul",
              "Ags",
              "Sep",
              "Okt",
              "Nov",
              "Des",
            ];

            const getMonthIndexFromTimestamp = (timestamp: any): number => {
              if (!timestamp) return new Date().getMonth();
              if (timestamp instanceof Date) return timestamp.getMonth();

              const str = String(timestamp).trim();

              if (str.includes("/")) {
                const parts = str.split(/[\s/:]+/);
                if (parts.length >= 2) {
                  const mVal = parseInt(parts[1], 10);
                  if (!isNaN(mVal) && mVal >= 1 && mVal <= 12) {
                    return mVal - 1;
                  }
                  const months = [
                    "jan",
                    "feb",
                    "mar",
                    "apr",
                    "may",
                    "jun",
                    "jul",
                    "aug",
                    "sep",
                    "oct",
                    "nov",
                    "dec",
                  ];
                  const indMonths = [
                    "jan",
                    "feb",
                    "mar",
                    "apr",
                    "mei",
                    "jun",
                    "jul",
                    "agu",
                    "ags",
                    "sep",
                    "okt",
                    "nov",
                    "des",
                  ];
                  const lowerM = parts[1].toLowerCase();
                  let m = months.findIndex((name) => lowerM.startsWith(name));
                  if (m === -1)
                    m = indMonths.findIndex((name) => lowerM.startsWith(name));
                  if (m !== -1) return m;
                }
              }

              if (str.includes("-")) {
                const parts = str.split(/[\s\-:]+/);
                if (parts.length >= 2) {
                  const mVal = parseInt(parts[1], 10);
                  if (!isNaN(mVal) && mVal >= 1 && mVal <= 12) {
                    return mVal - 1;
                  }
                  const months = [
                    "jan",
                    "feb",
                    "mar",
                    "apr",
                    "may",
                    "jun",
                    "jul",
                    "aug",
                    "sep",
                    "oct",
                    "nov",
                    "dec",
                  ];
                  const indMonths = [
                    "jan",
                    "feb",
                    "mar",
                    "apr",
                    "mei",
                    "jun",
                    "jul",
                    "agu",
                    "ags",
                    "sep",
                    "okt",
                    "nov",
                    "des",
                  ];
                  const lowerM = parts[1].toLowerCase();
                  let m = months.findIndex((name) => lowerM.startsWith(name));
                  if (m === -1)
                    m = indMonths.findIndex((name) => lowerM.startsWith(name));
                  if (m !== -1) return m;
                }
              }

              const d = new Date(str);
              return !isNaN(d.getTime()) ? d.getMonth() : new Date().getMonth();
            };

            const groupedMap: Record<
              string,
              {
                id: string;
                checker: string;
                channel: string;
                hybrid: string;
                lot: string;
                shippingDate: string;
                expDate: string;
                inputs: Array<{
                  tanggalInput: string;
                  qty: number;
                  id: string;
                }>;
                monthlyQty: number[];
                totalQty: number;
              }
            > = {};

            const monthsKeys = [
              "jan",
              "feb",
              "mar",
              "apr",
              "mei",
              "jun",
              "jul",
              "ags",
              "sep",
              "okt",
              "nov",
              "des",
            ];

            (rawWorkingData && rawWorkingData.length > 0
              ? rawWorkingData
              : workingData
            ).forEach((item, index) => {
              const checker = item.user || item.pic || "Unknown";
              const channel =
                item.kiosk || item.channel || item.toko || "Unknown";
              const hybrid = item.hybrid || item.hybrids || "Unknown";
              const lot = item.lot || "Unknown";
              const key = `${checker}_${channel}_${hybrid}_${lot}`;

              let shippingDate =
                item.drDate ||
                item.shipping_date ||
                item.shippingDate ||
                item.dr_date ||
                "N/A";
              let expDate =
                item.expired ||
                item.exp_date ||
                item.expDate ||
                item.expired_date ||
                "N/A";

              const isValidVal = (v: any) =>
                v && v !== "N/A" && v !== "-" && String(v).trim() !== "";

              if (!isValidVal(shippingDate) || !isValidVal(expDate)) {
                const matchedDr = drSalesData.find(
                  (dr) =>
                    dr &&
                    dr.lot &&
                    cleanForMatch(dr.lot) === cleanForMatch(lot),
                );
                if (matchedDr) {
                  if (!isValidVal(shippingDate) && matchedDr.drDate) {
                    shippingDate = matchedDr.drDate;
                  }
                }
              }
              if (!isValidVal(expDate)) {
                expDate = extractExpFromLot(lot);
              }

              // Determine monthly quantites for this item
              const itemMonthlyQty = Array(12).fill(0);
              let hasDbMonths = false;
              monthsKeys.forEach((mName, mIdx) => {
                if (
                  item[mName] !== undefined &&
                  item[mName] !== null &&
                  String(item[mName]).trim() !== ""
                ) {
                  itemMonthlyQty[mIdx] = Number(item[mName]) || 0;
                  hasDbMonths = true;
                }
              });

              const qtyVal = Number(item.stock) || Number(item.qty) || 0;
              if (!hasDbMonths) {
                const mIdx = getMonthIndexFromTimestamp(item.timestamp);
                itemMonthlyQty[mIdx] = qtyVal;
              }

              const itemTotal = itemMonthlyQty.reduce(
                (acc, val) => acc + val,
                0,
              );
              const tanggalInput = item.timestamp || "N/A";
              const inputId = item.id || `input_${index}`;

              if (!groupedMap[key]) {
                groupedMap[key] = {
                  id: `group_${index}`,
                  checker,
                  channel,
                  hybrid,
                  lot,
                  shippingDate,
                  expDate,
                  inputs: [],
                  monthlyQty: Array(12).fill(0),
                  totalQty: 0,
                };
              } else {
                const g = groupedMap[key];
                const isValidVal = (v: any) =>
                  v && v !== "N/A" && v !== "-" && String(v).trim() !== "";
                if (!isValidVal(g.shippingDate) && isValidVal(shippingDate)) {
                  g.shippingDate = shippingDate;
                }
                if (!isValidVal(g.expDate) && isValidVal(expDate)) {
                  g.expDate = expDate;
                }
              }

              // Save inputs details
              groupedMap[key].inputs.push({
                id: inputId,
                tanggalInput,
                qty: qtyVal,
              });

              // Sum monthly values
              for (let m = 0; m < 12; m++) {
                groupedMap[key].monthlyQty[m] += itemMonthlyQty[m];
              }
              groupedMap[key].totalQty += itemTotal;
            });

            const list = Object.values(groupedMap);

            const filteredList = list.filter((item) => {
              const q = tempSearchQuery.trim().toLowerCase();
              if (!q) return true;
              return (
                String(item.checker).toLowerCase().includes(q) ||
                String(item.channel).toLowerCase().includes(q) ||
                String(item.hybrid).toLowerCase().includes(q) ||
                String(item.lot).toLowerCase().includes(q)
              );
            });

            const uniqueCheckers = new Set(filteredList.map((i) => i.checker))
              .size;
            const uniqueChannels = new Set(filteredList.map((i) => i.channel))
              .size;
            const uniqueHybrids = new Set(filteredList.map((i) => i.hybrid))
              .size;
            const uniqueLots = new Set(filteredList.map((i) => i.lot)).size;

            // Handle Sorting
            const sortedList = [...filteredList].sort((a, b) => {
              const valA = String(a[tempSortBy] || "").toLowerCase();
              const valB = String(b[tempSortBy] || "").toLowerCase();
              if (valA < valB) return tempSortOrder === "asc" ? -1 : 1;
              if (valA > valB) return tempSortOrder === "asc" ? 1 : -1;
              return 0;
            });

            const toggleSort = (col: string) => {
              if (tempSortBy === col) {
                setTempSortOrder((prev) => (prev === "asc" ? "desc" : "asc"));
              } else {
                setTempSortBy(col);
                setTempSortOrder("asc");
              }
            };

            return (
              <>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
                  <div className="bg-white p-4 rounded-[20px] border border-[#edecff] shadow-sm flex items-center gap-3">
                    <div className="size-10 rounded-xl bg-indigo-50 flex items-center justify-center text-indigo-500 shrink-0">
                      <span className="material-symbols-outlined text-md">
                        assignment_ind
                      </span>
                    </div>
                    <div className="min-w-0">
                      <p className="text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider">
                        Unik Checker
                      </p>
                      <p className="text-sm font-bold text-[#181a2c] mt-0.5">
                        {uniqueCheckers}
                      </p>
                    </div>
                  </div>

                  <div className="bg-white p-4 rounded-[20px] border border-[#edecff] shadow-sm flex items-center gap-3">
                    <div className="size-10 rounded-xl bg-sky-50 flex items-center justify-center text-sky-500 shrink-0">
                      <span className="material-symbols-outlined text-md">
                        store
                      </span>
                    </div>
                    <div className="min-w-0">
                      <p className="text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider">
                        Unik Channel
                      </p>
                      <p className="text-sm font-bold text-[#181a2c] mt-0.5">
                        {uniqueChannels}
                      </p>
                    </div>
                  </div>

                  <div className="bg-white p-4 rounded-[20px] border border-[#edecff] shadow-sm flex items-center gap-3">
                    <div className="size-10 rounded-xl bg-emerald-50 flex items-center justify-center text-emerald-500 shrink-0">
                      <span className="material-symbols-outlined text-md">
                        category
                      </span>
                    </div>
                    <div className="min-w-0">
                      <p className="text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider">
                        Unik Hybrid
                      </p>
                      <p className="text-sm font-bold text-[#181a2c] mt-0.5">
                        {uniqueHybrids}
                      </p>
                    </div>
                  </div>

                  <div className="bg-white p-4 rounded-[20px] border border-[#edecff] shadow-sm flex items-center gap-3">
                    <div className="size-10 rounded-xl bg-purple-50 flex items-center justify-center text-purple-500 shrink-0">
                      <span className="material-symbols-outlined text-md">
                        layers
                      </span>
                    </div>
                    <div className="min-w-0">
                      <p className="text-[10px] font-bold text-[#8E94B7] uppercase tracking-wider">
                        Unik Lot No
                      </p>
                      <p className="text-sm font-bold text-[#181a2c] mt-0.5">
                        {uniqueLots}
                      </p>
                    </div>
                  </div>
                </div>

                {/* Table Container */}
                <div className="bg-white rounded-[24px] border border-[#edecff] shadow-[0_8px_24px_rgba(24,26,44,0.02)] overflow-hidden">
                  {isTempProceeded && (
                    <div className="bg-amber-50 px-6 py-2.5 border-b border-amber-100 flex items-center gap-2 text-amber-800 text-[11px] font-bold">
                      <span className="material-symbols-outlined text-[16px] animate-pulse text-amber-600">
                        info
                      </span>
                      Geser tabel ke samping untuk melihat seluruh kolom bulan ‚Üî
                    </div>
                  )}
                  <div className="overflow-x-auto min-w-full">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="border-b border-[#edecff] bg-[#fafbfe]/80">
                          {isTempProceeded ? (
                            <>
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider min-w-[130px]">
                                <button
                                  type="button"
                                  onClick={() => toggleSort("checker")}
                                  className="flex items-center gap-1.5 hover:text-primary transition-colors cursor-pointer text-left focus:outline-none"
                                >
                                  Checker
                                  <span className="material-symbols-outlined text-[14px]">
                                    {tempSortBy === "checker"
                                      ? tempSortOrder === "asc"
                                        ? "arrow_upward"
                                        : "arrow_downward"
                                      : "unfold_more"}
                                  </span>
                                </button>
                              </th>
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider min-w-[150px]">
                                <button
                                  type="button"
                                  onClick={() => toggleSort("channel")}
                                  className="flex items-center gap-1.5 hover:text-primary transition-colors cursor-pointer text-left focus:outline-none"
                                >
                                  Channel Partner
                                  <span className="material-symbols-outlined text-[14px]">
                                    {tempSortBy === "channel"
                                      ? tempSortOrder === "asc"
                                        ? "arrow_upward"
                                        : "arrow_downward"
                                      : "unfold_more"}
                                  </span>
                                </button>
                              </th>
                              <th className="py-4 px-6 text-[10px] font-extrabold text-[#8E94B7] uppercase tracking-wider min-w-[130px]">
                                xúÏ=is€∆íﬂÛ+&L*EæêIQ≤¨Hˆ*>^ºÒı$øº≠R©dê ID ¿@IWˇ}ªÁ fÉã¶m%]e‡‹}O˜Ùçñq¯ﬂê“OºZ8«VºQ°|‡?Û‹Ò’Ò∫Ÿ"«OHLßûsÑq≥1[çB◊n¥Ó*43ˆ¨(zkÕ°ÔâÁ‹7vÊQgÏ¯±í©µËÙª{d\;·aÏ‹∆ùEËŒ≠pE‚–Ú#7vø3º å»xFAÿY.≠KK{Œ$&ì`ºåÉeÏπæ”Òﬂ)üﬂì
CˇÖNìÏêÁN4ÆP˛(ZXæ<·πu-Ø≠Ê£¿ã:|å6¸y∏∏ΩhT
!kX∑.ˇœ+r||L*U&‰)ıﬂÖ6¨m¬ä∆UÎc+ÉõÀÂ‚∆
+wL»°®h7~ù™PqÈOœæú°”®ÇmG;ÉÚ=⁄aÑPVÚh'ûïñâg2–´Œê,n;˚»= 2‡®w‡9¥ ¸ø;xÒx¯Û£≤\,úplE¢˝¯ ıßù¡4w˝Œ¥A©Ä)G_ÉxA¸óÁØÉòºÓ¿•‡ ÚÁkpÄı´∑œﬂ]æy˜ˆ√/g›πµh6Ám‚⁄∑îRöÂCâg&vÂ¨é◊–hMBÏhwÏàñ„Ù Y”˛¢ÔVhl=/ü[p¥Jô—ÁdŸ¥\ËNgq¬Ω{π˜á ∂<ÚØxı…8y¥SÙ{´-èæ¢–ªØbn<s∆W0ºø∫®{∆ÊyÔdùXˇy'˛bÔ˝%~À˜Ôo@¸tû‰Ω∆˛ΩdL@˛<0Å¸œ√æó¸yÿ˜z` †Û∞”ı∞”ı@ÛÊÖª
˘ÉÜ¶√º™ÿ≠cŸπ?«£¿^…hfª◊¿#:+¬øú7L˙ì…C≥€à1ù»ôª2œÈÙ≠¡∏ˇ÷@h±cøv£∏Î9˛¯€“d(⁄+ëÍ–ù?§t 7÷ÖK9¸(&nÙ‚`hâá˝ »{‹º≤)‚b{]◊˛©∞µ–âó°Õa›˜a0vl≥x‰luÀU|∂Ò»ÜQéó†MB#¸wëÂoeÙQΩc€,Ü"«s∆q¡_âÈ>ï3ÓAwØ"”€ã‹?úŒ#KÑr«õX◊∑›i–ŸÎ14Âè˚ΩCb	Åπ<ÕÍ˜eªìïxLEj4]ˇ™”É/∞£N4'ñÔ"ìÜn»ƒ≤Èﬂ?Ç`;è˜àΩ-
ö˛^Ø‚ÅhŒbËfJÒæÀ∑hZ›h9äÿ˚^õ*âL‰∞b◊Vó=q∏Ù«052∑nÈûÎ^™~‘„⁄Ù‰yUùF5æ\y¬¿!sŸ£T*è¯Ã˜ú˝˛#IP[©∫Ru©E,5«+Ì€ j”≤é@j\«⁄Ÿ#™ ≈©´ãéÊ)Å~˚2ù¶X‰˙T€y¡¯äx ∂PeE›ãå™k∞?I_˝™ÑƒñìŸ6€˛_w›È“Õ?Î‹7X–∞U∂Ï ®»C,Q]Sˆ∫¨ëÉÜ5…^?Ç◊äë≠@Õ¸ÚÀœ∫ÜUägﬁÍ_ÒäiˇØ⁄d˛™≤ëÇ±“T©D«ñ´1¥|o"ì9Ut‹ä0Ä9W–π‰œîŒÈ|Ädü†ìÜï
≠GcÀc»«¥Ò,KÕûé±…8*ã∂›8x`\∂VîüuDO'†‘nF‹2ô2Óèáˆà”≥ÑsÀ´1’ŒÁòU5Í≠BïÂn‰2°Àº¿¶-ÅíúøZ^µucå"F7ÒøLhB~ùnÖ%ŸkÏS{éNkw_Ü÷tåB5$™•*Óm∑&r‚∫’îÃ¨ßƒ_zûè£ÊnŒ˙£¡®·‘ÓÏ%˚`™oﬁphS}·˚b˚P|î°ÀºÌº◊Ì.4N”¿}˛äeaÖ∞ÚUÙ›«Úï®§lKïÿæA∆˘Xe˛Tckk@Öä∂ª)!˝:ôäã›^O2—RÃ®æŸ≈+P≤"±O2éC≥B‡)±”|‹≥ùiã¢Kªr€wUeTu˘ V◊u¯óîç~Yq/MÙ–}÷ΩvÍ®€3¡iså5˙ãe•˚`}Ú√ïı∫∂∫Œ`ç¢¿[Ç·Ÿâ∂Œ$˝ör6∂ñ73_÷\’Ygä±à*¸<E“›:‰Í•+üåm1‘XÁú≈©æ≤u—≥¢:√ÆºwÚwO6⁄?©∑Jï'^MÈ˚‚˚(µwR∂0ìDÒüm?•˛éJΩçªØÅ?œŒJÕΩïÌ¢‹úaÉL5¢ä≤Lç(¿*∞¿*p.vÜUQï"I‡ùA◊«Î·ùÜ0ªa™k∏öF≥ËPŸ ƒ© ¡¡≠ßa¿˘wéÌå—áftDûKΩùIÃ©¯ﬁMŸAØ:Q¬ g√ÏFÇ‡E¶rN¨@ì˘àjπéˆZ"ææW{ø∫W[|–h˜m+ºå€*ãˆVFWS ‰π[ÆG>X˛tjy‰j.‰á
AÊrß≥a k8≈X"¸=~°h~ò>(êÜ√a=8U2w∏^äÛÇ.Ä,†JMXmàÕÜyK'± Ëµ ÙUa‘“&X|PãÒ„·‘e4s†»∫£¨çÕ5]^EÉ3"6X‘Eñs=˛¨AN˙ÁåØ3y‡¨Ωÿã-LXDWå<òäfêyå6ŸÛb˙Å¿"ú˘øˇ#ç∑;'ï"aî1◊ùfmÑ™_·O≈™¬ JY}¯ã∞*ÁFz9ZFu.·õ™˜5˙^q(ä≈h«lŒ†æ• π]Ê§<”C‰VÊ_:’˜t”œSBA5F∆KPª_qSòQ≠⁄;Tˆßg´uãkƒÖ^úâÜŒÃµm«ó˘jÎfT=55∂Fû#˜|√ˆ*”@^ﬁ®√ûµàÍ5O£∞¢'ÛkÈÜ´&y:£ÏpVío™}fí3∞Ìﬂ!πÈÙJ ¬ì∑Aï»ı:ù4û(¶”∂õó≈ç'gqpEv–6#Õ_ß≠M˙™∂’°’®èA_0*◊¸Qˆ·Y‡≠øh¯ØN¿å>©ê˚®˝EåV‰Â¥=eO&q äŸŸÎ‚ZÎ#âiªpWB…±(√7n6û¨Ÿ»…è§Wu∑≤ÚpÙ>ì=q9úÅÆbÃháíŒgá∆ÉXÃ≈ßYd?t*∆∏ú*Ü	m¥ıIªZã÷v]f0	Çx{"hg7ëBqV
Â9•üŸHƒH;∞}Ve’‹…õ
∏bAòà˘çPè%Ñ8πZŒó0$wîŸ—2ÙfêV√ö∏'aæ—Ù∑Óîô‹K≤ëP≠KPU≈œ§◊â≠ØÍ‹™¥*ÂπXv‘0±í∏≥¢É0w˘vQYº⁄QŸ\ ckí=Ü”›ÀVB•π˛æ"Á5æ¬D∞cªÀy±qY°Æm]À∂àm≈YÅ¿ÜôåA≥hj˘‰*tÈ~Y8˛ÿÇ/~∑x°J–ßmr—•D2e˝áà˘8∫ˆö%'¥ÿÈ, (e¿9&'ah≠ö˝A´;q=ØŸ+¿C©€Iæ∞∆≥äß¬Ù(§6ç˘§ ŒïÒ1üC˘Ú„1Å ≈4ìˇk—olâ¶†˜⁄LJÀ{9vöMk<nìdÚ
©¬‘€§h)˘π∂b¬-g¥En⁄Éœ™kTÒÌ~vΩÇ¡ÁW∞KÊ•º∏úıØcv@≤^≤3;MefZè›¨yìYÖ˛ﬁp‰.TìßÇ»Z√`˘˘¯ñQ(0Ï±S¡„Peô™$!3ö5ƒé◊P}◊)≈n¢OïœºL+(’årŸ¡]+Ô DÅÍî´˜‰ˆÕ¸»V¨D‚~—¨ô“M]7cko£1Rí¸–0√;Ú3(gs ∞Ê∆≤˙·ñ/≠¸ñ|?zLÉe¸,'n8ÿñ'Ut£w†{†^≈ä—ﬂÒï‹:û"G:J/ﬁú ”rZZ-÷ÁÒ:YÒÙÁdB“»)vﬁ˝ÑCß1˝kÏ®Ë˙†z,".ùi[Æ*ÁÚê∞(‹6YàØtÄP|ÓFŒ—öDÀÒÿâ¢C¿û¿s,ˇ'‚Ña>≈…›⁄)ˆ…§.mns-z"wá“($=ÅU8_FNËﬁ∂qa˛Õ.†	¯·åF~7úxç`˙†3≠Òû?÷†£~Miç¸°∞(z˙ o'cœJ=∑üæëÍŒÄãyùÚŸr4wc(oE+LT›NJ⁄=—:k∆·“I~ãÅ∑¶JÎ.tP	¥n,Ëà/y3]S±VÎr'§˘-‘ÍrË∂ΩM¢KG]ﬁ6Hxô4L"À≥f›Ü‘Å@Sê°V<ûIÕ+S˛‡Ñø[∂"ó6˙¸Ô"hÒikwd‚˙ GVj+Ú
	»ùﬂqàhZö…è9ÎD„–q|@ˆ≈ÿR‘9ßØdsá˝^œæû]≤íÁ¸1ç—.˜>/êO; ‹lﬁt‚†3Â£’Œø{9xÒËÂãrÌZÏ‘Œ›ñR>-Ò”L¯~ÌÄ˛ÊxÑªnX¿/∆Ísß—˘`Hc];"|ÓºwâGB.á¯_8YÕ¡∞=ÿoáÌ^∑∑ﬂ∫–ºL‘±¥≥ﬂ´zBOãπ3zˆ√º5IbuÉŒ,æŒ0˝Ñ…–x˚˘m«Z∆fÌI≥Í`BªÉdV˝ˆ£Ωˆ f÷ÎÓ∂*ú}êÇ˚≥ÒG'ˆµÂ«2b’evæ?dá ¯ó¥ôQ‰NF4Õ˙Y±ßã[’»N¥|!tcñŒu‘Èß"Y˚È…ÛìSrÚ¸∑ì∑NîÕ˙ xﬁyI!(∞ˆ◊Ia$‰~Fç'In=H˜Ù≈Ø'o≈8…s¸˛ˆ‰ı´≥¯Â˝ª” –Ö8∂ Ÿ◊Ye‚ˆ÷Bòî)Ω8§K‚MË˜û…Wt±-Hæ©¸Pâ	ÈcVËy1ç…ß≤ U›‡i<]éºí‹£ç'…:ﬁô+¥0MW´
1V“08œ€À‚kﬂx:§<õΩ«"=Ñ¸µX!UJ'çaî‘hhä%˛¢1ˆÜ_Ø-o	Í¢ËœdYÄvsäJÂµ–*ÖÄn^wc+ú:qó6d¥Ll7BªƒÊ™+•¶ÇŸÄYß?L"§Mäûä¯‘¸0 †∆À‰Da<wXÆùZMπ=⁄K∑ 2˚ﬁr{⁄LZKå≈˚Xc|ì59 ›x‰nÇŒ¬Tù¡|ú∏ÒÇõ„+ﬂÌv≥5ô©æˆ
Eykÿ+4¿E™b∞„◊∆‡ÕëW ©ÚöÏúåá…	¯åMîT
œÑö¿∆ú£ÕÑÅj4~-jÊóBê…À(†û–4 ı¸\˙*X≠1ûåÂπ7Œ{›«˙˚Rî© ∆”˜ü∏$∆mvhçû”õXô]>z?†˚x¡(ÛèÇ“&eµæ@´óÙm∏"0˜qF´7„˝Ó^ÀpΩC6[@ñ≥…,#'°KŒ5G—uﬁF§|ûò´q—∏X)7ù=Ÿ0»i˝7«ç¢ìˆ◊ÆsÛsp{‹Ëë#j04óÃ€Ñ;ª·ÿÀè~ï√94{˘^º1å§?(¯}U¸;†æëM±Oá¡PsE<√•≤“ˇqÌxv‹»Y£4fü£ÖUpŸâaa,Édµq€«ç7C“X‰ †⁄Ît~Î=€ÎÓ>⁄≈g"æı≥·Û<˜OuÔ˚kaøÙ∆=≤€Ì§ﬂÌÔÓA—¿í],∂{‡ÌÇy≤?|ÙGÕµˇzöG®˝sZ"Ø|ê!˘!È¶ùyŒo÷0PoÉõºF3öÅ9£h÷ÙV_(èÊ-U∂Ÿˆ8v0ˇüì[7˙ Ç7:∏≠yñ€*≥∑π&∑m≤¬M∏ïX6πÉÚ¥∏º’H5¸Åb›pk»≈Pø¿ç>‘ΩÇâ®ÖÃûm≈6Äè—ó›h·πq≥A∏∂Á‹É¬Ù®ãüåªb”4˘HH˙=ÕÔ◊∑wÌÔ◊´ª÷G))Œ22i-oè◊=+Ìãü¯„Y‚ÜõzÄååRæ:{{èNaL!ÇB˙Ã˝√AsÄ¡æZ¸˘?l„é<Íı≤?æ¥ÊÆ∑Ç⁄∏w“>t'JR™]÷tıôSÖm∫sÓ⁄∑áƒ_ŒGNhpÅ„V+î O»†%Vì¸®~$`–@9[˝˜Tè±]
›ÓT∞ﬁW–.ÄönΩ˛H®¥òÒQm4«?}S´>usRËõ=“.—y mˆ{ÜLHkiòY⁄ã≥T™¯ÃÓ$zÖ“Ä)—MsH¨ÄS'
ºk«~Ôéë‚hπ+7àÆ(Ÿ1¯-‹±ÙÑ˙íÙË±ZÙµíKT∆Ë{x›J|	zÿÇÄì>∂Ñ™ı-SµÃøRÊç>›Qî]˙W~p„7“w%π>[Øªd h•Œî˜˘c2€dpjCÚ(3Âòª@¥ó‘π≈vb;kÈ≈d≤Ù«®@ìì≈¢Ÿ zÖ–WòxÖAˆºÿûh—2&˜H‰DˆqÃŒÉû≈AhMù.`’+òG≥∫™^bwóºhCÛï◊	ˇ˜Ÿª∑›Ö¬“âüRoÛx4ùñ6ê¿sò+•ŸxiπÓ∆Ñ
äa¢Î6—=$À7ÓÚ˝M1=ÿìåYK˙)”ìYó≤VQ¡ZµŸr0“[&iN°{•π–ôÉ∫/∑ËÅ˛‡ÿó`E7îj<÷	=èÔ~ÙÎÀ |É´,∫Í.fà1°ô◊*oπMPøV^u7:±ÁÆèXà°DºäÁ\;ûDh {ç¬„Ã¬∆r7≤‘s‚jJ~∞F*ÔNó-#
 ∆,o≈heK≥)≥>Acµn‹®hnU+ù)%≠É∆˘¡îD≠]∑Y;Êä»Wñ&¸“ı@Õ˚Ÿ#ÛhGÕ∆âÁ)˛D‚x`ÄÁ‚g°Ú»π±%4l‹Ìv˙PSÿ∞Y’û få"≤ÆùrF!qJç3›ôÃôÖ¬ tn“Dæì„ÄŒÖÇµXúPﬂÚÈ“s¢∆¶¸—2$jcË‹£Cœ[©p3Ì§Ò3ßr¬IÊÙs‚!ÊK¢ÆrzØîxE|ï¸L≈WÅË‚/tﬂ-.ÅO‰.U,ø9 pp{+"g˛ˇwÍÑ_tœÇ˘‹	«ÆÂë◊ée·æó>"œ/_∫„ì–±æ‚ö≥ûO¶a‡cPÕó≈∫Âœ0ª≤ç/÷˚ùëÛú[Ç∂n™(vLó¯¬∫ù,\X„T√CS¨@ﬂKK£Ò•hF≈`Ìîi?ü@E©≥–6óhdï®p4•*ì™ÛZ∫ #øﬁX„°6N2‹›Ü≠˝õå´•€<ä
â<å»LÈPh}Ωi'Nyb∞Zzx(´4«ÖáÂ˙’YA•w£ﬂùq‹<jJÕaæogw¬öPÈ¿∞E¶ñï´¢◊Œ—˝™T U/À*#vh2kâ∞ëW∞PZæsÈßãß]¡Øh,>Ãp8ÅVLì”˚S77‘'Q1—yï=éö ﬁç˙≤“c;†A¬ä8ƒıÅFAXÉ‚pXd –l√s7z„¯Àﬂ‹»ÂÉG2JoîXF)‘0©Mo‹Öu=cúè∑†Ω5µ"7#÷Èç\aÃtd⁄“;√jX&ÍîxÙJœú§zˆ<ï“÷îo˙RmI“ƒ≥çªEıf¯ÎΩ·—7ôÒ˜5ö˙‡XsΩ|W£	T~Ù&]ùï3&≥,ŒÿÑä;gÆÌå¨Pd.‰∏£Ω-∆@hƒÒ&'îΩ∞Ç‘º°·[*??QAõÀñH’J⁄lÿÓé˝^HN:ö7`Ùhä
Ø'õFçÑG“∂Ê)∞Åt(B4™ùnÜ]¥tm*emÕÚQ4™SgÑˆëÿ≥∆ù≥*Í’¶¢T£Ÿ-3URn≤a∆fÕGá;í  ãÂÉ5¯`˛E≠¡€¯ÚﬁõÑFË‡=Í;Tëà™r’îíiÛ9PeOöÊ≥ø‚4¯˜t¸õ3°èb)_Iå-I6i∂’OJ—‹Ωl®‘2}õ»ùLç¨HJ‡<ß»¯êÑét}«Maq‰ÈÎb
Io4R?2t>S˜È≤dXΩ4(Õ\êzTÃ]Ú§Ç/Ø∆OÏìÙÜá¥â.~'Oü2≤ê $Ë…äÒGSIé∏¨}0∂áÕ€
¶¶)ö≥bâÖd(À(Äï√Ô¶2Ç0X).í3Â2£QÎg‚8YÕ©m–‡⁄2|5’jfEø¿?X#YÎ|Jæ˝VÂW¬ú?¸§6ûM´º~s3g∏Âç0‹»I0≠0
@sua•ï∑ë`ïπ!LÉPﬁE9sL·Ã4—Ã¥¡ín•D+mñ∂ª≥C^¯ì ;B©aD(;ÖäÑc≠dv7ò>ÄÔ]Ú´ﬂ—G.Ûmä:-yDØ¿Åû‘I±•∞Z¥ú≥€îx5Å≈}”¥äÖ≈ Iy≥¬
…>à®$aLaE∂ÍIµº-ñ-%ÂÃ(  rﬁ»¢AW©XE(…+.Rˆu®ëóVÜ'◊»lË»µd…ï$MuR»5çΩp¯(= ¸Leÿ»Ö)ÙL•”’U sh∫÷IJxíÔ+í@QùäwT◊èe2∫zz1ôÄ*`PÚíˆ≤Dád'ëìÈY*l©€èMï∫µ MÓõ≠&†hÏèq\C_ JS¡≥5(<MUd∂ú≠ñ†≤©j G≥9>¥D¬"∆:µ{ƒsIN≠ñé‘\U¶;µf8s≈î¯‘j	‡r˙„®ıÖÄ3WH»P≠¡ gÆ"—¢ZI¨æF7ä*#9—R`¥Uï£-´ÌT	hKr∏≠	˜∂"bç(èZì∏Td#v0^‚>mwvÏºËÆm≥Å?ß≥¡ßÓ,t&î=«Ò":‹Ÿ¡∏Ã®;Ç)®ˆ7ÍéÉ˘Œ8äO'4ZÛXÏù˛x∆N4˛¯éüh<—ÌõÈ,nø|ı˙u˚üß'œˇk0l{Ω6¸k(›ÜéGy/ÜïF3«âìüì±c≤—ÆµX8æ˝lÊzv+J˚p%C%–%≤c˛™ﬂ,O Ø`EëÙ¢|ÁM:”?¶!«/¥Ç$:˛¯˝˙ÏŸÈ´˜.ˇ}˙˙Ó©E#Èéß,‰}L\œ˘r¸˝⁄Ò«ÅÌ¸˚Ù’≥`æ|bÎÓc“¶Ñ™<"πØÊ¯=
¸¶∂—áÖ@˘¥‚e$ÿb*‘Èè∂f±'f3$Ù∞&›n7)õZ“PF‚hﬁGVøõd"`æLÓyî€¿q&.K
	µ,ji∫Ø(1Zít|wÄÓOíF26ûAùÅÏ; Låy~ïÓ≤[‘®Íq\ÖUK™à5ïº’t\“J≈⁄4ﬁi€µì®˛>Y*:$Ëñ: ≤ª»
‹≥ak∆VtOÄ^Dç†˛FzPY]ãtˇ‚4áá‰2i7ˇ£Ê~ïö}Æ0ﬁQÜ3àÖL™‰≤à¥ÑÕ;Ú.ë˙kKJáîuËj°_ZU›øõ];¸ò√5WG&ÊUÎ*”óŸÀ!/n~‘ü1˙ÕU S‰»{ Aûz:ù	"åÍo±ó4Fûé·03;3ëÑ±ïYﬁ*o·∞°g»®å!SQ2«à©£¶ÆâiJICÿóWòhFÈIË\¸ıL‚èÒË'ÎT[ãRôQî˝Fﬂ7ª+—ƒö¶ÙBYﬂF–TAÿTVÆóDN©´&Ïêb)rXÖ …í:	4F®ò≤	cHÉñUHBéV”§2©°5B˜5≤NãÃ—„-<¯fl/<2ºq#<k¸Iyâä≤‰ÎÛ…π8ÿú(9ç?Ãøbl7Ù„äUõgëEj<˘:rıH∏rå7π#ÿıa:Ω“˛ı
•¿≥©OÍﬁ?ÑŸè.‘Tåbã•áâÅêQòSñùêÃNOŒ	§u¢∂<€Õ&∫H2'hπ.¢9Me‘xÚ∆ôèÄCìˇ·ÕCOñŒvµ∂ã‘<§“ıBÂIä2 ÄQ¿z_Eäv…{–>¶Kﬂ“VnQıdirjä˛·l@ﬁ«˛∆º‘S9K'À©7ÕY∆“s‚˘ ñ ¿H—DÂBs˚ê~Éõ$Ÿœ√%»[œâêìπhΩÛ¬ÉÄ»s'∫äÉ˘«Œù¬e/%o‡7‚±§É”r&dBèÄ Úozx”Ÿ“[‰ÒÖvxóaáÇÚúÆ;ÁΩnÔ±`äÚÀ!Âç,ù»#‡8ÄUv,:≈0Ω∑ß#›‹#j˜$ÖlÏéØ0G¬ò≈ıÓ}A	fÿÌÙ>ö¨jk(e‚êrPjã‘¸6≥n⁄ÒÍ¬ã‚4…+-ﬁÄÍ∆¸2Õ¶~√“± ¸=0\Íì¨MOO°ë\p¬ñzh∫›Dq˘–Ôß7›ÁJ√$ßØ!'›mÖª›µú/˚¶dLfë 9YÁË–˜ÛSÕ”¥Pi¢u ó∆äΩ)·‰¶Âö”ÚïÈ‡Õí!ÌÈ·æz“}!ÚRÿ√x|Í¢Ke4ÌJwo`€∑R3‰≤B}=ãú‰É»&)ÕÖìP`fÎº8sÕf$∑ã©nvetÕc*É^ÜﬁåπhL>C1ë`∆ê&KQ{j>-∑tJ?H:É^^äG¥ü:w¸ì≤`˚<-Oø€2∆‡'7sç¬é$ÄèåÈ≈ˆ¯{•ûN6˙Un“h”KÃ3C≠í‹¶Ü÷8˜#∆¸aç'ÿ≈ó~;U3rû°%øç„¯~]Å‹}4§ ¸6hJ"Æ·BËººÉ¶ÏËˆ◊¸e["€‘uıß$›º3ªÏÛ@æ.Ú•Á}@ÖåÓ˘æ∏u∆Kj≤˚d∆É» cV∞£ÈÉˆﬁh™~æº¬ÖÈNU•å<◊úÈÓ∞-†–ÜFöfkLæ‹‰“64X@-¯9$∫Yüµ8Ω™˚Q[°Ωı«jƒkùãHÁ»â&°‡ä¥íK“ô4Ü˜ê¢Ô≤D®FÈlÖìë{KÇbÑHÅº5
L,k ∫¨`AØL)ûø:“MBN"['=9Œm+ÑóY›[¬#| <
‰≠^¢\ÉÏ¿_∆óÊ|æè§G©Éú±âm_Ï%¡¢€y©xo©G˜@q∏€u¡¥µ≈°„„=óÀ≈}§∂˜Ô˛I>=Á≠õhΩjcQæ˜ñ‹ËËçÇwkÙ∆»‘1Ï¢»ù˙c|ÈÌ‘°g∂öHŸkä?ï‡î≥[!9#oâéÅÏ(à∑Fv‚¥X¬√‹BóÙä¥K@¢ñÍ^nÆ0"!òôÁS	P.RÊE|GW'ﬁ]52ﬁön¢Ía=F Õ^¯°\[…>üF@ât ¬PÚi5À˘µ^é„õ˝ûv¡'Ñc©;Û,_eøí˝ÍxK+$g´ñ;ÉL¶à∞,ÇôÉõ“gÉüîr˝Œ,ª£p ‡¬¢Cx∏ı¯√ﬁmZƒy!û€á“˝#jhSÅÑùxPÜ/`ı`àU·ÀÅ£¯‹äf£¿“.D—XÏ¸¶bo¬íéSÓ§ñóµ—@@ÿù›∏Òx& %Áà‘¢RàÏÒZz»î:IG&?È√OB…q…É°5©†˙¨M≈ê¬ÎxmzõÈ√î˝ãvf˙A≠≠Áˇ:^Îo2ΩÈŸ¡hO˙À‹^x&0•˛Æ®ß§öÒu˛ú8cQß≈_Œ,©h~ü€#&Sz√E=±
Ÿwπ=`Í•|Q‘´ê}ó£0X® ÇÖ–°≤Ô‰:;˘Ïﬂ#åuk]ªSÍãØ <âÛ®â{L| ≥ù£<∆∏æ!°∑l„≠n‚FW¡%ë	V|"/GN|ÉAÑL>ñFf"&•§Ä(PòÀØ™◊›‡w\–qÆA:Gt<T‚∞“è˜í¬ô¢˙•FwÅã¢FlUà”ÏXúfr≈[ü	£Ãƒg.–?¨ﬂ$k”cä∑zπnŒ}πyQÔ≈äìü)Åƒp?∏äPi¥y„Gw9˛Ïß◊‹=÷„R”kÏTIÊ±ÜóöBkUvˆíãª∞m˚X€≠ˆ‘” &•ö Ei5»E7Hb7ˆ`)Œ0v‹_πW†X·nM´∫Xj_=πrVTâ∏¥¬0∏πÑπ˚Uï)Â‘aQê"Ü\W∆¢@a+.f‚G+D6‰€‡ü›`∞ƒ+gôa<.¡¨ˆƒbŸôNh¥ÕÛc≤∂¯@∑≈’ ôFÆÈ\f(õLeÛ-muÃe~i(Ω[\≠Ãbt+ƒ#£–.-äI»7úMw”°>¬;õ‰Cπ◊+äÈP∏Œ„NØª◊x"P9◊`7Rô·ñ0ô«lL%Ò˜öHdÔ˘ﬂÖF2I$?j‡ÀìG„œA!E!”GAò¿Ω¶Ÿ…˝w°èåcø"}9˜ø<ÖP4˛,$◊çøπÙ0ªÚÔ5egÙﬂÖ*Á{Eä(r¿ôÒÓüüÉ
<ÌDû∑˝^SD‚.˛+êƒ„
$°˙«´˘>Ú/OàªüÉ&
ù·SEæC¸^”Ö‰œ˝ªPÜÓ¬ÆJU‹ÿ_ûJ6"ù‰eûêﬁ∏≠∫ùZzÿ∂éW:92N7Gœ˜ˆqï“™îín:‰üç€k]dˆKµ„~“È;mw<.:µ∑›≠—Û.D…ARZT8°πúÛw(`Ω5Xû:ÃÉ5; óñ“¸jﬂ™Ó(MTö0°Ã}B/ Mfë„ãKùpÊ≥•"≈éúBÇ≈ó«*™T9
ãô  ›!L‡	€cÕ°rﬁÎ>>∏–1ì∫PDÆfkŸÙo‰πÙKáN+ôΩ‚0,ÚºpøÀkæp=ì◊≈êà¢™ç,L*§yZñ˝,ãn	Ã%)Éòd⁄ay‹hÃÜÎgC1Ñ˚ó]b˛Õˇ  ˇˇ ﬂó¯∆