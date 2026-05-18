"use client";

import { useCallback, useEffect, useMemo, useState, useSyncExternalStore } from "react";
import { useDropzone } from "react-dropzone";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  BarChart3,
  FileSpreadsheet,
  Languages,
  LogOut,
  RefreshCcw,
  Send,
  UploadCloud,
  X,
} from "lucide-react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Progress } from "@/components/ui/progress";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { cn } from "@/lib/utils";

type Language = "zh" | "en";
type ChartItem = { name: string; value: number };
type AggregatedChartData = {
  chartData: ChartItem[];
  excludedCount: number;
  totalCount: number;
};
type InsightLatestItem = {
  query: string;
  role: string;
  industry: string;
  sentiment: string;
};
type FeedbackLatestItem = {
  query: string;
  frictionSignal: string;
  featureRequest: string;
};
type InsightFilters = {
  userRole: string;
  industry: string;
  sentiment: string;
  queryLength: string;
};
type FeedbackFilters = {
  frictionSignal: string;
  featureRequest: string;
};
type FilterOption = {
  value: string;
  label: string;
};
type ActiveFilterChip = {
  key: string;
  label: string;
  value: string;
  onRemove: () => void;
};
type UserRole = "engineer" | "analyst";
type ProcessApiError = {
  error?: string;
  status?: number;
  details?: string;
  upstreamStatus?: number;
  upstreamStatusText?: string;
  queryPreview?: string;
};

const INSIGHT_CSV_URL =
  process.env.NEXT_PUBLIC_INSIGHT_CSV_URL ??
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vRjtyV41WFbZfq1QPC3vAfAiRMuUY19ebSKaHKZxDWaaeMhwWhQR6DE39KDXe2UkvoqsVX8jrS5LO0h/pub?gid=0&single=true&output=csv";
const FEEDBACK_CSV_URL =
  process.env.NEXT_PUBLIC_FEEDBACK_CSV_URL ??
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vRjtyV41WFbZfq1QPC3vAfAiRMuUY19ebSKaHKZxDWaaeMhwWhQR6DE39KDXe2UkvoqsVX8jrS5LO0h/pub?gid=813873150&single=true&output=csv";
const USERS_CSV_URL = process.env.NEXT_PUBLIC_USERS_CSV_URL ?? "";
const ROLE_STORAGE_KEY = "oak_role";
const ROLE_CHANGE_EVENT = "oak-role-change";

const ALL_FILTER_VALUE = "__ALL__";
const UNKNOWN_VALUE = "__UNKNOWN__";
const CHART_ANIMATION_MS = 650;
const INSIGHT_COLORS = ["#0f172a", "#1d4ed8", "#0f766e", "#0e7490"];
const FEEDBACK_COLORS = ["#1f2937", "#0f766e", "#e76f51", "#6d28d9"];
const INVALID_DIMENSION_LABELS = new Set(["unknown", "none", "ambiguous", "other", "", "null", "undefined"]);

const copy: Record<
  Language,
  {
    appTitle: string;
    appSubtitle: string;
    uploadTitle: string;
    uploadDesc: string;
    dropActive: string;
    dropIdle: string;
    chooseFile: string;
    selectedFile: string;
    submitButton: string;
    submitting: string;
    progress: (sent: number, total: number) => string;
    readSuccess: (count: number) => string;
    readFailed: (reason: string) => string;
    submitDone: (success: number, failed: number) => string;
    submitWarning: string;
    dashboardTitle: string;
    dashboardDesc: string;
    refreshButton: string;
    loadingDashboard: string;
    insightTab: string;
    feedbackTab: string;
    filterConsole: string;
    filterAll: string;
    filterUserRole: string;
    filterIndustry: string;
    filterSentiment: string;
    filterQueryLength: string;
    filterFrictionSignal: string;
    filterFeatureRequest: string;
    activeFilters: string;
    clearFilters: string;
    roleChart: string;
    queryTypeChart: string;
    queryLengthChart: string;
    industryChart: string;
    sentimentChart: string;
    frictionChart: string;
    featureChart: string;
    safetyChart: string;
    insightLatestTableTitle: string;
    feedbackLatestTableTitle: string;
    columnQuery: string;
    columnRole: string;
    columnIndustry: string;
    columnSentiment: string;
    columnFrictionSignal: string;
    columnFeatureRequest: string;
    emptyData: string;
    unknown: string;
    lastUpdated: string;
    excludedBadge: (excludedCount: number, percentage: string) => string;
    dashboardFailed: (reason: string) => string;
  }
> = {
  zh: {
    appTitle: "Oak 企业单页数据看板",
    appSubtitle: "Single Page Application · Client-Only Dashboard",
    uploadTitle: "模块 A：文件拖拽处理区",
    uploadDesc: "上传 .xlsx / .csv，提取 Original_Query 列并逐条提交给 AI 处理流程。",
    dropActive: "松开鼠标即可上传文件",
    dropIdle: "拖拽文件到此处，或点击选择文件",
    chooseFile: "重新选择文件",
    selectedFile: "当前文件",
    submitButton: "🚀 提交给 AI 处理",
    submitting: "处理中...",
    progress: (sent, total) => `已发送 ${sent} / 总数 ${total}`,
    readSuccess: (count) => `成功读取了 ${count} 条数据待处理`,
    readFailed: (reason) => `文件解析失败：${reason}`,
    submitDone: (success, failed) => `提交完成：成功 ${success} 条，失败 ${failed} 条`,
    submitWarning: "存在失败请求，请检查 /api/process 或 n8n Webhook 状态。",
    dashboardTitle: "模块 B：实时数据看板区",
    dashboardDesc: "页面加载或点击刷新后，拉取 Insight 与 Feedback 两张公开表并进行可视化统计。",
    refreshButton: "🔄 刷新看板",
    loadingDashboard: "正在加载看板数据...",
    insightTab: "Business Insight 业务洞察",
    feedbackTab: "Product Feedback 产品体验",
    filterConsole: "全局筛选控制台",
    filterAll: "全部 (All)",
    filterUserRole: "用户角色 (User_Role)",
    filterIndustry: "行业 (Industry)",
    filterSentiment: "情绪 (User_Sentiment)",
    filterQueryLength: "查询长度分层 (Query_Length_Category)",
    filterFrictionSignal: "痛点类型 (Friction_Signal)",
    filterFeatureRequest: "新功能需求 (Feature_Request)",
    activeFilters: "已生效筛选",
    clearFilters: "清空筛选",
    roleChart: "User_Role 分布",
    queryTypeChart: "Query_Type 分布",
    queryLengthChart: "Query_Length_Category 查询长度分层",
    industryChart: "Industry 分布",
    sentimentChart: "User_Sentiment 情绪大盘",
    frictionChart: "Friction_Signal 痛点分布",
    featureChart: "Feature_Request 功能需求",
    safetyChart: "Safety_or_Out_of_Scope_Flag 异常拦截",
    insightLatestTableTitle: "最新 5 条业务查询明细",
    feedbackLatestTableTitle: "最新 5 条产品反馈明细",
    columnQuery: "提问内容",
    columnRole: "User_Role",
    columnIndustry: "Industry",
    columnSentiment: "User_Sentiment",
    columnFrictionSignal: "Friction_Signal",
    columnFeatureRequest: "Feature_Request",
    emptyData: "暂无可展示的数据",
    unknown: "未标注",
    lastUpdated: "最近刷新时间",
    excludedBadge: (excludedCount, percentage) => `未分类/其他: ${excludedCount} (${percentage}%)`,
    dashboardFailed: (reason) => `看板加载失败：${reason}`,
  },
  en: {
    appTitle: "Oak Enterprise SPA Dashboard",
    appSubtitle: "Single Page Application · Client-Only Dashboard",
    uploadTitle: "Module A: File Drop Processing",
    uploadDesc:
      "Upload .xlsx / .csv, extract the Original_Query column, and submit each line to the AI pipeline.",
    dropActive: "Drop the file to upload",
    dropIdle: "Drag and drop a file here, or click to select",
    chooseFile: "Choose another file",
    selectedFile: "Selected file",
    submitButton: "🚀 Submit to AI Pipeline",
    submitting: "Submitting...",
    progress: (sent, total) => `Sent ${sent} / Total ${total}`,
    readSuccess: (count) => `Successfully loaded ${count} rows for processing`,
    readFailed: (reason) => `Failed to parse file: ${reason}`,
    submitDone: (success, failed) => `Done: ${success} succeeded, ${failed} failed`,
    submitWarning: "Some requests failed. Please check /api/process and n8n webhook status.",
    dashboardTitle: "Module B: Realtime Analytics Dashboard",
    dashboardDesc: "On page load or refresh, fetch both Insight and Feedback CSV sources for visualization.",
    refreshButton: "🔄 Refresh Dashboard",
    loadingDashboard: "Loading dashboard data...",
    insightTab: "Business Insight",
    feedbackTab: "Product Feedback",
    filterConsole: "Global Filter Console",
    filterAll: "All",
    filterUserRole: "User Role (User_Role)",
    filterIndustry: "Industry",
    filterSentiment: "User Sentiment",
    filterQueryLength: "Query Length (Query_Length_Category)",
    filterFrictionSignal: "Friction Signal",
    filterFeatureRequest: "Feature Request",
    activeFilters: "Active Filters",
    clearFilters: "Clear Filters",
    roleChart: "User_Role Distribution",
    queryTypeChart: "Query_Type Distribution",
    queryLengthChart: "Query_Length_Category Distribution",
    industryChart: "Industry Distribution",
    sentimentChart: "User_Sentiment Overview",
    frictionChart: "Friction_Signal Distribution",
    featureChart: "Feature_Request Demand",
    safetyChart: "Safety_or_Out_of_Scope_Flag",
    insightLatestTableTitle: "Latest 5 Insight Queries",
    feedbackLatestTableTitle: "Latest 5 Feedback Queries",
    columnQuery: "Query",
    columnRole: "User_Role",
    columnIndustry: "Industry",
    columnSentiment: "User_Sentiment",
    columnFrictionSignal: "Friction_Signal",
    columnFeatureRequest: "Feature_Request",
    emptyData: "No data available",
    unknown: "Unknown",
    lastUpdated: "Last updated",
    excludedBadge: (excludedCount, percentage) => `Unclassified / Other: ${excludedCount} (${percentage}%)`,
    dashboardFailed: (reason) => `Failed to load dashboard: ${reason}`,
  },
};

function normalizeRows(rows: Record<string, string | undefined>[]) {
  return rows
    .map((row) => {
      const normalizedRow: Record<string, string> = {};
      Object.entries(row).forEach(([key, value]) => {
        normalizedRow[key.trim()] = String(value ?? "").trim();
      });
      return normalizedRow;
    })
    .filter((row) => Object.values(row).some((value) => value.length > 0));
}

function parseCsvByUrl(url: string) {
  return new Promise<Record<string, string>[]>((resolve, reject) => {
    Papa.parse<Record<string, string | undefined>>(url, {
      download: true,
      header: true,
      skipEmptyLines: true,
      transformHeader: (header: string) => header.trim(),
      complete: (result) => {
        if (result.errors.length > 0 && result.data.length === 0) {
          reject(new Error(result.errors[0].message));
          return;
        }
        resolve(normalizeRows(result.data));
      },
      error: (error) => {
        reject(error);
      },
    });
  });
}

function getValueByAliases(row: Record<string, string>, aliases: string[]) {
  const loweredEntries = Object.entries(row).map(([key, value]) => [key.trim().toLowerCase(), String(value ?? "").trim()] as const);
  for (const alias of aliases) {
    const hit = loweredEntries.find(([key]) => key === alias);
    if (hit && hit[1]) {
      return hit[1];
    }
  }
  return "";
}

function parseUsersCsv(url: string) {
  return new Promise<Record<string, string>[]>((resolve, reject) => {
    Papa.parse<Record<string, string | undefined>>(url, {
      download: true,
      header: true,
      skipEmptyLines: true,
      transformHeader: (header: string) => header.trim(),
      complete: (result) => {
        if (result.errors.length > 0 && result.data.length === 0) {
          reject(new Error(result.errors[0].message));
          return;
        }
        resolve(normalizeRows(result.data));
      },
      error: (error) => reject(error),
    });
  });
}

function getStoredRole(): UserRole | null {
  if (typeof window === "undefined") return null;
  const savedRole = window.localStorage.getItem(ROLE_STORAGE_KEY);
  if (savedRole === "engineer" || savedRole === "analyst") {
    return savedRole;
  }
  return null;
}

function subscribeRole(onStoreChange: () => void) {
  if (typeof window === "undefined") {
    return () => {};
  }

  const handleChange = () => onStoreChange();
  window.addEventListener("storage", handleChange);
  window.addEventListener(ROLE_CHANGE_EVENT, handleChange);

  return () => {
    window.removeEventListener("storage", handleChange);
    window.removeEventListener(ROLE_CHANGE_EVENT, handleChange);
  };
}

function isValidSignal(value: string | undefined) {
  if (!value) return false;
  const normalized = value.trim();
  if (!normalized) return false;
  const lower = normalized.toLowerCase();
  return lower !== "none" && lower !== "undefined";
}

function normalizeDimensionValue(value: string | undefined) {
  const normalized = String(value ?? "").trim();
  return normalized ? normalized : UNKNOWN_VALUE;
}

function isInvalidDimensionLabel(value: string) {
  const normalized = value.trim().toLowerCase();
  if (normalized === UNKNOWN_VALUE.toLowerCase()) {
    return true;
  }
  return INVALID_DIMENSION_LABELS.has(normalized) || normalized.startsWith("other ");
}

function aggregateAndSortData(data: Record<string, string>[], key: string): AggregatedChartData {
  const counter = new Map<string, number>();
  let totalCount = 0;

  data.forEach((row) => {
    const value = normalizeDimensionValue(row[key]);
    totalCount += 1;
    counter.set(value, (counter.get(value) ?? 0) + 1);
  });

  const sorted = Array.from(counter.entries())
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);

  let excludedCount = 0;
  const validItems: ChartItem[] = [];

  sorted.forEach((item) => {
    if (isInvalidDimensionLabel(item.name)) {
      excludedCount += item.value;
      return;
    }
    validItems.push(item);
  });

  if (validItems.length > 3) {
    excludedCount += validItems.slice(3).reduce((sum, item) => sum + item.value, 0);
  }

  return {
    chartData: validItems.slice(0, 3),
    excludedCount,
    totalCount,
  };
}

function toDisplayLabel(value: string, unknownLabel: string) {
  return value === UNKNOWN_VALUE ? unknownLabel : value;
}

function formatMetricLabel(value: string) {
  return value.replace(/_/g, " ").replace(/\s+/g, " ").trim();
}

function buildFilterOptions(rows: Record<string, string>[], key: string, allLabel: string, unknownLabel: string): FilterOption[] {
  const uniqueValues = Array.from(new Set(rows.map((row) => normalizeDimensionValue(row[key])))).sort((a, b) =>
    a.localeCompare(b, "zh-CN")
  );

  return [
    { value: ALL_FILTER_VALUE, label: allLabel },
    ...uniqueValues.map((value) => ({
      value,
      label: toDisplayLabel(value, unknownLabel),
    })),
  ];
}

function sleep(ms: number) {
  return new Promise<void>((resolve) => {
    setTimeout(resolve, ms);
  });
}

export default function HomePage() {
  const [language, setLanguage] = useState<Language>("zh");
  const role = useSyncExternalStore(subscribeRole, getStoredRole, () => null);
  const [authTab, setAuthTab] = useState<"login" | "signup">("login");
  const [authUsername, setAuthUsername] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authError, setAuthError] = useState("");
  const [authSuccess, setAuthSuccess] = useState("");
  const [isAuthSubmitting, setIsAuthSubmitting] = useState(false);
  const [queries, setQueries] = useState<string[]>([]);
  const [fileName, setFileName] = useState<string>("");
  const [readMessage, setReadMessage] = useState<string>("");
  const [readError, setReadError] = useState<string>("");

  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitProgress, setSubmitProgress] = useState(0);
  const [submittedCount, setSubmittedCount] = useState(0);
  const [submitResult, setSubmitResult] = useState<{ success: number; failed: number } | null>(null);

  const [isDashboardLoading, setIsDashboardLoading] = useState(false);
  const [dashboardError, setDashboardError] = useState("");
  const [insightRows, setInsightRows] = useState<Record<string, string>[]>([]);
  const [feedbackRows, setFeedbackRows] = useState<Record<string, string>[]>([]);
  const [lastUpdated, setLastUpdated] = useState("");
  const [activeTab, setActiveTab] = useState("insight");
  const [insightFilters, setInsightFilters] = useState<InsightFilters>({
    userRole: ALL_FILTER_VALUE,
    industry: ALL_FILTER_VALUE,
    sentiment: ALL_FILTER_VALUE,
    queryLength: ALL_FILTER_VALUE,
  });
  const [feedbackFilters, setFeedbackFilters] = useState<FeedbackFilters>({
    frictionSignal: ALL_FILTER_VALUE,
    featureRequest: ALL_FILTER_VALUE,
  });

  const t = copy[language];

  const onDrop = useCallback(
    async (acceptedFiles: File[]) => {
      const selectedFile = acceptedFiles[0];
      if (!selectedFile) return;

      setFileName(selectedFile.name);
      setQueries([]);
      setReadMessage("");
      setReadError("");
      setSubmitProgress(0);
      setSubmittedCount(0);
      setSubmitResult(null);

      try {
        const buffer = await selectedFile.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) {
          throw new Error("未找到工作表");
        }

        const firstSheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, {
          header: 1,
          defval: "",
          blankrows: false,
        });

        if (rows.length < 2) {
          throw new Error("文件缺少可处理的数据行");
        }

        const header = (rows[0] as unknown[]).map((cell) => String(cell ?? "").trim());
        const queryColIndex = header.findIndex((name) => name === "Original_Query");
        const targetIndex = queryColIndex >= 0 ? queryColIndex : 0;

        const extractedQueries = rows
          .slice(1)
          .map((row) => {
            const cells = Array.isArray(row) ? row : [];
            return String(cells[targetIndex] ?? "").trim();
          })
          .filter((text) => text.length > 0);

        if (extractedQueries.length === 0) {
          throw new Error("未读取到 Original_Query 或第一列文本");
        }

        setQueries(extractedQueries);
        setReadMessage(t.readSuccess(extractedQueries.length));
      } catch (error) {
        const reason = error instanceof Error ? error.message : "未知错误";
        setReadError(t.readFailed(reason));
      }
    },
    [t]
  );

  const { getRootProps, getInputProps, isDragActive, open } = useDropzone({
    onDrop,
    multiple: false,
    noClick: true,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
      "text/csv": [".csv"],
      "application/vnd.ms-excel": [".csv"],
    },
  });

  const handleLogin = useCallback(async () => {
    if (isAuthSubmitting) return;
    setAuthError("");
    setAuthSuccess("");
    setIsAuthSubmitting(true);

    try {
      const inputUsername = authUsername.trim();
      const inputPassword = authPassword.trim();

      if (!inputUsername || !inputPassword || !USERS_CSV_URL) {
        setAuthError("Invalid credentials");
        return;
      }

      const rows = await parseUsersCsv(USERS_CSV_URL);
      const matched = rows.find((row) => {
        const username = getValueByAliases(row, ["username", "user_name", "user", "account"]);
        const password = getValueByAliases(row, ["password", "pass"]);
        return username === inputUsername && password === inputPassword;
      });

      if (!matched) {
        setAuthError("Invalid credentials");
        return;
      }

      const matchedRole = getValueByAliases(matched, ["role"]).toLowerCase();
      const nextRole: UserRole = matchedRole === "engineer" ? "engineer" : "analyst";
      window.localStorage.setItem(ROLE_STORAGE_KEY, nextRole);
      window.dispatchEvent(new Event(ROLE_CHANGE_EVENT));
      setAuthPassword("");
      setAuthError("");
    } catch {
      setAuthError("Invalid credentials");
    } finally {
      setIsAuthSubmitting(false);
    }
  }, [authPassword, authUsername, isAuthSubmitting]);

  const handleSignUp = useCallback(async () => {
    if (isAuthSubmitting) return;
    setAuthError("");
    setAuthSuccess("");

    const username = authUsername.trim();
    const password = authPassword.trim();
    if (!username || !password) {
      setAuthError("Username and password are required");
      return;
    }

    setIsAuthSubmitting(true);
    try {
      const response = await fetch("/api/register", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          username,
          password,
          role: "analyst",
        }),
      });

      if (!response.ok) {
        let errorMessage = "Registration failed";
        try {
          const data = (await response.json()) as { error?: string };
          errorMessage = data.error ?? errorMessage;
        } catch {
          // 保持默认错误提示，避免解析失败导致用户无反馈。
        }
        setAuthError(errorMessage);
        return;
      }

      setAuthSuccess("Registration successful, please login");
      setAuthTab("login");
      setAuthPassword("");
    } catch {
      setAuthError("Registration failed");
    } finally {
      setIsAuthSubmitting(false);
    }
  }, [authPassword, authUsername, isAuthSubmitting]);

  const handleLogout = useCallback(() => {
    window.localStorage.removeItem(ROLE_STORAGE_KEY);
    window.dispatchEvent(new Event(ROLE_CHANGE_EVENT));
    window.location.reload();
  }, []);

  const submitQueries = useCallback(async () => {
    if (queries.length === 0 || isSubmitting) return;

    setIsSubmitting(true);
    setSubmitResult(null);
    setSubmitProgress(0);
    setSubmittedCount(0);

    let successCount = 0;
    let failedCount = 0;

    for (const [index, query] of queries.entries()) {
      try {
        const response = await fetch("/api/process", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ query }),
        });

        if (!response.ok) {
          let errorDetail: ProcessApiError | string | null = null;
          try {
            errorDetail = (await response.json()) as ProcessApiError;
          } catch {
            try {
              errorDetail = await response.text();
            } catch {
              errorDetail = null;
            }
          }

          console.error("[AI提交失败]", {
            index: index + 1,
            total: queries.length,
            query,
            httpStatus: response.status,
            errorDetail,
          });

          failedCount += 1;
        } else {
          successCount += 1;
        }

        await sleep(4500);

        const done = index + 1;
        setSubmittedCount(done);
        setSubmitProgress(Math.round((done / queries.length) * 100));
      } catch (error) {
        console.error("[AI提交异常]", {
          index: index + 1,
          total: queries.length,
          query,
          error: error instanceof Error ? error.message : error,
        });

        failedCount += 1;
        await sleep(4500);

        const done = index + 1;
        setSubmittedCount(done);
        setSubmitProgress(Math.round((done / queries.length) * 100));
      }
    }

    setSubmitResult({ success: successCount, failed: failedCount });
    setIsSubmitting(false);
  }, [isSubmitting, queries]);

  const fetchDashboard = useCallback(async () => {
    setIsDashboardLoading(true);
    setDashboardError("");

    try {
      const [insightRows, feedbackRows] = await Promise.all([
        parseCsvByUrl(INSIGHT_CSV_URL),
        parseCsvByUrl(FEEDBACK_CSV_URL),
      ]);

      setInsightRows(insightRows);
      setFeedbackRows(feedbackRows);
      setLastUpdated(
        new Date().toLocaleString(language === "zh" ? "zh-CN" : "en-US", {
          hour12: false,
        })
      );
    } catch (error) {
      const reason = error instanceof Error ? error.message : "Unknown error";
      setDashboardError(t.dashboardFailed(reason));
    } finally {
      setIsDashboardLoading(false);
    }
  }, [language, t]);

  useEffect(() => {
    if (!role) return;
    const timer = setTimeout(() => {
      void fetchDashboard();
    }, 0);

    return () => clearTimeout(timer);
  }, [fetchDashboard, role]);

  const insightRowsForRoleOptions = useMemo(
    () =>
      insightRows.filter((row) => {
        const byIndustry =
          insightFilters.industry === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Industry) === insightFilters.industry;
        const bySentiment =
          insightFilters.sentiment === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.User_Sentiment) === insightFilters.sentiment;
        const byQueryLength =
          insightFilters.queryLength === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Query_Length_Category) === insightFilters.queryLength;
        return byIndustry && bySentiment && byQueryLength;
      }),
    [insightFilters.industry, insightFilters.queryLength, insightFilters.sentiment, insightRows]
  );
  const insightRowsForIndustryOptions = useMemo(
    () =>
      insightRows.filter((row) => {
        const byRole =
          insightFilters.userRole === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.User_Role) === insightFilters.userRole;
        const bySentiment =
          insightFilters.sentiment === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.User_Sentiment) === insightFilters.sentiment;
        const byQueryLength =
          insightFilters.queryLength === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Query_Length_Category) === insightFilters.queryLength;
        return byRole && bySentiment && byQueryLength;
      }),
    [insightFilters.queryLength, insightFilters.sentiment, insightFilters.userRole, insightRows]
  );
  const insightRowsForSentimentOptions = useMemo(
    () =>
      insightRows.filter((row) => {
        const byRole =
          insightFilters.userRole === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.User_Role) === insightFilters.userRole;
        const byIndustry =
          insightFilters.industry === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Industry) === insightFilters.industry;
        const byQueryLength =
          insightFilters.queryLength === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Query_Length_Category) === insightFilters.queryLength;
        return byRole && byIndustry && byQueryLength;
      }),
    [insightFilters.industry, insightFilters.queryLength, insightFilters.userRole, insightRows]
  );
  const insightRowsForQueryLengthOptions = useMemo(
    () =>
      insightRows.filter((row) => {
        const byRole =
          insightFilters.userRole === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.User_Role) === insightFilters.userRole;
        const byIndustry =
          insightFilters.industry === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Industry) === insightFilters.industry;
        const bySentiment =
          insightFilters.sentiment === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.User_Sentiment) === insightFilters.sentiment;
        return byRole && byIndustry && bySentiment;
      }),
    [insightFilters.industry, insightFilters.sentiment, insightFilters.userRole, insightRows]
  );

  const feedbackRowsForFrictionOptions = useMemo(
    () =>
      feedbackRows.filter((row) => {
        if (!isValidSignal(row.Friction_Signal)) return false;
        const byFeature =
          feedbackFilters.featureRequest === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Feature_Request) === feedbackFilters.featureRequest;
        return byFeature;
      }),
    [feedbackFilters.featureRequest, feedbackRows]
  );
  const feedbackRowsForFeatureOptions = useMemo(
    () =>
      feedbackRows.filter((row) => {
        if (!isValidSignal(row.Feature_Request)) return false;
        const byFriction =
          feedbackFilters.frictionSignal === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Friction_Signal) === feedbackFilters.frictionSignal;
        return byFriction;
      }),
    [feedbackFilters.frictionSignal, feedbackRows]
  );

  const insightRoleOptions = useMemo(
    () => buildFilterOptions(insightRowsForRoleOptions, "User_Role", t.filterAll, t.unknown),
    [insightRowsForRoleOptions, t.filterAll, t.unknown]
  );
  const insightIndustryOptions = useMemo(
    () => buildFilterOptions(insightRowsForIndustryOptions, "Industry", t.filterAll, t.unknown),
    [insightRowsForIndustryOptions, t.filterAll, t.unknown]
  );
  const insightSentimentOptions = useMemo(
    () => buildFilterOptions(insightRowsForSentimentOptions, "User_Sentiment", t.filterAll, t.unknown),
    [insightRowsForSentimentOptions, t.filterAll, t.unknown]
  );
  const insightQueryLengthOptions = useMemo(
    () => buildFilterOptions(insightRowsForQueryLengthOptions, "Query_Length_Category", t.filterAll, t.unknown),
    [insightRowsForQueryLengthOptions, t.filterAll, t.unknown]
  );
  const feedbackFrictionOptions = useMemo(
    () => buildFilterOptions(feedbackRowsForFrictionOptions, "Friction_Signal", t.filterAll, t.unknown),
    [feedbackRowsForFrictionOptions, t.filterAll, t.unknown]
  );
  const feedbackFeatureOptions = useMemo(
    () => buildFilterOptions(feedbackRowsForFeatureOptions, "Feature_Request", t.filterAll, t.unknown),
    [feedbackRowsForFeatureOptions, t.filterAll, t.unknown]
  );

  const sanitizeInsightFilters = useCallback(
    (next: InsightFilters): InsightFilters => {
      const validRoles = new Set(
        insightRows
          .filter((row) => {
            const byIndustry =
              next.industry === ALL_FILTER_VALUE || normalizeDimensionValue(row.Industry) === next.industry;
            const bySentiment =
              next.sentiment === ALL_FILTER_VALUE || normalizeDimensionValue(row.User_Sentiment) === next.sentiment;
            const byQueryLength =
              next.queryLength === ALL_FILTER_VALUE ||
              normalizeDimensionValue(row.Query_Length_Category) === next.queryLength;
            return byIndustry && bySentiment && byQueryLength;
          })
          .map((row) => normalizeDimensionValue(row.User_Role))
      );
      const validIndustries = new Set(
        insightRows
          .filter((row) => {
            const byRole =
              next.userRole === ALL_FILTER_VALUE || normalizeDimensionValue(row.User_Role) === next.userRole;
            const bySentiment =
              next.sentiment === ALL_FILTER_VALUE || normalizeDimensionValue(row.User_Sentiment) === next.sentiment;
            const byQueryLength =
              next.queryLength === ALL_FILTER_VALUE ||
              normalizeDimensionValue(row.Query_Length_Category) === next.queryLength;
            return byRole && bySentiment && byQueryLength;
          })
          .map((row) => normalizeDimensionValue(row.Industry))
      );
      const validSentiments = new Set(
        insightRows
          .filter((row) => {
            const byRole =
              next.userRole === ALL_FILTER_VALUE || normalizeDimensionValue(row.User_Role) === next.userRole;
            const byIndustry =
              next.industry === ALL_FILTER_VALUE || normalizeDimensionValue(row.Industry) === next.industry;
            const byQueryLength =
              next.queryLength === ALL_FILTER_VALUE ||
              normalizeDimensionValue(row.Query_Length_Category) === next.queryLength;
            return byRole && byIndustry && byQueryLength;
          })
          .map((row) => normalizeDimensionValue(row.User_Sentiment))
      );
      const validQueryLengths = new Set(
        insightRows
          .filter((row) => {
            const byRole =
              next.userRole === ALL_FILTER_VALUE || normalizeDimensionValue(row.User_Role) === next.userRole;
            const byIndustry =
              next.industry === ALL_FILTER_VALUE || normalizeDimensionValue(row.Industry) === next.industry;
            const bySentiment =
              next.sentiment === ALL_FILTER_VALUE ||
              normalizeDimensionValue(row.User_Sentiment) === next.sentiment;
            return byRole && byIndustry && bySentiment;
          })
          .map((row) => normalizeDimensionValue(row.Query_Length_Category))
      );

      return {
        userRole: next.userRole !== ALL_FILTER_VALUE && !validRoles.has(next.userRole) ? ALL_FILTER_VALUE : next.userRole,
        industry:
          next.industry !== ALL_FILTER_VALUE && !validIndustries.has(next.industry) ? ALL_FILTER_VALUE : next.industry,
        sentiment:
          next.sentiment !== ALL_FILTER_VALUE && !validSentiments.has(next.sentiment) ? ALL_FILTER_VALUE : next.sentiment,
        queryLength:
          next.queryLength !== ALL_FILTER_VALUE && !validQueryLengths.has(next.queryLength)
            ? ALL_FILTER_VALUE
            : next.queryLength,
      };
    },
    [insightRows]
  );

  const sanitizeFeedbackFilters = useCallback(
    (next: FeedbackFilters): FeedbackFilters => {
      const validFrictions = new Set(
        feedbackRows
          .filter((row) => {
            if (!isValidSignal(row.Friction_Signal)) return false;
            const byFeature =
              next.featureRequest === ALL_FILTER_VALUE ||
              normalizeDimensionValue(row.Feature_Request) === next.featureRequest;
            return byFeature;
          })
          .map((row) => normalizeDimensionValue(row.Friction_Signal))
      );
      const validFeatures = new Set(
        feedbackRows
          .filter((row) => {
            if (!isValidSignal(row.Feature_Request)) return false;
            const byFriction =
              next.frictionSignal === ALL_FILTER_VALUE ||
              normalizeDimensionValue(row.Friction_Signal) === next.frictionSignal;
            return byFriction;
          })
          .map((row) => normalizeDimensionValue(row.Feature_Request))
      );

      return {
        frictionSignal:
          next.frictionSignal !== ALL_FILTER_VALUE && !validFrictions.has(next.frictionSignal)
            ? ALL_FILTER_VALUE
            : next.frictionSignal,
        featureRequest:
          next.featureRequest !== ALL_FILTER_VALUE && !validFeatures.has(next.featureRequest)
            ? ALL_FILTER_VALUE
            : next.featureRequest,
      };
    },
    [feedbackRows]
  );

  const filteredInsightRows = useMemo(
    () =>
      insightRows.filter((row) => {
        const byRole =
          insightFilters.userRole === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.User_Role) === insightFilters.userRole;
        const byIndustry =
          insightFilters.industry === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Industry) === insightFilters.industry;
        const bySentiment =
          insightFilters.sentiment === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.User_Sentiment) === insightFilters.sentiment;
        const byQueryLength =
          insightFilters.queryLength === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Query_Length_Category) === insightFilters.queryLength;
        return byRole && byIndustry && bySentiment && byQueryLength;
      }),
    [insightRows, insightFilters]
  );

  const filteredFeedbackRows = useMemo(
    () =>
      feedbackRows.filter((row) => {
        const byFriction =
          feedbackFilters.frictionSignal === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Friction_Signal) === feedbackFilters.frictionSignal;
        const byFeature =
          feedbackFilters.featureRequest === ALL_FILTER_VALUE ||
          normalizeDimensionValue(row.Feature_Request) === feedbackFilters.featureRequest;
        return byFriction && byFeature;
      }),
    [feedbackFilters, feedbackRows]
  );

  const roleData = useMemo(() => aggregateAndSortData(filteredInsightRows, "User_Role"), [filteredInsightRows]);
  const queryTypeData = useMemo(() => aggregateAndSortData(filteredInsightRows, "Query_Type"), [filteredInsightRows]);
  const queryLengthData = useMemo(
    () => aggregateAndSortData(filteredInsightRows, "Query_Length_Category"),
    [filteredInsightRows]
  );
  const industryData = useMemo(() => aggregateAndSortData(filteredInsightRows, "Industry"), [filteredInsightRows]);
  const sentimentData = useMemo(
    () => aggregateAndSortData(filteredInsightRows, "User_Sentiment"),
    [filteredInsightRows]
  );

  const frictionData = useMemo(
    () => aggregateAndSortData(filteredFeedbackRows.filter((row) => isValidSignal(row.Friction_Signal)), "Friction_Signal"),
    [filteredFeedbackRows]
  );
  const featureData = useMemo(
    () => aggregateAndSortData(filteredFeedbackRows.filter((row) => isValidSignal(row.Feature_Request)), "Feature_Request"),
    [filteredFeedbackRows]
  );
  const safetyData = useMemo(
    () =>
      aggregateAndSortData(
        filteredFeedbackRows.filter((row) => isValidSignal(row.Safety_or_Out_of_Scope_Flag)),
        "Safety_or_Out_of_Scope_Flag"
      ),
    [filteredFeedbackRows]
  );

  const insightLatestRows = useMemo<InsightLatestItem[]>(
    () =>
      insightRows
        .slice(-5)
        .reverse()
        .map((row) => ({
          query: row.Original_Query || t.emptyData,
          role: toDisplayLabel(normalizeDimensionValue(row.User_Role), t.unknown),
          industry: toDisplayLabel(normalizeDimensionValue(row.Industry), t.unknown),
          sentiment: toDisplayLabel(normalizeDimensionValue(row.User_Sentiment), t.unknown),
        })),
    [insightRows, t.emptyData, t.unknown]
  );
  const feedbackLatestRows = useMemo<FeedbackLatestItem[]>(
    () =>
      feedbackRows
        .slice(-5)
        .reverse()
        .map((row) => ({
          query: row.Original_Query || t.emptyData,
          frictionSignal: isValidSignal(row.Friction_Signal)
            ? row.Friction_Signal
            : toDisplayLabel(normalizeDimensionValue(row.Friction_Signal), t.unknown),
          featureRequest: isValidSignal(row.Feature_Request)
            ? row.Feature_Request
            : toDisplayLabel(normalizeDimensionValue(row.Feature_Request), t.unknown),
        })),
    [feedbackRows, t.emptyData, t.unknown]
  );

  const insightPalette = INSIGHT_COLORS;
  const feedbackPalette = FEEDBACK_COLORS;

  const insightOptionMap = useMemo(
    () => ({
      userRole: new Map(insightRoleOptions.map((option) => [option.value, option.label])),
      industry: new Map(insightIndustryOptions.map((option) => [option.value, option.label])),
      sentiment: new Map(insightSentimentOptions.map((option) => [option.value, option.label])),
      queryLength: new Map(insightQueryLengthOptions.map((option) => [option.value, option.label])),
    }),
    [insightIndustryOptions, insightQueryLengthOptions, insightRoleOptions, insightSentimentOptions]
  );
  const feedbackOptionMap = useMemo(
    () => ({
      frictionSignal: new Map(feedbackFrictionOptions.map((option) => [option.value, option.label])),
      featureRequest: new Map(feedbackFeatureOptions.map((option) => [option.value, option.label])),
    }),
    [feedbackFeatureOptions, feedbackFrictionOptions]
  );

  const activeFilterChips = useMemo(() => {
    if (activeTab === "insight") {
      return [
        insightFilters.userRole !== ALL_FILTER_VALUE
          ? {
              key: "userRole",
              label: t.filterUserRole,
              value: insightOptionMap.userRole.get(insightFilters.userRole) ?? insightFilters.userRole,
              onRemove: () => setInsightFilters((prev) => ({ ...prev, userRole: ALL_FILTER_VALUE })),
            }
          : null,
        insightFilters.industry !== ALL_FILTER_VALUE
          ? {
              key: "industry",
              label: t.filterIndustry,
              value: insightOptionMap.industry.get(insightFilters.industry) ?? insightFilters.industry,
              onRemove: () => setInsightFilters((prev) => ({ ...prev, industry: ALL_FILTER_VALUE })),
            }
          : null,
        insightFilters.sentiment !== ALL_FILTER_VALUE
          ? {
              key: "sentiment",
              label: t.filterSentiment,
              value: insightOptionMap.sentiment.get(insightFilters.sentiment) ?? insightFilters.sentiment,
              onRemove: () => setInsightFilters((prev) => ({ ...prev, sentiment: ALL_FILTER_VALUE })),
            }
          : null,
        insightFilters.queryLength !== ALL_FILTER_VALUE
          ? {
              key: "queryLength",
              label: t.filterQueryLength,
              value: insightOptionMap.queryLength.get(insightFilters.queryLength) ?? insightFilters.queryLength,
              onRemove: () => setInsightFilters((prev) => ({ ...prev, queryLength: ALL_FILTER_VALUE })),
            }
          : null,
      ].filter((chip): chip is ActiveFilterChip => chip !== null);
    }

    return [
      feedbackFilters.frictionSignal !== ALL_FILTER_VALUE
        ? {
            key: "frictionSignal",
            label: t.filterFrictionSignal,
            value: feedbackOptionMap.frictionSignal.get(feedbackFilters.frictionSignal) ?? feedbackFilters.frictionSignal,
            onRemove: () => setFeedbackFilters((prev) => ({ ...prev, frictionSignal: ALL_FILTER_VALUE })),
          }
        : null,
      feedbackFilters.featureRequest !== ALL_FILTER_VALUE
        ? {
            key: "featureRequest",
            label: t.filterFeatureRequest,
            value: feedbackOptionMap.featureRequest.get(feedbackFilters.featureRequest) ?? feedbackFilters.featureRequest,
            onRemove: () => setFeedbackFilters((prev) => ({ ...prev, featureRequest: ALL_FILTER_VALUE })),
          }
        : null,
    ].filter((chip): chip is ActiveFilterChip => chip !== null);
  }, [
    activeTab,
    feedbackFilters.featureRequest,
    feedbackFilters.frictionSignal,
    feedbackOptionMap.featureRequest,
    feedbackOptionMap.frictionSignal,
    insightFilters.industry,
    insightFilters.queryLength,
    insightFilters.sentiment,
    insightFilters.userRole,
    insightOptionMap.industry,
    insightOptionMap.queryLength,
    insightOptionMap.sentiment,
    insightOptionMap.userRole,
    t.filterFeatureRequest,
    t.filterFrictionSignal,
    t.filterIndustry,
    t.filterQueryLength,
    t.filterSentiment,
    t.filterUserRole,
  ]);

  const renderGlassTooltip = (props: {
    active?: boolean;
    label?: string | number;
    total: number;
    payload?: ReadonlyArray<{
      value?: string | number | ReadonlyArray<string | number>;
      name?: string | number;
      payload?: {
        name?: string | number;
        value?: string | number | ReadonlyArray<string | number>;
      };
    }>;
  }) => {
    const { active, payload, label, total } = props;
    if (!active || !payload || payload.length === 0) return null;

    const current = payload[0];
    const value = Number(current.value ?? current.payload?.value ?? 0);
    const labelText = label === undefined ? "-" : String(label);
    const rawName = current.name ?? current.payload?.name ?? labelText;
    const name = formatMetricLabel(String(rawName));
    const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : "0.0";

    return (
      <div
        className="rounded-lg border border-white/60 bg-white/70 px-3 py-2 text-xs text-zinc-900 shadow-xl backdrop-blur-md"
        style={{ WebkitBackdropFilter: "blur(12px)" }}
      >
        <p className="mb-1 font-medium text-zinc-800">{name}</p>
        <p className="text-zinc-700">数量：{value}</p>
        <p className="text-zinc-500">占比：{percentage}%</p>
      </div>
    );
  };

  const renderExcludedBadge = (stats: AggregatedChartData) => {
    if (stats.excludedCount <= 0 || stats.totalCount <= 0) return null;
    const percentage = ((stats.excludedCount / stats.totalCount) * 100).toFixed(1);
    return (
      <div className="rounded-full bg-secondary/30 px-2 py-0.5 text-[11px] font-medium text-muted-foreground/60">
        {t.excludedBadge(stats.excludedCount, percentage)}
      </div>
    );
  };

  const renderChartLegend = (items: ChartItem[], palette: string[]) => {
    return (
      <div className="mt-3 grid w-full grid-cols-1 gap-2 rounded-md bg-zinc-50/70 p-2 text-xs text-zinc-700 sm:grid-cols-2">
        {items.map((item, index) => (
          <div key={`legend-${item.name}-${index}`} className="flex min-w-0 items-center gap-1.5">
            <span className="h-2.5 w-2.5 shrink-0 rounded-sm" style={{ backgroundColor: palette[index % palette.length] }} />
            <span className="truncate leading-4" title={formatMetricLabel(item.name)}>
              {formatMetricLabel(item.name)}
            </span>
          </div>
        ))}
      </div>
    );
  };

  if (!role) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-[radial-gradient(circle_at_15%_20%,rgba(148,163,184,0.28),transparent_36%),radial-gradient(circle_at_85%_10%,rgba(59,130,246,0.18),transparent_38%),linear-gradient(135deg,#f8fafc,#e2e8f0,#f1f5f9)] px-4">
        <Card
          className="w-full max-w-md border border-white/50 bg-white/35 shadow-2xl backdrop-blur-xl"
          style={{ WebkitBackdropFilter: "blur(18px)" }}
        >
          <CardHeader className="space-y-2 text-center">
            <CardTitle className="text-2xl font-semibold text-zinc-900">Welcome to Oak Dashboard</CardTitle>
            <CardDescription className="text-zinc-700">Login or Sign Up to continue</CardDescription>
          </CardHeader>
          <CardContent className="space-y-5">
            <Tabs value={authTab} onValueChange={(value) => setAuthTab(value as "login" | "signup")}>
              <TabsList className="grid w-full grid-cols-2 bg-white/70">
                <TabsTrigger value="login">Login</TabsTrigger>
                <TabsTrigger value="signup">Sign Up</TabsTrigger>
              </TabsList>

              <TabsContent value="login" className="mt-4 space-y-4">
                <div className="space-y-2">
                  <label className="text-sm font-medium text-zinc-700">Username</label>
                  <input
                    value={authUsername}
                    onChange={(event) => setAuthUsername(event.target.value)}
                    className="h-10 w-full rounded-md border border-zinc-300/80 bg-white/80 px-3 text-sm outline-none transition focus:border-zinc-400 focus:ring-2 focus:ring-zinc-200"
                    autoComplete="username"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-sm font-medium text-zinc-700">Password</label>
                  <input
                    type="password"
                    value={authPassword}
                    onChange={(event) => setAuthPassword(event.target.value)}
                    className="h-10 w-full rounded-md border border-zinc-300/80 bg-white/80 px-3 text-sm outline-none transition focus:border-zinc-400 focus:ring-2 focus:ring-zinc-200"
                    autoComplete="current-password"
                  />
                </div>
                <Button className="w-full" disabled={isAuthSubmitting} onClick={() => void handleLogin()}>
                  {isAuthSubmitting ? <RefreshCcw className="h-4 w-4 animate-spin" /> : null}
                  Login
                </Button>
              </TabsContent>

              <TabsContent value="signup" className="mt-4 space-y-4">
                <div className="space-y-2">
                  <label className="text-sm font-medium text-zinc-700">Username</label>
                  <input
                    value={authUsername}
                    onChange={(event) => setAuthUsername(event.target.value)}
                    className="h-10 w-full rounded-md border border-zinc-300/80 bg-white/80 px-3 text-sm outline-none transition focus:border-zinc-400 focus:ring-2 focus:ring-zinc-200"
                    autoComplete="username"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-sm font-medium text-zinc-700">Password</label>
                  <input
                    type="password"
                    value={authPassword}
                    onChange={(event) => setAuthPassword(event.target.value)}
                    className="h-10 w-full rounded-md border border-zinc-300/80 bg-white/80 px-3 text-sm outline-none transition focus:border-zinc-400 focus:ring-2 focus:ring-zinc-200"
                    autoComplete="new-password"
                  />
                </div>
                <Button className="w-full" disabled={isAuthSubmitting} onClick={() => void handleSignUp()}>
                  {isAuthSubmitting ? <RefreshCcw className="h-4 w-4 animate-spin" /> : null}
                  Sign Up
                </Button>
              </TabsContent>
            </Tabs>

            {authError ? <p className="text-sm text-red-600">{authError}</p> : null}
            {authSuccess ? <p className="text-sm text-emerald-600">{authSuccess}</p> : null}
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-zinc-100 text-zinc-900">
      <header className="border-b border-zinc-200 bg-white/90 backdrop-blur">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-4 py-4 md:px-8">
          <div className="flex items-center gap-3">
            <div className="rounded-lg bg-zinc-900 p-2 text-white">
              <BarChart3 className="h-5 w-5" />
            </div>
            <div>
              <h1 className="text-lg font-semibold">{t.appTitle}</h1>
              <p className="text-xs text-zinc-500">{t.appSubtitle}</p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <div className="flex items-center gap-2 rounded-lg border border-zinc-200 bg-zinc-50 p-1">
              <Languages className="h-4 w-4 text-zinc-500" />
              <Button
                size="sm"
                variant={language === "zh" ? "default" : "ghost"}
                onClick={() => setLanguage("zh")}
              >
                中文
              </Button>
              <Button
                size="sm"
                variant={language === "en" ? "default" : "ghost"}
                onClick={() => setLanguage("en")}
              >
                EN
              </Button>
            </div>
            <Button variant="secondary" className="bg-zinc-200 text-zinc-800 hover:bg-zinc-300" onClick={handleLogout}>
              <LogOut className="h-4 w-4" />
              Log out
            </Button>
          </div>
        </div>
      </header>

      <main className="mx-auto flex max-w-7xl flex-col gap-6 px-4 py-6 md:px-8 md:py-8">
        <section className={cn("grid gap-6", role === "engineer" ? "lg:grid-cols-3" : "grid-cols-1")}>
          {role === "engineer" ? (
            <Card className="lg:col-span-1">
            <CardHeader>
              <CardTitle>{t.uploadTitle}</CardTitle>
              <CardDescription>{t.uploadDesc}</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div
                {...getRootProps({
                  onClick: open,
                })}
                className={cn(
                  "cursor-pointer rounded-lg border border-dashed p-6 text-center transition",
                  isDragActive
                    ? "border-zinc-900 bg-zinc-100"
                    : "border-zinc-300 bg-zinc-50 hover:border-zinc-500"
                )}
              >
                <input {...getInputProps()} />
                <UploadCloud className="mx-auto mb-3 h-8 w-8 text-zinc-600" />
                <p className="text-sm font-medium">{isDragActive ? t.dropActive : t.dropIdle}</p>
                <p className="mt-2 text-xs text-zinc-500">.xlsx / .csv</p>
              </div>

              {fileName ? (
                <div className="rounded-lg border border-zinc-200 bg-zinc-50 p-3 text-sm">
                  <p className="flex items-center gap-2 font-medium">
                    <FileSpreadsheet className="h-4 w-4" />
                    {t.selectedFile}：{fileName}
                  </p>
                </div>
              ) : null}

              {readMessage ? <p className="text-sm text-emerald-600">{readMessage}</p> : null}
              {readError ? <p className="text-sm text-red-600">{readError}</p> : null}

              <Button variant="secondary" className="w-full" onClick={submitQueries} disabled={queries.length === 0 || isSubmitting}>
                <Send className="h-4 w-4" />
                {isSubmitting ? t.submitting : t.submitButton}
              </Button>

              <Button variant="outline" className="w-full" onClick={open}>
                {t.chooseFile}
              </Button>

              <div className="space-y-2">
                <Progress value={submitProgress} />
                <p className="text-xs text-zinc-500">{t.progress(submittedCount, queries.length)}</p>
              </div>

              {submitResult ? (
                <div className="rounded-lg border border-zinc-200 bg-zinc-50 p-3 text-sm">
                  <p>{t.submitDone(submitResult.success, submitResult.failed)}</p>
                  {submitResult.failed > 0 ? (
                    <p className="mt-1 text-xs text-amber-600">{t.submitWarning}</p>
                  ) : null}
                </div>
              ) : null}
            </CardContent>
            </Card>
          ) : null}

          <Card className={cn(role === "engineer" ? "lg:col-span-2" : "w-full")}>
            <CardHeader className="flex flex-row items-center justify-between">
              <div>
                <CardTitle>{t.dashboardTitle}</CardTitle>
                <CardDescription>{t.dashboardDesc}</CardDescription>
              </div>
              <Button variant="outline" onClick={() => void fetchDashboard()} disabled={isDashboardLoading}>
                <RefreshCcw className={cn("h-4 w-4", isDashboardLoading && "animate-spin")} />
                {t.refreshButton}
              </Button>
            </CardHeader>
            <CardContent>
              {dashboardError ? (
                <p className="rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-600">
                  {dashboardError}
                </p>
              ) : null}

              {isDashboardLoading ? <p className="text-sm text-zinc-500">{t.loadingDashboard}</p> : null}

              <Tabs value={activeTab} onValueChange={setActiveTab} className="transition-all duration-300">
                <TabsList className="max-w-xl">
                  <TabsTrigger value="insight">{t.insightTab}</TabsTrigger>
                  <TabsTrigger value="feedback">{t.feedbackTab}</TabsTrigger>
                </TabsList>

                <Card className="mt-4 border-zinc-200/70 bg-zinc-50/50 shadow-none">
                  <CardHeader className="pb-3">
                    <CardTitle className="text-sm">{t.filterConsole}</CardTitle>
                  </CardHeader>
                  <CardContent className="grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                    {activeTab === "insight" ? (
                      <>
                        <div className="space-y-2">
                          <p className="text-xs text-zinc-500">{t.filterUserRole}</p>
                          <Select
                            value={insightFilters.userRole}
                            onValueChange={(value) =>
                              setInsightFilters((prev) => sanitizeInsightFilters({ ...prev, userRole: value }))
                            }
                          >
                            <SelectTrigger>
                              <SelectValue placeholder={t.filterAll} />
                            </SelectTrigger>
                            <SelectContent>
                              {insightRoleOptions.map((option) => (
                                <SelectItem key={`role-option-${option.value}`} value={option.value}>
                                  {option.label}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>
                        <div className="space-y-2">
                          <p className="text-xs text-zinc-500">{t.filterIndustry}</p>
                          <Select
                            value={insightFilters.industry}
                            onValueChange={(value) =>
                              setInsightFilters((prev) => sanitizeInsightFilters({ ...prev, industry: value }))
                            }
                          >
                            <SelectTrigger>
                              <SelectValue placeholder={t.filterAll} />
                            </SelectTrigger>
                            <SelectContent>
                              {insightIndustryOptions.map((option) => (
                                <SelectItem key={`industry-option-${option.value}`} value={option.value}>
                                  {option.label}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>
                        <div className="space-y-2">
                          <p className="text-xs text-zinc-500">{t.filterSentiment}</p>
                          <Select
                            value={insightFilters.sentiment}
                            onValueChange={(value) =>
                              setInsightFilters((prev) => sanitizeInsightFilters({ ...prev, sentiment: value }))
                            }
                          >
                            <SelectTrigger>
                              <SelectValue placeholder={t.filterAll} />
                            </SelectTrigger>
                            <SelectContent>
                              {insightSentimentOptions.map((option) => (
                                <SelectItem key={`sentiment-option-${option.value}`} value={option.value}>
                                  {option.label}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>
                        <div className="space-y-2">
                          <p className="text-xs text-zinc-500">{t.filterQueryLength}</p>
                          <Select
                            value={insightFilters.queryLength}
                            onValueChange={(value) =>
                              setInsightFilters((prev) => sanitizeInsightFilters({ ...prev, queryLength: value }))
                            }
                          >
                            <SelectTrigger>
                              <SelectValue placeholder={t.filterAll} />
                            </SelectTrigger>
                            <SelectContent>
                              {insightQueryLengthOptions.map((option) => (
                                <SelectItem key={`query-length-option-${option.value}`} value={option.value}>
                                  {option.label}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>
                      </>
                    ) : (
                      <>
                        <div className="space-y-2">
                          <p className="text-xs text-zinc-500">{t.filterFrictionSignal}</p>
                          <Select
                            value={feedbackFilters.frictionSignal}
                            onValueChange={(value) =>
                              setFeedbackFilters((prev) => sanitizeFeedbackFilters({ ...prev, frictionSignal: value }))
                            }
                          >
                            <SelectTrigger>
                              <SelectValue placeholder={t.filterAll} />
                            </SelectTrigger>
                            <SelectContent>
                              {feedbackFrictionOptions.map((option) => (
                                <SelectItem key={`friction-option-${option.value}`} value={option.value}>
                                  {option.label}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>
                        <div className="space-y-2">
                          <p className="text-xs text-zinc-500">{t.filterFeatureRequest}</p>
                          <Select
                            value={feedbackFilters.featureRequest}
                            onValueChange={(value) =>
                              setFeedbackFilters((prev) => sanitizeFeedbackFilters({ ...prev, featureRequest: value }))
                            }
                          >
                            <SelectTrigger>
                              <SelectValue placeholder={t.filterAll} />
                            </SelectTrigger>
                            <SelectContent>
                              {feedbackFeatureOptions.map((option) => (
                                <SelectItem key={`feature-option-${option.value}`} value={option.value}>
                                  {option.label}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>
                      </>
                    )}

                    <div className="space-y-2 md:col-span-2 xl:col-span-3">
                      <div className="flex items-center justify-between">
                        <p className="text-xs text-zinc-500">{t.activeFilters}</p>
                        {activeFilterChips.length > 0 ? (
                          <Button
                            size="sm"
                            variant="ghost"
                            className="h-7 px-2 text-xs"
                            onClick={() => {
                              if (activeTab === "insight") {
                                setInsightFilters({
                                  userRole: ALL_FILTER_VALUE,
                                  industry: ALL_FILTER_VALUE,
                                  sentiment: ALL_FILTER_VALUE,
                                  queryLength: ALL_FILTER_VALUE,
                                });
                                return;
                              }
                              setFeedbackFilters({
                                frictionSignal: ALL_FILTER_VALUE,
                                featureRequest: ALL_FILTER_VALUE,
                              });
                            }}
                          >
                            {t.clearFilters}
                          </Button>
                        ) : null}
                      </div>
                      {activeFilterChips.length === 0 ? (
                        <div className="text-xs text-zinc-500">{t.filterAll}</div>
                      ) : (
                        <div className="flex flex-wrap gap-2">
                          {activeFilterChips.map((chip) => (
                            <div
                              key={chip.key}
                              className="inline-flex items-center gap-1 rounded-full border border-sky-200 bg-sky-50 px-2 py-1 text-xs text-sky-700"
                            >
                              <span className="max-w-[220px] break-all">{chip.label}: {chip.value}</span>
                              <button type="button" onClick={chip.onRemove} className="opacity-80 hover:opacity-100">
                                <X className="h-3.5 w-3.5" />
                              </button>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </CardContent>
                </Card>

                <TabsContent value="insight">
                  <div className="grid gap-6 grid-cols-1 xl:grid-cols-2">
                    <Card className="h-full border-zinc-200/80 bg-white/80 shadow-sm">
                      <CardHeader className="flex flex-row items-center justify-between gap-2 pb-2">
                        <CardTitle className="text-sm">{t.roleChart}</CardTitle>
                        {renderExcludedBadge(roleData)}
                      </CardHeader>
                      <CardContent className="h-full">
                        <div className="flex min-h-[300px] h-full items-center justify-center">
                          {roleData.chartData.length === 0 ? (
                            <div className="flex h-full w-full items-center justify-center text-sm text-zinc-500">{t.emptyData}</div>
                          ) : (
                            <div className="min-h-[300px] h-full w-full">
                              <div className="h-[250px] w-full">
                                <ResponsiveContainer width="100%" height="100%">
                                  <PieChart margin={{ top: 12, right: 16, bottom: 12, left: 16 }}>
                                    <Pie
                                      data={roleData.chartData}
                                      dataKey="value"
                                      nameKey="name"
                                      outerRadius={95}
                                      paddingAngle={2}
                                      isAnimationActive
                                      animationDuration={CHART_ANIMATION_MS}
                                      animationEasing="ease-out"
                                    >
                                      {roleData.chartData.map((item, index) => (
                                        <Cell key={`role-${item.name}`} fill={insightPalette[index % insightPalette.length]} />
                                      ))}
                                    </Pie>
                                    <Tooltip content={(props) => renderGlassTooltip({ ...props, total: roleData.totalCount })} />
                                  </PieChart>
                                </ResponsiveContainer>
                              </div>
                              {renderChartLegend(roleData.chartData, insightPalette)}
                            </div>
                          )}
                        </div>
                      </CardContent>
                    </Card>

                    <Card className="h-full border-zinc-200/80 bg-white/80 shadow-sm">
                      <CardHeader className="flex flex-row items-center justify-between gap-2 pb-2">
                        <CardTitle className="text-sm">{t.queryTypeChart}</CardTitle>
                        {renderExcludedBadge(queryTypeData)}
                      </CardHeader>
                      <CardContent className="h-full">
                        <div className="flex min-h-[300px] h-full items-center justify-center">
                          {queryTypeData.chartData.length === 0 ? (
                            <div className="flex h-full w-full items-center justify-center text-sm text-zinc-500">{t.emptyData}</div>
                          ) : (
                            <div className="min-h-[300px] h-full w-full">
                              <div className="h-[250px] w-full">
                                <ResponsiveContainer width="100%" height="100%">
                                  <PieChart margin={{ top: 12, right: 16, bottom: 12, left: 16 }}>
                                    <Pie
                                      data={queryTypeData.chartData}
                                      dataKey="value"
                                      nameKey="name"
                                      innerRadius={56}
                                      outerRadius={95}
                                      paddingAngle={2}
                                      isAnimationActive
                                      animationDuration={CHART_ANIMATION_MS}
                                      animationEasing="ease-out"
                                    >
                                      {queryTypeData.chartData.map((item, index) => (
                                        <Cell
                                          key={`query-type-${item.name}`}
                                          fill={insightPalette[index % insightPalette.length]}
                                        />
                                      ))}
                                    </Pie>
                                    <Tooltip
                                      content={(props) => renderGlassTooltip({ ...props, total: queryTypeData.totalCount })}
                                    />
                                  </PieChart>
                                </ResponsiveContainer>
                              </div>
                              {renderChartLegend(queryTypeData.chartData, insightPalette)}
                            </div>
                          )}
                        </div>
                      </CardContent>
                    </Card>

                    <Card className="h-full border-zinc-200/80 bg-white/80 shadow-sm">
                      <CardHeader className="flex flex-row items-center justify-between gap-2 pb-2">
                        <CardTitle className="text-sm">{t.queryLengthChart}</CardTitle>
                        {renderExcludedBadge(queryLengthData)}
                      </CardHeader>
                      <CardContent className="h-full">
                        <div className="flex min-h-[300px] h-full items-center justify-center">
                          {queryLengthData.chartData.length === 0 ? (
                            <div className="flex h-full w-full items-center justify-center text-sm text-zinc-500">{t.emptyData}</div>
                          ) : (
                            <div className="min-h-[300px] h-full w-full">
                              <div className="h-[250px] w-full">
                                <ResponsiveContainer width="100%" height="100%">
                                  <PieChart margin={{ top: 12, right: 16, bottom: 12, left: 16 }}>
                                    <Pie
                                      data={queryLengthData.chartData}
                                      dataKey="value"
                                      nameKey="name"
                                      innerRadius={56}
                                      outerRadius={95}
                                      paddingAngle={2}
                                      isAnimationActive
                                      animationDuration={CHART_ANIMATION_MS}
                                      animationEasing="ease-out"
                                    >
                                      {queryLengthData.chartData.map((item, index) => (
                                        <Cell
                                          key={`query-length-${item.name}`}
                                          fill={insightPalette[index % insightPalette.length]}
                                        />
                                      ))}
                                    </Pie>
                                    <Tooltip
                                      content={(props) => renderGlassTooltip({ ...props, total: queryLengthData.totalCount })}
                                    />
                                  </PieChart>
                                </ResponsiveContainer>
                              </div>
                              {renderChartLegend(queryLengthData.chartData, insightPalette)}
                            </div>
                          )}
                        </div>
                      </CardContent>
                    </Card>

                    <Card className="h-full border-zinc-200/80 bg-white/80 shadow-sm">
                      <CardHeader className="flex flex-row items-center justify-between gap-2 pb-2">
                        <CardTitle className="text-sm">{t.industryChart}</CardTitle>
                        {renderExcludedBadge(industryData)}
                      </CardHeader>
                      <CardContent className="h-full">
                        <div className="flex min-h-[300px] h-full items-center justify-center">
                          {industryData.chartData.length === 0 ? (
                            <div className="flex h-full w-full items-center justify-center text-sm text-zinc-500">{t.emptyData}</div>
                          ) : (
                            <div className="min-h-[300px] h-full w-full">
                              <ResponsiveContainer width="100%" height="100%">
                                <BarChart
                                  data={industryData.chartData}
                                  layout="vertical"
                                  margin={{ top: 16, right: 18, bottom: 16, left: 32 }}
                                >
                                  <defs>
                                    <linearGradient id="industryGradient" x1="0" y1="0" x2="1" y2="0">
                                      <stop offset="0%" stopColor="#1d4ed8" />
                                      <stop offset="100%" stopColor="#0f766e" />
                                    </linearGradient>
                                  </defs>
                                  <CartesianGrid vertical={false} strokeDasharray="3 3" stroke="#e4e4e7" />
                                  <XAxis
                                    type="number"
                                    allowDecimals={false}
                                    axisLine={false}
                                    tickLine={false}
                                    tick={{ fontSize: 12, fill: "#52525b" }}
                                  />
                                  <YAxis
                                    dataKey="name"
                                    type="category"
                                    width={170}
                                    axisLine={false}
                                    tickLine={false}
                                    tick={{ fontSize: 12, fill: "#52525b" }}
                                  />
                                  <Tooltip content={(props) => renderGlassTooltip({ ...props, total: industryData.totalCount })} />
                                  <Bar
                                    dataKey="value"
                                    radius={[0, 4, 4, 0]}
                                    fill="url(#industryGradient)"
                                    animationDuration={CHART_ANIMATION_MS}
                                    animationEasing="ease-out"
                                  />
                                </BarChart>
                              </ResponsiveContainer>
                            </div>
                          )}
                        </div>
                      </CardContent>
                    </Card>

                    <Card className="h-full border-zinc-200/80 bg-white/80 shadow-sm">
                      <CardHeader className="flex flex-row items-center justify-between gap-2 pb-2">
                        <CardTitle className="text-sm">{t.sentimentChart}</CardTitle>
                        {renderExcludedBadge(sentimentData)}
                      </CardHeader>
                      <CardContent className="h-full">
                        <div className="flex min-h-[300px] h-full items-center justify-center">
                          {sentimentData.chartData.length === 0 ? (
                            <div className="flex h-full w-full items-center justify-center text-sm text-zinc-500">{t.emptyData}</div>
                          ) : (
                            <div className="min-h-[300px] h-full w-full">
                              <div className="h-[250px] w-full">
                                <ResponsiveContainer width="100%" height="100%">
                                  <PieChart margin={{ top: 12, right: 16, bottom: 12, left: 16 }}>
                                    <Pie
                                      data={sentimentData.chartData}
                                      dataKey="value"
                                      nameKey="name"
                                      outerRadius={95}
                                      paddingAngle={2}
                                      isAnimationActive
                                      animationDuration={CHART_ANIMATION_MS}
                                      animationEasing="ease-out"
                                    >
                                      {sentimentData.chartData.map((item, index) => (
                                        <Cell
                                          key={`sentiment-${item.name}`}
                                          fill={insightPalette[index % insightPalette.length]}
                                        />
                                      ))}
                                    </Pie>
                                    <Tooltip
                                      content={(props) => renderGlassTooltip({ ...props, total: sentimentData.totalCount })}
                                    />
                                  </PieChart>
                                </ResponsiveContainer>
                              </div>
                              {renderChartLegend(sentimentData.chartData, insightPalette)}
                            </div>
                          )}
                        </div>
                      </CardContent>
                    </Card>
                  </div>

                  <Card className="mt-6 border-zinc-200/80 bg-white/80 shadow-sm">
                    <CardHeader className="pb-2">
                      <CardTitle className="text-sm">{t.insightLatestTableTitle}</CardTitle>
                      <CardDescription>
                        {t.lastUpdated}：{lastUpdated || "-"}
                      </CardDescription>
                    </CardHeader>
                    <CardContent>
                      <Table>
                        <TableHeader>
                          <TableRow>
                            <TableHead>{t.columnQuery}</TableHead>
                            <TableHead>{t.columnRole}</TableHead>
                            <TableHead>{t.columnIndustry}</TableHead>
                            <TableHead>{t.columnSentiment}</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {insightLatestRows.length === 0 ? (
                            <TableRow>
                              <TableCell colSpan={4} className="text-center text-zinc-500">
                                {t.emptyData}
                              </TableCell>
                            </TableRow>
                          ) : (
                            insightLatestRows.map((row, index) => (
                              <TableRow key={`${row.query}-${index}`}>
                                <TableCell className="max-w-[420px] truncate">{row.query}</TableCell>
                                <TableCell>{row.role}</TableCell>
                                <TableCell>{row.industry}</TableCell>
                                <TableCell>{row.sentiment}</TableCell>
                              </TableRow>
                            ))
                          )}
                        </TableBody>
                      </Table>
                    </CardContent>
                  </Card>
                </TabsContent>

                <TabsContent value="feedback">
                  <div className="grid gap-6 grid-cols-1 xl:grid-cols-2">
                    <Card className="h-full border-zinc-200/80 bg-white/80 shadow-sm xl:col-span-2">
                      <CardHeader className="flex flex-row items-center justify-between gap-2 pb-2">
                        <CardTitle className="text-sm">{t.frictionChart}</CardTitle>
                        {renderExcludedBadge(frictionData)}
                      </CardHeader>
                      <CardContent className="h-full">
                        <div className="flex min-h-[300px] h-full items-center justify-center">
                          {frictionData.chartData.length === 0 ? (
                            <div className="flex h-full w-full items-center justify-center text-sm text-zinc-500">{t.emptyData}</div>
                          ) : (
                            <div className="min-h-[300px] h-full w-full">
                              <ResponsiveContainer width="100%" height="100%">
                                <BarChart
                                  data={frictionData.chartData}
                                  layout="vertical"
                                  margin={{ top: 16, right: 18, bottom: 16, left: 32 }}
                                >
                                  <defs>
                                    <linearGradient id="frictionGradient" x1="0" y1="0" x2="1" y2="0">
                                      <stop offset="0%" stopColor="#1d3557" />
                                      <stop offset="100%" stopColor="#2a9d8f" />
                                    </linearGradient>
                                  </defs>
                                  <CartesianGrid vertical={false} strokeDasharray="3 3" stroke="#e4e4e7" />
                                  <XAxis
                                    type="number"
                                    allowDecimals={false}
                                    axisLine={false}
                                    tickLine={false}
                                    tick={{ fontSize: 12, fill: "#52525b" }}
                                  />
                                  <YAxis
                                    dataKey="name"
                                    type="category"
                                    width={180}
                                    axisLine={false}
                                    tickLine={false}
                                    tick={{ fontSize: 12, fill: "#52525b" }}
                                  />
                                  <Tooltip content={(props) => renderGlassTooltip({ ...props, total: frictionData.totalCount })} />
                                  <Bar
                                    dataKey="value"
                                    radius={[0, 4, 4, 0]}
                                    fill="url(#frictionGradient)"
                                    animationDuration={CHART_ANIMATION_MS}
                                    animationEasing="ease-out"
                                  />
                                </BarChart>
                              </ResponsiveContainer>
                            </div>
                          )}
                        </div>
                      </CardContent>
                    </Card>

                    <Card className="h-full border-zinc-200/80 bg-white/80 shadow-sm">
                      <CardHeader className="flex flex-row items-center justify-between gap-2 pb-2">
                        <CardTitle className="text-sm">{t.featureChart}</CardTitle>
                        {renderExcludedBadge(featureData)}
                      </CardHeader>
                      <CardContent className="h-full">
                        <div className="flex min-h-[300px] h-full items-center justify-center">
                          {featureData.chartData.length === 0 ? (
                            <div className="flex h-full w-full items-center justify-center text-sm text-zinc-500">{t.emptyData}</div>
                          ) : (
                            <div className="min-h-[300px] h-full w-full">
                              <div className="h-[250px] w-full">
                                <ResponsiveContainer width="100%" height="100%">
                                  <PieChart margin={{ top: 12, right: 16, bottom: 12, left: 16 }}>
                                    <Pie
                                      data={featureData.chartData}
                                      dataKey="value"
                                      nameKey="name"
                                      innerRadius={56}
                                      outerRadius={95}
                                      paddingAngle={2}
                                      isAnimationActive
                                      animationDuration={CHART_ANIMATION_MS}
                                      animationEasing="ease-out"
                                    >
                                      {featureData.chartData.map((item, index) => (
                                        <Cell key={`feature-${item.name}`} fill={feedbackPalette[index % feedbackPalette.length]} />
                                      ))}
                                    </Pie>
                                    <Tooltip content={(props) => renderGlassTooltip({ ...props, total: featureData.totalCount })} />
                                  </PieChart>
                                </ResponsiveContainer>
                              </div>
                              {renderChartLegend(featureData.chartData, feedbackPalette)}
                            </div>
                          )}
                        </div>
                      </CardContent>
                    </Card>

                    <Card className="h-full border-zinc-200/80 bg-white/80 shadow-sm">
                      <CardHeader className="flex flex-row items-center justify-between gap-2 pb-2">
                        <CardTitle className="text-sm">{t.safetyChart}</CardTitle>
                        {renderExcludedBadge(safetyData)}
                      </CardHeader>
                      <CardContent className="h-full">
                        <div className="flex min-h-[300px] h-full items-center justify-center">
                          {safetyData.chartData.length === 0 ? (
                            <div className="flex h-full w-full items-center justify-center text-sm text-zinc-500">{t.emptyData}</div>
                          ) : (
                            <div className="min-h-[300px] h-full w-full">
                              <div className="h-[250px] w-full">
                                <ResponsiveContainer width="100%" height="100%">
                                  <PieChart margin={{ top: 12, right: 16, bottom: 12, left: 16 }}>
                                    <Pie
                                      data={safetyData.chartData}
                                      dataKey="value"
                                      nameKey="name"
                                      outerRadius={95}
                                      paddingAngle={2}
                                      isAnimationActive
                                      animationDuration={CHART_ANIMATION_MS}
                                      animationEasing="ease-out"
                                    >
                                      {safetyData.chartData.map((item, index) => (
                                        <Cell key={`safety-${item.name}`} fill={feedbackPalette[index % feedbackPalette.length]} />
                                      ))}
                                    </Pie>
                                    <Tooltip content={(props) => renderGlassTooltip({ ...props, total: safetyData.totalCount })} />
                                  </PieChart>
                                </ResponsiveContainer>
                              </div>
                              {renderChartLegend(safetyData.chartData, feedbackPalette)}
                            </div>
                          )}
                        </div>
                      </CardContent>
                    </Card>
                  </div>

                  <Card className="mt-6 border-zinc-200/80 bg-white/80 shadow-sm">
                    <CardHeader className="pb-2">
                      <CardTitle className="text-sm">{t.feedbackLatestTableTitle}</CardTitle>
                      <CardDescription>
                        {t.lastUpdated}：{lastUpdated || "-"}
                      </CardDescription>
                    </CardHeader>
                    <CardContent>
                      <Table>
                        <TableHeader>
                          <TableRow>
                            <TableHead>{t.columnQuery}</TableHead>
                            <TableHead>{t.columnFrictionSignal}</TableHead>
                            <TableHead>{t.columnFeatureRequest}</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {feedbackLatestRows.length === 0 ? (
                            <TableRow>
                              <TableCell colSpan={3} className="text-center text-zinc-500">
                                {t.emptyData}
                              </TableCell>
                            </TableRow>
                          ) : (
                            feedbackLatestRows.map((row, index) => (
                              <TableRow key={`${row.query}-${index}`}>
                                <TableCell className="max-w-[460px] truncate">{row.query}</TableCell>
                                <TableCell>{row.frictionSignal}</TableCell>
                                <TableCell>{row.featureRequest}</TableCell>
                              </TableRow>
                            ))
                          )}
                        </TableBody>
                      </Table>
                    </CardContent>
                  </Card>
                </TabsContent>
              </Tabs>
            </CardContent>
          </Card>
        </section>
      </main>
    </div>
  );
}
