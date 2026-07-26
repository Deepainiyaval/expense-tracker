import React, { useState, useMemo, useEffect, useRef } from "react";
import { supabase } from "./lib/supabase";
import {
  LineChart, Line, AreaChart, Area, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend,
} from "recharts";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, List, ArrowDownCircle, ArrowUpCircle, Target, BarChart3,
  FileText, PiggyBank, Bell, Settings as SettingsIcon, Plus, Search, X,
  Edit2, Trash2, Sun, Moon, Monitor, Mail, Lock, User, Eye, EyeOff,
  Menu, Download, AlertTriangle, CheckCircle2, Info, Sparkles, LogOut,
  Utensils, ShoppingBag, Users, PartyPopper, Compass, Plane, Home, Receipt,
  Fuel, Stethoscope, GraduationCap, Film, RefreshCw, MoreHorizontal,
  Briefcase, Wallet, TrendingUp, TrendingDown, ChevronRight, Feather,
} from "lucide-react";

/* ============================================================
   BAHI — a ledger for a household's money
   Design language: the Indian bahi-khata (ledger book). Warm
   paper, ink typography, a brass/forest/rust palette tied to
   ink-and-brass ledger stationery, tabular mono numerals for
   every amount, and a stitched left margin on every card —
   the binding of a bound account book.
   ============================================================ */

const PALETTE = ["#1B4D3E", "#2F7A5F", "#B8923F", "#D4AF6A", "#B5502D", "#C96A46", "#3E6B8C", "#6B7F3E"];

const ICONS = {
  Utensils, ShoppingBag, Users, PartyPopper, Compass, Plane, Home, Receipt,
  Fuel, Stethoscope, GraduationCap, Film, RefreshCw, MoreHorizontal,
  Briefcase, Wallet, TrendingUp,
};

const DEFAULT_EXPENSE_CATEGORIES = [
  { name: "Food", icon: "Utensils" },
  { name: "Shopping", icon: "ShoppingBag" },
  { name: "Family", icon: "Users" },
  { name: "Friends", icon: "PartyPopper" },
  { name: "Outing", icon: "Compass" },
  { name: "Travel", icon: "Plane" },
  { name: "Rent", icon: "Home" },
  { name: "Bills", icon: "Receipt" },
  { name: "Fuel", icon: "Fuel" },
  { name: "Medical", icon: "Stethoscope" },
  { name: "Education", icon: "GraduationCap" },
  { name: "Entertainment", icon: "Film" },
  { name: "Subscriptions", icon: "RefreshCw" },
  { name: "Others", icon: "MoreHorizontal" },
].map((c, i) => ({ ...c, color: PALETTE[i % PALETTE.length], custom: false }));

const DEFAULT_INCOME_CATEGORIES = [
  { name: "My Salary", icon: "Briefcase", color: PALETTE[0], custom: false },
  { name: "Brother Income", icon: "Wallet", color: PALETTE[2], custom: false },
];

/* ---------------- seed data ---------------- */
function mulberry32(seed) {
  return function () {
    seed |= 0; seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function monthKey(d) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; }
function monthLabel(key) {
  const [y, m] = key.split("-").map(Number);
  return new Date(y, m - 1, 1).toLocaleDateString("en-IN", { month: "short", year: "numeric" });
}
function fmtDate(iso) {
  return new Date(iso).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
}
function fmtINR(n) {
  const sign = n < 0 ? "-" : "";
  return sign + "₹" + Math.abs(Math.round(n)).toLocaleString("en-IN");
}
function uid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4); }

function genSeedData() {
  const rand = mulberry32(42);
  const today = new Date();
  const txns = [];
  const expenseWeights = {
    Food: 0.9, Shopping: 0.7, Family: 0.4, Friends: 0.35, Outing: 0.4,
    Travel: 0.2, Rent: 1, Bills: 0.8, Fuel: 0.55, Medical: 0.25,
    Education: 0.2, Entertainment: 0.45, Subscriptions: 0.5, Others: 0.4,
  };
  const expenseRange = {
    Food: [120, 650], Shopping: [300, 4200], Family: [200, 3000], Friends: [150, 1800],
    Outing: [200, 2200], Travel: [800, 9000], Rent: [12000, 12000], Bills: [600, 3200],
    Fuel: [300, 1600], Medical: [150, 3500], Education: [500, 5000], Entertainment: [150, 1200],
    Subscriptions: [99, 799], Others: [100, 1500],
  };

  for (let m = 5; m >= 0; m--) {
    const monthDate = new Date(today.getFullYear(), today.getMonth() - m, 1);
    const daysInMonth = new Date(monthDate.getFullYear(), monthDate.getMonth() + 1, 0).getDate();
    const lastDay = m === 0 ? Math.min(today.getDate(), daysInMonth) : daysInMonth;

    // income
    const salary = 42000 + Math.floor(rand() * 6000);
    txns.push({ id: uid(), type: "income", amount: salary, category: "My Salary", date: new Date(monthDate.getFullYear(), monthDate.getMonth(), 1).toISOString().slice(0, 10), notes: "Monthly salary credit" });
    if (rand() > 0.3) {
      const bro = 4000 + Math.floor(rand() * 5000);
      txns.push({ id: uid(), type: "income", amount: bro, category: "Brother Income", date: new Date(monthDate.getFullYear(), monthDate.getMonth(), 3 + Math.floor(rand() * 5)).toISOString().slice(0, 10), notes: "" });
    }

    // rent, fixed
    txns.push({ id: uid(), type: "expense", amount: expenseRange.Rent[0], category: "Rent", date: new Date(monthDate.getFullYear(), monthDate.getMonth(), 5).toISOString().slice(0, 10), notes: "Monthly rent" });

    Object.keys(expenseWeights).forEach((cat) => {
      if (cat === "Rent") return;
      const density = expenseWeights[cat];
      const [lo, hi] = expenseRange[cat];
      const count = Math.round(density * (2 + rand() * 5));
      for (let i = 0; i < count; i++) {
        const day = 1 + Math.floor(rand() * lastDay);
        if (day > lastDay) continue;
        const weekendBoost = [0, 6].includes(new Date(monthDate.getFullYear(), monthDate.getMonth(), day).getDay()) ? 1.35 : 1;
        const amount = Math.round((lo + rand() * (hi - lo)) * weekendBoost);
        txns.push({
          id: uid(), type: "expense", amount, category: cat,
          date: new Date(monthDate.getFullYear(), monthDate.getMonth(), day).toISOString().slice(0, 10),
          notes: "",
        });
      }
    });
  }
  return txns.sort((a, b) => new Date(b.date) - new Date(a.date));
}

/* ---------------- insight + notification engine ---------------- */
function buildInsights(currentTx, prevTx, budgets, catColor) {
  const insights = [];
  const sum = (arr) => arr.reduce((s, t) => s + t.amount, 0);
  const income = sum(currentTx.filter((t) => t.type === "income"));
  const expense = sum(currentTx.filter((t) => t.type === "expense"));
  const prevIncome = sum(prevTx.filter((t) => t.type === "income"));
  const prevExpense = sum(prevTx.filter((t) => t.type === "expense"));

  const byCat = {}; currentTx.filter((t) => t.type === "expense").forEach((t) => byCat[t.category] = (byCat[t.category] || 0) + t.amount);
  const prevByCat = {}; prevTx.filter((t) => t.type === "expense").forEach((t) => prevByCat[t.category] = (prevByCat[t.category] || 0) + t.amount);
  const sortedCats = Object.entries(byCat).sort((a, b) => b[1] - a[1]);

  if (sortedCats.length) {
    const [topCat, topAmt] = sortedCats[0];
    const pct = expense ? ((topAmt / expense) * 100).toFixed(0) : 0;
    insights.push({ type: "info", text: `${topCat} is your largest expense — ${fmtINR(topAmt)}, about ${pct}% of this month's spending.` });
    if (income > 0) {
      const pctOfIncome = ((topAmt / income) * 100).toFixed(0);
      if (pctOfIncome >= 20) insights.push({ type: "warning", text: `You've put roughly ${pctOfIncome}% of your income into ${topCat} this month.` });
    }
    const saveEstimate = Math.round(topAmt * 0.15);
    if (saveEstimate > 200) insights.push({ type: "info", text: `Trimming ${topCat} by 15% would free up about ${fmtINR(saveEstimate)} this month.` });
  }

  sortedCats.forEach(([cat, amt]) => {
    const prev = prevByCat[cat] || 0;
    if (prev > 300 && amt > prev * 1.15) {
      const pct = (((amt - prev) / prev) * 100).toFixed(0);
      insights.push({ type: "warning", text: `${cat} spending rose ${pct}% compared to last month.` });
    } else if (prev > 300 && amt < prev * 0.8) {
      const pct = (((prev - amt) / prev) * 100).toFixed(0);
      insights.push({ type: "positive", text: `Nice — ${cat} spending dropped ${pct}% from last month.` });
    }
  });

  const weekend = currentTx.filter((t) => t.type === "expense" && [0, 6].includes(new Date(t.date).getDay()));
  const weekday = currentTx.filter((t) => t.type === "expense" && ![0, 6].includes(new Date(t.date).getDay()));
  const weekendAvg = weekend.length ? sum(weekend) / weekend.length : 0;
  const weekdayAvg = weekday.length ? sum(weekday) / weekday.length : 0;
  if (weekendAvg > weekdayAvg * 1.3 && weekend.length >= 3) {
    insights.push({ type: "info", text: `Weekend spending runs higher — average ${fmtINR(weekendAvg)} per entry versus ${fmtINR(weekdayAvg)} on weekdays.` });
  }

  const savingsRate = income ? (((income - expense) / income) * 100) : 0;
  if (income > 0) {
    if (savingsRate >= 30) insights.push({ type: "positive", text: `Your savings rate is excellent at ${savingsRate.toFixed(0)}% this month.` });
    else if (savingsRate >= 10) insights.push({ type: "info", text: `Your savings rate stands at ${savingsRate.toFixed(0)}% so far this month.` });
    else if (savingsRate >= 0) insights.push({ type: "warning", text: `Savings rate is thin this month, only ${savingsRate.toFixed(0)}%. Consider trimming a discretionary category.` });
    else insights.push({ type: "warning", text: `Spending has outpaced income this month by ${fmtINR(Math.abs(income - expense))}.` });
  }

  if (prevExpense > 0) {
    const days = new Date().getDate();
    const prevDays = new Date(new Date().getFullYear(), new Date().getMonth(), 0).getDate();
    const avgNow = expense / Math.max(days, 1);
    const avgPrev = prevExpense / prevDays;
    if (avgNow > avgPrev * 1.1) insights.push({ type: "warning", text: `Average daily expense is up to ${fmtINR(avgNow)}, compared with ${fmtINR(avgPrev)} last month.` });
  }

  const biggest = [...currentTx].filter((t) => t.type === "expense").sort((a, b) => b.amount - a.amount)[0];
  if (biggest) insights.push({ type: "info", text: `Largest single expense this month: ${fmtINR(biggest.amount)} on ${biggest.category}, ${fmtDate(biggest.date)}.` });

  budgets.forEach((b) => {
    const used = byCat[b.category] || 0;
    const pct = b.amount ? (used / b.amount) * 100 : 0;
    if (pct >= 100) insights.push({ type: "warning", text: `${b.category} budget exceeded — ${fmtINR(used)} spent against a ${fmtINR(b.amount)} budget.` });
    else if (pct >= 90) insights.push({ type: "warning", text: `${b.category} budget is at ${pct.toFixed(0)}%, close to its limit.` });
  });

  if (income > 0 && prevIncome > 0) {
    const growth = (((income - prevIncome) / prevIncome) * 100).toFixed(0);
    if (Math.abs(growth) >= 5) insights.push({ type: growth > 0 ? "positive" : "warning", text: `Income is ${growth > 0 ? "up" : "down"} ${Math.abs(growth)}% versus last month.` });
  }

  insights.push({ type: "info", text: `${currentTx.length} transactions logged so far this month.` });

  return insights.slice(0, 12);
}

/* ============================================================
   ROOT APP
   ============================================================ */
export default function App() {
 const [session, setSession] = useState(null);
const [loadingSession, setLoadingSession] = useState(true);
const user = session?.user;
  const [authMode, setAuthMode] = useState("login");
  const [theme, setTheme] = useState("system");
  const [systemDark, setSystemDark] = useState(false);
  const [view, setView] = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(false);

  const [transactions, setTransactions] = useState([]);
const [loadingTransactions, setLoadingTransactions] = useState(true);
  const [incomeCategories, setIncomeCategories] = useState(DEFAULT_INCOME_CATEGORIES);
  const [expenseCategories, setExpenseCategories] = useState(DEFAULT_EXPENSE_CATEGORIES);
  const [budgets, setBudgets] = useState([]);
const [loadingBudgets, setLoadingBudgets] = useState(true);
  const [txModal, setTxModal] = useState(null); // {mode:'add'|'edit', type, data}
  const [budgetModal, setBudgetModal] = useState(null);
  const [catModal, setCatModal] = useState(false);
  const [settings, setSettings] = useState({
    currency: "INR", notifBudget: true, notifDaily: true, notifReport: true,
  });
  useEffect(() => {
  const loadSession = async () => {
    const {
      data: { session },
    } = await supabase.auth.getSession();

    setSession(session);
    setLoadingSession(false);
  };
   loadSession();

  const {
    data: { subscription },
  } = supabase.auth.onAuthStateChange((_event, session) => {
    setSession(session);
  });

  return () => {
    subscription.unsubscribe();
  };
}, []);
const loadTransactions = async () => {
  if (!user) return;

  setLoadingTransactions(true);

  const { data, error } = await supabase
    .from("transactions")
    .select("*")
    .order("date", { ascending: false });

  if (error) {
    console.error(error);
  } else {
    setTransactions(data);
  }

  setLoadingTransactions(false);
};
useEffect(() => {
  if (user) {
    loadTransactions();
  }
}, [user]);
const loadBudgets = async () => {
  if (!user) return;

  setLoadingBudgets(true);

  const { data, error } = await supabase
    .from("budgets")
    .select("*")
    .order("category");

  if (error) {
    console.error(error);
  } else {
   setBudgets(
  data.map((b) => ({
    ...b,
    amount: b.monthly_limit,
  }))
);
  }

  setLoadingBudgets(false);
};
useEffect(() => {
  if (user) {
    loadBudgets();
  }
}, [user]);
async function addBudget(budget) {
  if (!user) return;

  const { error } = await supabase
    .from("budgets")
    .insert([
      {
        user_id: user.id,
        category: budget.category,
        monthly_limit: budget.amount,
      },
    ]);

  if (error) {
    console.error(error);
    alert(error.message);
    return;
  }

  await loadBudgets();
}
async function updateBudget(budget) {
  const { error } = await supabase
    .from("budgets")
    .update({
      category: budget.category,
      monthly_limit: budget.amount,
    })
    .eq("id", budget.id);

  if (error) {
    console.error(error);
    alert(error.message);
    return;
  }

  await loadBudgets();
}
async function deleteBudget(id) {
  const { error } = await supabase
    .from("budgets")
    .delete()
    .eq("id", id);

  if (error) {
    console.error(error);
    alert(error.message);
    return;
  }

  await loadBudgets();
}
  useEffect(() => {
    const mq = window.matchMedia("(prefers-color-scheme: dark)");
    setSystemDark(mq.matches);
    const fn = (e) => setSystemDark(e.matches);
    mq.addEventListener?.("change", fn);
    return () => mq.removeEventListener?.("change", fn);
  }, []);
  const isDark = theme === "dark" || (theme === "system" && systemDark);

  const now = new Date();
  const curKey = monthKey(now);
  const prevDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const prevKey = monthKey(prevDate);

  const currentMonthTx = useMemo(() => transactions.filter((t) => monthKey(new Date(t.date)) === curKey), [transactions, curKey]);
  const prevMonthTx = useMemo(() => transactions.filter((t) => monthKey(new Date(t.date)) === prevKey), [transactions, prevKey]);

  const allCategories = [...incomeCategories, ...expenseCategories];
  const catColor = (name) => allCategories.find((c) => c.name === name)?.color || PALETTE[0];
  const catIcon = (name) => ICONS[allCategories.find((c) => c.name === name)?.icon] || MoreHorizontal;

  const monthlyAgg = useMemo(() => {
    const map = {};
    transactions.forEach((t) => {
      const k = monthKey(new Date(t.date));
      if (!map[k]) map[k] = { key: k, income: 0, expense: 0 };
      map[k][t.type] += t.amount;
    });
    return Object.values(map).sort((a, b) => a.key.localeCompare(b.key)).map((m) => ({
      ...m, label: monthLabel(m.key), savings: m.income - m.expense,
    }));
  }, [transactions]);

  const insights = useMemo(() => buildInsights(currentMonthTx, prevMonthTx, budgets, catColor), [currentMonthTx, prevMonthTx, budgets]);

  const budgetStatus = useMemo(() => budgets.map((b) => {
    const used = currentMonthTx.filter((t) => t.type === "expense" && t.category === b.category).reduce((s, t) => s + t.amount, 0);
    const pct = b.amount ? Math.min((used / b.amount) * 100, 999) : 0;
    const status = pct >= 100 ? "red" : pct >= 75 ? "yellow" : "green";
    return { ...b, used, pct, status };
  }), [budgets, currentMonthTx]);

  const notifications = useMemo(() => {
    const list = [];
    const today = new Date().toISOString().slice(0, 10);
    const hasToday = transactions.some((t) => t.date === today);
    if (settings.notifDaily && !hasToday) list.push({ type: "info", text: "Add today's expense — nothing logged yet today.", icon: Feather });
    budgetStatus.forEach((b) => {
      if (!settings.notifBudget) return;
      if (b.status === "red") list.push({ type: "warning", text: `${b.category} budget exceeded (${b.pct.toFixed(0)}%).`, icon: AlertTriangle });
      else if (b.status === "yellow" && b.pct >= 90) list.push({ type: "warning", text: `${b.category} budget reached ${b.pct.toFixed(0)}%.`, icon: AlertTriangle });
    });
    const daysLeft = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate() - now.getDate();
    if (settings.notifReport && daysLeft <= 2) list.push({ type: "positive", text: "End of month report is ready to review.", icon: CheckCircle2 });
    return list;
  }, [transactions, budgetStatus, settings]);

  async function addTransaction(tx) {
  if (!user) return;

  const { error } = await supabase
    .from("transactions")
    .insert([
      {
        user_id: user.id,
        type: tx.type,
        amount: tx.amount,
        category: tx.category,
        date: tx.date,
        notes: tx.notes,
      },
    ]);

  if (error) {
    console.error(error);
    alert(error.message);
    return;
  }

  await loadTransactions();
}
  async function updateTransaction(tx) {
  const { error } = await supabase
    .from("transactions")
    .update({
      type: tx.type,
      amount: tx.amount,
      category: tx.category,
      date: tx.date,
      notes: tx.notes,
    })
    .eq("id", tx.id);

  if (error) {
    console.error(error);
    alert(error.message);
    return;
  }

  await loadTransactions();
}
  async function deleteTransaction(id) {
  const { error } = await supabase
    .from("transactions")
    .delete()
    .eq("id", id);

  if (error) {
    console.error(error);
    alert(error.message);
    return;
  }

  await loadTransactions();
}

  function exportExcel(monthKeyToExport) {
    const txForMonth = transactions.filter((t) => monthKey(new Date(t.date)) === monthKeyToExport);
    const income = txForMonth.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
    const expense = txForMonth.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
    const byCat = {};
    txForMonth.filter((t) => t.type === "expense").forEach((t) => byCat[t.category] = (byCat[t.category] || 0) + t.amount);

    const wb = XLSX.utils.book_new();
    const txSheet = XLSX.utils.json_to_sheet(txForMonth.map((t) => ({
      Date: t.date, Type: t.type, Category: t.category, Amount: t.amount, Notes: t.notes || "",
    })));
    XLSX.utils.book_append_sheet(wb, txSheet, "Transactions");

    const catSheet = XLSX.utils.json_to_sheet(Object.entries(byCat).map(([Category, Amount]) => ({
      Category, Amount, "% of Expense": expense ? ((Amount / expense) * 100).toFixed(1) + "%" : "0%",
    })));
    XLSX.utils.book_append_sheet(wb, catSheet, "Category Summary");

    const summarySheet = XLSX.utils.json_to_sheet([{
      Month: monthLabel(monthKeyToExport), "Total Income": income, "Total Expense": expense,
      "Remaining Balance": income - expense, Savings: income - expense,
      "Savings Rate": income ? (((income - expense) / income) * 100).toFixed(1) + "%" : "0%",
      "Transactions": txForMonth.length,
    }]);
    XLSX.utils.book_append_sheet(wb, summarySheet, "Monthly Summary");

    const savingsSheet = XLSX.utils.json_to_sheet(monthlyAgg.map((m) => ({ Month: m.label, Income: m.income, Expense: m.expense, Savings: m.savings })));
    XLSX.utils.book_append_sheet(wb, savingsSheet, "Savings");

    XLSX.writeFile(wb, `Bahi_Report_${monthKeyToExport}.xlsx`);
  }
if (loadingSession) {
  return <div style={{ padding: 40 }}>Loading...</div>;
}
if (loadingTransactions && session) {
  return <div style={{ padding: 40 }}>Loading transactions...</div>;
}
 if (!session) {
  return (
    <LoginScreen
      isDark={isDark}
      authMode={authMode}
      setAuthMode={setAuthMode}
      onLogin={() => {}}
    />
  );
}

  const NAV = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "transactions", label: "Transactions", icon: List },
    { id: "income", label: "Income", icon: ArrowUpCircle },
    { id: "expenses", label: "Expenses", icon: ArrowDownCircle },
    { id: "budgets", label: "Budgets", icon: Target },
    { id: "analytics", label: "Analytics", icon: BarChart3 },
    { id: "reports", label: "Monthly Reports", icon: FileText },
    { id: "savings", label: "Savings", icon: PiggyBank },
    { id: "notifications", label: "Notifications", icon: Bell, badge: notifications.length },
    { id: "settings", label: "Settings", icon: SettingsIcon },
  ];

  return (
    <div className={`bahi ${isDark ? "dark" : ""}`}>
      <style>{CSS}</style>
      <div className="shell">
        {sidebarOpen && <div className="scrim" onClick={() => setSidebarOpen(false)} />}
        <aside className={`sidebar ${sidebarOpen ? "open" : ""}`}>
          <div className="brand">
            <div className="brand-mark"><Feather size={18} /></div>
            <div>
              <div className="brand-name">Bahi</div>
              <div className="brand-sub">the household ledger</div>
            </div>
          </div>
          <nav className="nav">
            {NAV.map((n) => (
              <button key={n.id} className={`nav-item ${view === n.id ? "active" : ""}`} onClick={() => { setView(n.id); setSidebarOpen(false); }}>
                <n.icon size={17} strokeWidth={2} />
                <span>{n.label}</span>
                {!!n.badge && <span className="nav-badge">{n.badge}</span>}
              </button>
            ))}
          </nav>
          <button
  className="nav-item logout"
  onClick={async () => {
    await supabase.auth.signOut();
  }}
>
            <LogOut size={17} /> <span>Sign out</span>
          </button>
        </aside>

        <div className="main">
          <header className="topbar">
            <button className="icon-btn hamburger" onClick={() => setSidebarOpen(true)}><Menu size={20} /></button>
            <div className="topbar-title">{NAV.find((n) => n.id === view)?.label}</div>
            <div className="topbar-right">
              <ThemeToggle theme={theme} setTheme={setTheme} />
              <div className="avatar">A</div>
            </div>
          </header>

          <main className="content">
            {view === "dashboard" && (
              <Dashboard
                currentMonthTx={currentMonthTx} prevMonthTx={prevMonthTx} monthlyAgg={monthlyAgg}
                catColor={catColor} catIcon={catIcon} insights={insights}
                onAdd={(type) => setTxModal({ mode: "add", type })}
              />
            )}
            {(view === "transactions" || view === "income" || view === "expenses") && (
              <TransactionsView
                filterType={view === "income" ? "income" : view === "expenses" ? "expense" : "all"}
                transactions={transactions} catColor={catColor} catIcon={catIcon}
                incomeCategories={incomeCategories} expenseCategories={expenseCategories}
                onAdd={(type) => setTxModal({ mode: "add", type })}
                onEdit={(t) => setTxModal({ mode: "edit", type: t.type, data: t })}
                onDelete={deleteTransaction}
              />
            )}
            {view === "budgets" && (
              <BudgetsView budgetStatus={budgetStatus} catColor={catColor} catIcon={catIcon}
                expenseCategories={expenseCategories}
                onAdd={() => setBudgetModal({ mode: "add" })}
                onEdit={(b) => setBudgetModal({ mode: "edit", data: b })}
                onDelete={deleteBudget}
              />
            )}
            {view === "analytics" && (
              <AnalyticsView monthlyAgg={monthlyAgg} currentMonthTx={currentMonthTx} catColor={catColor} isDark={isDark} />
            )}
            {view === "reports" && (
              <ReportsView transactions={transactions} monthlyAgg={monthlyAgg} catColor={catColor} onExport={exportExcel} />
            )}
            {view === "savings" && <SavingsView monthlyAgg={monthlyAgg} isDark={isDark} />}
            {view === "notifications" && <NotificationsView notifications={notifications} />}
            {view === "settings" && (
              <SettingsView theme={theme} setTheme={setTheme} settings={settings} setSettings={setSettings}
                incomeCategories={incomeCategories} expenseCategories={expenseCategories}
                setIncomeCategories={setIncomeCategories} setExpenseCategories={setExpenseCategories}
              />
            )}
          </main>
        </div>
      </div>

      {txModal && (
       <TransactionModal
  modal={txModal}
  onClose={() => setTxModal(null)}
  incomeCategories={incomeCategories}
  expenseCategories={expenseCategories}
  onSave={async (tx) => {
  if (txModal.mode === "add") {
    await addTransaction(tx);
  } else {
    await updateTransaction(tx);
  }

  await loadTransactions();

  setTxModal(null);
}}
/>
      )}
      {budgetModal && (
        <BudgetModal
  modal={budgetModal}
  onClose={() => setBudgetModal(null)}
  expenseCategories={expenseCategories}
  onSave={async (b) => {
    if (budgetModal.mode === "add") {
      await addBudget(b);
    } else {
      await updateBudget(b);
    }

    setBudgetModal(null);
  }}
/>
      )}
    </div>
  );
}

/* ============================================================
   LOGIN
   ============================================================ */
function LoginScreen({ isDark, authMode, setAuthMode, onLogin }) {
  const [showPw, setShowPw] = useState(false);
  const [fullName, setFullName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const handleAuth = async () => {
  setLoading(true);

  try {
    if (authMode === "register") {
      const { data, error } = await supabase.auth.signUp({
        email,
        password,
        options: {
          data: {
            full_name: fullName,
          },
        },
      });

      if (error) {
        alert(error.message);
        return;
      }

      alert("Registration successful! Please check your email to verify your account.");

      console.log(data);

      setAuthMode("login");
      return;
    }
if (authMode === "login") {
  const { data, error } = await supabase.auth.signInWithPassword({
    email,
    password,
  });

  if (error) {
    alert(error.message);
    return;
  }

  console.log("Logged in:", data);

  //onLogin();
  return;
}
    if (authMode === "forgot") {
      alert("We'll implement forgot password later.");
      return;
    }
  } finally {
    setLoading(false);
  }
};
  return (
    <div className={`bahi ${isDark ? "dark" : ""}`}>
      <style>{CSS}</style>
      <div className="auth-wrap">
        <div className="auth-side">
          <div className="brand" style={{ marginBottom: 28 }}>
            <div className="brand-mark"><Feather size={20} /></div>
            <div><div className="brand-name" style={{ fontSize: 22 }}>Bahi</div><div className="brand-sub">the household ledger</div></div>
          </div>
          <h1 className="auth-headline">Every rupee,<br />entered by hand.</h1>
          <p className="auth-copy">Bahi keeps income, spending, budgets and savings in one bound
            book for your household — five accounts, each private to the person keeping it.</p>
          <div className="auth-stitch" />
          <p className="auth-demo-note">Demo mode: sign-in is illustrative here. Wire this screen to
            Supabase Auth to make it real — see the schema notes shared alongside this file.</p>
        </div>
        <div className="auth-card">
          <div className="auth-tabs">
            <button className={authMode === "login" ? "active" : ""} onClick={() => setAuthMode("login")}>Sign in</button>
            <button className={authMode === "register" ? "active" : ""} onClick={() => setAuthMode("register")}>Register</button>
            <button className={authMode === "forgot" ? "active" : ""} onClick={() => setAuthMode("forgot")}>Forgot password</button>
          </div>

          <form onSubmit={(e) => { e.preventDefault(); handleAuth(); }} className="auth-form">
            {authMode === "register" && (
              <label className="field">
                <span>Full name</span>
                <div className="input-wrap"><User size={15} /><input
                  value={fullName}
                  onChange={(e) => setFullName(e.target.value)}
                  placeholder="Anita Sharma"
                  required
                /></div>
              </label>
            )}
            <label className="field">
              <span>Email</span>
              <div className="input-wrap"><Mail size={15} /><input
  type="email"
  value={email}
  onChange={(e) => setEmail(e.target.value)}
  placeholder="you@example.com"
  required
/></div>
            </label>
            {authMode !== "forgot" && (
              <label className="field">
                <span>Password</span>
                <div className="input-wrap">
                  <Lock size={15} />
                 <input
  type={showPw ? "text" : "password"}
  value={password}
  onChange={(e) => setPassword(e.target.value)}
  placeholder="••••••••"
  required
/>
                  <button type="button" className="ghost-icon" onClick={() => setShowPw((s) => !s)}>
                    {showPw ? <EyeOff size={15} /> : <Eye size={15} />}
                  </button>
                </div>
              </label>
            )}
            {authMode === "login" && (
              <div className="row-between">
                <label className="remember"><input type="checkbox" defaultChecked /> Remember me</label>
                <button type="button" className="link" onClick={() => setAuthMode("forgot")}>Forgot password?</button>
              </div>
            )}
            <button type="submit"  disabled={loading} className="btn-primary full">
              {loading
  ? "Please wait..."
  : authMode === "login"
  ? "Sign in"
  : authMode === "register"
  ? "Create account"
  : "Send reset link"}
            </button>
          </form>
          {authMode === "login" && <p className="auth-switch">New here? <button className="link" onClick={() => setAuthMode("register")}>Create an account</button></p>}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   SHARED BITS
   ============================================================ */
function ThemeToggle({ theme, setTheme }) {
  const opts = [{ id: "light", icon: Sun }, { id: "system", icon: Monitor }, { id: "dark", icon: Moon }];
  return (
    <div className="theme-toggle">
      {opts.map((o) => (
        <button key={o.id} className={theme === o.id ? "active" : ""} onClick={() => setTheme(o.id)} title={o.id}>
          <o.icon size={14} />
        </button>
      ))}
    </div>
  );
}

function StatCard({ label, value, sub, icon: Icon, tone = "ink" }) {
  return (
    <div className="stat-card">
      <div className="stitch" />
      <div className="stat-top">
        <span className="stat-label">{label}</span>
        {Icon && <Icon size={16} className={`stat-icon tone-${tone}`} />}
      </div>
      <div className={`stat-value tone-${tone}`}>{value}</div>
      {sub && <div className="stat-sub">{sub}</div>}
    </div>
  );
}

function Pill({ tone, children }) {
  return <span className={`pill pill-${tone}`}>{children}</span>;
}

function InsightRow({ insight }) {
  const Icon = insight.type === "positive" ? CheckCircle2 : insight.type === "warning" ? AlertTriangle : Info;
  return (
    <div className={`insight-row tone-${insight.type}`}>
      <Icon size={15} />
      <span>{insight.text}</span>
    </div>
  );
}

/* ============================================================
   DASHBOARD
   ============================================================ */
function Dashboard({ currentMonthTx, prevMonthTx, monthlyAgg, catColor, catIcon, insights, onAdd }) {
  const income = currentMonthTx.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const expense = currentMonthTx.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  const balance = income - expense;
  const savingsRate = income ? ((balance / income) * 100).toFixed(0) : 0;

  const today = new Date();
  const dayOfMonth = today.getDate();
  const avgDaily = expense / dayOfMonth;

  const weekAgo = new Date(); weekAgo.setDate(today.getDate() - 7);
  const weeklyExpense = currentMonthTx.filter((t) => t.type === "expense" && new Date(t.date) >= weekAgo).reduce((s, t) => s + t.amount, 0);

  const byCat = {};
  currentMonthTx.filter((t) => t.type === "expense").forEach((t) => byCat[t.category] = (byCat[t.category] || 0) + t.amount);
  const topCat = Object.entries(byCat).sort((a, b) => b[1] - a[1])[0];
  const highestExpense = [...currentMonthTx].filter((t) => t.type === "expense").sort((a, b) => b.amount - a.amount)[0];

  const recent = [...currentMonthTx].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 6);
  const trend = monthlyAgg.slice(-6);

  return (
    <div className="stack-lg">
      <div className="page-head">
        <div>
          <h2 className="page-title">This month, at a glance</h2>
          <p className="page-sub">{today.toLocaleDateString("en-IN", { month: "long", year: "numeric" })}</p>
        </div>
        <div className="head-actions">
          <button className="btn-secondary" onClick={() => onAdd("income")}><ArrowUpCircle size={15} /> Add income</button>
          <button className="btn-primary" onClick={() => onAdd("expense")}><Plus size={15} /> Add expense</button>
        </div>
      </div>

      <div className="grid-stats">
        <StatCard label="Income this month" value={fmtINR(income)} icon={ArrowUpCircle} tone="forest" />
        <StatCard label="Expense this month" value={fmtINR(expense)} icon={ArrowDownCircle} tone="rust" />
        <StatCard label="Current balance" value={fmtINR(balance)} icon={Wallet} tone={balance >= 0 ? "forest" : "rust"} />
        <StatCard label="Savings rate" value={`${savingsRate}%`} icon={PiggyBank} tone="brass" />
        <StatCard label="Spent till date" value={fmtINR(expense)} sub={`${dayOfMonth} days in`} icon={Receipt} />
        <StatCard label="Weekly expense" value={fmtINR(weeklyExpense)} sub="last 7 days" icon={Compass} />
        <StatCard label="Average daily expense" value={fmtINR(avgDaily)} icon={BarChart3} />
        <StatCard label="Highest single expense" value={highestExpense ? fmtINR(highestExpense.amount) : "—"} sub={highestExpense?.category} icon={AlertTriangle} tone="rust" />
        <StatCard label="Top spending category" value={topCat ? topCat[0] : "—"} sub={topCat ? fmtINR(topCat[1]) : ""} icon={Target} tone="brass" />
      </div>

      <div className="grid-two">
        <div className="panel">
          <div className="panel-head"><h3>Income vs expense, six-month trend</h3></div>
          <ResponsiveContainer width="100%" height={260}>
            <AreaChart data={trend}>
              <defs>
                <linearGradient id="inc" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="var(--forest)" stopOpacity={0.35} /><stop offset="100%" stopColor="var(--forest)" stopOpacity={0} /></linearGradient>
                <linearGradient id="exp" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="var(--rust)" stopOpacity={0.35} /><stop offset="100%" stopColor="var(--rust)" stopOpacity={0} /></linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 5" stroke="var(--border)" vertical={false} />
              <XAxis dataKey="label" tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `₹${v / 1000}k`} />
              <Tooltip content={<ChartTooltip />} />
              <Area type="monotone" dataKey="income" stroke="var(--forest)" fill="url(#inc)" strokeWidth={2} name="Income" />
              <Area type="monotone" dataKey="expense" stroke="var(--rust)" fill="url(#exp)" strokeWidth={2} name="Expense" />
            </AreaChart>
          </ResponsiveContainer>
        </div>

        <div className="panel">
          <div className="panel-head"><h3><Sparkles size={15} style={{ marginRight: 6, verticalAlign: -2 }} />Insights</h3></div>
          <div className="insight-list">
            {insights.slice(0, 6).map((ins, i) => <InsightRow key={i} insight={ins} />)}
          </div>
        </div>
      </div>

      <div className="panel">
        <div className="panel-head"><h3>Recent transactions</h3></div>
        <div className="ledger-list">
          {recent.map((t) => <LedgerRow key={t.id} t={t} catColor={catColor} catIcon={catIcon} />)}
          {recent.length === 0 && <EmptyState text="Nothing logged this month yet." />}
        </div>
      </div>
    </div>
  );
}

function ChartTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div className="chart-tip">
      <div className="chart-tip-label">{label}</div>
      {payload.map((p, i) => (
        <div key={i} className="chart-tip-row"><span style={{ color: p.color }}>●</span> {p.name}: <b>{fmtINR(p.value)}</b></div>
      ))}
    </div>
  );
}

function LedgerRow({ t, catColor, catIcon, onEdit, onDelete }) {
  const Icon = catIcon(t.category);
  return (
    <div className="ledger-row">
      <div className={`ledger-bar tone-${t.type === "income" ? "forest" : "rust"}`} />
      <div className="ledger-icon" style={{ background: catColor(t.category) + "22", color: catColor(t.category) }}>
        <Icon size={15} />
      </div>
      <div className="ledger-mid">
        <div className="ledger-cat">{t.category}</div>
        <div className="ledger-notes">{t.notes || fmtDate(t.date)}</div>
      </div>
      <div className="ledger-date">{fmtDate(t.date)}</div>
      <div className={`ledger-amount tone-${t.type === "income" ? "forest" : "rust"}`}>
        {t.type === "income" ? "+" : "−"}{fmtINR(t.amount)}
      </div>
      {(onEdit || onDelete) && (
        <div className="ledger-actions">
          {onEdit && <button className="icon-btn" onClick={() => onEdit(t)}><Edit2 size={14} /></button>}
          {onDelete && <button className="icon-btn" onClick={() => onDelete(t.id)}><Trash2 size={14} /></button>}
        </div>
      )}
    </div>
  );
}

function EmptyState({ text }) {
  return (
    <div className="empty-state">
      <Feather size={22} />
      <p>{text}</p>
    </div>
  );
}

/* ============================================================
   TRANSACTIONS / INCOME / EXPENSES
   ============================================================ */
function TransactionsView({ filterType, transactions, catColor, catIcon, incomeCategories, expenseCategories, onAdd, onEdit, onDelete }) {
  const [search, setSearch] = useState("");
  const [month, setMonth] = useState("all");
  const [category, setCategory] = useState("all");
  const [typeFilter, setTypeFilter] = useState(filterType);

  useEffect(() => setTypeFilter(filterType), [filterType]);

  const months = useMemo(() => {
    const set = new Set(transactions.map((t) => monthKey(new Date(t.date))));
    return Array.from(set).sort().reverse();
  }, [transactions]);

  const cats = typeFilter === "income" ? incomeCategories : typeFilter === "expense" ? expenseCategories : [...incomeCategories, ...expenseCategories];

  const filtered = transactions.filter((t) => {
    if (typeFilter !== "all" && t.type !== typeFilter) return false;
    if (month !== "all" && monthKey(new Date(t.date)) !== month) return false;
    if (category !== "all" && t.category !== category) return false;
    if (search && !(t.category.toLowerCase().includes(search.toLowerCase()) || (t.notes || "").toLowerCase().includes(search.toLowerCase()))) return false;
    return true;
  }).sort((a, b) => new Date(b.date) - new Date(a.date));

  const total = filtered.reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
  const catBreakdown = useMemo(() => {
    const map = {};
    filtered.forEach((t) => map[t.category] = (map[t.category] || 0) + t.amount);
    return Object.entries(map).sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value }));
  }, [filtered]);

  const heading = filterType === "income" ? "Income" : filterType === "expense" ? "Expenses" : "All transactions";

  return (
    <div className="stack-lg">
      <div className="page-head">
        <div><h2 className="page-title">{heading}</h2><p className="page-sub">{filtered.length} entries</p></div>
        <div className="head-actions">
          {filterType !== "expense" && <button className="btn-secondary" onClick={() => onAdd("income")}><ArrowUpCircle size={15} /> Add income</button>}
          {filterType !== "income" && <button className="btn-primary" onClick={() => onAdd("expense")}><Plus size={15} /> Add expense</button>}
        </div>
      </div>

      <div className="filter-bar">
        <div className="input-wrap search"><Search size={15} /><input placeholder="Search category or notes…" value={search} onChange={(e) => setSearch(e.target.value)} /></div>
        <select value={month} onChange={(e) => setMonth(e.target.value)}>
          <option value="all">All months</option>
          {months.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
        </select>
        <select value={category} onChange={(e) => setCategory(e.target.value)}>
          <option value="all">All categories</option>
          {cats.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
        </select>
        {filterType === "all" && (
          <select value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)}>
            <option value="all">Income & expense</option>
            <option value="income">Income only</option>
            <option value="expense">Expense only</option>
          </select>
        )}
      </div>

      <div className="grid-two align-start">
        <div className="panel">
          <div className="panel-head row-between"><h3>Entries</h3><span className={`net tone-${total >= 0 ? "forest" : "rust"}`}>Net {fmtINR(total)}</span></div>
          <div className="ledger-list">
            {filtered.map((t) => <LedgerRow key={t.id} t={t} catColor={catColor} catIcon={catIcon} onEdit={onEdit} onDelete={onDelete} />)}
            {filtered.length === 0 && <EmptyState text="No entries match these filters." />}
          </div>
        </div>
        <div className="panel">
          <div className="panel-head"><h3>By category</h3></div>
          {catBreakdown.length === 0 ? <EmptyState text="Nothing to break down yet." /> : (
            <ResponsiveContainer width="100%" height={260}>
              <PieChart>
                <Pie data={catBreakdown} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={2}>
                  {catBreakdown.map((c, i) => <Cell key={i} fill={catColor(c.name)} />)}
                </Pie>
                <Tooltip content={<ChartTooltip />} />
              </PieChart>
            </ResponsiveContainer>
          )}
          <div className="legend-list">
            {catBreakdown.slice(0, 8).map((c) => (
              <div key={c.name} className="legend-item">
                <span className="dot" style={{ background: catColor(c.name) }} />
                <span className="legend-name">{c.name}</span>
                <span className="legend-value mono">{fmtINR(c.value)}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   BUDGETS
   ============================================================ */
function BudgetsView({ budgetStatus, catColor, catIcon, expenseCategories, onAdd, onEdit, onDelete }) {
  return (
    <div className="stack-lg">
      <div className="page-head">
        <div><h2 className="page-title">Budgets</h2><p className="page-sub">Category limits for this month</p></div>
        <button className="btn-primary" onClick={onAdd}><Plus size={15} /> New budget</button>
      </div>
      <div className="grid-budgets">
        {budgetStatus.map((b) => {
          const Icon = catIcon(b.category);
          return (
            <div className="budget-card" key={b.id}>
              <div className="stitch" />
              <div className="budget-top">
                <div className="ledger-icon" style={{ background: catColor(b.category) + "22", color: catColor(b.category) }}><Icon size={16} /></div>
                <div className="budget-name">{b.category}</div>
                <Pill tone={b.status}>{b.status === "red" ? "Exceeded" : b.status === "yellow" ? "Watch" : "On track"}</Pill>
              </div>
              <div className="budget-figures">
                <span className="mono">{fmtINR(b.used)}</span>
                <span className="dim"> of {fmtINR(b.amount)}</span>
              </div>
              <div className="progress-track"><div className={`progress-fill tone-${b.status}`} style={{ width: `${Math.min(b.pct, 100)}%` }} /></div>
              <div className="budget-bottom">
                <span className="dim">{b.pct.toFixed(0)}% used · {fmtINR(Math.max(b.amount - b.used, 0))} left</span>
                <div className="row-actions">
                  <button className="icon-btn" onClick={() => onEdit(b)}><Edit2 size={13} /></button>
                  <button className="icon-btn" onClick={() => onDelete(b.id)}><Trash2 size={13} /></button>
                </div>
              </div>
            </div>
          );
        })}
        {budgetStatus.length === 0 && <EmptyState text="No budgets set yet — add one to start tracking limits." />}
      </div>
    </div>
  );
}

/* ============================================================
   ANALYTICS
   ============================================================ */
function AnalyticsView({ monthlyAgg, currentMonthTx, catColor, isDark }) {
  const byCat = {};
  currentMonthTx.filter((t) => t.type === "expense").forEach((t) => byCat[t.category] = (byCat[t.category] || 0) + t.amount);
  const topCats = Object.entries(byCat).sort((a, b) => b[1] - a[1]).slice(0, 8).map(([name, value]) => ({ name, value }));

  return (
    <div className="stack-lg">
      <div className="page-head"><div><h2 className="page-title">Analytics</h2><p className="page-sub">Trends across the last six months</p></div></div>

      <div className="panel">
        <div className="panel-head"><h3>Monthly trend — income vs expense</h3></div>
        <ResponsiveContainer width="100%" height={280}>
          <LineChart data={monthlyAgg.slice(-6)}>
            <CartesianGrid strokeDasharray="3 5" stroke="var(--border)" vertical={false} />
            <XAxis dataKey="label" tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `₹${v / 1000}k`} />
            <Tooltip content={<ChartTooltip />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Line type="monotone" dataKey="income" stroke="var(--forest)" strokeWidth={2.5} dot={{ r: 3 }} name="Income" />
            <Line type="monotone" dataKey="expense" stroke="var(--rust)" strokeWidth={2.5} dot={{ r: 3 }} name="Expense" />
          </LineChart>
        </ResponsiveContainer>
      </div>

      <div className="grid-two">
        <div className="panel">
          <div className="panel-head"><h3>Savings trend</h3></div>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={monthlyAgg.slice(-6)}>
              <CartesianGrid strokeDasharray="3 5" stroke="var(--border)" vertical={false} />
              <XAxis dataKey="label" tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `₹${v / 1000}k`} />
              <Tooltip content={<ChartTooltip />} />
              <Bar dataKey="savings" name="Savings" radius={[4, 4, 0, 0]}>
                {monthlyAgg.slice(-6).map((m, i) => <Cell key={i} fill={m.savings >= 0 ? "var(--forest)" : "var(--rust)"} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="panel">
          <div className="panel-head"><h3>Top spending categories</h3></div>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={topCats} layout="vertical" margin={{ left: 10 }}>
              <CartesianGrid strokeDasharray="3 5" stroke="var(--border)" horizontal={false} />
              <XAxis type="number" tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `₹${v / 1000}k`} />
              <YAxis type="category" dataKey="name" width={90} tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={false} tickLine={false} />
              <Tooltip content={<ChartTooltip />} />
              <Bar dataKey="value" name="Spent" radius={[0, 4, 4, 0]}>
                {topCats.map((c, i) => <Cell key={i} fill={catColor(c.name)} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   REPORTS
   ============================================================ */
function ReportsView({ transactions, monthlyAgg, catColor, onExport }) {
  const months = useMemo(() => Array.from(new Set(transactions.map((t) => monthKey(new Date(t.date))))).sort().reverse(), [transactions]);
  const [selected, setSelected] = useState(months[0]);
  const tx = transactions.filter((t) => monthKey(new Date(t.date)) === selected);
  const income = tx.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const expense = tx.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  const balance = income - expense;
  const savingsRate = income ? ((balance / income) * 100).toFixed(1) : "0";
  const daysInMonth = selected ? new Date(...selected.split("-").map((n, i) => (i === 1 ? n - 1 : Number(n)))).getMonth() : 0;
  const dayCount = selected ? new Date(Number(selected.split("-")[0]), Number(selected.split("-")[1]), 0).getDate() : 30;
  const avgDaily = expense / dayCount;

  const byDay = {};
  tx.filter((t) => t.type === "expense").forEach((t) => byDay[t.date] = (byDay[t.date] || 0) + t.amount);
  const highestDay = Object.entries(byDay).sort((a, b) => b[1] - a[1])[0];

  const byCat = {};
  tx.filter((t) => t.type === "expense").forEach((t) => byCat[t.category] = (byCat[t.category] || 0) + t.amount);
  const catRows = Object.entries(byCat).sort((a, b) => b[1] - a[1]);
  const topCat = catRows[0];

  return (
    <div className="stack-lg">
      <div className="page-head">
        <div><h2 className="page-title">Monthly reports</h2><p className="page-sub">Auto-generated summary, ready to export</p></div>
        <div className="head-actions">
          <select value={selected || ""} onChange={(e) => setSelected(e.target.value)}>
            {months.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
          </select>
          <button className="btn-primary" onClick={() => onExport(selected)}><Download size={15} /> Export .xlsx</button>
        </div>
      </div>

      <div className="grid-stats">
        <StatCard label="Total income" value={fmtINR(income)} tone="forest" icon={ArrowUpCircle} />
        <StatCard label="Total expense" value={fmtINR(expense)} tone="rust" icon={ArrowDownCircle} />
        <StatCard label="Remaining balance" value={fmtINR(balance)} tone={balance >= 0 ? "forest" : "rust"} icon={Wallet} />
        <StatCard label="Savings" value={fmtINR(balance)} tone="brass" icon={PiggyBank} />
        <StatCard label="Savings rate" value={`${savingsRate}%`} icon={TrendingUp} />
        <StatCard label="Average daily expense" value={fmtINR(avgDaily)} icon={BarChart3} />
        <StatCard label="Highest spending day" value={highestDay ? fmtDate(highestDay[0]) : "—"} sub={highestDay ? fmtINR(highestDay[1]) : ""} icon={AlertTriangle} tone="rust" />
        <StatCard label="Top category" value={topCat ? topCat[0] : "—"} sub={topCat ? fmtINR(topCat[1]) : ""} icon={Target} tone="brass" />
        <StatCard label="Transactions" value={tx.length} icon={List} />
      </div>

      <div className="panel">
        <div className="panel-head"><h3>Category-wise summary</h3></div>
        <div className="table">
          <div className="table-head"><span>Category</span><span>Amount</span><span>Share</span></div>
          {catRows.map(([cat, amt]) => (
            <div className="table-row" key={cat}>
              <span><span className="dot" style={{ background: catColor(cat) }} /> {cat}</span>
              <span className="mono">{fmtINR(amt)}</span>
              <span className="dim">{expense ? ((amt / expense) * 100).toFixed(1) : 0}%</span>
            </div>
          ))}
          {catRows.length === 0 && <EmptyState text="No expenses recorded for this month." />}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   SAVINGS
   ============================================================ */
function SavingsView({ monthlyAgg, isDark }) {
  const lifetime = monthlyAgg.reduce((s, m) => s + m.savings, 0);
  const growth = monthlyAgg.length >= 2 ? monthlyAgg[monthlyAgg.length - 1].savings - monthlyAgg[monthlyAgg.length - 2].savings : 0;
  return (
    <div className="stack-lg">
      <div className="page-head"><div><h2 className="page-title">Savings</h2><p className="page-sub">What's been set aside, month over month</p></div></div>

      <div className="grid-stats" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))" }}>
        <StatCard label="Lifetime savings" value={fmtINR(lifetime)} tone="forest" icon={PiggyBank} />
        <StatCard label="Savings growth (last month)" value={fmtINR(growth)} tone={growth >= 0 ? "forest" : "rust"} icon={growth >= 0 ? TrendingUp : TrendingDown} />
        <StatCard label="Months tracked" value={monthlyAgg.length} icon={FileText} />
      </div>

      <div className="panel">
        <div className="panel-head"><h3>Savings by month</h3></div>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={monthlyAgg}>
            <CartesianGrid strokeDasharray="3 5" stroke="var(--border)" vertical={false} />
            <XAxis dataKey="label" tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: "var(--ink-soft)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `₹${v / 1000}k`} />
            <Tooltip content={<ChartTooltip />} />
            <Bar dataKey="savings" name="Saved" radius={[4, 4, 0, 0]}>
              {monthlyAgg.map((m, i) => <Cell key={i} fill={m.savings >= 0 ? "var(--forest)" : "var(--rust)"} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="panel">
        <div className="panel-head"><h3>Ledger of months</h3></div>
        <div className="table">
          <div className="table-head"><span>Month</span><span>Income</span><span>Expense</span><span>Saved</span></div>
          {[...monthlyAgg].reverse().map((m) => (
            <div className="table-row cols-4" key={m.key}>
              <span>{m.label}</span>
              <span className="mono">{fmtINR(m.income)}</span>
              <span className="mono">{fmtINR(m.expense)}</span>
              <span className={`mono tone-${m.savings >= 0 ? "forest" : "rust"}`}>{fmtINR(m.savings)}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   NOTIFICATIONS
   ============================================================ */
function NotificationsView({ notifications }) {
  return (
    <div className="stack-lg">
      <div className="page-head"><div><h2 className="page-title">Notifications</h2><p className="page-sub">Reminders generated from your ledger</p></div></div>
      <div className="panel">
        <div className="insight-list">
          {notifications.map((n, i) => (
            <div key={i} className={`insight-row tone-${n.type}`}><n.icon size={15} /><span>{n.text}</span></div>
          ))}
          {notifications.length === 0 && <EmptyState text="All clear — nothing needs your attention." />}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   SETTINGS
   ============================================================ */
function SettingsView({ theme, setTheme, settings, setSettings, incomeCategories, expenseCategories, setIncomeCategories, setExpenseCategories }) {
  const [newIncomeCat, setNewIncomeCat] = useState("");
  const [newExpenseCat, setNewExpenseCat] = useState("");

  function addCat(kind) {
    if (kind === "income" && newIncomeCat.trim()) {
      setIncomeCategories((p) => [...p, { name: newIncomeCat.trim(), icon: "Wallet", color: PALETTE[p.length % PALETTE.length], custom: true }]);
      setNewIncomeCat("");
    }
    if (kind === "expense" && newExpenseCat.trim()) {
      setExpenseCategories((p) => [...p, { name: newExpenseCat.trim(), icon: "MoreHorizontal", color: PALETTE[p.length % PALETTE.length], custom: true }]);
      setNewExpenseCat("");
    }
  }

  return (
    <div className="stack-lg">
      <div className="page-head"><div><h2 className="page-title">Settings</h2><p className="page-sub">Preferences for this ledger</p></div></div>

      <div className="grid-two align-start">
        <div className="panel">
          <div className="panel-head"><h3>Appearance</h3></div>
          <div className="settings-row">
            <div><div className="settings-label">Theme</div><div className="dim">Follows your system by default</div></div>
            <ThemeToggle theme={theme} setTheme={setTheme} />
          </div>
          <div className="settings-row">
            <div><div className="settings-label">Currency</div><div className="dim">Amounts are shown in Indian Rupees</div></div>
            <select value={settings.currency} onChange={(e) => setSettings((s) => ({ ...s, currency: e.target.value }))}>
              <option value="INR">₹ INR — Indian Rupee</option>
              <option value="USD">$ USD — US Dollar</option>
              <option value="EUR">€ EUR — Euro</option>
            </select>
          </div>
          <div className="panel-head" style={{ marginTop: 8 }}><h3>Notifications</h3></div>
          <label className="settings-row checkbox-row">
            <span><div className="settings-label">Budget alerts</div><div className="dim">Warn at 90% and when exceeded</div></span>
            <input type="checkbox" checked={settings.notifBudget} onChange={(e) => setSettings((s) => ({ ...s, notifBudget: e.target.checked }))} />
          </label>
          <label className="settings-row checkbox-row">
            <span><div className="settings-label">Daily entry reminder</div><div className="dim">Nudge if nothing's logged today</div></span>
            <input type="checkbox" checked={settings.notifDaily} onChange={(e) => setSettings((s) => ({ ...s, notifDaily: e.target.checked }))} />
          </label>
          <label className="settings-row checkbox-row">
            <span><div className="settings-label">Report ready</div><div className="dim">Notify near month end</div></span>
            <input type="checkbox" checked={settings.notifReport} onChange={(e) => setSettings((s) => ({ ...s, notifReport: e.target.checked }))} />
          </label>
        </div>

        <div className="panel">
          <div className="panel-head"><h3>Income categories</h3></div>
          <div className="cat-chip-list">
            {incomeCategories.map((c) => (
              <div className="cat-chip" key={c.name}>
                <span className="dot" style={{ background: c.color }} /> {c.name}
                {c.custom && <button onClick={() => setIncomeCategories((p) => p.filter((x) => x.name !== c.name))}><X size={12} /></button>}
              </div>
            ))}
          </div>
          <div className="add-cat-row">
            <input placeholder="New income category…" value={newIncomeCat} onChange={(e) => setNewIncomeCat(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addCat("income")} />
            <button className="btn-secondary" onClick={() => addCat("income")}><Plus size={14} /></button>
          </div>

          <div className="panel-head" style={{ marginTop: 20 }}><h3>Expense categories</h3></div>
          <div className="cat-chip-list">
            {expenseCategories.map((c) => (
              <div className="cat-chip" key={c.name}>
                <span className="dot" style={{ background: c.color }} /> {c.name}
                {c.custom && <button onClick={() => setExpenseCategories((p) => p.filter((x) => x.name !== c.name))}><X size={12} /></button>}
              </div>
            ))}
          </div>
          <div className="add-cat-row">
            <input placeholder="New expense category…" value={newExpenseCat} onChange={(e) => setNewExpenseCat(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addCat("expense")} />
            <button className="btn-secondary" onClick={() => addCat("expense")}><Plus size={14} /></button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   MODALS
   ============================================================ */
function TransactionModal({ modal, onClose, incomeCategories, expenseCategories, onSave }) {
  const isEdit = modal.mode === "edit";
  const [type, setType] = useState(modal.data?.type || modal.type || "expense");
  const [amount, setAmount] = useState(modal.data?.amount || "");
  const [category, setCategory] = useState(modal.data?.category || (type === "income" ? incomeCategories[0]?.name : expenseCategories[0]?.name) || "");
  const [date, setDate] = useState(modal.data?.date || new Date().toISOString().slice(0, 10));
  const [notes, setNotes] = useState(modal.data?.notes || "");
  const cats = type === "income" ? incomeCategories : expenseCategories;

  useEffect(() => { if (!cats.find((c) => c.name === category)) setCategory(cats[0]?.name || ""); }, [type]);

  return (
    <div className="modal-scrim" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h3>{isEdit ? "Edit entry" : "New entry"}</h3>
          <button className="icon-btn" onClick={onClose}><X size={16} /></button>
        </div>
        <div className="type-switch">
          <button className={type === "income" ? "active" : ""} onClick={() => setType("income")}><ArrowUpCircle size={14} /> Income</button>
          <button className={type === "expense" ? "active" : ""} onClick={() => setType("expense")}><ArrowDownCircle size={14} /> Expense</button>
        </div>
        <label className="field">
          <span>Amount</span>
          <div className="input-wrap"><span className="rupee">₹</span><input type="number" min="0" value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" /></div>
        </label>
        <label className="field">
          <span>Category</span>
          <select value={category} onChange={(e) => setCategory(e.target.value)}>
            {cats.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
          </select>
        </label>
        <label className="field">
          <span>Date</span>
          <input type="date" value={date} onChange={(e) => setDate(e.target.value)} />
        </label>
        <label className="field">
          <span>Notes</span>
          <textarea rows={2} value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="Optional" />
        </label>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>Cancel</button>
          <button className="btn-primary" onClick={() => amount && category && onSave({ id: modal.data?.id, type, amount: Number(amount), category, date, notes })}>
            {isEdit ? "Save changes" : "Add entry"}
          </button>
        </div>
      </div>
    </div>
  );
}

function BudgetModal({ modal, onClose, expenseCategories, onSave }) {
  const isEdit = modal.mode === "edit";
  const [category, setCategory] = useState(modal.data?.category || expenseCategories[0]?.name || "");
  const [amount, setAmount] = useState(modal.data?.amount || "");
  return (
    <div className="modal-scrim" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head"><h3>{isEdit ? "Edit budget" : "New budget"}</h3><button className="icon-btn" onClick={onClose}><X size={16} /></button></div>
        <label className="field">
          <span>Category</span>
          <select value={category} onChange={(e) => setCategory(e.target.value)}>
            {expenseCategories.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
          </select>
        </label>
        <label className="field">
          <span>Monthly limit</span>
          <div className="input-wrap"><span className="rupee">₹</span><input type="number" min="0" value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" /></div>
        </label>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>Cancel</button>
          <button className="btn-primary" onClick={() => amount && category && onSave({ id: modal.data?.id, category, amount: Number(amount) })}>
            {isEdit ? "Save changes" : "Create budget"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   STYLES — ledger-book design tokens
   ============================================================ */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');

.bahi {
  --bg: #F2F5F0; --surface: #FFFFFF; --surface-2: #EAEEE7;
  --ink: #16241C; --ink-soft: #5C6B60; --border: #DCE1D8;
  --forest: #1B4D3E; --forest-light: #2F7A5F;
  --brass: #97742B; --brass-light: #B8923F;
  --rust: #A5462A; --rust-light: #C96A46;
  --shadow: 0 1px 2px rgba(22,36,28,0.04), 0 6px 20px rgba(22,36,28,0.06);
  font-family: 'Inter', sans-serif; color: var(--ink); background: var(--bg);
  min-height: 100vh; -webkit-font-smoothing: antialiased;
}
.bahi.dark {
  --bg: #0C1712; --surface: #12201A; --surface-2: #182920;
  --ink: #E9EFE9; --ink-soft: #93A69B; --border: #22362B;
  --forest: #4FAE85; --forest-light: #6BC79E;
  --brass: #D9B571; --brass-light: #E8C98E;
  --rust: #E0825E; --rust-light: #EC9C7C;
  --shadow: 0 1px 2px rgba(0,0,0,0.3), 0 8px 24px rgba(0,0,0,0.35);
}
.bahi * { box-sizing: border-box; }
.bahi h1, .bahi h2, .bahi h3 { font-family: 'Fraunces', serif; margin: 0; letter-spacing: -0.01em; }
.mono { font-family: 'JetBrains Mono', monospace; font-variant-numeric: tabular-nums; }
.dim { color: var(--ink-soft); font-size: 13px; }

/* layout */
.shell { display: flex; min-height: 100vh; }
.sidebar {
  width: 232px; background: var(--surface); border-right: 1px solid var(--border);
  display: flex; flex-direction: column; padding: 20px 14px; position: sticky; top: 0; height: 100vh; flex-shrink: 0;
}
.brand { display: flex; align-items: center; gap: 10px; padding: 6px 8px 22px; }
.brand-mark { width: 34px; height: 34px; border-radius: 9px; background: var(--forest); color: #fff; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
.brand-name { font-family: 'Fraunces', serif; font-weight: 600; font-size: 18px; line-height: 1.1; }
.brand-sub { font-size: 11px; color: var(--ink-soft); letter-spacing: 0.02em; }
.nav { display: flex; flex-direction: column; gap: 2px; flex: 1; overflow-y: auto; }
.nav-item {
  display: flex; align-items: center; gap: 10px; padding: 9px 10px; border-radius: 8px; border: none;
  background: transparent; color: var(--ink-soft); font-size: 13.5px; font-weight: 500; cursor: pointer; text-align: left; width: 100%;
  transition: background .12s, color .12s;
}
.nav-item:hover { background: var(--surface-2); color: var(--ink); }
.nav-item.active { background: var(--forest); color: #fff; }
.nav-badge { margin-left: auto; background: var(--rust); color: #fff; font-size: 10px; font-weight: 700; padding: 1px 6px; border-radius: 999px; }
.nav-item.logout { margin-top: 8px; border-top: 1px solid var(--border); padding-top: 14px; border-radius: 0; }
.scrim { position: fixed; inset: 0; background: rgba(0,0,0,.4); z-index: 20; display: none; }

.main { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.topbar { height: 60px; display: flex; align-items: center; padding: 0 24px; border-bottom: 1px solid var(--border); background: var(--bg); position: sticky; top: 0; z-index: 5; gap: 12px; }
.topbar-title { font-family: 'Fraunces', serif; font-weight: 600; font-size: 17px; }
.topbar-right { margin-left: auto; display: flex; align-items: center; gap: 12px; }
.hamburger { display: none; }
.avatar { width: 30px; height: 30px; border-radius: 999px; background: var(--brass); color: #fff; font-weight: 700; font-size: 12px; display: flex; align-items: center; justify-content: center; font-family: 'Fraunces', serif; }
.content { padding: 24px; max-width: 1240px; width: 100%; margin: 0 auto; }

.stack-lg { display: flex; flex-direction: column; gap: 20px; }
.page-head { display: flex; align-items: flex-end; justify-content: space-between; gap: 16px; flex-wrap: wrap; }
.page-title { font-size: 24px; font-weight: 600; }
.page-sub { color: var(--ink-soft); font-size: 13px; margin-top: 2px; }
.head-actions { display: flex; gap: 8px; }

/* buttons + inputs */
.btn-primary, .btn-secondary, .btn-ghost {
  display: inline-flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 600;
  padding: 8px 14px; border-radius: 8px; border: 1px solid transparent; cursor: pointer; font-family: 'Inter', sans-serif;
  transition: transform .08s, opacity .12s;
}
.btn-primary { background: var(--forest); color: #fff; }
.btn-primary:hover { opacity: .92; }
.btn-secondary { background: var(--surface-2); color: var(--ink); border-color: var(--border); }
.btn-secondary:hover { background: var(--surface); }
.btn-ghost { background: transparent; color: var(--ink-soft); }
.btn-ghost:hover { color: var(--ink); }
.btn-primary.full, .btn-secondary.full { width: 100%; justify-content: center; padding: 10px; }
.icon-btn { border: none; background: transparent; color: var(--ink-soft); cursor: pointer; padding: 6px; border-radius: 6px; display: inline-flex; }
.icon-btn:hover { background: var(--surface-2); color: var(--ink); }
.ghost-icon { border: none; background: transparent; color: var(--ink-soft); cursor: pointer; margin-left: auto; display: flex; }

select, input[type=text], input[type=number], input[type=date], input[type=email], input[type=password], textarea {
  font-family: 'Inter', sans-serif; font-size: 13.5px; background: var(--surface); color: var(--ink);
  border: 1px solid var(--border); border-radius: 8px; padding: 8px 10px; outline: none;
}
select { cursor: pointer; }
input:focus, select:focus, textarea:focus { border-color: var(--forest); }
.input-wrap { display: flex; align-items: center; gap: 6px; border: 1px solid var(--border); border-radius: 8px; padding: 0 10px; background: var(--surface); }
.input-wrap input, .input-wrap select { border: none; padding: 8px 4px; flex: 1; background: transparent; }
.input-wrap svg { color: var(--ink-soft); flex-shrink: 0; }
.input-wrap.search { min-width: 200px; flex: 1; }
.rupee { color: var(--ink-soft); font-weight: 600; }

.filter-bar { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }

.theme-toggle { display: flex; background: var(--surface-2); border-radius: 8px; padding: 3px; gap: 2px; }
.theme-toggle button { border: none; background: transparent; color: var(--ink-soft); padding: 5px 8px; border-radius: 6px; cursor: pointer; display: flex; }
.theme-toggle button.active { background: var(--surface); color: var(--forest); box-shadow: var(--shadow); }

/* stat cards */
.grid-stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; }
.stat-card { position: relative; background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 16px 16px 16px 20px; box-shadow: var(--shadow); overflow: hidden; }
.stitch { position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background-image: repeating-linear-gradient(to bottom, var(--brass) 0 6px, transparent 6px 12px); }
.stat-top { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }
.stat-label { font-size: 12px; color: var(--ink-soft); font-weight: 500; }
.stat-value { font-family: 'JetBrains Mono', monospace; font-size: 21px; font-weight: 600; font-variant-numeric: tabular-nums; }
.stat-sub { font-size: 11.5px; color: var(--ink-soft); margin-top: 3px; }
.tone-forest { color: var(--forest); }
.tone-rust { color: var(--rust); }
.tone-brass { color: var(--brass); }
.tone-ink { color: var(--ink); }
.stat-icon.tone-forest { color: var(--forest); } .stat-icon.tone-rust { color: var(--rust); } .stat-icon.tone-brass { color: var(--brass); } .stat-icon.tone-ink { color: var(--ink-soft); }

/* panels */
.grid-two { display: grid; grid-template-columns: 1.3fr 1fr; gap: 16px; }
.grid-two.align-start { align-items: start; }
.panel { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 18px; box-shadow: var(--shadow); }
.panel-head { margin-bottom: 12px; }
.panel-head h3 { font-size: 15px; font-weight: 600; }
.row-between { display: flex; align-items: center; justify-content: space-between; }
.net { font-family: 'JetBrains Mono', monospace; font-weight: 600; font-size: 13px; }

/* insights */
.insight-list { display: flex; flex-direction: column; gap: 8px; }
.insight-row { display: flex; gap: 9px; align-items: flex-start; font-size: 13px; padding: 9px 10px; border-radius: 8px; background: var(--surface-2); line-height: 1.45; }
.insight-row svg { flex-shrink: 0; margin-top: 1px; }
.insight-row.tone-positive { color: var(--forest); } .insight-row.tone-positive svg { color: var(--forest); }
.insight-row.tone-warning { color: var(--rust); } .insight-row.tone-warning svg { color: var(--rust); }
.insight-row.tone-info { color: var(--ink); } .insight-row.tone-info svg { color: var(--brass); }

/* ledger rows */
.ledger-list { display: flex; flex-direction: column; }
.ledger-row { display: flex; align-items: center; gap: 12px; padding: 10px 4px; border-bottom: 1px dashed var(--border); }
.ledger-row:last-child { border-bottom: none; }
.ledger-bar { width: 3px; align-self: stretch; border-radius: 3px; background: var(--forest); }
.ledger-bar.tone-rust { background: var(--rust); }
.ledger-icon { width: 32px; height: 32px; border-radius: 9px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
.ledger-mid { flex: 1; min-width: 0; }
.ledger-cat { font-size: 13.5px; font-weight: 600; }
.ledger-notes { font-size: 12px; color: var(--ink-soft); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.ledger-date { font-size: 12px; color: var(--ink-soft); width: 90px; text-align: right; display: none; }
.ledger-amount { font-family: 'JetBrains Mono', monospace; font-weight: 600; font-size: 13.5px; white-space: nowrap; }
.ledger-amount.tone-forest { color: var(--forest); } .ledger-amount.tone-rust { color: var(--rust); }
.ledger-actions { display: flex; gap: 2px; }
@media (min-width: 720px) { .ledger-date { display: block; } }

.empty-state { display: flex; flex-direction: column; align-items: center; gap: 8px; padding: 32px 0; color: var(--ink-soft); text-align: center; }

/* budgets */
.grid-budgets { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 14px; }
.budget-card { position: relative; background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 16px 16px 16px 20px; box-shadow: var(--shadow); }
.budget-top { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
.budget-name { font-weight: 600; font-size: 14px; flex: 1; }
.budget-figures { font-size: 18px; margin-bottom: 8px; }
.progress-track { height: 7px; background: var(--surface-2); border-radius: 999px; overflow: hidden; margin-bottom: 8px; }
.progress-fill { height: 100%; border-radius: 999px; transition: width .3s; }
.progress-fill.tone-green { background: var(--forest); }
.progress-fill.tone-yellow { background: var(--brass); }
.progress-fill.tone-red { background: var(--rust); }
.budget-bottom { display: flex; align-items: center; justify-content: space-between; }
.row-actions { display: flex; gap: 2px; }

.pill { font-size: 10.5px; font-weight: 700; padding: 3px 8px; border-radius: 999px; text-transform: uppercase; letter-spacing: .03em; }
.pill-green { background: color-mix(in srgb, var(--forest) 16%, transparent); color: var(--forest); }
.pill-yellow { background: color-mix(in srgb, var(--brass) 18%, transparent); color: var(--brass); }
.pill-red { background: color-mix(in srgb, var(--rust) 16%, transparent); color: var(--rust); }

/* tables */
.table { display: flex; flex-direction: column; }
.table-head, .table-row { display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 8px; padding: 9px 4px; font-size: 13px; align-items: center; }
.table-row.cols-4 { grid-template-columns: 1.4fr 1fr 1fr 1fr; }
.table-head { color: var(--ink-soft); font-size: 11.5px; text-transform: uppercase; letter-spacing: .04em; border-bottom: 1px solid var(--border); }
.table-row { border-bottom: 1px dashed var(--border); }
.table-row:last-child { border-bottom: none; }
.dot { display: inline-block; width: 8px; height: 8px; border-radius: 999px; margin-right: 6px; }

/* legend */
.legend-list { display: flex; flex-direction: column; gap: 6px; margin-top: 10px; }
.legend-item { display: flex; align-items: center; gap: 8px; font-size: 12.5px; }
.legend-name { flex: 1; }
.legend-value { color: var(--ink-soft); }

/* settings */
.settings-row { display: flex; align-items: center; justify-content: space-between; padding: 12px 2px; border-bottom: 1px dashed var(--border); gap: 12px; }
.settings-row:last-child { border-bottom: none; }
.settings-label { font-size: 13.5px; font-weight: 600; }
.checkbox-row input { width: 17px; height: 17px; accent-color: var(--forest); cursor: pointer; }
.cat-chip-list { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 10px; }
.cat-chip { display: flex; align-items: center; gap: 6px; background: var(--surface-2); border-radius: 999px; padding: 5px 10px; font-size: 12.5px; }
.cat-chip button { border: none; background: transparent; color: var(--ink-soft); cursor: pointer; display: flex; padding: 0; }
.add-cat-row { display: flex; gap: 8px; }
.add-cat-row input { flex: 1; }

/* modals */
.modal-scrim { position: fixed; inset: 0; background: rgba(12,20,16,.5); display: flex; align-items: center; justify-content: center; z-index: 50; padding: 16px; }
.modal { background: var(--surface); border-radius: 14px; padding: 22px; width: 100%; max-width: 380px; box-shadow: 0 20px 50px rgba(0,0,0,.25); display: flex; flex-direction: column; gap: 12px; max-height: 90vh; overflow-y: auto; }
.modal-head { display: flex; align-items: center; justify-content: space-between; }
.modal-head h3 { font-size: 17px; }
.type-switch { display: flex; background: var(--surface-2); border-radius: 8px; padding: 3px; gap: 2px; }
.type-switch button { flex: 1; border: none; background: transparent; padding: 8px; border-radius: 6px; cursor: pointer; font-size: 12.5px; font-weight: 600; color: var(--ink-soft); display: flex; align-items: center; justify-content: center; gap: 5px; }
.type-switch button.active { background: var(--surface); color: var(--forest); box-shadow: var(--shadow); }
.field { display: flex; flex-direction: column; gap: 5px; }
.field span { font-size: 12px; font-weight: 600; color: var(--ink-soft); }
.modal-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 6px; }

/* auth screen */
.auth-wrap { min-height: 100vh; display: grid; grid-template-columns: 1fr 1fr; background: var(--bg); }
.auth-side { padding: 60px 56px; display: flex; flex-direction: column; justify-content: center; background: var(--surface-2); border-right: 1px solid var(--border); }
.auth-headline { font-size: 36px; line-height: 1.15; font-weight: 600; margin-bottom: 14px; }
.auth-copy { color: var(--ink-soft); font-size: 14.5px; line-height: 1.6; max-width: 380px; }
.auth-stitch { height: 1px; margin: 26px 0 18px; background-image: repeating-linear-gradient(to right, var(--brass) 0 8px, transparent 8px 16px); max-width: 200px; }
.auth-demo-note { font-size: 12px; color: var(--ink-soft); max-width: 360px; padding: 10px 12px; border: 1px dashed var(--border); border-radius: 8px; background: var(--bg); }
.auth-card { display: flex; flex-direction: column; justify-content: center; padding: 56px; max-width: 420px; margin: 0 auto; width: 100%; }
.auth-tabs { display: flex; gap: 4px; margin-bottom: 22px; border-bottom: 1px solid var(--border); }
.auth-tabs button { border: none; background: transparent; padding: 8px 4px; margin-right: 16px; font-size: 13px; font-weight: 600; color: var(--ink-soft); cursor: pointer; border-bottom: 2px solid transparent; }
.auth-tabs button.active { color: var(--forest); border-color: var(--forest); }
.auth-form { display: flex; flex-direction: column; gap: 14px; }
.remember { display: flex; align-items: center; gap: 6px; font-size: 12.5px; color: var(--ink-soft); }
.remember input { accent-color: var(--forest); }
.link { border: none; background: transparent; color: var(--forest); font-weight: 600; font-size: 12.5px; cursor: pointer; padding: 0; }
.auth-switch { text-align: center; font-size: 13px; color: var(--ink-soft); margin-top: 16px; }

@media (max-width: 980px) {
  .auth-wrap { grid-template-columns: 1fr; }
  .auth-side { display: none; }
  .grid-two { grid-template-columns: 1fr; }
}
@media (max-width: 760px) {
  .sidebar { position: fixed; z-index: 30; transform: translateX(-100%); transition: transform .2s; box-shadow: 0 0 0 rgba(0,0,0,0); }
  .sidebar.open { transform: translateX(0); box-shadow: 10px 0 30px rgba(0,0,0,.2); }
  .scrim { display: block; }
  .hamburger { display: flex; }
  .content { padding: 16px; }
  .grid-stats { grid-template-columns: repeat(auto-fit, minmax(150px,1fr)); }
}
`;
