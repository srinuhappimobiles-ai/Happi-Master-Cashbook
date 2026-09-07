import React, { useEffect, useRef, useState } from 'react';
import Handsontable from 'handsontable';
import { registerAllModules } from 'handsontable/registry';
import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';

import 'handsontable/styles/handsontable.min.css';
import 'handsontable/styles/ht-theme-main.min.css';

registerAllModules();

const SUPABASE_URL = 'https://qcubsjcttikjtlsicaii.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFjdWJzamN0dGlranRsc2ljYWlpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODgwODc5MjIsImV4cCI6MjEwMzY2MzkyMn0.tZ0Yekxc2b9BM_mgPOmQ1Gv8cqTBRUQArmznP_WDCek';
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

const DEFAULT_BRANCHES = [
  { CODE: 'ADBD', BRANCH: 'ADILABAD' }, { CODE: 'AMP', BRANCH: 'AMALAPURAM' },
  { CODE: 'AMPT', BRANCH: 'AMEERPET' }, { CODE: 'ANTP', BRANCH: 'ANANTAPUR' },
  { CODE: 'ARMU', BRANCH: 'ARMOOR' }, { CODE: 'ATMKR', BRANCH: 'ATMAKUR' },
  { CODE: 'BDHN', BRANCH: 'BODHAN' }, { CODE: 'BDPL', BRANCH: 'BODUPPAL' },
  { CODE: 'BG', BRANCH: 'BHUVANAGIRI' }, { CODE: 'BVRM', BRANCH: 'BHIMAVARAM' },
  { CODE: 'CHND', BRANCH: 'CHANDANAGAR' }, { CODE: 'CHNT', BRANCH: 'CHINTAL' },
  { CODE: 'DBGS', BRANCH: 'DABAGARDENS' }, { CODE: 'DBGS-2', BRANCH: 'DABAGARDENS-2' },
  { CODE: 'DVKD', BRANCH: 'DEVARAKONDA' }, { CODE: 'DVRM', BRANCH: 'DHARMAVRAM' },
  { CODE: 'ECIL', BRANCH: 'ECIL' }, { CODE: 'ELR', BRANCH: 'ELURU' },
  { CODE: 'GDVK', BRANCH: 'GODHAVARIKHANI' }, { CODE: 'GJWK', BRANCH: 'GAJUWAKA' },
  { CODE: 'GJWL', BRANCH: 'GAJWEL' }, { CODE: 'GNT', BRANCH: 'GUNTUR' },
  { CODE: 'GNT2', BRANCH: 'GUNTUR2' }, { CODE: 'GTKL', BRANCH: 'GUNTAKAL' },
  { CODE: 'GWD', BRANCH: 'GADWAL' }, { CODE: 'HAL', BRANCH: 'HALIYA' },
  { CODE: 'HNMK', BRANCH: 'HANUMAKONDA' }, { CODE: 'HUP', BRANCH: 'HINDUPUR' },
  { CODE: 'JCL', BRANCH: 'JADCHERLA' }, { CODE: 'JNGN', BRANCH: 'JANGAON' },
  { CODE: 'JTL', BRANCH: 'JAGTIAL' }, { CODE: 'KDGM', BRANCH: 'KALYANADURGAM' },
  { CODE: 'KDR', BRANCH: 'KADIRI' }, { CODE: 'KDR2', BRANCH: 'KADIRI-2' },
  { CODE: 'KKP', BRANCH: 'KUKATPALLY' }, { CODE: 'KMGH', BRANCH: 'KHARMANGHAT' },
  { CODE: 'KMM', BRANCH: 'KHAMMAM' }, { CODE: 'KMM2', BRANCH: 'KHAMMAM 2' },
  { CODE: 'KPM', BRANCH: 'KUPPAM' }, { CODE: 'KRKH', BRANCH: 'KHARKHANA' },
  { CODE: 'KRLA', BRANCH: 'KORUTLA' }, { CODE: 'KRMN', BRANCH: 'KARIMNAGAR' },
  { CODE: 'KRNL', BRANCH: 'KURNOOL' }, { CODE: 'KRNL2', BRANCH: 'KURNOOL2' },
  { CODE: 'KZP', BRANCH: 'KAZIPET' }, { CODE: 'MCI', BRANCH: 'MANCHERIAL' },
  { CODE: 'MDPL', BRANCH: 'MADANAPALLI' }, { CODE: 'MDPR', BRANCH: 'MADHAPUR' },
  { CODE: 'MDPT', BRANCH: 'MANDAPETA' }, { CODE: 'MHBR', BRANCH: 'MAHABUBNAGAR' },
  { CODE: 'MLKJ', BRANCH: 'MALKAJGIRI' }, { CODE: 'MRGA', BRANCH: 'MIRYALAGUDA' },
  { CODE: 'MTM', BRANCH: 'MACHILIPATNAM' }, { CODE: 'MVP', BRANCH: 'MVP COLONY' },
  { CODE: 'NDD', BRANCH: 'NIDADAVOLE' }, { CODE: 'NDL', BRANCH: 'NANDYALA' },
  { CODE: 'NGKL', BRANCH: 'NAGARKURNOOL' }, { CODE: 'NKRL', BRANCH: 'NAKREKAL' },
  { CODE: 'NLG', BRANCH: 'NALGONDA' }, { CODE: 'NRKD', BRANCH: 'NARAYANKHED' },
  { CODE: 'NRML', BRANCH: 'NIRMAL' }, { CODE: 'NRSP', BRANCH: 'NARASANNAPETA' },
  { CODE: 'NSMP', BRANCH: 'NARSAMPET' }, { CODE: 'NSPT', BRANCH: 'NARSIPATNAM' },
  { CODE: 'NZVD', BRANCH: 'NUZVID' }, { CODE: 'ONG', BRANCH: 'ONGOLE' },
  { CODE: 'PDPL', BRANCH: 'PEDDAPALLI' }, { CODE: 'PDPM', BRANCH: 'PEDDAPURAM' },
  { CODE: 'PIL', BRANCH: 'PILERU' }, { CODE: 'PLM', BRANCH: 'PALAMANER' },
  { CODE: 'PSA', BRANCH: 'PALASA' }, { CODE: 'PVA', BRANCH: 'PALAVANCHA' },
  { CODE: 'RCT', BRANCH: 'RAYACHOTI' }, { CODE: 'RJY', BRANCH: 'RAJAMUNDRY' },
  { CODE: 'RMTP', BRANCH: 'RAMANTHAPUR' }, { CODE: 'RTCX', BRANCH: 'RTC X ROAD' },
  { CODE: 'SDPT', BRANCH: 'SIDDIPET' }, { CODE: 'SDR', BRANCH: 'S.D.ROAD' },
  { CODE: 'SHDR', BRANCH: 'SHADNAGAR' }, { CODE: 'SHPR', BRANCH: 'SHAPUR' },
  { CODE: 'SKKM', BRANCH: 'SRIKAKULAM' }, { CODE: 'SMBD', BRANCH: 'SHAMSHABAD' },
  { CODE: 'SNGR', BRANCH: 'SANGAREDDY' }, { CODE: 'SPT', BRANCH: 'SOMPETA' },
  { CODE: 'SRN', BRANCH: 'S.R.NAGAR' }, { CODE: 'SRNR', BRANCH: 'SAROORNAGAR' },
  { CODE: 'SRPT', BRANCH: 'SURYAPET' }, { CODE: 'STNR', BRANCH: 'SANTOSHNAGAR' },
  { CODE: 'TDPG', BRANCH: 'TADEPALLIGUDEM' }, { CODE: 'TDPT', BRANCH: 'TADIPATRI' },
  { CODE: 'TDU', BRANCH: 'TANDUR' }, { CODE: 'TEK', BRANCH: 'TEKKALI' },
  { CODE: 'TN', BRANCH: 'TUNI' }, { CODE: 'TNK', BRANCH: 'TANUKU' },
  { CODE: 'TNL', BRANCH: 'TENALI' }, { CODE: 'TPT', BRANCH: 'TIRUPATHI' },
  { CODE: 'TPT2', BRANCH: 'TIRUPATHI 2' }, { CODE: 'UPL', BRANCH: 'UPPAL' },
  { CODE: 'VIJ-1', BRANCH: 'VIJAYAWADA 1' }, { CODE: 'VIJ-3', BRANCH: 'VIJAYAWADA 3' },
  { CODE: 'VIJ-4', BRANCH: 'VIJAYAWADA 4' }, { CODE: 'VNSP', BRANCH: 'VANASTALIPURAM' },
  { CODE: 'VZM', BRANCH: 'VIZIANAGARAM' }, { CODE: 'VZM2', BRANCH: 'VIZIANAGARAM 2' },
  { CODE: 'WGL', BRANCH: 'WARANGAL' }, { CODE: 'WGL2', BRANCH: 'WARANGAL 2' },
  { CODE: 'ZB', BRANCH: 'ZAHEERABAD' }
];

const COL_HEADERS = [
  'A: SL.No.', 'B: CODE', 'C: BRANCH', 'D: OPENING BALANCE', 'E: DEPOSIT', 'F: DENOMINATION',
  'G: AddinGS', 'H: PENDING APPRVLS', 'I: FINANCE AMNT', 'J: SR', 'K: SWEEPER SALARY',
  'L: EDITS', 'M: APX SHORTAGE', "N: (KSP)'Sir's Approvals", 'O: CLOSING BALANCE', 'P: REMARKS'
];

const EXCEL_HEADERS = [
  'SL.No.', 'CODE', 'BRANCH', 'OPENING BALANCE', 'DEPOSIT', 'DENOMINATION',
  'AddinGS', 'PENDING APPRVLS', 'FINANCE AMNT', 'SR', 'SWEEPER SALARY',
  'EDITS', 'APX SHORTAGE', "(KSP) Sir's Approvals", 'CLOSING BALANCE', 'REMARKS'
];

const DB_FIELD_MAP = {
  3: 'opening_balance',
  4: 'deposit',
  5: 'denomination',
  6: 'addings',
  7: 'pending_apprvls',
  8: 'finance_amnt',
  9: 'sr',
  10: 'sweeper_salary',
  11: 'edits',
  12: 'apx_shortage',
  13: 'ksp_approvals',
  15: 'remarks'
};

const FIELD_TO_COL_MAP = {
  opening_balance: 3,
  deposit: 4,
  denomination: 5,
  addings: 6,
  pending_apprvls: 7,
  finance_amnt: 8,
  sr: 9,
  sweeper_salary: 10,
  edits: 11,
  apx_shortage: 12,
  ksp_approvals: 13,
  remarks: 15
};

function colLetterToIndex(letter) {
  let col = 0;
  letter = letter.toUpperCase();
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col - 1;
}

function parseToNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  const str = val.toString().trim();
  if (str.startsWith('=')) return 0;
  const clean = str.replace(/,/g, '');
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : Math.round(num);
}

function evaluateFormula(formulaStr, hotInstance, currentRowIdx, visited = new Set()) {
  if (!formulaStr || !formulaStr.toString().startsWith('=')) return formulaStr;
  const rawFormula = formulaStr.toString();
  
  if (visited.has(`${currentRowIdx}-${rawFormula}`)) return 0;
  visited.add(`${currentRowIdx}-${rawFormula}`);

  let expr = rawFormula.substring(1).trim();

  const resolvedExpr = expr.replace(/\b([A-Z]+)([0-9]+)\b/gi, (match, colLetters, rowNumStr) => {
    const colIdx = colLetterToIndex(colLetters);
    const targetRowIdx = parseInt(rowNumStr, 10) - 1;
    if (hotInstance && targetRowIdx >= 0 && colIdx >= 0) {
      const cellVal = hotInstance.getDataAtCell(targetRowIdx, colIdx);
      if (cellVal && cellVal.toString().startsWith('=')) {
        return evaluateFormula(cellVal, hotInstance, targetRowIdx, visited).toString();
      }
      return parseToNum(cellVal).toString();
    }
    return '0';
  });

  if (resolvedExpr.toUpperCase().startsWith('SUM')) {
    try {
      const match = resolvedExpr.match(/\(([^)]+)\)/);
      if (match && match[1]) {
        const inner = match[1];
        const parts = inner.split(/[,+]/);
        let total = 0;
        parts.forEach(p => {
          total += parseFloat(p.trim()) || 0;
        });
        return Math.round(total);
      }
    } catch {
      return 0;
    }
  }

  try {
    const sanitized = resolvedExpr.replace(/[^0-9+\-*/().\s]/g, '');
    if (sanitized === '') return 0;
    const evaluated = Function('return (' + sanitized + ')')();
    return Math.round(Number(evaluated)) || 0;
  } catch {
    return 0;
  }
}

export default function App() {
  const containerRef = useRef(null);
  const hotInstanceRef = useRef(null);
  const jetChannelRef = useRef(null);
  const isBusyRef = useRef(false);
  const debounceTimerRef = useRef(null);

  const [syncStatus, setSyncStatus] = useState('🟢 Jet Synced');
  const [cashierFilter, setCashierFilter] = useState('ALL');
  const [totalStores, setTotalStores] = useState(DEFAULT_BRANCHES.length);
  
  const [pendingStoresList, setPendingStoresList] = useState([]);
  const [pendingDepositSlipsList, setPendingDepositSlipsList] = useState([]);
  const [showCombinedHover, setShowCombinedHover] = useState(false);

  const [pendingApprovalsList, setPendingApprovalsList] = useState([]);
  const [showApprovalsHover, setShowApprovalsHover] = useState(false);

  const [editsList, setEditsList] = useState([]);
  const [showEditsHover, setShowEditsHover] = useState(false);

  const [lowNoCashList, setLowNoCashList] = useState([]);
  const [showLowCashModal, setShowLowCashModal] = useState(false);
  const [lowCashRemarksState, setLowCashRemarksState] = useState({});

  const [showFinanceModal, setShowFinanceModal] = useState(false);
  const [financeRows, setFinanceRows] = useState([]);

  const [showSrModal, setShowSrModal] = useState(false);
  const [srRows, setSrRows] = useState([]);

  const [showFindModal, setShowFindModal] = useState(false);
  const [findText, setFindText] = useState('');
  const [findResults, setFindResults] = useState([]);
  const [currentResultIdx, setCurrentResultIdx] = useState(-1);
  const findInputRef = useRef(null);

  const updatePendingStores = (hotInstance) => {
    if (!hotInstance) return;
    const data = hotInstance.getData();
    const pendingCash = [];
    const pendingDepSlips = [];
    const pendingApprovals = [];
    const editsStores = [];
    const lowNoCash = [];

    data.forEach((row) => {
      const code = row[1];
      const branch = row[2];
      const rowText = (code + ' ' + branch).toUpperCase();
      
      const depositVal = row[4];
      const denomVal = row[5];
      const apprvlsVal = row[7];
      const editsVal = row[11];

      const isDepositEmpty = depositVal === '' || depositVal === null || depositVal === undefined || parseToNum(depositVal) === 0;
      const isDenomEmpty = denomVal === '' || denomVal === null || denomVal === undefined;
      const denomNum = parseToNum(denomVal);

      if (!rowText.includes('TOTAL')) {
        const isDenomStrictEmpty = denomVal === '' || denomVal === null || denomVal === undefined;
        if (isDenomStrictEmpty && isDepositEmpty) {
          pendingCash.push({ code, branch });
          pendingDepSlips.push({ code, branch, denomVal: '0' });
        } else if (!isDenomStrictEmpty && isDepositEmpty) {
          pendingDepSlips.push({ code, branch, denomVal: denomVal });
        }

        if (isDepositEmpty && !isDenomEmpty) {
          lowNoCash.push({ code, branch, denom: denomNum });
        }

        const apprvlsAmt = parseToNum(apprvlsVal);
        if (apprvlsAmt > 0) {
          pendingApprovals.push({ code, branch, amount: apprvlsAmt });
        }

        const editsAmt = parseToNum(editsVal);
        if (editsAmt > 0) {
          editsStores.push({ code, branch, amount: editsAmt });
        }
      }
    });

    lowNoCash.sort((a, b) => a.denom - b.denom);

    setPendingStoresList(pendingCash);
    setPendingDepositSlipsList(pendingDepSlips);
    setPendingApprovalsList(pendingApprovals);
    setEditsList(editsStores);
    setLowNoCashList(lowNoCash);
  };

  const exportPendingApprovalsToExcel = () => {
    const headers = ['Sl.no', 'Store Code', 'Branch Name', 'Pending Approvals Amount'];
    const data = pendingApprovalsList.map((item, idx) => [idx + 1, item.code, item.branch, item.amount]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pending Approvals');
    XLSX.writeFile(wb, 'PENDING_APPROVALS_REPORT.xlsx');
  };

  const exportEditsToExcel = () => {
    const headers = ['Sl.no', 'Store Code', 'Branch Name', 'Edits Amount'];
    const data = editsList.map((item, idx) => [idx + 1, item.code, item.branch, item.amount]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Edits');
    XLSX.writeFile(wb, 'EDITS_REPORT.xlsx');
  };

  const exportLowCashToExcel = () => {
    const headers = ['Sl.no', 'Store Code', 'Branch Name', 'Denom Amt', 'Status / Remarks'];
    const data = lowNoCashList.map((item, idx) => {
      const denomNum = item.denom;
      let status = lowCashRemarksState[item.code] || '';
      if (denomNum === 0) status = 'No Cash';
      else if (denomNum < 500) status = 'Low Cash';
      return [idx + 1, item.code, item.branch, denomNum, status];
    });

    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Low-No Cash Report');
    XLSX.writeFile(wb, 'LOW_NO_CASH_REPORT.xlsx');
  };

  const openFinanceReportModal = () => {
    const hot = hotInstanceRef.current;
    if (!hot) return;
    const data = hot.getData();
    const rows = [];
    const todayFormatted = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: '2-digit' }).replace(/\//g, '-');
    const todayInputFormat = new Date().toISOString().split('T')[0];

    data.forEach((row) => {
      const code = row[1];
      const branch = row[2];
      const rowText = (code + ' ' + branch).toUpperCase();
      const financeAmt = parseToNum(row[8]);

      if (financeAmt > 0 && !rowText.includes('TOTAL')) {
        const existing = financeRows.find(r => r.code === code);
        rows.push({
          code,
          branch,
          amount: financeAmt,
          billNo: existing ? existing.billNo : '',
          remarks: existing ? existing.remarks : 'CASH TO CARD',
          date: todayFormatted,
          billDate: existing ? existing.billDate : todayInputFormat
        });
      }
    });

    setFinanceRows(rows);
    setShowFinanceModal(true);
  };

  const openSrReportModal = () => {
    const hot = hotInstanceRef.current;
    if (!hot) return;
    const data = hot.getData();
    const rows = [];
    const todayFormatted = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: '2-digit' }).replace(/\//g, '-');
    const todayInputFormat = new Date().toISOString().split('T')[0];

    data.forEach((row) => {
      const code = row[1];
      const branch = row[2];
      const rowText = (code + ' ' + branch).toUpperCase();
      const srAmt = parseToNum(row[9]);

      if (srAmt > 0 && !rowText.includes('TOTAL')) {
        const existing = srRows.find(r => r.code === code);
        rows.push({
          code,
          branch,
          amount: srAmt,
          billNo: existing ? existing.billNo : '',
          remarks: existing ? existing.remarks : 'SR REFUND',
          date: todayFormatted,
          billDate: existing ? existing.billDate : todayInputFormat
        });
      }
    });

    setSrRows(rows);
    setShowSrModal(true);
  };

  const exportFinanceToExcel = () => {
    const headers = ['Sl.no', 'Branch', 'Amount', 'Bill No.', 'Remarks', 'Date', 'Bill date', 'Ageing'];
    const data = financeRows.map((item, idx) => {
      const d1 = new Date();
      const d2 = new Date(item.billDate);
      const diffTime = d1 - d2;
      const ageingDays = Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
      return [idx + 1, item.branch, item.amount, item.billNo, item.remarks, item.date, item.billDate, ageingDays];
    });

    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Finance Report');
    XLSX.writeFile(wb, 'FINANCE_PENDING_REPORT.xlsx');
  };

  const exportSrToExcel = () => {
    const headers = ['Sl.no', 'Branch', 'Amount', 'Bill No.', 'Remarks', 'Date', 'Bill date', 'Ageing'];
    const data = srRows.map((item, idx) => {
      const d1 = new Date();
      const d2 = new Date(item.billDate);
      const diffTime = d1 - d2;
      const ageingDays = Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
      return [idx + 1, item.branch, item.amount, item.billNo, item.remarks, item.date, item.billDate, ageingDays];
    });

    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'SR Report');
    XLSX.writeFile(wb, 'SR_PENDING_REPORT.xlsx');
  };

  const sendFinanceEmail = () => {
    const storeEmails = financeRows.map(item => {
      const cleanBranch = item.branch.toLowerCase().replace(/[^a-z0-9]/g, '');
      return `${cleanBranch}@happimobiles.com`;
    });
    const uniqueToEmails = [...new Set(storeEmails)].join(';');

    const ccEmails = 'dfm@happimobiles.com;jfm@happimobiles.com;agmfinance@happimobiles.com;financesupport@happimobiles.com;financesupport1@happimobiles.com;financesupport3@happimobiles.com;affordability@happimobiles.com';
    const subject = encodeURIComponent('Finance and card amount pending in cash book');

    let htmlTable = `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; font-family:Segoe UI, sans-serif; font-size:12px;">`;
    htmlTable += `<tr style="background:#facc15; font-weight:bold; text-align:center;">
      <th>Sl.no</th><th>Branch</th><th>Amount</th><th>Bill No.</th><th>Remarks</th><th>Date</th><th>Bill date</th><th>Ageing</th>
    </tr>`;

    financeRows.forEach((item, idx) => {
      const d1 = new Date();
      const d2 = new Date(item.billDate);
      const diffTime = d1 - d2;
      const ageingDays = Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
      htmlTable += `<tr>
        <td style="text-align:center;">${idx + 1}</td>
        <td>${item.branch}</td>
        <td style="text-align:right;">₹${item.amount.toLocaleString('en-IN')}</td>
        <td>${item.billNo || ''}</td>
        <td>${item.remarks || ''}</td>
        <td style="text-align:center;">${item.date}</td>
        <td style="text-align:center;">${item.billDate}</td>
        <td style="text-align:center; font-weight:bold; color:red;">${ageingDays}</td>
      </tr>`;
    });
    htmlTable += `</table>`;

    if (navigator.clipboard && window.ClipboardItem) {
      const blobHtml = new Blob([htmlTable], { type: 'text/html' });
      const blobText = new Blob([htmlTable.replace(/<[^>]*>?/gm, '')], { type: 'text/plain' });
      navigator.clipboard.write([
        new ClipboardItem({ 'text/html': blobHtml, 'text/plain': blobText })
      ]).then(() => {
        alert('✅ Finance table copied to clipboard! Outlook is opening. Just press Ctrl+V in the mail body to paste the table.');
        window.location.href = `mailto:${uniqueToEmails}?cc=${ccEmails}&subject=${subject}`;
      }).catch(() => {
        fallbackCopyAndMail(htmlTable, uniqueToEmails, ccEmails, subject);
      });
    } else {
      fallbackCopyAndMail(htmlTable, uniqueToEmails, ccEmails, subject);
    }
  };

  const sendSrEmail = () => {
    const storeEmails = srRows.map(item => {
      const cleanBranch = item.branch.toLowerCase().replace(/[^a-z0-9]/g, '');
      return `${cleanBranch}@happimobiles.com`;
    });
    const uniqueToEmails = [...new Set(storeEmails)].join(';');

    const ccEmails = 'doatg@happimobiles.com;doaap@happimobiles.com;dfm@happimobiles.com;affordability@happimobiles.com';
    const subject = encodeURIComponent('SR amount pending in cash book');

    let htmlTable = `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse; font-family:Segoe UI, sans-serif; font-size:12px;">`;
    htmlTable += `<tr style="background:#facc15; font-weight:bold; text-align:center;">
      <th>Sl.no</th><th>Branch</th><th>Amount</th><th>Bill No.</th><th>Remarks</th><th>Date</th><th>Bill date</th><th>Ageing</th>
    </tr>`;

    srRows.forEach((item, idx) => {
      const d1 = new Date();
      const d2 = new Date(item.billDate);
      const diffTime = d1 - d2;
      const ageingDays = Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
      htmlTable += `<tr>
        <td style="text-align:center;">${idx + 1}</td>
        <td>${item.branch}</td>
        <td style="text-align:right;">₹${item.amount.toLocaleString('en-IN')}</td>
        <td>${item.billNo || ''}</td>
        <td>${item.remarks || ''}</td>
        <td style="text-align:center;">${item.date}</td>
        <td style="text-align:center;">${item.billDate}</td>
        <td style="text-align:center; font-weight:bold; color:red;">${ageingDays}</td>
      </tr>`;
    });
    htmlTable += `</table>`;

    if (navigator.clipboard && window.ClipboardItem) {
      const blobHtml = new Blob([htmlTable], { type: 'text/html' });
      const blobText = new Blob([htmlTable.replace(/<[^>]*>?/gm, '')], { type: 'text/plain' });
      navigator.clipboard.write([
        new ClipboardItem({ 'text/html': blobHtml, 'text/plain': blobText })
      ]).then(() => {
        alert('✅ SR table copied to clipboard! Outlook is opening. Just press Ctrl+V in the mail body to paste the table.');
        window.location.href = `mailto:${uniqueToEmails}?cc=${ccEmails}&subject=${subject}`;
      }).catch(() => {
        fallbackCopyAndMail(htmlTable, uniqueToEmails, ccEmails, subject);
      });
    } else {
      fallbackCopyAndMail(htmlTable, uniqueToEmails, ccEmails, subject);
    }
  };

  const fallbackCopyAndMail = (html, to, cc, subj) => {
    const container = document.createElement('div');
    container.innerHTML = html;
    container.style.position = 'fixed';
    container.style.opacity = '0';
    document.body.appendChild(container);
    window.getSelection().removeAllRanges();
    const range = document.createRange();
    range.selectNode(container);
    window.getSelection().addRange(range);
    try {
      document.execCommand('copy');
      alert('✅ Table copied to clipboard! Outlook is opening. Press Ctrl+V in the mail body.');
      window.location.href = `mailto:${to}?cc=${cc}&subject=${subj}`;
    } catch {
      alert('❌ Failed to copy table automatically. Please use Download button.');
    }
    document.body.removeChild(container);
  };

  const exportContraSheet = () => {
    if (!hotInstanceRef.current) return;
    const hot = hotInstanceRef.current;
    const data = hot.getData();

    const now = new Date();
    const yyyymmdd = now.getFullYear().toString() +
      String(now.getMonth() + 1).padStart(2, '0') +
      String(now.getDate()).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');

    const contraHeaders = [
      'Voucher Date', 'Credit Account', 'Debit Account', 'Payment Mode',
      'Amount', 'TXN No', 'TXN Date', 'Bank Narration', 'Branch Short Code',
      'Ref Doc No', 'Ref Doc Date', 'Salesperson Name', 'Product Category', 'Deposited Date'
    ];

    const sampleRow = [
      'ABCD1234', 'ABCD', 'ABCD1234', 'ABCD', 1234,
      'ABCD1234', 'ABCD1234', 'ABCD', 'ABCD1234', 'ABCD1234',
      'ABCD1234', 'ABCD', 'ABCD', 'ABCD1234'
    ];

    const contraRows = [sampleRow];

    data.forEach(row => {
      const code = row[1] ? row[1].toString() : '';
      const branch = row[2] ? row[2].toString() : '';
      const rowText = (code + ' ' + branch).toUpperCase();
      const depositVal = parseToNum(row[4]);

      if (depositVal > 0 && !rowText.includes('TOTAL')) {
        contraRows.push([
          yyyymmdd,
          'Cash A/c.',
          'HDFC BANK OD A/C: 57500001930168',
          'CASH',
          depositVal,
          dd,
          yyyymmdd,
          'Being Entry Passing words Cash deposited into bank',
          code,
          dd,
          yyyymmdd,
          '',
          '',
          yyyymmdd
        ]);
      }
    });

    const ws = XLSX.utils.aoa_to_sheet([contraHeaders, ...contraRows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Contra Export');
    XLSX.writeFile(wb, 'CONTRA_UPLOAD_SHEET.xlsx');
  };

  const exportToExcel = () => {
    if (!hotInstanceRef.current) return;
    const hot = hotInstanceRef.current;
    const rawData = hot.getData();

    const excelRows = rawData.map((row, idx) => {
      const newRow = [...row];
      
      for (let c = 3; c <= 13; c++) {
        const val = newRow[c];
        if (val === null || val === undefined || val === '' || (typeof val === 'string' && val.trim() === '')) {
          newRow[c] = '';
        } else if (val.toString().startsWith('=')) {
          newRow[c] = evaluateFormula(val, hot, idx);
        } else {
          const clean = val.toString().replace(/,/g, '');
          const num = parseFloat(clean);
          newRow[c] = isNaN(num) ? '' : Math.round(num);
        }
      }

      const dVal = parseToNum(newRow[3]);
      const eVal = parseToNum(newRow[4]);
      const fVal = parseToNum(newRow[5]);
      const gVal = parseToNum(newRow[6]);
      const hVal = parseToNum(newRow[7]);
      const iVal = parseToNum(newRow[8]);
      const jVal = parseToNum(newRow[9]);
      const kVal = parseToNum(newRow[10]);
      const lVal = parseToNum(newRow[11]);
      const mVal = parseToNum(newRow[12]);
      const nVal = parseToNum(newRow[13]);

      newRow[14] = dVal - eVal - fVal - gVal - hVal - iVal - jVal - kVal - lVal - mVal - nVal;

      return newRow;
    });

    const ws = XLSX.utils.aoa_to_sheet([EXCEL_HEADERS, ...excelRows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'MASTER REPORT');
    XLSX.writeFile(wb, 'MASTER_CASHBOOK_REPORT.xlsx');
  };

  const saveMasterDataToCloud = async () => {
    if (!hotInstanceRef.current) return;
    isBusyRef.current = true;
    setSyncStatus('💾 Saving...');

    try {
      const hot = hotInstanceRef.current;
      const data = hot.getData();
      const allPayloads = [];

      data.forEach((row) => {
        const branchName = row[2];
        const branchCode = row[1];
        if (branchName) {
          allPayloads.push({
            branch: branchName,
            code: branchCode || '',
            updated_at: new Date().toISOString(),
            opening_balance: row[3]?.toString() || '',
            deposit: row[4]?.toString() || '',
            denomination: row[5]?.toString() || '',
            addings: row[6]?.toString() || '',
            pending_apprvls: row[7]?.toString() || '',
            finance_amnt: row[8]?.toString() || '',
            sr: row[9]?.toString() || '',
            sweeper_salary: row[10]?.toString() || '',
            edits: row[11]?.toString() || '',
            apx_shortage: row[12]?.toString() || '',
            ksp_approvals: row[13]?.toString() || '',
            remarks: row[15]?.toString() || ''
          });
        }
      });

      const { error } = await supabase.from('master_cashbook').upsert(allPayloads, { onConflict: 'branch' });
      if (error) {
        console.error('Supabase Upsert Error:', error);
        alert('❌ Supabase Error: ' + error.message);
        return;
      }

      if (jetChannelRef.current) {
        jetChannelRef.current.send({
          type: 'broadcast',
          event: 'table_jet_update',
          payload: { payloadList: allPayloads }
        });
      }

      alert('💾 Data Saved Permanently to Cloud Successfully!');
    } catch (err) {
      console.error('Save Exception:', err);
      alert('❌ Error saving data: ' + err.message);
    } finally {
      isBusyRef.current = false;
      setSyncStatus('🟢 Jet Synced');
    }
  };

  const clearDrBalances = async () => {
    if (!window.confirm("Are you sure you want to clear all Dr Balances (Opening Balances)?")) return;
    if (!hotInstanceRef.current) return;
    isBusyRef.current = true;
    setSyncStatus('🧹 Clearing...');

    try {
      const hot = hotInstanceRef.current;
      const currentData = hot.getData();
      const updates = [];
      const cloudPayload = [];

      currentData.forEach((row, rIdx) => {
        updates.push([rIdx, 3, '']);

        cloudPayload.push({
          branch: row[2],
          code: row[1],
          updated_at: new Date().toISOString(),
          opening_balance: '',
          deposit: row[4]?.toString() || '',
          denomination: row[5]?.toString() || '',
          addings: row[6]?.toString() || '',
          pending_apprvls: row[7]?.toString() || '',
          finance_amnt: row[8]?.toString() || '',
          sr: row[9]?.toString() || '',
          sweeper_salary: row[10]?.toString() || '',
          edits: row[11]?.toString() || '',
          apx_shortage: row[12]?.toString() || '',
          ksp_approvals: row[13]?.toString() || '',
          remarks: row[15]?.toString() || ''
        });
      });

      if (updates.length > 0) {
        hot.setDataAtCell(updates);
        hot.render();
        updatePendingStores(hot);

        const { error } = await supabase.from('master_cashbook').upsert(cloudPayload, { onConflict: 'branch' });
        if (error) throw error;

        if (jetChannelRef.current) {
          jetChannelRef.current.send({
            type: 'broadcast',
            event: 'table_jet_update',
            payload: { payloadList: cloudPayload }
          });
        }

        alert('✅ All Dr Balances Cleared Successfully!');
      }
    } catch (err) {
      console.error('Clear Dr Balances Error:', err);
      alert('❌ Error clearing balances: ' + err.message);
    } finally {
      isBusyRef.current = false;
      setSyncStatus('🟢 Jet Synced');
    }
  };

  useEffect(() => {
    if (!containerRef.current) return;

    const initialData = DEFAULT_BRANCHES.map((b, idx) => {
      const rNum = idx + 1;
      const closingFormula = `=D${rNum}-E${rNum}-F${rNum}-G${rNum}-H${rNum}-I${rNum}-J${rNum}-K${rNum}-L${rNum}-M${rNum}-N${rNum}`;
      return [
        idx + 1,
        b.CODE,
        b.BRANCH,
        '', '', '', '', '', '', '', '', '', '', '',
        closingFormula,
        ''
      ];
    });

    const formulaDisplayRenderer = (instance, td, row, col, prop, value, cellProperties) => {
      td.innerHTML = '';
      if (value !== null && value !== undefined && value !== '') {
        const strVal = value.toString().trim();
        if (strVal.startsWith('=')) {
          const calculated = evaluateFormula(strVal, instance, row);
          td.innerText = calculated.toLocaleString('en-IN');
          td.style.color = '#004085';
          td.style.fontWeight = '500';
        } else if (!isNaN(parseFloat(strVal.replace(/,/g, '')))) {
          const numVal = parseToNum(strVal);
          td.innerText = numVal.toLocaleString('en-IN');
        } else {
          td.innerText = strVal;
        }
      } else {
        td.innerText = '';
      }
      td.style.textAlign = col >= 3 && col <= 14 ? 'right' : 'left';
    };

    const hot = new Handsontable(containerRef.current, {
      data: initialData,
      colHeaders: COL_HEADERS,
      rowHeaders: true,
      height: '100%',
      width: '100%',
      undo: true,
      fillHandle: {
        direction: 'vertical',
        autoInsertRow: false,
      },
      autoWrapRow: false,
      autoWrapCol: false,
      outsideClickDeselect: false,
      manualColumnResize: true,
      manualRowResize: true,
      contextMenu: true,
      customBorders: true,
      hiddenRows: {
        rows: [],
        indicators: false
      },

      afterChange: async (changes, source) => {
        const hotInstance = hotInstanceRef.current;
        if (!hotInstance) return;
        updatePendingStores(hotInstance);
        if (!changes || source === 'loadData' || source === 'autoCalc' || source === 'cloudSync' || source === 'jetBroadcast' || isBusyRef.current) return;

        setSyncStatus('🟡 Syncing...');
        hotInstance.render();

        if (debounceTimerRef.current) clearTimeout(debounceTimerRef.current);

        debounceTimerRef.current = setTimeout(async () => {
          const branchPayloadMap = {};
          const currentData = hotInstance.getData();

          currentData.forEach((row) => {
            const branchName = row[2];
            const branchCode = row[1];

            if (branchName) {
              branchPayloadMap[branchName] = {
                branch: branchName,
                code: branchCode || '',
                updated_at: new Date().toISOString(),
                opening_balance: row[3]?.toString() || '',
                deposit: row[4]?.toString() || '',
                denomination: row[5]?.toString() || '',
                addings: row[6]?.toString() || '',
                pending_apprvls: row[7]?.toString() || '',
                finance_amnt: row[8]?.toString() || '',
                sr: row[9]?.toString() || '',
                sweeper_salary: row[10]?.toString() || '',
                edits: row[11]?.toString() || '',
                apx_shortage: row[12]?.toString() || '',
                ksp_approvals: row[13]?.toString() || '',
                remarks: row[15]?.toString() || ''
              };
            }
          });

          const payloadList = Object.values(branchPayloadMap);
          if (payloadList.length > 0) {
            isBusyRef.current = true;
            try {
              await supabase.from('master_cashbook').upsert(payloadList, { onConflict: 'branch' });

              if (jetChannelRef.current) {
                jetChannelRef.current.send({
                  type: 'broadcast',
                  event: 'table_jet_update',
                  payload: { payloadList }
                });
              }
            } catch (err) {
              console.error('Batch Sync Error:', err);
            } finally {
              isBusyRef.current = false;
              setSyncStatus('🟢 Jet Synced');
            }
          }
        }, 1000);
      },

      cells: (row, col) => {
        const cellProperties = {};
        if (col === 0 || col === 1 || col === 2 || col === 14) {
          cellProperties.readOnly = true;
          cellProperties.className = 'htCenter htBold';
          if (col === 14) cellProperties.className = 'htRight htBold';
        }
        if (col >= 3 && col <= 14) {
          cellProperties.renderer = formulaDisplayRenderer;
        }
        return cellProperties;
      },
      licenseKey: 'non-commercial-and-evaluation'
    });

    hotInstanceRef.current = hot;
    setTotalStores(hot.countRows());
    updatePendingStores(hot);

    const loadFromCloud = async () => {
      if (isBusyRef.current) return;
      try {
        const { data, error } = await supabase.from('master_cashbook').select('*').limit(1000);
        if (!error && data && data.length > 0 && !isBusyRef.current) {
          const cloudMap = {};
          data.forEach(item => { 
            if (item.code) {
              cloudMap[item.code.toString().trim().toUpperCase()] = item;
            }
            if (item.branch) {
              cloudMap[item.branch.toString().trim().toUpperCase()] = item; 
            }
          });

          const currentGrid = hot.getData();
          let cellUpdates = [];

          currentGrid.forEach((row, idx) => {
            const storeCode = row[1] ? row[1].toString().trim().toUpperCase() : '';
            const bName = row[2] ? row[2].toString().trim().toUpperCase() : '';
            const cloudRow = cloudMap[storeCode] || cloudMap[bName];
            const rNum = idx + 1;
            const closingFormula = `=D${rNum}-E${rNum}-F${rNum}-G${rNum}-H${rNum}-I${rNum}-J${rNum}-K${rNum}-L${rNum}-M${rNum}-N${rNum}`;

            if (cloudRow) {
              const targetValues = [
                cloudRow.opening_balance || '',
                cloudRow.deposit || '',
                cloudRow.denomination || '',
                cloudRow.addings || '',
                cloudRow.pending_apprvls || '',
                cloudRow.finance_amnt || '',
                cloudRow.sr || '',
                cloudRow.sweeper_salary || '',
                cloudRow.edits || '',
                cloudRow.apx_shortage || '',
                cloudRow.ksp_approvals || ''
              ];

              const targetCols = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13];
              targetCols.forEach((cIdx, i) => {
                if (row[cIdx] !== targetValues[i]) {
                  cellUpdates.push([idx, cIdx, targetValues[i]]);
                }
              });

              row[14] = closingFormula;

              if (row[15] !== (cloudRow.remarks || '')) {
                cellUpdates.push([idx, 15, cloudRow.remarks || '']);
              }
            }
          });

          if (cellUpdates.length > 0 && !isBusyRef.current) {
            hot.setDataAtCell(cellUpdates, 'cloudSync');
            hot.render();
          }
          setTotalStores(hot.countRows());
          updatePendingStores(hot);
          setSyncStatus('🟢 Jet Synced');
        }
      } catch (err) {
        console.warn('Initial load handled smoothly:', err);
      }
    };

    loadFromCloud();

    if (!jetChannelRef.current) {
      jetChannelRef.current = supabase.channel('happi_jet_sync_v13', {
        config: { broadcast: { self: false } }
      });

      jetChannelRef.current
        .on('broadcast', { event: 'table_jet_update' }, ({ payload }) => {
          if (!payload || !payload.payloadList || !hotInstanceRef.current || isBusyRef.current) return;
          
          isBusyRef.current = true;
          const hotInstance = hotInstanceRef.current;
          const currentData = hotInstance.getData();
          let cellUpdates = [];

          payload.payloadList.forEach(item => {
            const rIdx = currentData.findIndex(row => {
              const rCode = row[1] ? row[1].toString().trim().toUpperCase() : '';
              const iCode = item.code ? item.code.toString().trim().toUpperCase() : '';
              const rBranch = row[2] ? row[2].toString().trim().toUpperCase() : '';
              const iBranch = item.branch ? item.branch.toString().trim().toUpperCase() : '';
              return (rCode && iCode && rCode === iCode) || (rBranch && iBranch && rBranch === iBranch);
            });

            if (rIdx !== -1) {
              Object.keys(FIELD_TO_COL_MAP).forEach(field => {
                const cIdx = FIELD_TO_COL_MAP[field];
                const newVal = item[field] || '';
                if (currentData[rIdx][cIdx] !== newVal) {
                  cellUpdates.push([rIdx, cIdx, newVal]);
                }
              });
            }
          });

          if (cellUpdates.length > 0) {
            hotInstance.setDataAtCell(cellUpdates, 'jetBroadcast');
            hotInstance.render();
            updatePendingStores(hotInstance);
          }
          isBusyRef.current = false;
        })
        .subscribe((status) => {
          if (status === 'SUBSCRIBED') {
            console.log('✅ Realtime Jet Channel Connected Successfully!');
          }
        });
    }

    let altSequence = [];
    let altTimer = null;

    const handleKeyDown = (e) => {
      const hotInstance = hotInstanceRef.current;
      if (!hotInstance) return;

      if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) {
        e.preventDefault();
        e.stopPropagation();
        setShowFindModal(true);
        setTimeout(() => findInputRef.current?.focus(), 50);
        return;
      }

      if (e.key === 'Escape') {
        setShowFindModal(false);
        setShowCombinedHover(false);
        setShowApprovalsHover(false);
        setShowEditsHover(false);
        setShowLowCashModal(false);
        setShowFinanceModal(false);
        setShowSrModal(false);
      }

      const selected = hotInstance.getSelected();
      if (!selected || selected.length === 0) return;

      const [r1, c1, r2, c2] = selected[0];
      const startRow = Math.min(r1, r2);
      const endRow = Math.max(r1, r2);
      const startCol = Math.min(c1, c2);
      const endCol = Math.max(c1, c2);

      if (e.key === 'Delete' || e.key === 'Backspace') {
        const isEditing = hotInstance.getActiveEditor()?.isOpened();
        if (!isEditing) {
          e.preventDefault();
          e.stopPropagation();

          const clearUpdates = [];
          const branchPayloadMap = {};

          for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
              if ((c >= 3 && c <= 13) || c === 15) {
                clearUpdates.push([r, c, '']);

                const branchName = hotInstance.getDataAtCell(r, 2);
                const branchCode = hotInstance.getDataAtCell(r, 1);
                if (branchName) {
                  if (!branchPayloadMap[branchName]) {
                    branchPayloadMap[branchName] = {
                      branch: branchName,
                      code: branchCode,
                      updated_at: new Date().toISOString(),
                      opening_balance: hotInstance.getDataAtCell(r, 3)?.toString() || '',
                      deposit: hotInstance.getDataAtCell(r, 4)?.toString() || '',
                      denomination: hotInstance.getDataAtCell(r, 5)?.toString() || '',
                      addings: hotInstance.getDataAtCell(r, 6)?.toString() || '',
                      pending_apprvls: hotInstance.getDataAtCell(r, 7)?.toString() || '',
                      finance_amnt: hotInstance.getDataAtCell(r, 8)?.toString() || '',
                      sr: hotInstance.getDataAtCell(r, 9)?.toString() || '',
                      sweeper_salary: hotInstance.getDataAtCell(r, 10)?.toString() || '',
                      edits: hotInstance.getDataAtCell(r, 11)?.toString() || '',
                      apx_shortage: hotInstance.getDataAtCell(r, 12)?.toString() || '',
                      ksp_approvals: hotInstance.getDataAtCell(r, 13)?.toString() || '',
                      remarks: hotInstance.getDataAtCell(r, 15)?.toString() || ''
                    };
                    for (let cl = startCol; cl <= endCol; cl++) {
                      const fName = DB_FIELD_MAP[cl];
                      if (fName) branchPayloadMap[branchName][fName] = '';
                    }
                  }
                }
              }
            }
          }

          if (clearUpdates.length > 0) {
            hotInstance.setDataAtCell(clearUpdates);
            hotInstance.render();
            updatePendingStores(hotInstance);

            const payloadList = Object.values(branchPayloadMap);
            isBusyRef.current = true;
            supabase.from('master_cashbook').upsert(payloadList, { onConflict: 'branch' }).then(() => {
              isBusyRef.current = false;
            });

            if (jetChannelRef.current) {
              jetChannelRef.current.send({
                type: 'broadcast',
                event: 'table_jet_update',
                payload: { payloadList }
              });
            }
          }
          return;
        }
      }

      if ((e.ctrlKey || e.metaKey) && (e.key === 'd' || e.key === 'D')) {
        e.preventDefault();
        e.stopPropagation();
        const updates = [];
        if (startRow === endRow && startRow > 0) {
          for (let c = startCol; c <= endCol; c++) {
            if (c !== 14) updates.push([startRow, c, hotInstance.getDataAtCell(startRow - 1, c)]);
          }
        } else if (endRow > startRow) {
          for (let c = startCol; c <= endCol; c++) {
            if (c !== 14) {
              const topVal = hotInstance.getDataAtCell(startRow, c);
              for (let r = startRow + 1; r <= endRow; r++) updates.push([r, c, topVal]);
            }
          }
        }
        if (updates.length > 0) {
          hotInstance.setDataAtCell(updates);
          hotInstance.render();
          updatePendingStores(hotInstance);
        }
        return;
      }

      if ((e.ctrlKey || e.metaKey) && (e.key === 'r' || e.key === 'R')) {
        e.preventDefault();
        e.stopPropagation();
        const updates = [];
        for (let r = startRow; r <= endRow; r++) {
          const leftVal = hotInstance.getDataAtCell(r, startCol);
          for (let c = startCol + 1; c <= endCol; c++) {
            if (c !== 14) updates.push([r, c, leftVal]);
          }
        }
        if (updates.length > 0) {
          hotInstance.setDataAtCell(updates);
          hotInstance.render();
          updatePendingStores(hotInstance);
        }
        return;
      }

      if (e.altKey) {
        altSequence = ['ALT'];
        clearTimeout(altTimer);
        altTimer = setTimeout(() => { altSequence = []; }, 2000);
      } else if (altSequence.length > 0) {
        altSequence.push(e.key.toUpperCase());
        const seq = altSequence.join('');
        if (seq.includes('ALTHBA')) {
          e.preventDefault();
          const customBordersPlugin = hotInstance.getPlugin('customBorders');
          const range = [{ from: { row: startRow, col: startCol }, to: { row: endRow, col: endCol } }];
          customBordersPlugin.setBorder(range, {
            top: { width: 1, color: '#000' }, bottom: { width: 1, color: '#000' },
            left: { width: 1, color: '#000' }, right: { width: 1, color: '#000' }
          });
          altSequence = [];
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown, true);
    return () => {
      window.removeEventListener('keydown', handleKeyDown, true);
      if (jetChannelRef.current) {
        supabase.removeChannel(jetChannelRef.current);
        jetChannelRef.current = null;
      }
      hot.destroy();
    };
  }, []);

  const handleCashierFilterChange = (filterVal) => {
    setCashierFilter(filterVal);
    const hot = hotInstanceRef.current;
    if (!hot) return;

    const total = hot.countRows();
    setTotalStores(total);

    const half = Math.ceil(total / 2);

    let rowsToHide = [];
    if (filterVal === 'C1') {
      for (let i = half; i < total; i++) {
        rowsToHide.push(i);
      }
    } else if (filterVal === 'C2') {
      for (let i = 0; i < half; i++) {
        rowsToHide.push(i);
      }
    }

    hot.updateSettings({
      hiddenRows: {
        rows: rowsToHide,
        indicators: false
      }
    });
    hot.render();
  };

  const handleFindNext = (searchQuery) => {
    const q = searchQuery !== undefined ? searchQuery : findText;
    if (!q.trim() || !hotInstanceRef.current) return;

    const hot = hotInstanceRef.current;
    const data = hot.getData();
    const query = q.trim().toUpperCase();
    const matches = [];

    data.forEach((row, r) => {
      row.forEach((cellVal, c) => {
        if (cellVal !== null && cellVal !== undefined) {
          const strVal = cellVal.toString().toUpperCase();
          if (strVal.includes(query)) {
            matches.push({ r, c });
          }
        }
      });
    });

    if (matches.length === 0) {
      alert('Cannot find matching data: ' + q);
      setFindResults([]);
      setCurrentResultIdx(-1);
      return;
    }

    let nextIdx = (currentResultIdx + 1) % matches.length;
    setFindResults(matches);
    setCurrentResultIdx(nextIdx);

    const target = matches[nextIdx];
    hot.selectCell(target.r, target.c);
    hot.scrollViewportTo(target.r, target.c);
  };

  const handleApexUpload = (e) => {
    const file = e.target.files[0];
    if (!file || !hotInstanceRef.current) return;
    isBusyRef.current = true;
    setSyncStatus('🟡 Uploading Apex...');

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws);
        const hot = hotInstanceRef.current;
        const currentData = hot.getData();

        const dumpMap = {};
        rawData.forEach((row) => {
          const values = Object.values(row);
          const branchKey = values[0]?.toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
          const rawAmt = values[1];
          if (branchKey && rawAmt !== undefined) {
            const rawStr = rawAmt.toString().trim();
            const isCr = /cr\.?/i.test(rawStr);
            const cleanStr = rawStr.replace(/dr\.?/gi, '').replace(/cr\.?/gi, '').replace(/,/g, '').trim();
            const num = parseFloat(cleanStr);
            if (!isNaN(num)) {
              const finalNum = isCr ? -Math.round(num) : Math.round(num);
              dumpMap[branchKey] = finalNum;
            }
          }
        });

        const updates = [];
        const cloudPayload = [];

        currentData.forEach((row, rIdx) => {
          const bNorm = row[2]?.toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
          const cNorm = row[1]?.toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
          const val = dumpMap[bNorm] || dumpMap[cNorm];
          if (val !== undefined) {
            const roundedVal = val.toString();
            updates.push([rIdx, 3, roundedVal]);

            cloudPayload.push({
              branch: row[2],
              code: row[1],
              updated_at: new Date().toISOString(),
              opening_balance: roundedVal,
              deposit: row[4]?.toString() || '',
              denomination: row[5]?.toString() || '',
              addings: row[6]?.toString() || '',
              pending_apprvls: row[7]?.toString() || '',
              finance_amnt: row[8]?.toString() || '',
              sr: row[9]?.toString() || '',
              sweeper_salary: row[10]?.toString() || '',
              edits: row[11]?.toString() || '',
              apx_shortage: row[12]?.toString() || '',
              ksp_approvals: row[13]?.toString() || '',
              remarks: row[15]?.toString() || ''
            });
          }
        });

        if (updates.length > 0) {
          hot.setDataAtCell(updates);
          hot.render();
          updatePendingStores(hot);

          if (cloudPayload.length > 0) {
            await supabase.from('master_cashbook').upsert(cloudPayload, { onConflict: 'branch' });
          }

          if (jetChannelRef.current) {
            jetChannelRef.current.send({
              type: 'broadcast',
              event: 'table_jet_update',
              payload: { payloadList: cloudPayload }
            });
          }

          alert('✅ Apex Dr/Cr Balances Uploaded & Cleaned Successfully!');
        } else {
          alert('⚠️ No matching branch opening balances found in file.');
        }
      } catch (err) {
        console.error('Apex Upload Error:', err);
        alert('❌ Apex Upload Error: ' + err.message);
      } finally {
        isBusyRef.current = false;
        setSyncStatus('🟢 Jet Synced');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleAddinsUpload = (e) => {
    const file = e.target.files[0];
    if (!file || !hotInstanceRef.current) return;
    isBusyRef.current = true;
    setSyncStatus('🟡 Uploading Addins...');

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const hot = hotInstanceRef.current;
        const currentData = hot.getData();

        let branchColIdx = 1;
        let valColIdx = 2;

        for (let i = 0; i < Math.min(5, rawData.length); i++) {
          const r = rawData[i];
          if (Array.isArray(r)) {
            r.forEach((cell, idx) => {
              const str = cell ? cell.toString().trim().toUpperCase() : '';
              if (str.includes('BRANCH')) branchColIdx = idx;
              if (str.includes('VOUCHER') || str.includes('VALUE') || str.includes('AMOUNT')) valColIdx = idx;
            });
          }
        }

        const branchVouchersMap = {};
        rawData.forEach((row, idx) => {
          if (idx === 0 && (row[branchColIdx]?.toString().toUpperCase().includes('BRANCH'))) return;
          const branchRaw = row[branchColIdx]?.toString().trim();
          const valRaw = row[valColIdx];

          if (branchRaw && valRaw !== undefined && valRaw !== null && valRaw !== '') {
            const bKey = branchRaw.replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
            const cleanNum = Math.round(parseFloat(valRaw.toString().replace(/,/g, '')) || 0);
            if (cleanNum > 0) {
              if (!branchVouchersMap[bKey]) branchVouchersMap[bKey] = [];
              branchVouchersMap[bKey].push(cleanNum);
            }
          }
        });

        const updates = [];
        const cloudPayload = [];

        currentData.forEach((row, rIdx) => {
          const bNorm = row[2]?.toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
          const cNorm = row[1]?.toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
          const vouchers = branchVouchersMap[bNorm] || branchVouchersMap[cNorm];

          if (vouchers && vouchers.length > 0) {
            const formulaStr = vouchers.length > 1 ? '=' + vouchers.join('+') : vouchers[0].toString();
            updates.push([rIdx, 6, formulaStr]);

            cloudPayload.push({
              branch: row[2],
              code: row[1],
              updated_at: new Date().toISOString(),
              opening_balance: row[3]?.toString() || '',
              deposit: row[4]?.toString() || '',
              denomination: row[5]?.toString() || '',
              addings: formulaStr,
              pending_apprvls: row[7]?.toString() || '',
              finance_amnt: row[8]?.toString() || '',
              sr: row[9]?.toString() || '',
              sweeper_salary: row[10]?.toString() || '',
              edits: row[11]?.toString() || '',
              apx_shortage: row[12]?.toString() || '',
              ksp_approvals: row[13]?.toString() || '',
              remarks: row[15]?.toString() || ''
            });
          }
        });

        if (updates.length > 0) {
          hot.setDataAtCell(updates);
          hot.render();
          updatePendingStores(hot);

          if (cloudPayload.length > 0) {
            await supabase.from('master_cashbook').upsert(cloudPayload, { onConflict: 'branch' });
          }

          if (jetChannelRef.current) {
            jetChannelRef.current.send({
              type: 'broadcast',
              event: 'table_jet_update',
              payload: { payloadList }
            });
          }

          alert('✅ Addins Mapped & Preserved Successfully!');
        } else {
          alert('⚠️ No matching branch vouchers found.');
        }
      } catch (err) {
        console.error('Addins Upload Error:', err);
      } finally {
        isBusyRef.current = false;
        setSyncStatus('🟢 Jet Synced');
      }
    };
    reader.readAsBinaryString(file);
  };

  const halfStores = Math.ceil(totalStores / 2);

  return (
    <div style={{ padding: '10px', fontFamily: 'Segoe UI, sans-serif', height: '100vh', display: 'flex', flexDirection: 'column' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px', flexWrap: 'wrap', gap: '8px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
          <h2 style={{ margin: 0, color: '#107c41', fontSize: '18px', fontWeight: 'bold' }}>
            📊 HAPPI MASTER CASHBOOK
          </h2>
          <span style={{ fontSize: '12px', fontWeight: '600', padding: '4px 8px', borderRadius: '4px', background: '#f1f5f9' }}>
            {syncStatus}
          </span>

          <div style={{ display: 'flex', alignItems: 'center', gap: '6px', background: '#f1f5f9', padding: '3px 8px', borderRadius: '4px' }}>
            <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#334155' }}>👤 Filter:</span>
            <select
              value={cashierFilter}
              onChange={(e) => handleCashierFilterChange(e.target.value)}
              style={{
                padding: '4px 8px',
                borderRadius: '4px',
                border: '1px solid #cbd5e1',
                fontSize: '12px',
                fontWeight: 'bold',
                background: '#ffffff',
                cursor: 'pointer',
                outline: 'none',
                color: '#0f172a'
              }}
            >
              <option value="ALL">🏢 All Stores ({totalStores})</option>
              <option value="C1">👨‍💼 Cashier 1 (1 to {halfStores})</option>
              <option value="C2">👨‍💼 Cashier 2 ({halfStores + 1} to {totalStores})</option>
            </select>
          </div>

          <label style={{ background: '#0e4c92', color: '#fff', padding: '6px 12px', borderRadius: '4px', cursor: 'pointer', fontSize: '13px', fontWeight: 'bold' }}>
            📂 Upload Apex Dr
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleApexUpload} style={{ display: 'none' }} />
          </label>
          <label style={{ background: '#7c3aed', color: '#fff', padding: '6px 12px', borderRadius: '4px', cursor: 'pointer', fontSize: '13px', fontWeight: 'bold' }}>
            📥 Upload Addins Dump
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleAddinsUpload} style={{ display: 'none' }} />
          </label>
          <button
            onClick={() => {
              setShowFindModal(true);
              setTimeout(() => findInputRef.current?.focus(), 50);
            }}
            style={{ background: '#475569', color: '#fff', border: 'none', padding: '6px 12px', borderRadius: '4px', cursor: 'pointer', fontSize: '13px', fontWeight: 'bold' }}
          >
            🔍 Find (Ctrl+F)
          </button>
        </div>

        <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap' }}>
          {/* Explicit Save Master Data Button */}
          <button
            onClick={saveMasterDataToCloud}
            style={{ background: '#2563eb', color: '#fff', border: 'none', padding: '6px 14px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}
          >
            💾 Save Data
          </button>

          {/* Clear Dr Balances Button */}
          <button
            onClick={clearDrBalances}
            style={{ background: '#dc2626', color: '#fff', border: 'none', padding: '6px 14px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}
          >
            🧹 Clear Dr Balances
          </button>

          {/* Finance Report Button */}
          <button
            onClick={openFinanceReportModal}
            style={{ background: '#0f766e', color: '#fff', border: 'none', padding: '6px 14px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}
          >
            💳 Finance Report
          </button>

          {/* SR Report Button */}
          <button
            onClick={openSrReportModal}
            style={{ background: '#7c3aed', color: '#fff', border: 'none', padding: '6px 14px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}
          >
            📦 SR Report
          </button>

          {/* Pending Approvals Hover Preview Button */}
          <div
            style={{ position: 'relative', display: 'inline-block' }}
            onMouseEnter={() => setShowApprovalsHover(true)}
            onMouseLeave={() => setShowApprovalsHover(false)}
          >
            <button
              style={{ background: '#0369a1', color: '#fff', border: 'none', padding: '6px 14px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}
            >
              📌 Pending Approvals ({pendingApprovalsList.length})
            </button>
            {showApprovalsHover && (
              <div style={{
                position: 'absolute', top: '100%', right: 0, marginTop: '4px',
                background: '#fff', border: '1px solid #cbd5e1', borderRadius: '6px',
                boxShadow: '0 10px 25px rgba(0,0,0,0.2)', width: '320px', maxHeight: '350px',
                overflowY: 'auto', zIndex: 99999, padding: '10px'
              }}>
                <div style={{ fontWeight: 'bold', fontSize: '12px', color: '#0369a1', marginBottom: '6px', borderBottom: '1px solid #e2e8f0', paddingBottom: '4px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span>📌 Pending Approvals List</span>
                  <button
                    onClick={exportPendingApprovalsToExcel}
                    style={{ background: '#107c41', color: '#fff', border: 'none', padding: '3px 8px', borderRadius: '4px', cursor: 'pointer', fontSize: '11px', fontWeight: 'bold' }}
                  >
                    📥 Excel
                  </button>
                </div>
                {pendingApprovalsList.length === 0 ? (
                  <div style={{ fontSize: '11px', color: '#16a34a', padding: '10px 0', textAlign: 'center', fontWeight: 'bold' }}>
                    🎉 No pending approvals!
                  </div>
                ) : (
                  pendingApprovalsList.map((item, idx) => (
                    <div key={idx} style={{ fontSize: '11px', padding: '4px 6px', borderBottom: '1px solid #f8fafc', color: '#1e293b', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span><b>{item.code}</b> - {item.branch}</span>
                      <span style={{ color: '#0369a1', fontWeight: 'bold' }}>₹{item.amount.toLocaleString('en-IN')}</span>
                    </div>
                  ))
                )}
              </div>
            )}
          </div>

          {/* Edits Hover Preview Button */}
          <div
            style={{ position: 'relative', display: 'inline-block' }}
            onMouseEnter={() => setShowEditsHover(true)}
            onMouseLeave={() => setShowEditsHover(false)}
          >
            <button
              style={{ background: '#475569', color: '#fff', border: 'none', padding: '6px 14px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}
            >
              ✏️ Edits ({editsList.length})
            </button>
            {showEditsHover && (
              <div style={{
                position: 'absolute', top: '100%', right: 0, marginTop: '4px',
                background: '#fff', border: '1px solid #cbd5e1', borderRadius: '6px',
                boxShadow: '0 10px 25px rgba(0,0,0,0.2)', width: '320px', maxHeight: '350px',
                overflowY: 'auto', zIndex: 99999, padding: '10px'
              }}>
                <div style={{ fontWeight: 'bold', fontSize: '12px', color: '#475569', marginBottom: '6px', borderBottom: '1px solid #e2e8f0', paddingBottom: '4px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span>✏️ Edits List</span>
                  <button
                    onClick={exportEditsToExcel}
                    style={{ background: '#107c41', color: '#fff', border: 'none', padding: '3px 8px', borderRadius: '4px', cursor: 'pointer', fontSize: '11px', fontWeight: 'bold' }}
                  >
                    📥 Excel
                  </button>
                </div>
                {editsList.length === 0 ? (
                  <div style={{ fontSize: '11px', color: '#16a34a', padding: '10px 0', textAlign: 'center', fontWeight: 'bold' }}>
                    🎉 No edits pending!
                  </div>
                ) : (
                  editsList.map((item, idx) => (
                    <div key={idx} style={{ fontSize: '11px', padding: '4px 6px', borderBottom: '1px solid #f8fafc', color: '#1e293b', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span><b>{item.code}</b> - {item.branch}</span>
                      <span style={{ color: '#475569', fontWeight: 'bold' }}>₹{item.amount.toLocaleString('en-IN')}</span>
                    </div>
                  ))
                )}
              </div>
            )}
          </div>

          <button
            onClick={exportContraSheet}
            style={{ background: '#d97706', color: '#fff', border: 'none', padding: '6px 14px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}
          >
            📥 Download Contra Sheet
          </button>

          {/* Combined Pending Status Hover Preview */}
          <div
            style={{ position: 'relative', display: 'inline-block' }}
            onMouseEnter={() => setShowCombinedHover(true)}
            onMouseLeave={() => setShowCombinedHover(false)}
          >
            <button
              style={{ background: '#7c2d12', color: '#fff', border: 'none', padding: '6px 14px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}
            >
              📋 Pending Status (Cash: {pendingStoresList.length} | Slips: {pendingDepositSlipsList.length})
            </button>
            {showCombinedHover && (
              <div style={{
                position: 'absolute', top: '100%', right: 0, marginTop: '4px',
                background: '#fff', border: '1px solid #cbd5e1', borderRadius: '6px',
                boxShadow: '0 10px 25px rgba(0,0,0,0.2)', width: '340px', maxHeight: '380px',
                overflowY: 'auto', zIndex: 99999, padding: '10px'
              }}>
                {/* Section 1: Pending Cashbooks */}
                <div style={{ fontWeight: 'bold', fontSize: '12px', color: '#b91c1c', marginBottom: '4px', borderBottom: '1px solid #e2e8f0', paddingBottom: '4px', display: 'flex', justifyContent: 'space-between' }}>
                  <span>📌 Pending Cashbooks</span>
                  <span style={{ background: '#fee2e2', color: '#b91c1c', padding: '1px 6px', borderRadius: '10px', fontSize: '11px' }}>{pendingStoresList.length}</span>
                </div>
                {pendingStoresList.length === 0 ? (
                  <div style={{ fontSize: '11px', color: '#16a34a', padding: '6px 0', textAlign: 'center', fontWeight: 'bold' }}>
                    🎉 All cashbooks submitted!
                  </div>
                ) : (
                  pendingStoresList.map((item, idx) => (
                    <div key={idx} style={{ fontSize: '11px', padding: '3px 6px', borderBottom: '1px solid #f8fafc', color: '#1e293b', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span><b>{item.code}</b> - {item.branch}</span>
                      <span style={{ color: '#ef4444', fontWeight: '500' }}>Not Mailed</span>
                    </div>
                  ))
                )}

                {/* Section 2: Pending Deposit Slips */}
                <div style={{ fontWeight: 'bold', fontSize: '12px', color: '#ca8a04', marginTop: '10px', marginBottom: '4px', borderBottom: '1px solid #e2e8f0', paddingBottom: '4px', display: 'flex', justifyContent: 'space-between' }}>
                  <span>📄 Pending Deposit Slips (Denom Amt)</span>
                  <span style={{ background: '#fef9c3', color: '#ca8a04', padding: '1px 6px', borderRadius: '10px', fontSize: '11px' }}>{pendingDepositSlipsList.length}</span>
                </div>
                {pendingDepositSlipsList.length === 0 ? (
                  <div style={{ fontSize: '11px', color: '#16a34a', padding: '6px 0', textAlign: 'center', fontWeight: 'bold' }}>
                    🎉 No deposit slips pending!
                  </div>
                ) : (
                  pendingDepositSlipsList.map((item, idx) => {
                    const amtStr = item.denomVal ? parseToNum(item.denomVal).toLocaleString('en-IN') : '0';
                    return (
                      <div key={idx} style={{ fontSize: '11px', padding: '3px 6px', borderBottom: '1px solid #f8fafc', color: '#1e293b', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <span><b>{item.code}</b> - {item.branch}</span>
                        <span style={{ color: '#ca8a04', fontWeight: 'bold' }}>₹{amtStr}</span>
                      </div>
                    );
                  })
                )}
              </div>
            )}
          </div>

          {/* Low Cash & No Cash Report Button (Click to Open) */}
          <button
            onClick={() => {
              updatePendingStores(hotInstanceRef.current);
              setShowLowCashModal(true);
            }}
            style={{ background: '#b45309', color: '#fff', border: 'none', padding: '6px 14px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}
          >
            ⚠️ Low/No Cash Report ({lowNoCashList.length})
          </button>

          {showLowCashModal && (
            <div style={{
              position: 'fixed',
              top: '50%',
              left: '50%',
              transform: 'translate(-50%, -50%)',
              background: '#ffffff',
              border: '1px solid #cbd5e1',
              borderRadius: '8px',
              boxShadow: '0 25px 50px rgba(0,0,0,0.35)',
              width: '600px',
              maxHeight: '75vh',
              overflowY: 'auto',
              zIndex: 999999,
              padding: '20px',
              fontFamily: 'Segoe UI, sans-serif'
            }}>
              <div style={{ fontSize: '16px', fontWeight: 'bold', color: '#0f172a', marginBottom: '16px', borderBottom: '2px solid #e2e8f0', paddingBottom: '8px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>⚠️ Low Cash & No Cash Report</span>
                <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                  <button
                    onClick={exportLowCashToExcel}
                    style={{ background: '#107c41', color: '#fff', border: 'none', padding: '4px 10px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '11px' }}
                  >
                    📥 Download Excel
                  </button>
                  <button
                    onClick={() => setShowLowCashModal(false)}
                    style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontWeight: 'bold', fontSize: '16px', color: '#64748b' }}
                  >
                    ✕
                  </button>
                </div>
              </div>

              <div style={{ fontSize: '12px', color: '#64748b', marginBottom: '12px' }}>
                Press <b>Esc</b> on keyboard to close this window.
              </div>

              <div style={{ border: '1px solid #e2e8f0', borderRadius: '6px', overflow: 'hidden' }}>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 100px 1fr', background: '#f8fafc', padding: '8px 12px', fontWeight: 'bold', fontSize: '12px', color: '#334155', borderBottom: '1px solid #e2e8f0' }}>
                  <span>Store Name</span>
                  <span>Denom Amt</span>
                  <span>Status / Remarks</span>
                </div>
                <div style={{ maxHeight: '400px', overflowY: 'auto' }}>
                  {lowNoCashList.length === 0 ? (
                    <div style={{ fontSize: '12px', color: '#16a34a', padding: '20px', textAlign: 'center', fontWeight: 'bold' }}>
                      🎉 No low/no cash stores pending deposit!
                    </div>
                  ) : (
                    lowNoCashList.map((item, idx) => {
                      const denomNum = item.denom;
                      const isZero = denomNum === 0;
                      const isLow = denomNum < 500;

                      if (isZero) {
                        return (
                          <div key={idx} style={{ display: 'grid', gridTemplateColumns: '1fr 100px 1fr', padding: '8px 12px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', fontSize: '12px' }}>
                            <span style={{ fontWeight: '500' }}>{item.branch}</span>
                            <span style={{ fontWeight: '600', color: '#0f172a' }}>₹{denomNum.toLocaleString('en-IN')}</span>
                            <span>
                              <div style={{ width: '100%', padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: '4px', fontSize: '12px', fontWeight: '600', backgroundColor: '#fee2e2', color: '#dc2626' }}>
                                No Cash
                              </div>
                            </span>
                          </div>
                        );
                      } else if (isLow) {
                        return (
                          <div key={idx} style={{ display: 'grid', gridTemplateColumns: '1fr 100px 1fr', padding: '8px 12px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', fontSize: '12px' }}>
                            <span style={{ fontWeight: '500' }}>{item.branch}</span>
                            <span style={{ fontWeight: '600', color: '#0f172a' }}>₹{denomNum.toLocaleString('en-IN')}</span>
                            <span>
                              <div style={{ width: '100%', padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: '4px', fontSize: '12px', fontWeight: '600', backgroundColor: '#fef3c7', color: '#d97706' }}>
                                Low Cash
                              </div>
                            </span>
                          </div>
                        );
                      } else {
                        const currentText = lowCashRemarksState[item.code] || '';
                        const bgColor = currentText.trim() !== '' ? '#ffedd5' : '#ffffff';
                        return (
                          <div key={idx} style={{ display: 'grid', gridTemplateColumns: '1fr 100px 1fr', padding: '8px 12px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', fontSize: '12px' }}>
                            <span style={{ fontWeight: '500' }}>{item.branch}</span>
                            <span style={{ fontWeight: '600', color: '#0f172a' }}>₹{denomNum.toLocaleString('en-IN')}</span>
                            <span>
                              <input
                                type="text"
                                value={currentText}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  setLowCashRemarksState(prev => ({ ...prev, [item.code]: val }));
                                }}
                                placeholder="Type remarks..."
                                style={{ width: '100%', padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: '4px', fontSize: '12px', outline: 'none', backgroundColor: bgColor }}
                              />
                            </span>
                          </div>
                        );
                      }
                    })
                  )}
                </div>
              </div>
            </div>
          )}

          {/* Finance Report Modal */}
          {showFinanceModal && (
            <div style={{
              position: 'fixed',
              top: '50%',
              left: '50%',
              transform: 'translate(-50%, -50%)',
              background: '#ffffff',
              border: '1px solid #cbd5e1',
              borderRadius: '8px',
              boxShadow: '0 25px 50px rgba(0,0,0,0.35)',
              width: '880px',
              maxHeight: '85vh',
              overflowY: 'auto',
              zIndex: 999999,
              padding: '20px',
              fontFamily: 'Segoe UI, sans-serif'
            }}>
              <div style={{ fontSize: '16px', fontWeight: 'bold', color: '#0f172a', marginBottom: '16px', borderBottom: '2px solid #e2e8f0', paddingBottom: '8px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>💳 Finance and Card Amount Pending in Cash Book</span>
                <button
                  onClick={() => setShowFinanceModal(false)}
                  style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontWeight: 'bold', fontSize: '16px', color: '#64748b' }}
                >
                  ✕
                </button>
              </div>

              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
                <span style={{ fontSize: '12px', color: '#64748b' }}>
                  Press <b>Esc</b> to close. Edit Bill No, Remarks & Bill date directly.
                </span>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <button
                    onClick={exportFinanceToExcel}
                    style={{ background: '#107c41', color: '#fff', border: 'none', padding: '6px 12px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '12px' }}
                  >
                    📥 Download Excel
                  </button>
                  <button
                    onClick={sendFinanceEmail}
                    style={{ background: '#2563eb', color: '#fff', border: 'none', padding: '6px 12px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '12px' }}
                  >
                    ✉️ Send Mail (Outlook)
                  </button>
                </div>
              </div>

              <div style={{ border: '1px solid #e2e8f0', borderRadius: '6px', overflow: 'hidden' }}>
                <div style={{ display: 'grid', gridTemplateColumns: '50px 1.5fr 100px 130px 1.2fr 90px 110px 70px', background: '#facc15', padding: '8px 10px', fontWeight: 'bold', fontSize: '12px', color: '#1e293b', borderBottom: '1px solid #cbd5e1', textAlign: 'center' }}>
                  <span>Sl.no</span>
                  <span>Branch</span>
                  <span>Amount</span>
                  <span>Bill No.</span>
                  <span>Remarks</span>
                  <span>Date</span>
                  <span>Bill date</span>
                  <span>Ageing</span>
                </div>
                <div style={{ maxHeight: '380px', overflowY: 'auto' }}>
                  {financeRows.length === 0 ? (
                    <div style={{ fontSize: '12px', color: '#16a34a', padding: '25px', textAlign: 'center', fontWeight: 'bold' }}>
                      🎉 No finance/card amounts pending in cash book!
                    </div>
                  ) : (
                    financeRows.map((item, idx) => {
                      const d1 = new Date();
                      const d2 = new Date(item.billDate);
                      const diffTime = d1 - d2;
                      const ageingDays = Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));

                      return (
                        <div key={idx} style={{ display: 'grid', gridTemplateColumns: '50px 1.5fr 100px 130px 1.2fr 90px 110px 70px', padding: '6px 10px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', fontSize: '12px', textAlign: 'center' }}>
                          <span style={{ fontWeight: 'bold' }}>{idx + 1}</span>
                          <span style={{ fontWeight: '500', textAlign: 'left', paddingLeft: '4px' }}>{item.branch}</span>
                          <span style={{ fontWeight: '600', textAlign: 'right', paddingRight: '6px' }}>₹{item.amount.toLocaleString('en-IN')}</span>
                          <span>
                            <input
                              type="text"
                              value={item.billNo}
                              onChange={(e) => {
                                const val = e.target.value;
                                setFinanceRows(prev => prev.map((r, i) => i === idx ? { ...r, billNo: val } : r));
                              }}
                              placeholder="Bill No..."
                              style={{ width: '100%', padding: '4px 6px', border: '1px solid #cbd5e1', borderRadius: '4px', fontSize: '11px', outline: 'none' }}
                            />
                          </span>
                          <span>
                            <input
                              type="text"
                              value={item.remarks}
                              onChange={(e) => {
                                const val = e.target.value;
                                setFinanceRows(prev => prev.map((r, i) => i === idx ? { ...r, remarks: val } : r));
                              }}
                              placeholder="Remarks..."
                              style={{ width: '100%', padding: '4px 6px', border: '1px solid #cbd5e1', borderRadius: '4px', fontSize: '11px', outline: 'none' }}
                            />
                          </span>
                          <span style={{ fontSize: '11px', color: '#334155' }}>{item.date}</span>
                          <span>
                            <input
                              type="date"
                              value={item.billDate}
                              onChange={(e) => {
                                const val = e.target.value;
                                setFinanceRows(prev => prev.map((r, i) => i === idx ? { ...r, billDate: val } : r));
                              }}
                              style={{ width: '100%', padding: '3px 4px', border: '1px solid #cbd5e1', borderRadius: '4px', fontSize: '10px', outline: 'none' }}
                            />
                          </span>
                          <span style={{ fontWeight: 'bold', color: '#b91c1c' }}>{ageingDays}</span>
                        </div>
                      );
                    })
                  )}
                </div>
              </div>
            </div>
          )}

          {/* SR Report Modal */}
          {showSrModal && (
            <div style={{
              position: 'fixed',
              top: '50%',
              left: '50%',
              transform: 'translate(-50%, -50%)',
              background: '#ffffff',
              border: '1px solid #cbd5e1',
              borderRadius: '8px',
              boxShadow: '0 25px 50px rgba(0,0,0,0.35)',
              width: '880px',
              maxHeight: '85vh',
              overflowY: 'auto',
              zIndex: 999999,
              padding: '20px',
              fontFamily: 'Segoe UI, sans-serif'
            }}>
              <div style={{ fontSize: '16px', fontWeight: 'bold', color: '#0f172a', marginBottom: '16px', borderBottom: '2px solid #e2e8f0', paddingBottom: '8px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>📦 SR Amount Pending in Cash Book</span>
                <button
                  onClick={() => setShowSrModal(false)}
                  style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontWeight: 'bold', fontSize: '16px', color: '#64748b' }}
                >
                  ✕
                </button>
              </div>

              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
                <span style={{ fontSize: '12px', color: '#64748b' }}>
                  Press <b>Esc</b> to close. Edit Bill No, Remarks & Bill date directly.
                </span>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <button
                    onClick={exportSrToExcel}
                    style={{ background: '#107c41', color: '#fff', border: 'none', padding: '6px 12px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '12px' }}
                  >
                    📥 Download Excel
                  </button>
                  <button
                    onClick={sendSrEmail}
                    style={{ background: '#2563eb', color: '#fff', border: 'none', padding: '6px 12px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '12px' }}
                  >
                    ✉️ Send Mail (Outlook)
                  </button>
                </div>
              </div>

              <div style={{ border: '1px solid #e2e8f0', borderRadius: '6px', overflow: 'hidden' }}>
                <div style={{ display: 'grid', gridTemplateColumns: '50px 1.5fr 100px 130px 1.2fr 90px 110px 70px', background: '#facc15', padding: '8px 10px', fontWeight: 'bold', fontSize: '12px', color: '#1e293b', borderBottom: '1px solid #cbd5e1', textAlign: 'center' }}>
                  <span>Sl.no</span>
                  <span>Branch</span>
                  <span>Amount</span>
                  <span>Bill No.</span>
                  <span>Remarks</span>
                  <span>Date</span>
                  <span>Bill date</span>
                  <span>Ageing</span>
                </div>
                <div style={{ maxHeight: '380px', overflowY: 'auto' }}>
                  {srRows.length === 0 ? (
                    <div style={{ fontSize: '12px', color: '#16a34a', padding: '25px', textAlign: 'center', fontWeight: 'bold' }}>
                      🎉 No SR amounts pending in cash book!
                    </div>
                  ) : (
                    srRows.map((item, idx) => {
                      const d1 = new Date();
                      const d2 = new Date(item.billDate);
                      const diffTime = d1 - d2;
                      const ageingDays = Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));

                      return (
                        <div key={idx} style={{ display: 'grid', gridTemplateColumns: '50px 1.5fr 100px 130px 1.2fr 90px 110px 70px', padding: '6px 10px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', fontSize: '12px', textAlign: 'center' }}>
                          <span style={{ fontWeight: 'bold' }}>{idx + 1}</span>
                          <span style={{ fontWeight: '500', textAlign: 'left', paddingLeft: '4px' }}>{item.branch}</span>
                          <span style={{ fontWeight: '600', textAlign: 'right', paddingRight: '6px' }}>₹{item.amount.toLocaleString('en-IN')}</span>
                          <span>
                            <input
                              type="text"
                              value={item.billNo}
                              onChange={(e) => {
                                const val = e.target.value;
                                setSrRows(prev => prev.map((r, i) => i === idx ? { ...r, billNo: val } : r));
                              }}
                              placeholder="Bill No..."
                              style={{ width: '100%', padding: '4px 6px', border: '1px solid #cbd5e1', borderRadius: '4px', fontSize: '11px', outline: 'none' }}
                            />
                          </span>
                          <span>
                            <input
                              type="text"
                              value={item.remarks}
                              onChange={(e) => {
                                const val = e.target.value;
                                setSrRows(prev => prev.map((r, i) => i === idx ? { ...r, remarks: val } : r));
                              }}
                              placeholder="Remarks..."
                              style={{ width: '100%', padding: '4px 6px', border: '1px solid #cbd5e1', borderRadius: '4px', fontSize: '11px', outline: 'none' }}
                            />
                          </span>
                          <span style={{ fontSize: '11px', color: '#334155' }}>{item.date}</span>
                          <span>
                            <input
                              type="date"
                              value={item.billDate}
                              onChange={(e) => {
                                const val = e.target.value;
                                setSrRows(prev => prev.map((r, i) => i === idx ? { ...r, billDate: val } : r));
                              }}
                              style={{ width: '100%', padding: '3px 4px', border: '1px solid #cbd5e1', borderRadius: '4px', fontSize: '10px', outline: 'none' }}
                            />
                          </span>
                          <span style={{ fontWeight: 'bold', color: '#b91c1c' }}>{ageingDays}</span>
                        </div>
                      );
                    })
                  )}
                </div>
              </div>
            </div>
          )}

          <button
            onClick={exportToExcel}
            style={{ background: '#107c41', color: '#fff', border: 'none', padding: '6px 16px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '13px' }}
          >
            📥 Download Master Excel
          </button>
        </div>
      </div>

      <div style={{ flex: 1, overflow: 'hidden' }} ref={containerRef}></div>

      {showFindModal && (
        <div style={{
          position: 'fixed',
          top: '60px',
          right: '20px',
          background: '#f3f4f6',
          border: '1px solid #94a3b8',
          boxShadow: '0 10px 25px rgba(0,0,0,0.2)',
          borderRadius: '6px',
          padding: '12px 16px',
          zIndex: 999999,
          width: '320px',
          fontFamily: 'Segoe UI, sans-serif'
        }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
            <span style={{ fontWeight: 'bold', fontSize: '13px', color: '#1e293b' }}>🔍 Find in Sheet</span>
            <button
              onClick={() => setShowFindModal(false)}
              style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontWeight: 'bold', fontSize: '14px', color: '#64748b' }}
            >
              ✕
            </button>
          </div>

          <div style={{ display: 'flex', gap: '8px', marginBottom: '8px' }}>
            <input
              ref={findInputRef}
              type="text"
              placeholder="Search branch, code, amount..."
              value={findText}
              onChange={(e) => {
                setFindText(e.target.value);
                setCurrentResultIdx(-1);
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') handleFindNext();
              }}
              style={{
                flex: 1,
                padding: '6px 8px',
                border: '1px solid #cbd5e1',
                borderRadius: '4px',
                fontSize: '13px',
                outline: 'none'
              }}
            />
            <button
              onClick={() => handleFindNext()}
              style={{
                background: '#107c41',
                color: '#fff',
                border: 'none',
                padding: '6px 12px',
                borderRadius: '4px',
                cursor: 'pointer',
                fontWeight: 'bold',
                fontSize: '13px'
              }}
            >
              Next
            </button>
          </div>

          <div style={{ fontSize: '11px', color: '#64748b', display: 'flex', justifyContent: 'space-between' }}>
            <span>
              {findResults.length > 0
                ? `Match ${currentResultIdx + 1} of ${findResults.length}`
                : 'Press Enter to find'}
            </span>
            <span>Press Esc to close</span>
          </div>
        </div>
      )}
    </div>
  );
}