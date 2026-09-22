import React, { useState, useMemo } from 'react';
import { 
  Users, 
  Calendar, 
  Download, 
  Printer, 
  RefreshCw, 
  Filter, 
  ChevronRight, 
  X, 
  Search, 
  Baby, 
  GraduationCap, 
  Heart, 
  UserCheck, 
  TrendingUp, 
  PieChart as PieChartIcon, 
  CheckCircle2, 
  FileSpreadsheet,
  MapPin,
  Sparkles
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { motion, AnimatePresence } from 'motion/react';
import { UserProfile, MosqueLocation, Jamaah } from '../types';
import { useDataQuery, where } from '../hooks/useDataQuery';
import { MOSQUE_LOCATIONS as ALL_KELOMPOK } from '../constants';
import { normalizeDateToInputFormat } from '../lib/utils';

interface CacahJiwaRowData {
  no: number;
  kelompok: MosqueLocation;
  // Belum Menikah
  balitaL: number;
  balitaP: number;
  acrL: number;
  acrP: number;
  aprL: number;
  aprP: number;
  arL: number;
  arP: number;
  gpnAL: number;
  gpnAP: number;
  gpnBL: number;
  gpnBP: number;
  // Menikah
  menikahL: number;
  menikahP: number;
  duda: number;
  janda: number;
  // Totals
  jumlahL: number;
  jumlahP: number;
  total: number;
  balig: number;
}

type CellCategoryKey = 
  | 'balitaL' | 'balitaP'
  | 'acrL' | 'acrP'
  | 'aprL' | 'aprP'
  | 'arL' | 'arP'
  | 'gpnAL' | 'gpnAP'
  | 'gpnBL' | 'gpnBP'
  | 'menikahL' | 'menikahP'
  | 'duda' | 'janda'
  | 'jumlahL' | 'jumlahP'
  | 'total' | 'balig';

interface SelectedCellDetail {
  kelompok: MosqueLocation | 'ALL';
  categoryKey: CellCategoryKey;
  label: string;
  count: number;
  members: Array<{ jamaah: Jamaah; age: number; categoryLabel: string }>;
}

export function CacahJiwaView({ profile }: { profile: UserProfile }) {
  // Always query all jamaah if admin, or query all so we can display exactly what is permitted
  const { data: allJamaah, loading } = useDataQuery<Jamaah>('jamaah', []);
  
  // Cutoff date state (default: today or 31 Agustus 2026 as shown in template)
  const [cutoffDateStr, setCutoffDateStr] = useState<string>(() => {
    // Current date in YYYY-MM-DD format
    const now = new Date();
    return now.toISOString().split('T')[0];
  });
  
  const [selectedKelompokFilter, setSelectedKelompokFilter] = useState<string>(() => {
    if (profile.role === 'pengurus' && profile.location && profile.location !== ('Seluruh Lokasi' as any)) {
      return profile.location;
    }
    return 'ALL';
  });

  const [selectedCellDetail, setSelectedCellDetail] = useState<SelectedCellDetail | null>(null);
  const [detailSearch, setDetailSearch] = useState('');
  const [isPrintMode, setIsPrintMode] = useState(false);

  // Determine which Kelompok rows should be shown
  const visibleKelompokList = useMemo<MosqueLocation[]>(() => {
    if (profile.role === 'pengurus' && profile.location && profile.location !== ('Seluruh Lokasi' as any)) {
      // If pengurus, only their own location is displayed
      return [profile.location as MosqueLocation];
    }
    
    // If admin and filtered to specific kelompok
    if (selectedKelompokFilter !== 'ALL') {
      return [selectedKelompokFilter as MosqueLocation];
    }

    // Default admin view: Kramat Batu, Cipete, Karya Utama, Radio Dalam, Antena
    return ALL_KELOMPOK;
  }, [profile.role, profile.location, selectedKelompokFilter]);

  // Parse cutoff date object
  const cutoffDate = useMemo(() => {
    const d = new Date(cutoffDateStr);
    return isNaN(d.getTime()) ? new Date() : d;
  }, [cutoffDateStr]);

  // Formatted cutoff date display, e.g. "31 AGUSTUS 2026"
  const formattedCutoffDisplay = useMemo(() => {
    const months = [
      'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI',
      'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'
    ];
    const day = cutoffDate.getDate();
    const month = months[cutoffDate.getMonth()];
    const year = cutoffDate.getFullYear();
    return `${day} ${month} ${year}`;
  }, [cutoffDate]);


   const calculateAgeAtCutoff = (dobStr?: string, defaultAge = 25): number => {
    if (!dobStr) return defaultAge;
    const stdDob = normalizeDateToInputFormat(dobStr);
    const dob = stdDob ? new Date(stdDob) : new Date(dobStr);
    if (isNaN(dob.getTime())) return defaultAge;

    let age = cutoffDate.getFullYear() - dob.getFullYear();
    const m = cutoffDate.getMonth() - dob.getMonth();
    if (m < 0 || (m === 0 && cutoffDate.getDate() < dob.getDate())) {
      age--;
    }
    return Math.max(0, age);
  };

  const normalizeCategoryValue = (value?: string) => {
    return String(value ?? '').trim().toLowerCase().replace(/\s+/g, '_');
  };

  const getCategoryGroup = (value?: string): 'ACR' | 'APR' | 'AR' | 'GPN' | 'UMUM' | 'OTHER' => {
    const normalized = normalizeCategoryValue(value);

    if (!normalized) return 'OTHER';
    if (normalized === 'acr' || normalized.startsWith('acr')) return 'ACR';
    if (normalized === 'apr' || normalized.startsWith('apr')) return 'APR';
    if (normalized === 'ar' || normalized.startsWith('ar')) return 'AR';
    if (normalized === 'gpn' || normalized.startsWith('gpn')) return 'GPN';
    if (normalized.includes('umum')) return 'UMUM';
    return 'OTHER';
  };

  // Helper to categorize individual Jamaah
  const categorizeJamaah = (j: Jamaah) => {
    // 1. Precise Gender Extraction
    const genderStr = (
      j.gender || 
      (j as any).jenisKelamin || 
      (j as any).jk || 
      (j as any).sex || 
      (j as any).Gender || 
      ''
    ).toString().trim().toLowerCase();

    const relStr = (
      (j as any).relationshipInFamily || 
      (j as any).hubunganKeluarga || 
      (j as any).statusKeluarga || 
      (j as any).posisiKeluarga || 
      ''
    ).toString().trim().toLowerCase();

    let isMale = true;
    if (
      genderStr.startsWith('p') || 
      genderStr.includes('perempuan') || 
      genderStr.includes('wanita') || 
      genderStr === 'w' || 
      genderStr === 'f' || 
      genderStr.includes('ibu') || 
      genderStr.includes('istri')
    ) {
      isMale = false;
    } else if (
      genderStr.startsWith('l') || 
      genderStr.includes('laki') || 
      genderStr.includes('pria') || 
      genderStr === 'm' || 
      genderStr.includes('bapak') || 
      genderStr.includes('suami')
    ) {
      isMale = true;
    } else if (
      relStr.includes('istri') || 
      relStr.includes('ibu') || 
      relStr.includes('anak perempuan') || 
      relStr.includes('putri')
    ) {
      isMale = false;
    } else if (
      relStr.includes('suami') || 
      relStr.includes('bapak') || 
      relStr.includes('kepala') || 
      relStr.includes('putra')
    ) {
      isMale = true;
    }

    // 2. Extract marital status, category, dapukan, and spouse fields
    const marital = (
      j.maritalStatus || 
      (j as any).statusPernikahan || 
      (j as any).status_nikah || 
      (j as any).pernikahan || 
      (j as any).statusKawin || 
      (j as any).status_kawin || 
      ''
    ).toString().trim().toLowerCase();

    const categoryStr = (
      j.category || 
      (j as any).kategori || 
      ''
    ).toString().trim().toLowerCase();

    const statusStr = (
      (j as any).status || 
      ''
    ).toString().trim().toLowerCase();

    const spouseNameStr = (
      j.spouseName || 
      (j as any).namaPasangan || 
      (j as any).namaSuami || 
      (j as any).namaIstri || 
      (j as any).pasangan || 
      ''
    ).toString().trim();

    const dapukanStr = (
      (Array.isArray(j.dapukan) ? j.dapukan.join(' ') : (j.dapukan || '')) + ' ' +
      (Array.isArray(j.positions) ? j.positions.join(' ') : (j.positions || ''))
    ).toLowerCase();

    // Default fallback age based on legacy category if dateOfBirth is missing
    let defaultAge = 25;
    const categoryGroup = getCategoryGroup(j.category);
    if (categoryGroup === 'ACR') defaultAge = 8;
    else if (categoryGroup === 'APR') defaultAge = 14;
    else if (categoryGroup === 'GPN') defaultAge = 20;

    const age = calculateAgeAtCutoff(j.dateOfBirth, defaultAge);

    // 3. Check Duda / Janda (Top Priority)
    const isDudaKeyword = 
      marital.includes('duda') || 
      categoryStr.includes('duda') || 
      statusStr.includes('duda') || 
      dapukanStr.includes('duda') || 
      relStr.includes('duda');

    const isJandaKeyword = 
      marital.includes('janda') || 
      categoryStr.includes('janda') || 
      statusStr.includes('janda') || 
      dapukanStr.includes('janda') || 
      relStr.includes('janda');

    const isCeraiKeyword = 
      marital.includes('cerai') || 
      categoryStr.includes('cerai') || 
      statusStr.includes('cerai');

    if (isDudaKeyword || (isCeraiKeyword && isMale)) {
      return { key: 'duda' as CellCategoryKey, isMale, age, label: 'Duda' };
    }
    if (isJandaKeyword || (isCeraiKeyword && !isMale)) {
      return { key: 'janda' as CellCategoryKey, isMale, age, label: 'Janda' };
    }

    // 4. Check Explicitly Belum Menikah (Youth / Single)
    const isYouthCategory = 
      getCategoryGroup(j.category) === 'ACR' ||
      getCategoryGroup(j.category) === 'APR' ||
      getCategoryGroup(j.category) === 'GPN' ||
      categoryStr.includes('caberawit') || categoryStr.includes('pra remaja') || categoryStr.includes('remaja') ||
      categoryStr === 'balita';

    const hasSpouse = Boolean(
      spouseNameStr && 
      spouseNameStr !== '-' && 
      spouseNameStr !== 'Tidak Ada' && 
      spouseNameStr !== 'Belum Ada' && 
      spouseNameStr !== 'null' && 
      spouseNameStr !== 'undefined'
    );

    const isExplicitlyMenikah = 
      marital.includes('menikah') || 
      marital.includes('kawin') || 
      marital.includes('nikah') || 
      marital.includes('berkeluarga') || 
      marital === 'k' || 
      marital === 'sudah' || 
      categoryStr.includes('menikah') || 
      categoryStr.includes('nikah') || 
      categoryStr.includes('kawin') || 
      statusStr.includes('menikah') ||
      relStr.includes('istri') || 
      relStr.includes('suami');

    const isExplicitlyBelumMenikah = 
      !isExplicitlyMenikah && !hasSpouse && (
        marital.includes('belum') || 
        marital.includes('bk') || 
        marital.includes('single') || 
        marital.includes('lajang') || 
        isYouthCategory
      );

    // In mosque records, UMUM is the category for adult married members (Bapak/Ibu)
    const isUmumAdult = (
      categoryStr.includes('umum') || 
      categoryStr.includes('dewasa') || 
      categoryStr.includes('orang tua') || 
      categoryStr.includes('bapak') || 
      categoryStr.includes('ibu') || 
      categoryStr.includes('keluarga') ||
      categoryStr === ''
    );

    const isKKHead = (j.isKK === true || (j.isKK as any) === 'true' || (j.isKK as any) === 'TRUE');

    const isMenikah = 
      isExplicitlyMenikah || 
      hasSpouse || 
      (isUmumAdult && !isExplicitlyBelumMenikah) || 
      (isKKHead && !isExplicitlyBelumMenikah);

    if (isMenikah) {
      if (isMale) {
        return { key: 'menikahL' as CellCategoryKey, isMale, age, label: 'Menikah (L)' };
      } else {
        return { key: 'menikahP' as CellCategoryKey, isMale, age, label: 'Menikah (P)' };
      }
    }

    // 5. Belum Menikah Age Brackets
    if (age <= 4) {
      return { 
        key: (isMale ? 'balitaL' : 'balitaP') as CellCategoryKey, 
        isMale, 
        age, 
        label: isMale ? 'Balita L (0-4 Th)' : 'Balita P (0-4 Th)' 
      };
    }
    if (age <= 12) {
      return { 
        key: (isMale ? 'acrL' : 'acrP') as CellCategoryKey, 
        isMale, 
        age, 
        label: isMale ? 'ACR L (5-12 Th)' : 'ACR P (5-12 Th)' 
      };
    }
    if (age <= 15) {
      return { 
        key: (isMale ? 'aprL' : 'aprP') as CellCategoryKey, 
        isMale, 
        age, 
        label: isMale ? 'APR L (13-15 Th)' : 'APR P (13-15 Th)' 
      };
    }
    if (age <= 18) {
      return { 
        key: (isMale ? 'arL' : 'arP') as CellCategoryKey, 
        isMale, 
        age, 
        label: isMale ? 'AR L (16-18 Th)' : 'AR P (16-18 Th)' 
      };
    }
    if (age <= 22) {
      return { 
        key: (isMale ? 'gpnAL' : 'gpnAP') as CellCategoryKey, 
        isMale, 
        age, 
        label: isMale ? 'GPN A L (19-22 Th)' : 'GPN A P (19-22 Th)' 
      };
    }
    return { 
      key: (isMale ? 'gpnBL' : 'gpnBP') as CellCategoryKey, 
      isMale, 
      age, 
      label: isMale ? 'GPN B L (23+ Th)' : 'GPN B P (23+ Th)' 
    };
  };

  // Grouped members per kelompok and category for detail drill-down
  const categorizedDataByKelompok = useMemo(() => {
    const result: Record<string, Record<CellCategoryKey, Array<{ jamaah: Jamaah; age: number; categoryLabel: string }>>> = {};

    ALL_KELOMPOK.forEach(loc => {
      result[loc] = {
        balitaL: [], balitaP: [],
        acrL: [], acrP: [],
        aprL: [], aprP: [],
        arL: [], arP: [],
        gpnAL: [], gpnAP: [],
        gpnBL: [], gpnBP: [],
        menikahL: [], menikahP: [],
        duda: [], janda: [],
        jumlahL: [], jumlahP: [],
        total: [], balig: []
      };
    });

    allJamaah.forEach(j => {
      const loc = (j.location || 'Kramat Batu') as MosqueLocation;
      if (!result[loc]) return;

      const cat = categorizeJamaah(j);
      const item = { jamaah: j, age: cat.age, categoryLabel: cat.label };

      // Add to specific sub-bracket
      result[loc][cat.key].push(item);

      // Add to aggregate collections
      if (cat.isMale) {
        result[loc].jumlahL.push(item);
      } else {
        result[loc].jumlahP.push(item);
      }
      result[loc].total.push(item);

      // Balig definition: APR (13+) + AR + GPN A + GPN B + Menikah L/P + Duda/Janda
      // (or Total minus Balita minus ACR)
      const isBalig = cat.key !== 'balitaL' && cat.key !== 'balitaP' && cat.key !== 'acrL' && cat.key !== 'acrP';
      if (isBalig) {
        result[loc].balig.push(item);
      }
    });

    return result;
  }, [allJamaah, cutoffDate]);

  // Compute table rows based on visible kelompok list
  const tableRows = useMemo<CacahJiwaRowData[]>(() => {
    return visibleKelompokList.map((loc, idx) => {
      const d = categorizedDataByKelompok[loc] || {
        balitaL: [], balitaP: [], acrL: [], acrP: [], aprL: [], aprP: [], arL: [], arP: [],
        gpnAL: [], gpnAP: [], gpnBL: [], gpnBP: [], menikahL: [], menikahP: [], duda: [], janda: [],
        jumlahL: [], jumlahP: [], total: [], balig: []
      };

      const balitaL = d.balitaL.length;
      const balitaP = d.balitaP.length;
      const acrL = d.acrL.length;
      const acrP = d.acrP.length;
      const aprL = d.aprL.length;
      const aprP = d.aprP.length;
      const arL = d.arL.length;
      const arP = d.arP.length;
      const gpnAL = d.gpnAL.length;
      const gpnAP = d.gpnAP.length;
      const gpnBL = d.gpnBL.length;
      const gpnBP = d.gpnBP.length;
      const menikahL = d.menikahL.length;
      const menikahP = d.menikahP.length;
      const duda = d.duda.length;
      const janda = d.janda.length;

      const jumlahL = balitaL + acrL + aprL + arL + gpnAL + gpnBL + menikahL + duda;
      const jumlahP = balitaP + acrP + aprP + arP + gpnAP + gpnBP + menikahP + janda;
      const total = jumlahL + jumlahP;
      const balig = total - (balitaL + balitaP) - (acrL + acrP);

      return {
        no: idx + 1,
        kelompok: loc,
        balitaL, balitaP,
        acrL, acrP,
        aprL, aprP,
        arL, arP,
        gpnAL, gpnAP,
        gpnBL, gpnBP,
        menikahL, menikahP,
        duda, janda,
        jumlahL, jumlahP,
        total,
        balig
      };
    });
  }, [visibleKelompokList, categorizedDataByKelompok]);

  // Compute grand summary row (JUMLAH / TOTAL)
  const summaryRow = useMemo(() => {
    return tableRows.reduce(
      (acc, r) => ({
        balitaL: acc.balitaL + r.balitaL,
        balitaP: acc.balitaP + r.balitaP,
        acrL: acc.acrL + r.acrL,
        acrP: acc.acrP + r.acrP,
        aprL: acc.aprL + r.aprL,
        aprP: acc.aprP + r.aprP,
        arL: acc.arL + r.arL,
        arP: acc.arP + r.arP,
        gpnAL: acc.gpnAL + r.gpnAL,
        gpnAP: acc.gpnAP + r.gpnAP,
        gpnBL: acc.gpnBL + r.gpnBL,
        gpnBP: acc.gpnBP + r.gpnBP,
        menikahL: acc.menikahL + r.menikahL,
        menikahP: acc.menikahP + r.menikahP,
        duda: acc.duda + r.duda,
        janda: acc.janda + r.janda,
        jumlahL: acc.jumlahL + r.jumlahL,
        jumlahP: acc.jumlahP + r.jumlahP,
        total: acc.total + r.total,
        balig: acc.balig + r.balig,
      }),
      {
        balitaL: 0, balitaP: 0,
        acrL: 0, acrP: 0,
        aprL: 0, aprP: 0,
        arL: 0, arP: 0,
        gpnAL: 0, gpnAP: 0,
        gpnBL: 0, gpnBP: 0,
        menikahL: 0, menikahP: 0,
        duda: 0, janda: 0,
        jumlahL: 0, jumlahP: 0,
        total: 0,
        balig: 0
      }
    );
  }, [tableRows]);

  // Cell click handler for drill-down modal
  const handleCellClick = (kelompok: MosqueLocation | 'ALL', key: CellCategoryKey, label: string) => {
    let members: Array<{ jamaah: Jamaah; age: number; categoryLabel: string }> = [];

    if (kelompok === 'ALL') {
      visibleKelompokList.forEach(loc => {
        const cat = categorizedDataByKelompok[loc];
        if (cat && cat[key]) {
          members = members.concat(cat[key]);
        }
      });
    } else {
      const cat = categorizedDataByKelompok[kelompok];
      if (cat && cat[key]) {
        members = cat[key];
      }
    }

    setSelectedCellDetail({
      kelompok,
      categoryKey: key,
      label,
      count: members.length,
      members
    });
    setDetailSearch('');
  };

  // Export to Excel handler
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();

    // Prepare multi-header sheet matching official LDII / Desa Gandaria format
    const title = visibleKelompokList.length === 1 
      ? `CACAH JIWA KELOMPOK ${visibleKelompokList[0].toUpperCase()}`
      : 'CACAH JIWA DESA GANDARIA';

    const wsData: any[][] = [
      [title],
      [`s/d ${formattedCutoffDisplay}`],
      [],
      [
        'NO',
        'KELOMPOK',
        'BELUM MENIKAH', '', '', '', '', '', '', '', '', '', '', '',
        'MENIKAH', '', '', '',
        'JUMLAH', '',
        'TOTAL',
        'JML KIRA2 BALIG'
      ],
      [
        '',
        '',
        'BALITA (0 - 4 TH)', '',
        'ACR (5 - 12 TH)', '',
        'APR (13 - 15 TH)', '',
        'AR (16 - 18 TH)', '',
        'GPN A (19 - 22 TH)', '',
        'GPN B (23 TH KE ATAS)', '',
        'L', 'P', 'DUDA', 'JANDA',
        'L', 'P',
        '',
        '(KEPERLUAN DESA)'
      ],
      [
        '',
        '',
        'L', 'P',
        'L', 'P',
        'L', 'P',
        'L', 'P',
        'L', 'P',
        'L', 'P',
        '', '', '', '',
        '', '',
        '',
        ''
      ]
    ];

    // Data rows
    tableRows.forEach(r => {
      wsData.push([
        r.no,
        r.kelompok.toUpperCase(),
        r.balitaL, r.balitaP,
        r.acrL, r.acrP,
        r.aprL, r.aprP,
        r.arL, r.arP,
        r.gpnAL, r.gpnAP,
        r.gpnBL, r.gpnBP,
        r.menikahL, r.menikahP,
        r.duda, r.janda,
        r.jumlahL, r.jumlahP,
        r.total,
        r.balig
      ]);
    });

    // Summary bottom row
    wsData.push([
      '',
      'JUMLAH',
      summaryRow.balitaL, summaryRow.balitaP,
      summaryRow.acrL, summaryRow.acrP,
      summaryRow.aprL, summaryRow.aprP,
      summaryRow.arL, summaryRow.arP,
      summaryRow.gpnAL, summaryRow.gpnAP,
      summaryRow.gpnBL, summaryRow.gpnBP,
      summaryRow.menikahL, summaryRow.menikahP,
      summaryRow.duda, summaryRow.janda,
      summaryRow.jumlahL, summaryRow.jumlahP,
      summaryRow.total,
      summaryRow.balig
    ]);

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Cacah Jiwa');
    XLSX.writeFile(wb, `Cacah_Jiwa_${visibleKelompokList.length === 1 ? visibleKelompokList[0].replace(/\s+/g, '_') : 'Desa_Gandaria'}_${cutoffDateStr}.xlsx`);
  };

  const handlePrint = () => {
    window.print();
  };

  const filteredDetailMembers = useMemo(() => {
    if (!selectedCellDetail) return [];
    if (!detailSearch.trim()) return selectedCellDetail.members;
    const q = detailSearch.toLowerCase();
    return selectedCellDetail.members.filter(m => 
      String(m.jamaah.name || '').toLowerCase().includes(q) ||
      String(m.jamaah.memberId || '').toLowerCase().includes(q) ||
      String(m.jamaah.phone || '').includes(q) ||
      String(m.jamaah.nickname || '').toLowerCase().includes(q)
    );
  }, [selectedCellDetail, detailSearch]);

  const pageTitle = visibleKelompokList.length === 1
    ? `CACAH JIWA KELOMPOK ${visibleKelompokList[0].toUpperCase()}`
    : 'CACAH JIWA DESA GANDARIA';

  return (
    <div className="space-y-8 animate-fadeIn">
      {/* Header & Controls */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-2 border-b border-slate-100 print:hidden">
        <div>
          <div className="flex items-center gap-2.5">
            <div className="p-2.5 bg-emerald-100 text-emerald-700 rounded-2xl shadow-sm">
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-2xl lg:text-3xl font-black text-slate-900 tracking-tight">Cacah Jiwa</h1>
              <p className="text-xs sm:text-sm text-slate-500 font-medium">
                Rekapitulasi sensus penduduk & status demografi jamaah berdasarkan kelompok dan kelompok umur
              </p>
            </div>
          </div>
        </div>

        {/* Action Controls */}
        <div className="flex flex-wrap items-center gap-2.5">
          {/* Admin Kelompok Filter Dropdown */}
          {profile.role === 'admin' && (
            <div className="flex items-center gap-2 bg-white px-3 py-2 rounded-xl border border-slate-200 shadow-sm text-xs font-bold">
              <Filter className="w-3.5 h-3.5 text-emerald-600" />
              <select
                value={selectedKelompokFilter}
                onChange={(e) => setSelectedKelompokFilter(e.target.value)}
                className="bg-transparent outline-none cursor-pointer text-slate-700 font-bold"
              >
                <option value="ALL">Semua Kelompok (Desa Gandaria)</option>
                {ALL_KELOMPOK.map(loc => (
                  <option key={loc} value={loc}>Kelompok {loc}</option>
                ))}
              </select>
            </div>
          )}

          {/* Cutoff Date Picker */}
          <div className="flex items-center gap-2 bg-white px-3 py-2 rounded-xl border border-slate-200 shadow-sm text-xs font-bold">
            <Calendar className="w-3.5 h-3.5 text-emerald-600" />
            <span className="text-slate-400 font-medium">s/d:</span>
            <input
              type="date"
              value={cutoffDateStr}
              onChange={(e) => setCutoffDateStr(e.target.value)}
              className="bg-transparent outline-none cursor-pointer text-slate-800 font-bold"
            />
          </div>

          <button
            type="button"
            onClick={handleExportExcel}
            className="flex items-center gap-2 px-4 py-2 rounded-xl bg-emerald-50 text-emerald-700 border border-emerald-200 hover:bg-emerald-100 font-bold text-xs transition-all shadow-sm active:scale-95"
            title="Download Excel (.xlsx)"
          >
            <Download className="w-3.5 h-3.5 text-emerald-600" />
            <span>Ekspor Excel</span>
          </button>

          <button
            type="button"
            onClick={handlePrint}
            className="flex items-center gap-2 px-4 py-2 rounded-xl bg-slate-900 text-white hover:bg-slate-800 font-bold text-xs transition-all shadow-sm active:scale-95"
            title="Cetak Laporan / Print Landscape"
          >
            <Printer className="w-3.5 h-3.5 text-emerald-400" />
            <span>Cetak Laporan</span>
          </button>
        </div>
      </div>

      {/* Quick Demographic Metrics */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 lg:gap-4 print:hidden">
        <div className="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-[11px] font-black uppercase tracking-wider text-slate-400">Total Jiwa</p>
            <p className="text-2xl lg:text-3xl font-black text-slate-900 mt-1">{summaryRow.total}</p>
            <p className="text-[10px] text-emerald-600 font-bold mt-0.5">
              {summaryRow.jumlahL} L / {summaryRow.jumlahP} P
            </p>
          </div>
          <div className="w-10 h-10 rounded-2xl bg-emerald-50 text-emerald-600 flex items-center justify-center">
            <Users className="w-5 h-5" />
          </div>
        </div>

        <div className="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-[11px] font-black uppercase tracking-wider text-slate-400">Kira-kira Balig</p>
            <p className="text-2xl lg:text-3xl font-black text-emerald-700 mt-1">{summaryRow.balig}</p>
            <p className="text-[10px] text-slate-400 font-medium mt-0.5">
              {summaryRow.total > 0 ? Math.round((summaryRow.balig / summaryRow.total) * 100) : 0}% dari Total Jiwa
            </p>
          </div>
          <div className="w-10 h-10 rounded-2xl bg-teal-50 text-teal-600 flex items-center justify-center">
            <UserCheck className="w-5 h-5" />
          </div>
        </div>

        <div className="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-[11px] font-black uppercase tracking-wider text-slate-400">Generasi Penerus (GPN)</p>
            <p className="text-2xl lg:text-3xl font-black text-slate-900 mt-1">
              {summaryRow.gpnAL + summaryRow.gpnAP + summaryRow.gpnBL + summaryRow.gpnBP}
            </p>
            <p className="text-[10px] text-slate-400 font-medium mt-0.5">
              GPN A: {summaryRow.gpnAL + summaryRow.gpnAP} | GPN B: {summaryRow.gpnBL + summaryRow.gpnBP}
            </p>
          </div>
          <div className="w-10 h-10 rounded-2xl bg-amber-50 text-amber-600 flex items-center justify-center">
            <GraduationCap className="w-5 h-5" />
          </div>
        </div>

        <div className="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm flex items-center justify-between">
          <div>
            <p className="text-[11px] font-black uppercase tracking-wider text-slate-400">Status Menikah / Keluarga</p>
            <p className="text-2xl lg:text-3xl font-black text-slate-900 mt-1">
              {summaryRow.menikahL + summaryRow.menikahP}
            </p>
            <p className="text-[10px] text-slate-400 font-medium mt-0.5">
              Duda: {summaryRow.duda} | Janda: {summaryRow.janda}
            </p>
          </div>
          <div className="w-10 h-10 rounded-2xl bg-rose-50 text-rose-600 flex items-center justify-center">
            <Heart className="w-5 h-5" />
          </div>
        </div>
      </div>

      {/* Official Sensus Table Container (Pixel-perfect replication of uploaded design) */}
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden print:border-none print:shadow-none">
        {/* Official Report Title on Table */}
        <div className="py-6 px-4 text-center border-b border-slate-200 bg-slate-50/50 print:bg-white">
          <h2 className="text-lg sm:text-xl md:text-2xl font-black text-slate-900 uppercase tracking-wide">
            {pageTitle}
          </h2>
          <p className="text-xs sm:text-sm font-bold text-slate-700 uppercase tracking-widest mt-1">
            s/d {formattedCutoffDisplay}
          </p>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-center border-collapse text-xs sm:text-sm border border-slate-900 print:text-[10px]">
            {/* Multi-tier Header */}
            <thead>
              {/* Row 1: Major Groupings */}
              <tr className="bg-white border-b border-slate-900">
                <th 
                  rowSpan={3} 
                  className="border border-slate-900 px-2 py-2 font-black text-slate-900 w-10 uppercase align-middle"
                >
                  NO
                </th>
                <th 
                  rowSpan={3} 
                  className="border border-slate-900 px-3 py-2 font-black text-slate-900 min-w-[140px] uppercase align-middle text-left"
                >
                  KELOMPOK
                </th>
                <th 
                  colSpan={12} 
                  className="border border-slate-900 py-1.5 font-black text-slate-900 uppercase tracking-wider bg-slate-50"
                >
                  BELUM MENIKAH
                </th>
                <th 
                  colSpan={4} 
                  className="border border-slate-900 py-1.5 font-black text-slate-900 uppercase tracking-wider bg-slate-50"
                >
                  MENIKAH
                </th>
                <th 
                  colSpan={2} 
                  className="border border-slate-900 py-1.5 font-black text-slate-900 uppercase tracking-wider bg-[#8cd34a] text-slate-950"
                >
                  JUMLAH
                </th>
                <th 
                  rowSpan={3} 
                  className="border border-slate-900 px-2 py-2 font-black text-slate-900 uppercase align-middle min-w-[55px]"
                >
                  TOTAL
                </th>
                <th 
                  rowSpan={3} 
                  className="border border-slate-900 px-2 py-2 font-black text-slate-900 uppercase align-middle min-w-[65px] bg-slate-50/70"
                >
                  <div className="flex flex-col items-center justify-center leading-tight">
                    <span className="text-[9px] font-black text-slate-500 uppercase tracking-tighter mb-0.5">KEPERLUAN DESA</span>
                    <span className="text-[11px] font-black text-slate-900">JML KIRA2 BALIG</span>
                  </div>
                </th>
              </tr>

              {/* Row 2: Sub-groups (Age Brackets & Categories) */}
              <tr className="bg-white border-b border-slate-900 text-[11px] font-bold">
                <th colSpan={2} className="border border-slate-900 px-1 py-1">
                  <div className="font-black">BALITA</div>
                  <div className="text-[10px] font-semibold text-slate-600">0 - 4 TH</div>
                </th>
                <th colSpan={2} className="border border-slate-900 px-1 py-1">
                  <div className="font-black">ACR</div>
                  <div className="text-[10px] font-semibold text-slate-600">5 - 12 TH</div>
                </th>
                <th colSpan={2} className="border border-slate-900 px-1 py-1">
                  <div className="font-black">APR</div>
                  <div className="text-[10px] font-semibold text-slate-600">13 - 15 TH</div>
                </th>
                <th colSpan={2} className="border border-slate-900 px-1 py-1">
                  <div className="font-black">AR</div>
                  <div className="text-[10px] font-semibold text-slate-600">16 - 18 TH</div>
                </th>
                <th colSpan={2} className="border border-slate-900 px-1 py-1">
                  <div className="font-black">GPN A</div>
                  <div className="text-[10px] font-semibold text-slate-600">19 - 22 TH</div>
                </th>
                <th colSpan={2} className="border border-slate-900 px-1 py-1">
                  <div className="font-black">GPN B</div>
                  <div className="text-[10px] font-semibold text-slate-600">23 TH KE ATAS</div>
                </th>
                {/* Under MENIKAH */}
                <th rowSpan={2} className="border border-slate-900 px-1.5 py-1 font-black align-middle">L</th>
                <th rowSpan={2} className="border border-slate-900 px-1.5 py-1 font-black align-middle">P</th>
                <th rowSpan={2} className="border border-slate-900 px-1.5 py-1 font-black align-middle">DUDA</th>
                <th rowSpan={2} className="border border-slate-900 px-1.5 py-1 font-black align-middle">JANDA</th>
                {/* Under JUMLAH */}
                <th rowSpan={2} className="border border-slate-900 px-2 py-1 font-black bg-[#8cd34a] text-slate-950 align-middle">L</th>
                <th rowSpan={2} className="border border-slate-900 px-2 py-1 font-black bg-[#8cd34a] text-slate-950 align-middle">P</th>
              </tr>

              {/* Row 3: L & P for Belum Menikah sub-columns */}
              <tr className="bg-white border-b border-slate-900 text-[11px] font-black">
                <th className="border border-slate-900 px-1.5 py-1">L</th>
                <th className="border border-slate-900 px-1.5 py-1">P</th>
                <th className="border border-slate-900 px-1.5 py-1">L</th>
                <th className="border border-slate-900 px-1.5 py-1">P</th>
                <th className="border border-slate-900 px-1.5 py-1">L</th>
                <th className="border border-slate-900 px-1.5 py-1">P</th>
                <th className="border border-slate-900 px-1.5 py-1">L</th>
                <th className="border border-slate-900 px-1.5 py-1">P</th>
                <th className="border border-slate-900 px-1.5 py-1">L</th>
                <th className="border border-slate-900 px-1.5 py-1">P</th>
                <th className="border border-slate-900 px-1.5 py-1">L</th>
                <th className="border border-slate-900 px-1.5 py-1">P</th>
              </tr>
            </thead>

            {/* Table Body */}
            <tbody>
              {tableRows.map((row) => (
                <tr 
                  key={row.kelompok} 
                  className="hover:bg-emerald-50/40 transition-colors border-b border-slate-900 font-semibold text-slate-900"
                >
                  <td className="border border-slate-900 px-2 py-2.5 font-bold">{row.no}</td>
                  <td className="border border-slate-900 px-3 py-2.5 text-left font-black uppercase text-slate-900">
                    {row.kelompok}
                  </td>
                  
                  {/* Belum Menikah: Balita (0-4) */}
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'balitaL', `${row.kelompok} - Balita Laki-laki (0-4 Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.balitaL}
                  </td>
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'balitaP', `${row.kelompok} - Balita Perempuan (0-4 Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.balitaP}
                  </td>

                  {/* ACR (5-12) */}
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'acrL', `${row.kelompok} - ACR Laki-laki (5-12 Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.acrL}
                  </td>
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'acrP', `${row.kelompok} - ACR Perempuan (5-12 Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.acrP}
                  </td>

                  {/* APR (13-15) */}
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'aprL', `${row.kelompok} - APR Laki-laki (13-15 Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.aprL}
                  </td>
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'aprP', `${row.kelompok} - APR Perempuan (13-15 Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.aprP}
                  </td>

                  {/* AR (16-18) */}
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'arL', `${row.kelompok} - AR Laki-laki (16-18 Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.arL}
                  </td>
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'arP', `${row.kelompok} - AR Perempuan (16-18 Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.arP}
                  </td>

                  {/* GPN A (19-22) */}
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'gpnAL', `${row.kelompok} - GPN A Laki-laki (19-22 Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.gpnAL}
                  </td>
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'gpnAP', `${row.kelompok} - GPN A Perempuan (19-22 Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.gpnAP}
                  </td>

                  {/* GPN B (23+) */}
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'gpnBL', `${row.kelompok} - GPN B Laki-laki (23+ Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.gpnBL}
                  </td>
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'gpnBP', `${row.kelompok} - GPN B Perempuan (23+ Th)`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.gpnBP}
                  </td>

                  {/* Menikah */}
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'menikahL', `${row.kelompok} - Menikah Laki-laki`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.menikahL}
                  </td>
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'menikahP', `${row.kelompok} - Menikah Perempuan`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.menikahP}
                  </td>
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'duda', `${row.kelompok} - Duda`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.duda}
                  </td>
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'janda', `${row.kelompok} - Janda`)}
                    className="border border-slate-900 px-1.5 py-2.5 cursor-pointer hover:bg-emerald-100 hover:text-emerald-900 font-bold"
                  >
                    {row.janda}
                  </td>

                  {/* Jumlah Total L & P (Green Column) */}
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'jumlahL', `${row.kelompok} - Total Laki-laki`)}
                    className="border border-slate-900 px-2 py-2.5 bg-[#8cd34a]/30 font-black cursor-pointer hover:bg-[#8cd34a]/50 text-slate-950"
                  >
                    {row.jumlahL}
                  </td>
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'jumlahP', `${row.kelompok} - Total Perempuan`)}
                    className="border border-slate-900 px-2 py-2.5 bg-[#8cd34a]/30 font-black cursor-pointer hover:bg-[#8cd34a]/50 text-slate-950"
                  >
                    {row.jumlahP}
                  </td>

                  {/* Row Total */}
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'total', `${row.kelompok} - Total Seluruh Jiwa`)}
                    className="border border-slate-900 px-2 py-2.5 font-black bg-slate-50 cursor-pointer hover:bg-slate-200 text-slate-900"
                  >
                    {row.total}
                  </td>

                  {/* Keperluan Desa: JML KIRA2 BALIG */}
                  <td 
                    onClick={() => handleCellClick(row.kelompok, 'balig', `${row.kelompok} - Kira-kira Balig`)}
                    className="border border-slate-900 px-2 py-2.5 font-black text-sm bg-slate-50/70 cursor-pointer hover:bg-emerald-100 text-slate-900"
                  >
                    {row.balig}
                  </td>
                </tr>
              ))}

              {/* Bottom Summary Row (JUMLAH / TOTAL DESA) */}
              <tr className="bg-white border-t-2 border-b border-slate-900 font-black text-slate-950">
                <td colSpan={2} className="border border-slate-900 px-3 py-3 text-center uppercase tracking-wider text-xs sm:text-sm">
                  JUMLAH
                </td>

                {/* Belum Menikah sub-totals */}
                <td colSpan={2} className="border border-slate-900 px-1 py-3 text-center bg-slate-50">
                  {summaryRow.balitaL + summaryRow.balitaP}
                </td>
                <td colSpan={2} className="border border-slate-900 px-1 py-3 text-center bg-slate-50">
                  {summaryRow.acrL + summaryRow.acrP}
                </td>
                <td colSpan={2} className="border border-slate-900 px-1 py-3 text-center bg-slate-50">
                  {summaryRow.aprL + summaryRow.aprP}
                </td>
                <td colSpan={2} className="border border-slate-900 px-1 py-3 text-center bg-slate-50">
                  {summaryRow.arL + summaryRow.arP}
                </td>
                <td colSpan={2} className="border border-slate-900 px-1 py-3 text-center bg-slate-50">
                  {summaryRow.gpnAL + summaryRow.gpnAP}
                </td>
                <td colSpan={2} className="border border-slate-900 px-1 py-3 text-center bg-slate-50">
                  {summaryRow.gpnBL + summaryRow.gpnBP}
                </td>

                {/* Menikah Sub-total */}
                <td colSpan={2} className="border border-slate-900 px-1 py-3 text-center bg-slate-50">
                  {summaryRow.menikahL + summaryRow.menikahP}
                </td>
                {/* Duda & Janda Sub-total */}
                <td colSpan={2} className="border border-slate-900 px-1 py-3 text-center bg-slate-50">
                  {summaryRow.duda + summaryRow.janda}
                </td>

                {/* Highlighted Total Cells */}
                <td 
                  colSpan={2} 
                  className="border border-slate-900 px-2 py-3 bg-[#ffff00] text-black font-black text-sm sm:text-base text-center"
                >
                  {summaryRow.total}
                </td>

                <td className="border border-slate-900 px-2 py-3 bg-slate-900 text-white font-black">
                  -
                </td>

                <td className="border border-slate-900 px-2 py-3 font-black text-sm sm:text-base text-center bg-emerald-100 text-emerald-950">
                  {summaryRow.balig}
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        {/* Footer Notes & Signatures for Print */}
        <div className="p-4 sm:p-6 bg-slate-50/50 border-t border-slate-100 text-xs text-slate-500 flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-2">
            <span className="w-2.5 h-2.5 rounded-full bg-[#8cd34a]"></span>
            <span className="font-semibold text-slate-700">Tips:</span> Klik angka pada tabel untuk melihat daftar nama jamaah secara detail.
          </div>
          <div className="text-[11px] text-slate-400 font-medium">
            Sistem Informasi Manajemen Jamaah Desa Gandaria
          </div>
        </div>

        {/* Printable Signatures Block (Visible in Print Mode) */}
        <div className="hidden print:grid grid-cols-2 gap-12 mt-12 p-6 text-center text-xs font-bold text-slate-900">
          <div>
            <p>Mengetahui,</p>
            <p className="mt-1 font-black">PENGURUS DESA GANDARIA</p>
            <div className="h-16"></div>
            <p className="border-t border-slate-900 pt-1 inline-block min-w-[180px]">( ........................................ )</p>
          </div>
          <div>
            <p>Jakarta, {formattedCutoffDisplay}</p>
            <p className="mt-1 font-black">
              {visibleKelompokList.length === 1 ? `KETUA KELOMPOK ${visibleKelompokList[0].toUpperCase()}` : 'KETUA DESA GANDARIA'}
            </p>
            <div className="h-16"></div>
            <p className="border-t border-slate-900 pt-1 inline-block min-w-[180px]">( ........................................ )</p>
          </div>
        </div>
      </div>

      {/* Detail Jamaah Slide-over / Modal */}
      <AnimatePresence>
        {selectedCellDetail && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4"
            onClick={() => setSelectedCellDetail(null)}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              onClick={(e) => e.stopPropagation()}
              className="bg-white w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[85vh]"
            >
              {/* Modal Header */}
              <div className="p-6 bg-gradient-to-r from-emerald-800 to-teal-800 text-white flex items-center justify-between">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="px-2.5 py-0.5 rounded-full bg-white/20 text-[11px] font-black uppercase tracking-wider">
                      Detail Cacah Jiwa
                    </span>
                    <span className="text-xs text-emerald-200 font-bold">
                      {selectedCellDetail.count} Jiwa
                    </span>
                  </div>
                  <h3 className="text-lg sm:text-xl font-black mt-1">
                    {selectedCellDetail.label}
                  </h3>
                </div>
                <button
                  type="button"
                  onClick={() => setSelectedCellDetail(null)}
                  className="p-2 rounded-2xl bg-white/10 hover:bg-white/20 text-white transition-all"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Search within detail */}
              <div className="p-4 border-b border-slate-100 bg-slate-50 flex items-center gap-2">
                <Search className="w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  placeholder="Cari nama, ID Jamaah, nomor HP..."
                  value={detailSearch}
                  onChange={(e) => setDetailSearch(e.target.value)}
                  className="w-full bg-transparent outline-none text-xs sm:text-sm font-medium text-slate-800 placeholder:text-slate-400"
                />
                {detailSearch && (
                  <button type="button" onClick={() => setDetailSearch('')} className="text-slate-400 hover:text-slate-600">
                    <X className="w-4 h-4" />
                  </button>
                )}
              </div>

              {/* List of Members */}
              <div className="p-4 sm:p-6 overflow-y-auto space-y-2.5 flex-grow">
                {filteredDetailMembers.length === 0 ? (
                  <div className="text-center py-12 text-slate-400 font-medium">
                    Tidak ada data jamaah pada kategori ini.
                  </div>
                ) : (
                  filteredDetailMembers.map(({ jamaah, age, categoryLabel }, idx) => (
                    <div
                      key={jamaah.id || jamaah.memberId || idx}
                      className="p-3.5 rounded-2xl border border-slate-100 hover:border-emerald-200 hover:bg-emerald-50/30 transition-all flex items-center justify-between gap-3"
                    >
                      <div className="flex items-center gap-3 min-w-0">
                        <div className="w-10 h-10 rounded-2xl bg-emerald-100 text-emerald-800 font-black flex items-center justify-center flex-shrink-0 text-sm overflow-hidden">
                          {jamaah.photoUrl ? (
                            <img src={jamaah.photoUrl} alt={jamaah.name} className="w-full h-full object-cover" />
                          ) : (
                            jamaah.name.charAt(0).toUpperCase()
                          )}
                        </div>
                        <div className="min-w-0">
                          <div className="flex items-center gap-2">
                            <span className="font-bold text-slate-900 text-sm truncate">
                              {jamaah.name}
                            </span>
                            {jamaah.nickname && (
                              <span className="text-xs text-slate-400 font-medium">
                                ({jamaah.nickname})
                              </span>
                            )}
                          </div>
                          <div className="flex flex-wrap items-center gap-x-3 gap-y-0.5 text-xs text-slate-500 mt-0.5">
                            <span className="font-mono font-bold text-emerald-700 bg-emerald-50 px-1.5 py-0.5 rounded text-[10px]">
                              {jamaah.memberId}
                            </span>
                            <span>{jamaah.location}</span>
                            <span>•</span>
                            <span>{jamaah.gender || 'Laki-laki'}</span>
                            <span>•</span>
                            <span className="font-semibold text-slate-700">{age} Tahun</span>
                          </div>
                        </div>
                      </div>

                      <div className="text-right flex-shrink-0">
                        <span className="px-2.5 py-1 rounded-xl bg-emerald-50 text-emerald-800 text-[11px] font-bold border border-emerald-200">
                          {categoryLabel || jamaah.maritalStatus || jamaah.category || 'Belum Menikah'}
                        </span>
                        {jamaah.phone && (
                          <p className="text-[11px] text-slate-400 font-mono mt-1">
                            {jamaah.phone}
                          </p>
                        )}
                      </div>
                    </div>
                  ))
                )}
              </div>

              {/* Modal Footer */}
              <div className="p-4 border-t border-slate-100 bg-slate-50 flex items-center justify-between">
                <span className="text-xs text-slate-500 font-medium">
                  Menampilkan {filteredDetailMembers.length} dari {selectedCellDetail.count} jamaah
                </span>
                <button
                  type="button"
                  onClick={() => setSelectedCellDetail(null)}
                  className="px-5 py-2 rounded-xl bg-slate-900 text-white font-bold text-xs hover:bg-slate-800 transition-all"
                >
                  Tutup
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}