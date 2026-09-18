import * as XLSX from 'xlsx';
import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { 
  LayoutDashboard, 
  Users, 
  Package, 
  Calendar, 
  LogOut, 
  Plus, 
  MapPin, 
  Search,
  CheckCircle2,
  AlertCircle,
  Clock,
  Menu,
  X,
  Map as MapIcon,
  Landmark,
  Image as ImageIcon,
  Home,
  ClipboardList,
  ChevronLeft,
  ChevronRight,
  ShoppingBag,
  Database,
  CloudCog,
  Loader2,
  Eye,
  FileText,
  HeartPulse,
  GraduationCap,
  Briefcase,
  UserCheck,
  PhoneCall,
  RefreshCw,
  Copy,
  Check,
  ExternalLink,
  ShieldCheck,
  QrCode,
  Download,
  Printer,
  ChevronDown,
  Share2,
  Link as LinkIcon,
  UserPlus,
  Send,
  CalendarDays,
  Timer,
  Sparkles,
  CheckCheck,
  FileSpreadsheet
} from 'lucide-react';
import { QRCodeCanvas } from 'qrcode.react';
import { ToastProvider, useToast } from './components/ToastContext';
import { CacahJiwaView } from './components/CacahJiwaView';
import { 
  LineChart, 
  Line, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer,
  BarChart,
  Bar,
  Cell,
  PieChart,
  Pie,
  Legend
} from 'recharts';
import { motion, AnimatePresence } from 'motion/react';
import { FirebaseProvider, useFirebase } from './components/FirebaseProvider';
import { DeleteConfirmation } from './components/DeleteConfirmation';
import { Pagination } from './components/Pagination';
import { 
  getAuth, 
  signInWithEmailAndPassword, 
  createUserWithEmailAndPassword, 
  signOut 
} from 'firebase/auth';
import { useDataQuery, where } from './hooks/useDataQuery';
import { createSpreadsheet, updateSheetValues, getSheetValues } from './lib/sheets';
import { saveData, deleteData, clearSheetMemoryCache } from './lib/dataService';
import { cn } from './lib/utils';
import { UserProfile, UserRole, MosqueLocation, Jamaah, Asset, Activity, FacilityStat, JamaahCategory, Attendance, UBShopping, RegistrationLink } from './types';

// --- Firestore Error Handling ---
enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null, auth: any) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth?.currentUser?.uid,
      email: auth?.currentUser?.email,
      emailVerified: auth?.currentUser?.emailVerified,
      isAnonymous: auth?.currentUser?.isAnonymous,
      tenantId: auth?.currentUser?.tenantId,
      providerInfo: auth?.currentUser?.providerData?.map((provider: any) => ({
        providerId: provider.providerId,
        email: provider.email,
      })) || []
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// --- Components ---

function StatCard({ title, value, total, icon: Icon, color }: any) {
  const percentage = total ? Math.round((value / total) * 100) : 0;
  
  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm"
    >
      <div className="flex items-center justify-between mb-4">
        <div className={cn("p-3 rounded-xl", color)}>
          <Icon className="w-6 h-6 text-white" />
        </div>
        {total && (
          <span className="text-xs font-medium text-slate-400">
            {percentage}% Kapasitas
          </span>
        )}
      </div>
      <div>
        <h3 className="text-sm font-medium text-slate-500">{title}</h3>
        <div className="flex items-baseline gap-2 mt-1">
          <p className="text-2xl font-bold text-slate-900">{value}</p>
          {total && <p className="text-sm text-slate-400">/ {total}</p>}
        </div>
      </div>
      {total && (
        <div className="w-full bg-slate-100 h-1.5 rounded-full mt-4 overflow-hidden">
          <motion.div 
            initial={{ width: 0 }}
            animate={{ width: `${percentage}%` }}
            className={cn("h-full", color.replace('bg-', 'bg-opacity-80 bg-'))}
          />
        </div>
      )}
    </motion.div>
  );
}

function SidebarItem({ icon: Icon, label, active, onClick }: { icon: any; label: string; active: boolean; onClick: () => void }) {
  return (
    <button
      onClick={onClick}
      className={cn(
        "w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all duration-200",
        active 
          ? "bg-emerald-50 text-emerald-700 shadow-sm" 
          : "text-slate-600 hover:bg-slate-50 hover:text-slate-900"
      )}
    >
      <Icon className={cn("w-5 h-5", active ? "text-emerald-600" : "text-slate-400")} />
      <span className="font-medium">{label}</span>
    </button>
  );
}

function PublicAttendanceView({ onBack, spreadsheetId }: { onBack: () => void, spreadsheetId?: string }) {
  const { accessToken } = useFirebase();
  const { showToast } = useToast();
  
  const [selectedLocation, setSelectedLocation] = useState<MosqueLocation | ''>(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      const loc = params.get('location');
      if (loc && ['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].includes(loc)) {
        return loc as MosqueLocation;
      }
    }
    return '';
  });

  const [isFromQR] = useState(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      return Boolean(params.get('location'));
    }
    return false;
  });

  const [qrDateParam] = useState(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      return params.get('date') || '';
    }
    return '';
  });

  const [searchTerm, setSearchTerm] = useState('');
  const [selectedJamaah, setSelectedJamaah] = useState<Jamaah | null>(null);
  
  const [sessionType, setSessionType] = useState<'Kelompok' | 'Desa' | 'Acara'>(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      const sess = params.get('session');
      if (sess && ['Kelompok', 'Desa', 'Acara'].includes(sess)) {
        return sess as 'Kelompok' | 'Desa' | 'Acara';
      }
    }
    return 'Kelompok';
  });

  const [status, setStatus] = useState<'hadir' | 'izin'>('hadir');
  const [reason, setReason] = useState('');
  const [loading, setLoading] = useState(false);
  const [success, setSuccess] = useState(false);
  const [showAlreadyAttended, setShowAlreadyAttended] = useState(false);

  // QR Code Generator Modal State
  const [showQRModal, setShowQRModal] = useState(false);
  const [qrTargetLocation, setQrTargetLocation] = useState<MosqueLocation>(() => {
    return (selectedLocation as MosqueLocation) || 'Kramat Batu';
  });
  const [qrTargetDate, setQrTargetDate] = useState(() => {
    const today = new Date();
    return today.toISOString().split('T')[0];
  });
  const [qrTargetSession, setQrTargetSession] = useState<'Kelompok' | 'Desa' | 'Acara'>('Kelompok');
  const [copiedLink, setCopiedLink] = useState(false);

  const { data: jamaahList } = useDataQuery<Jamaah>('jamaah');
  const { data: attendanceList } = useDataQuery<Attendance>('attendance');

  const filteredJamaah = useMemo(() => {
    if (!searchTerm || selectedJamaah) return [];
    return jamaahList.filter(j => {
      const matchesSearch = j.name.toLowerCase().includes(searchTerm.toLowerCase()) || 
        j.memberId.toLowerCase().includes(searchTerm.toLowerCase());
      if (selectedLocation) {
        return matchesSearch && j.location === selectedLocation;
      }
      return matchesSearch;
    }).slice(0, 5);
  }, [jamaahList, searchTerm, selectedJamaah, selectedLocation]);

  const qrPayloadUrl = useMemo(() => {
    if (typeof window === 'undefined') return '';
    const baseUrl = `${window.location.origin}${window.location.pathname}`;
    const params = new URLSearchParams();
    params.set('attendance', 'true');
    params.set('location', qrTargetLocation);
    params.set('date', qrTargetDate);
    params.set('session', qrTargetSession);
    return `${baseUrl}?${params.toString()}`;
  }, [qrTargetLocation, qrTargetDate, qrTargetSession]);

  const handleDownloadQR = () => {
    const canvas = document.getElementById('attendance-qr-canvas') as HTMLCanvasElement;
    if (!canvas) return;
    const url = canvas.toDataURL('image/png');
    const a = document.createElement('a');
    a.href = url;
    a.download = `QR_Presensi_${qrTargetLocation.replace(/\s+/g, '_')}_${qrTargetDate}.png`;
    a.click();
    showToast('Gambar QR Code berhasil diunduh!', 'success');
  };

  const handlePrintQR = () => {
    const canvas = document.getElementById('attendance-qr-canvas') as HTMLCanvasElement;
    const dataUrl = canvas ? canvas.toDataURL('image/png') : '';
    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      window.print();
      return;
    }
    printWindow.document.write(`
      <!DOCTYPE html>
      <html>
        <head>
          <title>Cetak QR Presensi - ${qrTargetLocation}</title>
          <style>
            body { font-family: system-ui, -apple-system, sans-serif; text-align: center; padding: 40px; margin: 0; background: #f8fafc; }
            .card { max-width: 440px; margin: 0 auto; background: #fff; border: 2px solid #059669; border-radius: 28px; padding: 36px 24px; box-shadow: 0 4px 20px rgba(0,0,0,0.06); }
            .tag { display: inline-block; background: #ecfdf5; color: #047857; font-weight: 800; font-size: 11px; padding: 6px 14px; border-radius: 9999px; text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 12px; }
            h1 { color: #0f172a; margin: 0 0 4px; font-size: 24px; font-weight: 900; }
            .loc { color: #059669; font-size: 18px; font-weight: 800; margin-bottom: 4px; }
            .meta { color: #64748b; font-size: 13px; font-weight: 600; margin-bottom: 24px; }
            .qr-wrapper { display: inline-block; padding: 14px; border: 1px solid #e2e8f0; border-radius: 20px; background: #fff; margin-bottom: 20px; }
            img { width: 240px; height: 240px; display: block; }
            .box { background: #f8fafc; border-radius: 16px; padding: 14px 18px; text-align: left; font-size: 12px; color: #334155; line-height: 1.6; border: 1px solid #f1f5f9; }
            .desc { margin-top: 18px; font-size: 12px; color: #059669; font-weight: 700; }
          </style>
        </head>
        <body>
          <div class="card">
            <div class="tag">Presensi Sambung Jamaah</div>
            <h1>Daftar Hadir Pengajian</h1>
            <div class="loc">Kelompok ${qrTargetLocation}</div>
            <div class="meta">Tanggal: ${qrTargetDate} • Sesi: ${qrTargetSession}</div>
            <div class="qr-wrapper">
              <img src="${dataUrl}" alt="QR Presensi" />
            </div>
            <div class="box">
              <div>📍 <strong>Lokasi:</strong> Kelompok ${qrTargetLocation}</div>
              <div>📅 <strong>Tanggal:</strong> ${qrTargetDate}</div>
              <div>📖 <strong>Sesi:</strong> ${qrTargetSession}</div>
            </div>
            <div class="desc">Arahkan kamera HP ke QR Code untuk mengisi absensi secara instan.</div>
          </div>
          <script>
            window.onload = function() {
              setTimeout(function() { window.print(); }, 250);
            };
          </script>
        </body>
      </html>
    `);
    printWindow.document.close();
  };

  const handleCopyLink = async () => {
    try {
      await navigator.clipboard.writeText(qrPayloadUrl);
      setCopiedLink(true);
      showToast('Tautan presensi berhasil disalin ke clipboard!', 'success');
      setTimeout(() => setCopiedLink(false), 2500);
    } catch {
      showToast('Gagal menyalin tautan', 'error');
    }
  };

  const handleSubmit = async () => {
    if (!selectedJamaah) return;
    
    setLoading(true);
    try {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const startOfDay = today.getTime();
      
      const nextDay = new Date(today);
      nextDay.setDate(today.getDate() + 1);
      const endOfDay = nextDay.getTime();

      // Duplicate check against fetched attendance
      const alreadyAttended = attendanceList.some(a => 
        a.jamaahId === selectedJamaah.id && 
        Number(a.date) >= startOfDay && 
        Number(a.date) < endOfDay
      );

      if (alreadyAttended) {
        setShowAlreadyAttended(true);
        setLoading(false);
        return;
      }

      const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
      const dayName = days[new Date().getDay()];

      const data = {
        jamaahId: selectedJamaah.id,
        jamaahName: selectedJamaah.name,
        location: selectedJamaah.location,
        category: selectedJamaah.category,
        date: Date.now(),
        sessionType,
        day: dayName,
        status,
        reason: status === 'izin' ? reason : ''
      };

      await saveData(null, accessToken, spreadsheetId, 'attendance', data);
      setSuccess(true);
      showToast('Data presensi berhasil disimpan!', 'success');
      setSelectedJamaah(null);
      setSearchTerm('');
      setStatus('hadir');
      setReason('');
    } catch (error: any) {
      console.error(error);
      showToast(`Gagal mengirim absensi: ${error.message || 'Error tidak diketahui'}`, 'error');
    } finally {
      setLoading(false);
    }
  };

  const getDayStatus = () => {
    const today = new Date().getDay(); // 0-6 (Sun-Sat)
    const schedules = [
      { day: 1, name: 'Senin', sessions: ['UMUM (Kelompok/Desa)', 'GPN'] },
      { day: 2, name: 'Selasa', sessions: ['GPN'] },
      { day: 3, name: 'Rabu', sessions: ['UMUM (Kelompok/Desa)', 'GPN'] },
      { day: 4, name: 'Kamis', sessions: ['UMUM (Kelompok/Desa)', 'GPN'] },
      { day: 6, name: 'Sabtu', sessions: ['Acara Kelompok/Desa'] },
      { day: 0, name: 'Minggu', sessions: ['Acara Kelompok/Desa'] },
    ];
    const current = schedules.find(s => s.day === today);
    return current ? current.sessions.join(', ') : 'Tidak ada jadwal rutin';
  };

  if (success) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50 p-4">
        <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} className="bg-white p-10 rounded-[2.5rem] shadow-2xl max-w-sm w-full text-center border border-slate-100">
          <div className="w-24 h-24 bg-emerald-100 text-emerald-600 rounded-3xl flex items-center justify-center mx-auto mb-8 shadow-inner">
            <CheckCircle2 className="w-12 h-12" />
          </div>
          <h2 className="text-3xl font-black text-slate-900 mb-3 tracking-tight">Sukses!</h2>
          <p className="text-slate-500 mb-10 font-medium">Absensi Anda telah dicatat untuk pengajian hari ini.</p>
          <div className="space-y-4">
            <button onClick={() => setSuccess(false)} className="w-full py-4 bg-emerald-600 text-white rounded-2xl font-black uppercase tracking-wider hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200 active:scale-95">Absen Lagi</button>
            <button onClick={onBack} className="w-full py-4 text-slate-400 font-bold hover:text-slate-600 transition-colors">Selesai & Keluar</button>
          </div>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-50 p-4">
      <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="bg-white p-8 lg:p-10 rounded-[2.5rem] shadow-2xl max-w-md w-full border border-slate-100 relative overflow-hidden">
        <div className="absolute top-0 left-0 w-full h-2 bg-emerald-500" />
        
        <div className="flex items-center justify-between mb-8">
          <button onClick={onBack} className="p-3 bg-slate-50 rounded-2xl text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition-all"><X className="w-6 h-6" /></button>
          <div className="text-center">
            <h2 className="text-lg font-black uppercase tracking-widest text-emerald-600 leading-none">Absensi</h2>
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-[0.3em] mt-1">Sambung Jamaah</p>
          </div>
          <button 
            type="button"
            onClick={() => {
              if (selectedLocation) {
                setQrTargetLocation(selectedLocation as MosqueLocation);
              }
              setShowQRModal(true);
            }}
            title="Generate QR Code Presensi" 
            className="p-3 bg-emerald-50 text-emerald-600 hover:text-emerald-700 hover:bg-emerald-100 rounded-2xl transition-all flex items-center justify-center shadow-sm"
          >
            <QrCode className="w-6 h-6" />
          </button>
        </div>

        {/* QR Code Quick Action Banner */}
        <div className="mb-6 p-3.5 bg-gradient-to-r from-emerald-50 to-teal-50 rounded-2xl border border-emerald-100 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-xl bg-emerald-600 text-white flex items-center justify-center shadow-sm">
              <QrCode className="w-4 h-4" />
            </div>
            <div>
              <p className="text-xs font-black text-slate-800">QR Code Presensi</p>
              <p className="text-[10px] text-slate-500 font-medium">Buat QR code lokasi & tanggal untuk jamaah</p>
            </div>
          </div>
          <button
            type="button"
            onClick={() => {
              if (selectedLocation) {
                setQrTargetLocation(selectedLocation as MosqueLocation);
              }
              setShowQRModal(true);
            }}
            className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-black transition-all shadow-sm flex items-center gap-1.5 active:scale-95"
          >
            <QrCode className="w-3.5 h-3.5" />
            Buat QR
          </button>
        </div>

        {/* Notifikasi aktif jika jamaah membuka lewat scan QR */}
        {isFromQR && selectedLocation && (
          <div className="mb-6 p-3.5 bg-emerald-600 text-white rounded-2xl shadow-md flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-8 h-8 rounded-xl bg-white/20 flex items-center justify-center backdrop-blur-sm">
                <CheckCircle2 className="w-4 h-4 text-white" />
              </div>
              <div>
                <p className="text-[10px] font-black uppercase tracking-wider text-emerald-100">Scan QR Terverifikasi</p>
                <p className="text-xs font-black text-white">Kelompok {selectedLocation} {qrDateParam ? `• ${qrDateParam}` : ''}</p>
              </div>
            </div>
            <button 
              type="button" 
              onClick={() => { setSelectedLocation(''); setSelectedJamaah(null); }}
              className="text-[10px] font-bold bg-white/20 hover:bg-white/30 text-white px-2.5 py-1 rounded-lg transition-colors"
            >
              Ganti Lokasi
            </button>
          </div>
        )}

        <div className="mb-8 p-4 bg-amber-50 rounded-2xl border border-amber-100">
          <div className="flex items-start gap-3">
            <Clock className="w-5 h-5 text-amber-500 mt-0.5" />
            <div>
              <p className="text-[10px] font-black text-amber-600 uppercase tracking-widest">Jadwal Hari Ini</p>
              <p className="text-sm font-bold text-slate-700 mt-0.5">{getDayStatus()}</p>
            </div>
          </div>
        </div>

        <div className="space-y-8">
          <div className="relative">
            <label className="block text-[11px] font-black text-slate-400 mb-2 uppercase tracking-[0.2em] ml-1">1. Pilih Lokasi Kelompok</label>
            <div className="relative">
              <MapPin className="absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5 text-emerald-500" />
              <select 
                value={selectedLocation} 
                onChange={(e) => {
                  setSelectedLocation(e.target.value as MosqueLocation);
                  setSelectedJamaah(null);
                  setSearchTerm('');
                }}
                className="w-full pl-12 pr-4 py-4 rounded-2xl border border-slate-200 outline-none focus:ring-4 focus:ring-emerald-500/10 focus:border-emerald-500 transition-all font-bold text-slate-700 bg-slate-50/50 appearance-none cursor-pointer"
              >
                <option value="">-- Pilih Lokasi --</option>
                {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(l => (
                  <option key={l} value={l}>{l}</option>
                ))}
              </select>
            </div>
          </div>

          {selectedLocation && (
            <motion.div initial={{ opacity: 0, height: 0 }} animate={{ opacity: 1, height: 'auto' }} className="space-y-8">
              <div className="relative">
                <label className="block text-[11px] font-black text-slate-400 mb-2 uppercase tracking-[0.2em] ml-1">2. Masukkan Nama Anda</label>
                <div className="relative">
                  <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
                  <input
                    type="text"
                    value={searchTerm}
                    onChange={(e) => {
                      setSearchTerm(e.target.value);
                      if (selectedJamaah) setSelectedJamaah(null);
                    }}
                    placeholder="Cari nama Anda..."
                    className="w-full pl-12 pr-4 py-4 rounded-2xl border border-slate-200 outline-none focus:ring-4 focus:ring-emerald-500/10 focus:border-emerald-500 transition-all font-bold text-slate-700 bg-slate-50/50"
                  />
                </div>
                
                <AnimatePresence>
                  {filteredJamaah.length > 0 && !selectedJamaah && (
                    <motion.div 
                      initial={{ opacity: 0, y: -10 }} 
                      animate={{ opacity: 1, y: 0 }} 
                      exit={{ opacity: 0, y: -10 }}
                      className="absolute z-10 left-0 right-0 mt-3 bg-white rounded-3xl border border-slate-100 shadow-2xl overflow-hidden max-h-60 overflow-y-auto ring-1 ring-slate-200"
                    >
                      {filteredJamaah.map((j, idx) => (
                        <button
                          key={j.id ? `${j.id}-${idx}` : `search-j-${idx}`}
                          onClick={() => {
                            setSelectedJamaah(j);
                            setSearchTerm(j.name);
                          }}
                          className="w-full px-6 py-4 text-left hover:bg-emerald-50 border-b border-slate-50 last:border-0 flex justify-between items-center group transition-colors"
                        >
                          <div>
                            <p className="font-black text-slate-900 group-hover:text-emerald-700 transition-colors">{j.name}</p>
                            <div className="flex items-center gap-2 mt-1">
                              <span className="px-2 py-0.5 rounded-md bg-slate-100 text-[9px] font-black uppercase text-slate-500">{j.category}</span>
                              <span className="text-[10px] font-bold text-slate-400">{j.memberId}</span>
                            </div>
                          </div>
                          <div className="w-8 h-8 rounded-full bg-slate-50 flex items-center justify-center group-hover:bg-emerald-100 group-hover:text-emerald-600 transition-all">
                            <Plus className="w-4 h-4" />
                          </div>
                        </button>
                      ))}
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              {selectedJamaah && (
                <motion.div initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} className="p-5 bg-emerald-600 rounded-[2rem] text-white shadow-xl shadow-emerald-200 relative overflow-hidden group">
                  <CheckCircle2 className="absolute -right-4 -bottom-4 w-24 h-24 text-white/10 rotate-12 group-hover:scale-110 transition-transform duration-500" />
                  <p className="text-[10px] font-black text-emerald-200 uppercase tracking-[0.2em] mb-2">Konfirmasi Identitas</p>
                  <p className="font-black text-xl leading-tight">{selectedJamaah.name}</p>
                  <div className="flex items-center gap-2 mt-3">
                    <span className="px-2 py-1 rounded-lg bg-white/20 text-[9px] font-black uppercase backdrop-blur-sm">{selectedJamaah.category}</span>
                    <span className="px-2 py-1 rounded-lg bg-white/20 text-[9px] font-black uppercase backdrop-blur-sm">{selectedJamaah.location}</span>
                  </div>
                </motion.div>
              )}

              <div className="space-y-3">
                <label className="block text-[11px] font-black text-slate-400 mb-2 uppercase tracking-[0.2em] ml-1">3. Jenis Pengajian</label>
                <div className="grid grid-cols-3 gap-3">
                  {(['Kelompok', 'Desa', 'Acara'] as const).map(type => (
                    <button
                      key={type}
                      onClick={() => setSessionType(type)}
                      className={cn(
                        "py-3 rounded-2xl border-2 font-black text-xs uppercase tracking-wider transition-all",
                        sessionType === type 
                          ? "bg-emerald-600 border-emerald-600 text-white shadow-lg shadow-emerald-100" 
                          : "bg-white border-slate-100 text-slate-400 hover:border-emerald-200 hover:text-emerald-600"
                      )}
                    >
                      {type}
                    </button>
                  ))}
                </div>
              </div>

              <div className="space-y-3">
                <label className="block text-[11px] font-black text-slate-400 mb-2 uppercase tracking-[0.2em] ml-1">4. Status Kehadiran</label>
                <div className="grid grid-cols-2 gap-3">
                  {(['hadir', 'izin'] as const).map(s => (
                    <button
                      key={s}
                      onClick={() => setStatus(s)}
                      className={cn(
                        "py-3 rounded-2xl border-2 font-black text-xs uppercase tracking-wider transition-all",
                        status === s 
                          ? "bg-emerald-600 border-emerald-600 text-white shadow-lg shadow-emerald-100" 
                          : "bg-white border-slate-100 text-slate-400 hover:border-emerald-200 hover:text-emerald-600"
                      )}
                    >
                      {s === 'hadir' ? 'Hadir' : 'Izin'}
                    </button>
                  ))}
                </div>
              </div>

              {status === 'izin' && (
                <motion.div initial={{ opacity: 0, height: 0 }} animate={{ opacity: 1, height: 'auto' }} className="space-y-3">
                  <label className="block text-[11px] font-black text-slate-400 mb-2 uppercase tracking-[0.2em] ml-1">Alasan Izin</label>
                  <textarea
                    value={reason}
                    onChange={(e) => setReason(e.target.value)}
                    placeholder="Sebutkan alasan izin..."
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 outline-none focus:ring-4 focus:ring-emerald-500/10 focus:border-emerald-500 transition-all font-bold text-slate-700 bg-slate-50/50 min-h-[100px]"
                    required
                  />
                  <p className="text-[10px] text-slate-400 font-bold italic leading-relaxed">*Anda harus sudah melakukan izin terlebih dahulu melalui WhatsApp kepada pengurus kelompok.</p>
                </motion.div>
              )}

              <button
                disabled={!selectedJamaah || loading}
                onClick={handleSubmit}
                className={cn(
                  "w-full py-5 rounded-[2rem] font-black uppercase tracking-[0.2em] text-white transition-all shadow-xl",
                  !selectedJamaah || loading 
                    ? "bg-slate-200 cursor-not-allowed shadow-none" 
                    : "bg-emerald-600 hover:bg-emerald-700 active:scale-95 shadow-emerald-200"
                )}
              >
                {loading ? (
                  <div className="flex items-center justify-center gap-3">
                    <motion.div animate={{ rotate: 360 }} transition={{ repeat: Infinity, duration: 1, ease: "linear" }} className="w-5 h-5 border-2 border-white border-t-transparent rounded-full" />
                    Memproses...
                  </div>
                ) : "Kirim Absensi"}
              </button>
            </motion.div>
          )}
        </div>
      </motion.div>

      <AnimatePresence>
        {showQRModal && (
          <motion.div 
            initial={{ opacity: 0 }} 
            animate={{ opacity: 1 }} 
            exit={{ opacity: 0 }} 
            className="fixed inset-0 bg-black/60 backdrop-blur-md flex items-center justify-center p-4 z-[100] overflow-y-auto"
          >
            <motion.div 
              initial={{ scale: 0.92, y: 15 }} 
              animate={{ scale: 1, y: 0 }} 
              exit={{ scale: 0.92, y: 15 }}
              className="bg-white rounded-[2.5rem] p-6 sm:p-8 max-w-md w-full shadow-2xl border border-slate-100 relative my-8"
            >
              <div className="flex items-center justify-between mb-6">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-2xl bg-emerald-100 text-emerald-600 flex items-center justify-center">
                    <QrCode className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-lg font-black text-slate-900 leading-tight">QR Code Presensi</h3>
                    <p className="text-xs text-slate-500 font-medium">Embed lokasi & tanggal pengajian</p>
                  </div>
                </div>
                <button 
                  onClick={() => setShowQRModal(false)}
                  className="p-2 rounded-xl text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Konfigurasi Lokasi, Tanggal & Sesi */}
              <div className="space-y-4 mb-6">
                <div>
                  <label className="block text-[10px] font-black uppercase tracking-wider text-slate-400 mb-1.5">
                    1. Lokasi Masjid
                  </label>
                  <div className="grid grid-cols-3 gap-1.5">
                    {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(loc => (
                      <button
                        key={loc}
                        type="button"
                        onClick={() => setQrTargetLocation(loc as MosqueLocation)}
                        className={cn(
                          "px-2 py-2 rounded-xl text-[11px] font-bold transition-all truncate text-center",
                          qrTargetLocation === loc 
                            ? "bg-emerald-600 text-white shadow-sm font-black" 
                            : "bg-slate-50 text-slate-600 hover:bg-slate-100 border border-slate-200/60"
                        )}
                      >
                        {loc}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-[10px] font-black uppercase tracking-wider text-slate-400 mb-1.5">
                      2. Tanggal Presensi
                    </label>
                    <input 
                      type="date"
                      value={qrTargetDate}
                      onChange={(e) => setQrTargetDate(e.target.value)}
                      className="w-full px-3 py-2 rounded-xl border border-slate-200 text-xs font-bold text-slate-700 outline-none focus:ring-2 focus:ring-emerald-500 bg-slate-50"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-black uppercase tracking-wider text-slate-400 mb-1.5">
                      3. Sesi Pengajian
                    </label>
                    <select
                      value={qrTargetSession}
                      onChange={(e) => setQrTargetSession(e.target.value as any)}
                      className="w-full px-3 py-2 rounded-xl border border-slate-200 text-xs font-bold text-slate-700 outline-none focus:ring-2 focus:ring-emerald-500 bg-slate-50"
                    >
                      <option value="Kelompok">Kelompok</option>
                      <option value="Desa">Desa</option>
                      <option value="Acara">Acara</option>
                    </select>
                  </div>
                </div>
              </div>

              {/* Preview QR Code Card */}
              <div className="p-5 bg-gradient-to-b from-slate-50 to-emerald-50/40 rounded-2xl border border-emerald-100 text-center mb-6">
                <div className="inline-block p-3 bg-white rounded-2xl shadow-md border border-slate-100 mb-3">
                  <QRCodeCanvas
                    id="attendance-qr-canvas"
                    value={qrPayloadUrl}
                    size={190}
                    level="H"
                    marginSize={2}
                  />
                </div>
                <p className="text-xs font-black text-slate-900">Kelompok {qrTargetLocation}</p>
                <p className="text-[11px] text-slate-500 font-semibold mt-0.5">
                  Tanggal: {qrTargetDate} • Sesi {qrTargetSession}
                </p>
                <p className="text-[10px] text-emerald-700 font-bold mt-2 bg-emerald-100/70 py-1 px-3 rounded-lg inline-block">
                  Arahkan kamera HP ke QR Code untuk membuka presensi instan
                </p>
              </div>

              {/* Action Buttons */}
              <div className="space-y-2.5">
                <div className="grid grid-cols-2 gap-2.5">
                  <button
                    type="button"
                    onClick={handleDownloadQR}
                    className="py-3 px-4 rounded-xl bg-slate-900 text-white text-xs font-bold hover:bg-slate-800 transition-all flex items-center justify-center gap-2 shadow-sm active:scale-95"
                  >
                    <Download className="w-4 h-4" />
                    Unduh Gambar
                  </button>
                  <button
                    type="button"
                    onClick={handlePrintQR}
                    className="py-3 px-4 rounded-xl bg-emerald-600 text-white text-xs font-bold hover:bg-emerald-700 transition-all flex items-center justify-center gap-2 shadow-sm active:scale-95"
                  >
                    <Printer className="w-4 h-4" />
                    Cetak QR
                  </button>
                </div>

                <button
                  type="button"
                  onClick={handleCopyLink}
                  className="w-full py-2.5 px-4 rounded-xl border border-slate-200 text-slate-700 text-xs font-bold hover:bg-slate-50 transition-all flex items-center justify-center gap-2"
                >
                  {copiedLink ? <Check className="w-4 h-4 text-emerald-600" /> : <Copy className="w-4 h-4 text-slate-500" />}
                  {copiedLink ? 'Tautan Disalin ke Clipboard!' : 'Salin Tautan Presensi'}
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}

        {showAlreadyAttended && (
          <motion.div 
            initial={{ opacity: 0 }} 
            animate={{ opacity: 1 }} 
            exit={{ opacity: 0 }} 
            className="fixed inset-0 bg-black/60 backdrop-blur-md flex items-center justify-center p-4 z-[100]"
          >
            <motion.div 
              initial={{ scale: 0.9, y: 20 }} 
              animate={{ scale: 1, y: 0 }} 
              className="bg-white rounded-[2.5rem] p-10 max-w-sm w-full text-center shadow-2xl border border-slate-100"
            >
              <div className="w-24 h-24 bg-amber-100 text-amber-600 rounded-3xl flex items-center justify-center mx-auto mb-8 shadow-inner">
                <AlertCircle className="w-12 h-12" />
              </div>
              <h2 className="text-2xl font-black text-slate-900 mb-4 tracking-tight leading-tight">Sudah Melakukan Absensi</h2>
              <p className="text-slate-500 mb-10 font-medium leading-relaxed">
                Anda sudah tercatat melakukan absensi hari ini. Jika ada perubahan data, mohon hubungi <span className="font-black text-slate-900">pengurus kelompok</span> Anda.
              </p>
              <button 
                onClick={() => setShowAlreadyAttended(false)} 
                className="w-full py-4 bg-slate-900 text-white rounded-2xl font-black uppercase tracking-wider hover:bg-slate-800 transition-all active:scale-95 shadow-xl shadow-slate-200"
              >
                Tutup
              </button>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

// --- Public Jamaah Self-Registration View ---
function PublicJamaahRegistrationView({ onBack, spreadsheetId }: { onBack: () => void, spreadsheetId?: string }) {
  const { accessToken } = useFirebase();
  const { showToast } = useToast();
  const { data: jamaahList, loading: loadingJamaah } = useDataQuery<Jamaah>('jamaah');

  // Decode URL token payload
  const [tokenData, setTokenData] = useState<{
    id?: string;
    location?: MosqueLocation | 'Seluruh Lokasi';
    expiresAt?: number;
    createdAt?: number;
    createdBy?: string;
    note?: string;
  } | null>(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      const rawToken = params.get('token') || params.get('reg_token');
      if (rawToken) {
        try {
          const decodedStr = decodeURIComponent(atob(rawToken));
          return JSON.parse(decodedStr);
        } catch (e) {
          try {
            return JSON.parse(atob(rawToken));
          } catch (e2) {
            console.warn('Failed to parse registration token:', e2);
          }
        }
      }
      // Fallback if individual query params were supplied
      const loc = params.get('location') as MosqueLocation;
      const exp = params.get('expires');
      if (loc || exp) {
        return {
          location: loc || 'Seluruh Lokasi',
          expiresAt: exp ? Number(exp) : undefined,
          note: params.get('note') || undefined
        };
      }
    }
    return null;
  });

  // Calculate expiration status & remaining time
  const [currentTime, setCurrentTime] = useState(Date.now());
  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(Date.now()), 1000);
    return () => clearInterval(timer);
  }, []);

  const isExpired = useMemo(() => {
    if (!tokenData?.expiresAt) return false;
    return currentTime > tokenData.expiresAt;
  }, [tokenData?.expiresAt, currentTime]);

  const remainingTimeString = useMemo(() => {
    if (!tokenData?.expiresAt) return null;
    const diff = tokenData.expiresAt - currentTime;
    if (diff <= 0) return 'Kedaluwarsa';
    const days = Math.floor(diff / (1000 * 60 * 60 * 24));
    const hours = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
    const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
    const seconds = Math.floor((diff % (1000 * 60)) / 1000);

    if (days > 0) return `${days} hari ${hours} jam lagi`;
    if (hours > 0) return `${hours} jam ${minutes} menit lagi`;
    return `${minutes} menit ${seconds} detik lagi`;
  }, [tokenData?.expiresAt, currentTime]);

  // Form State
  const [activeStep, setActiveStep] = useState<'identitas' | 'keluarga' | 'keilmuan'>('identitas');
  const [selectedLocation, setSelectedLocation] = useState<MosqueLocation>(() => {
    if (tokenData?.location && tokenData.location !== 'Seluruh Lokasi') {
      return tokenData.location as MosqueLocation;
    }
    return 'Kramat Batu';
  });

  const [isKK, setIsKK] = useState(true);
  const [selectedKKId, setSelectedKKId] = useState<string>('');
  const [kkSearchTerm, setKkSearchTerm] = useState<string>('');
  const [isKkDropdownOpen, setIsKkDropdownOpen] = useState<boolean>(false);
  const [base64Photo, setBase64Photo] = useState<string | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [registeredResult, setRegisteredResult] = useState<{
    memberId: string;
    name: string;
    location: string;
    isKK: boolean;
    phone: string;
  } | null>(null);

  // Filter Kepala Keluarga in the selected location
  const headsOfFamily = useMemo(() => {
    return jamaahList.filter(j => 
      (j.isKK === true || (j.isKK as any) === 'true' || (j.isKK as any) === 'TRUE') &&
      j.location === selectedLocation
    );
  }, [jamaahList, selectedLocation]);

  const filteredKKList = useMemo(() => {
    if (!kkSearchTerm.trim()) return headsOfFamily;
    const term = kkSearchTerm.toLowerCase();
    return headsOfFamily.filter(k => 
      (k.name && k.name.toLowerCase().includes(term)) ||
      (k.nickname && k.nickname.toLowerCase().includes(term)) ||
      (k.memberId && k.memberId.toLowerCase().includes(term)) ||
      (k.phone && k.phone.includes(term))
    );
  }, [headsOfFamily, kkSearchTerm]);

  const handlePhotoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 2 * 1024 * 1024) {
        showToast('Ukuran foto maksimal 2MB', 'error');
        return;
      }
      const reader = new FileReader();
      reader.onloadend = () => {
        setBase64Photo(reader.result as string);
      };
      reader.readAsDataURL(file);
    }
  };

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();

    if (isExpired) {
      showToast('Tautan pendaftaran ini sudah kedaluwarsa.', 'error');
      return;
    }

    const formData = new FormData(e.currentTarget);
    const nameVal = (formData.get('name') as string || '').trim();
    if (!nameVal) {
      setActiveStep('identitas');
      showToast('Nama lengkap wajib diisi.', 'error');
      return;
    }

    const phoneVal = (formData.get('phone') as string || '').trim();
    if (!phoneVal) {
      setActiveStep('identitas');
      showToast('Nomor telepon / WA wajib diisi.', 'error');
      return;
    }

    if (!isKK && !selectedKKId) {
      setActiveStep('keluarga');
      showToast('Silakan pilih Kepala Keluarga terlebih dahulu.', 'error');
      return;
    }

    // Automatic Member ID Calculation
    const prefix = LOCATION_PREFIXES[selectedLocation] || 'JM';
    let memberId = '';
    let familyOrder = 0;
    let finalKKId = '';

    if (isKK) {
      const kkInLocation = jamaahList.filter(j => 
        j.location === selectedLocation && 
        (j.isKK === true || (j.isKK as any) === 'true' || (j.isKK as any) === 'TRUE')
      );
      let maxNum = 0;
      kkInLocation.forEach(kk => {
        const numPart = (kk.memberId || '').replace(prefix, '');
        const num = parseInt(numPart);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      });
      const nextNum = maxNum + 1;
      memberId = `${prefix}${nextNum.toString().padStart(3, '0')}`;
      familyOrder = 0;
      finalKKId = memberId;
    } else {
      const kk = jamaahList.find(j => j.memberId === selectedKKId);
      const familyMembers = jamaahList.filter(j => j.kkId === selectedKKId);
      let maxOrder = 0;
      familyMembers.forEach(m => {
        if (m.familyOrder > maxOrder) maxOrder = m.familyOrder;
      });
      familyOrder = maxOrder + 1;
      memberId = `${selectedKKId}-${familyOrder.toString().padStart(2, '0')}`;
      finalKKId = selectedKKId;
    }

    const originAddress = (formData.get('originAddress') as string) || '';
    const currentAddress = (formData.get('currentAddress') as string) || '';

    const newJamaahData: Jamaah = {
      id: Math.random().toString(36).substring(2, 11),
      memberId,
      name: nameVal,
      nickname: (formData.get('nickname') as string) || '',
      gender: (formData.get('gender') as string) || 'Laki-laki',
      bloodType: (formData.get('bloodType') as string) || '',
      originAddress,
      currentAddress,
      address: currentAddress || originAddress,
      phone: phoneVal,
      location: selectedLocation,
      category: (formData.get('category') as JamaahCategory) || 'UMUM',
      isKK,
      kkId: finalKKId,
      familyOrder,
      dapukan: [],
      positions: [],
      photoUrl: base64Photo || undefined,
      registeredAt: Date.now(),
      placeOfBirth: (formData.get('placeOfBirth') as string) || '',
      dateOfBirth: (formData.get('dateOfBirth') as string) || '',
      fatherName: (formData.get('fatherName') as string) || '',
      motherName: (formData.get('motherName') as string) || '',
      parentPhone: (formData.get('parentPhone') as string) || '',
      lastEducation: (formData.get('lastEducation') as string) || '',
      majorOrClass: (formData.get('majorOrClass') as string) || '',
      schoolOrUniversity: (formData.get('schoolOrUniversity') as string) || '',
      currentJob: (formData.get('currentJob') as string) || '',
      workplaceAddress: (formData.get('workplaceAddress') as string) || '',
      maritalStatus: (formData.get('maritalStatus') as string) || 'Belum Menikah',
      marriageYear: (formData.get('marriageYear') as string) || '',
      spouseName: (formData.get('spouseName') as string) || '',
      hasJurusKeras: (formData.get('hasJurusKeras') as string) || 'Belum',
      hasJurusHalus: (formData.get('hasJurusHalus') as string) || 'Belum',
      hasUbShares: (formData.get('hasUbShares') as string) || 'Tidak',
      isMubaligh: (formData.get('isMubaligh') as string) || 'Bukan',
      hasHajj: (formData.get('hasHajj') as string) || 'Belum',
      hajjPortionNumber: (formData.get('hajjPortionNumber') as string) || '',
      plannedHajjYear: (formData.get('plannedHajjYear') as string) || '',
      hajjName: (formData.get('hajjName') as string) || '',
      hajjYear: (formData.get('hajjYear') as string) || '',
      medicalHistory: (formData.get('medicalHistory') as string) || ''
    };

    setIsSaving(true);
    try {
      await saveData(null, accessToken, spreadsheetId, 'jamaah', newJamaahData, newJamaahData.id);
      showToast(`Alhamdulillah! Data jamaah ${nameVal} berhasil didaftarkan.`, 'success');
      setRegisteredResult({
        memberId,
        name: nameVal,
        location: selectedLocation,
        isKK,
        phone: phoneVal
      });
    } catch (err: any) {
      showToast(`Gagal menyimpan data: ${err.message || 'Silakan periksa koneksi internet Anda.'}`, 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleResetForAnotherFamilyMember = () => {
    // Keep location and set isKK to false with the newly created KK or current KK
    const previousKK = registeredResult?.isKK ? registeredResult.memberId : selectedKKId;
    if (previousKK) {
      setIsKK(false);
      setSelectedKKId(previousKK);
      const kkItem = jamaahList.find(j => j.memberId === previousKK);
      if (kkItem) {
        setKkSearchTerm(`${kkItem.memberId} - ${kkItem.name}`);
      }
    }
    setRegisteredResult(null);
    setBase64Photo(null);
    setActiveStep('identitas');
  };

  // 1. EXPIRED LINK SCREEN
  if (isExpired) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4">
        <motion.div 
          initial={{ opacity: 0, scale: 0.95 }}
          animate={{ opacity: 1, scale: 1 }}
          className="bg-white rounded-[2.5rem] p-8 md:p-12 shadow-2xl max-w-lg w-full text-center border border-slate-100 relative overflow-hidden"
        >
          <div className="w-20 h-20 bg-amber-50 text-amber-500 rounded-3xl flex items-center justify-center mx-auto mb-6 shadow-inner">
            <Clock className="w-10 h-10" />
          </div>
          <span className="px-3.5 py-1 rounded-full text-xs font-black uppercase tracking-wider bg-amber-100 text-amber-800 inline-block mb-3">
            Tautan Kedaluwarsa
          </span>
          <h2 className="text-2xl md:text-3xl font-black text-slate-900 mb-3 tracking-tight">
            Masa Berlaku Telah Berakhir
          </h2>
          <p className="text-slate-600 text-sm leading-relaxed mb-6">
            Mohon maaf, tautan pengisian formulir pendaftaran jamaah mandiri ini telah melewati batas waktu kedaluwarsa pada:
          </p>

          <div className="bg-slate-50 border border-slate-200/80 rounded-2xl p-4 mb-8 text-left space-y-2">
            <div className="flex justify-between items-center text-xs">
              <span className="text-slate-400 font-medium">Batas Waktu:</span>
              <span className="font-bold text-slate-800">
                {tokenData?.expiresAt ? new Date(tokenData.expiresAt).toLocaleString('id-ID', { dateStyle: 'full', timeStyle: 'short' }) : '-'} WIB
              </span>
            </div>
            {tokenData?.location && (
              <div className="flex justify-between items-center text-xs">
                <span className="text-slate-400 font-medium">Lokasi Kelompok:</span>
                <span className="font-bold text-emerald-700">{tokenData.location}</span>
              </div>
            )}
            {tokenData?.note && (
              <div className="flex justify-between items-center text-xs">
                <span className="text-slate-400 font-medium">Keterangan:</span>
                <span className="font-bold text-slate-700">{tokenData.note}</span>
              </div>
            )}
          </div>

          <p className="text-xs text-slate-500 mb-8 leading-relaxed">
            Silakan hubungi pengurus kelompok atau admin masjid untuk meminta tautan pendaftaran mandiri yang baru.
          </p>

          <button
            type="button"
            onClick={onBack}
            className="w-full py-4 bg-slate-900 text-white rounded-2xl font-black text-sm uppercase tracking-wider hover:bg-slate-800 transition-all shadow-xl shadow-slate-200 active:scale-95"
          >
            Kembali ke Halaman Utama
          </button>
        </motion.div>
      </div>
    );
  }

  // 2. SUCCESS CONFIRMATION SCREEN
  if (registeredResult) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-emerald-900 via-slate-900 to-slate-950 flex items-center justify-center p-4">
        <motion.div 
          initial={{ opacity: 0, scale: 0.92, y: 20 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          className="bg-white rounded-[2.5rem] p-8 md:p-12 shadow-2xl max-w-lg w-full text-center border border-white/20 relative overflow-hidden"
        >
          <div className="w-20 h-20 bg-emerald-100 text-emerald-600 rounded-3xl flex items-center justify-center mx-auto mb-6 shadow-sm">
            <CheckCircle2 className="w-10 h-10" />
          </div>

          <span className="px-3.5 py-1 rounded-full text-xs font-black uppercase tracking-wider bg-emerald-100 text-emerald-800 inline-block mb-3">
            Pendaftaran Berhasil 🎉
          </span>

          <h2 className="text-2xl md:text-3xl font-black text-slate-900 mb-2 tracking-tight">
            Alhamdulillah, Data Tersimpan!
          </h2>
          <p className="text-slate-600 text-sm leading-relaxed mb-6">
            Data diri Anda telah sukses tersimpan di database jamaah Desa Gandaria.
          </p>

          {/* Member Badge Card */}
          <div className="bg-gradient-to-br from-emerald-500 to-teal-700 text-white rounded-3xl p-6 mb-8 text-left shadow-xl shadow-emerald-900/20 relative overflow-hidden">
            <div className="absolute right-[-20px] bottom-[-20px] opacity-10 pointer-events-none">
              <Users className="w-44 h-44" />
            </div>
            <div className="flex justify-between items-start mb-4">
              <div>
                <p className="text-[10px] uppercase font-black tracking-widest text-emerald-200">ID Anggota / Nomor Induk</p>
                <p className="text-2xl font-black tracking-tight">{registeredResult.memberId}</p>
              </div>
              <span className="px-3 py-1 bg-white/20 backdrop-blur-md rounded-full text-[10px] font-bold uppercase">
                {registeredResult.isKK ? 'Kepala Keluarga' : 'Anggota Keluarga'}
              </span>
            </div>

            <div className="space-y-1.5 pt-3 border-t border-white/20 text-xs">
              <div className="flex justify-between">
                <span className="text-emerald-100">Nama Lengkap:</span>
                <span className="font-black text-white">{registeredResult.name}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-emerald-100">Kelompok:</span>
                <span className="font-bold text-white">{registeredResult.location}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-emerald-100">No. WhatsApp:</span>
                <span className="font-medium text-white">{registeredResult.phone}</span>
              </div>
            </div>
          </div>

          <div className="space-y-3">
            <button
              type="button"
              onClick={handleResetForAnotherFamilyMember}
              className="w-full py-4 bg-emerald-600 text-white rounded-2xl font-black text-sm uppercase tracking-wider hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200 active:scale-95 flex items-center justify-center gap-2"
            >
              <UserPlus className="w-5 h-5" />
              <span>Daftarkan Anggota Keluarga Lain</span>
            </button>

            <button
              type="button"
              onClick={onBack}
              className="w-full py-3.5 bg-slate-100 text-slate-700 rounded-2xl font-bold text-sm hover:bg-slate-200 transition-all"
            >
              Selesai & Keluar
            </button>
          </div>
        </motion.div>
      </div>
    );
  }

  // 3. REGISTRATION FORM
  const isLocationLocked = Boolean(tokenData?.location && tokenData.location !== 'Seluruh Lokasi');

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 py-8 px-4 sm:px-6 lg:px-8">
      <div className="max-w-3xl mx-auto">
        {/* Top Header Card */}
        <div className="bg-slate-800/90 border border-slate-700/80 backdrop-blur-xl rounded-[2.5rem] p-6 sm:p-8 mb-8 shadow-2xl">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-700/60 pb-6 mb-6">
            <div className="flex items-center gap-3.5">
              <div className="w-12 h-12 rounded-2xl bg-emerald-500/20 text-emerald-400 border border-emerald-500/30 flex items-center justify-center font-bold shadow-inner">
                <Users className="w-6 h-6" />
              </div>
              <div>
                <span className="text-[10px] font-black uppercase tracking-[0.2em] text-emerald-400">Pendaftaran Mandiri Jamaah</span>
                <h1 className="text-xl sm:text-2xl font-black text-white tracking-tight">Formulir Data Jamaah</h1>
                <p className="text-xs text-slate-400">Desa Gandaria • Tanpa Perlu Login</p>
              </div>
            </div>

            <button
              type="button"
              onClick={onBack}
              className="self-start sm:self-auto px-4 py-2 bg-slate-700/60 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5"
            >
              <ChevronLeft className="w-4 h-4" />
              <span>Kembali</span>
            </button>
          </div>

          {/* Expiration & Location Notification Banner */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            {tokenData?.expiresAt && (
              <div className="p-3.5 bg-emerald-950/50 border border-emerald-500/30 rounded-2xl flex items-center gap-3">
                <div className="w-8 h-8 rounded-xl bg-emerald-500/20 text-emerald-400 flex items-center justify-center flex-shrink-0">
                  <Timer className="w-4 h-4 animate-pulse" />
                </div>
                <div className="overflow-hidden">
                  <p className="text-[10px] font-black uppercase tracking-wider text-emerald-400">Masa Aktif Tautan</p>
                  <p className="text-xs font-bold text-white truncate">{remainingTimeString}</p>
                </div>
              </div>
            )}

            <div className="p-3.5 bg-slate-700/40 border border-slate-600/50 rounded-2xl flex items-center gap-3">
              <div className="w-8 h-8 rounded-xl bg-blue-500/20 text-blue-400 flex items-center justify-center flex-shrink-0">
                <MapPin className="w-4 h-4" />
              </div>
              <div className="overflow-hidden">
                <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">Lokasi Kelompok</p>
                <p className="text-xs font-bold text-white truncate">
                  {selectedLocation} {isLocationLocked ? '(Terkunci dari Tautan)' : ''}
                </p>
              </div>
            </div>
          </div>
        </div>

        {/* Main Form Box */}
        <div className="bg-white text-slate-900 rounded-[2.5rem] p-6 sm:p-10 shadow-2xl border border-slate-100">
          {/* Step Tabs Navigation */}
          <div className="flex border-b border-slate-100 mb-8 overflow-x-auto gap-2 sm:gap-4 pb-2">
            {[
              { id: 'identitas', label: '1. Identitas Diri', icon: Users },
              { id: 'keluarga', label: '2. Keluarga & Domisili', icon: Home },
              { id: 'keilmuan', label: '3. Keilmuan & Ibadah', icon: HeartPulse }
            ].map(tab => (
              <button
                key={tab.id}
                type="button"
                onClick={() => setActiveStep(tab.id as any)}
                className={cn(
                  "px-4 py-3 rounded-2xl font-bold text-xs sm:text-sm flex items-center gap-2 whitespace-nowrap transition-all",
                  activeStep === tab.id
                    ? "bg-emerald-600 text-white shadow-lg shadow-emerald-200"
                    : "text-slate-500 hover:bg-slate-100"
                )}
              >
                <tab.icon className="w-4 h-4" />
                <span>{tab.label}</span>
              </button>
            ))}
          </div>

          <form onSubmit={handleSubmit} className="space-y-6">
            {/* STEP 1: IDENTITAS DIRI */}
            {activeStep === 'identitas' && (
              <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="space-y-5">
                <div className="p-4 bg-emerald-50/60 rounded-2xl border border-emerald-100/80 mb-4">
                  <h3 className="font-black text-emerald-900 text-sm mb-0.5">Data Pribadi Jamaah</h3>
                  <p className="text-xs text-emerald-700 font-medium">Mohon isi identitas diri sesuai data KTP / Kartu Keluarga</p>
                </div>

                {/* Upload Foto Profil */}
                <div className="flex items-center gap-5 p-4 bg-slate-50 rounded-2xl border border-slate-200/80">
                  <div className="w-20 h-20 rounded-2xl bg-white border-2 border-dashed border-slate-300 flex items-center justify-center overflow-hidden relative shadow-sm">
                    {base64Photo ? (
                      <img src={base64Photo} alt="Foto Profil" className="w-full h-full object-cover" />
                    ) : (
                      <Users className="w-8 h-8 text-slate-300" />
                    )}
                  </div>
                  <div className="flex-1">
                    <label className="block text-xs font-bold text-slate-700 mb-1">Pas Foto (Opsional)</label>
                    <label className="inline-flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-xl text-xs font-bold hover:bg-emerald-700 transition-all cursor-pointer shadow-sm">
                      <ImageIcon className="w-4 h-4" />
                      <span>{base64Photo ? 'Ganti Foto' : 'Pilih Foto'}</span>
                      <input type="file" accept="image/*" onChange={handlePhotoUpload} className="hidden" />
                    </label>
                    {base64Photo && (
                      <button
                        type="button"
                        onClick={() => setBase64Photo(null)}
                        className="ml-2 text-xs text-red-500 font-semibold hover:underline"
                      >
                        Hapus
                      </button>
                    )}
                    <p className="text-[11px] text-slate-400 mt-1">Format JPG, PNG maksimal 2MB</p>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="sm:col-span-2">
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      Nama Lengkap <span className="text-red-500">*</span>
                    </label>
                    <input
                      name="name"
                      required
                      placeholder="Contoh: Muhammad Fulan"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Nama Panggilan / Akrab</label>
                    <input
                      name="nickname"
                      placeholder="Contoh: Fulan"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      No. Telepon / WhatsApp <span className="text-red-500">*</span>
                    </label>
                    <input
                      name="phone"
                      type="tel"
                      required
                      placeholder="Contoh: 081234567890"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Jenis Kelamin</label>
                    <select
                      name="gender"
                      defaultValue="Laki-laki"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium bg-white"
                    >
                      <option value="Laki-laki">Laki-laki</option>
                      <option value="Perempuan">Perempuan</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Golongan Darah</label>
                    <select
                      name="bloodType"
                      defaultValue="Belum Tahu"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium bg-white"
                    >
                      <option value="Belum Tahu">Belum Tahu</option>
                      <option value="A">A</option>
                      <option value="B">B</option>
                      <option value="AB">AB</option>
                      <option value="O">O</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Tempat Lahir</label>
                    <input
                      name="placeOfBirth"
                      placeholder="Kota Tempat Lahir"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Tanggal Lahir</label>
                    <input
                      name="dateOfBirth"
                      type="date"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium bg-white"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Pekerjaan / Profesi</label>
                    <input
                      name="currentJob"
                      placeholder="Contoh: Karyawan Swasta, Wiraswasta, Pelajar"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Pendidikan Terakhir</label>
                    <select
                      name="lastEducation"
                      defaultValue="SMA/SMK"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium bg-white"
                    >
                      <option value="SD">SD / Sederajat</option>
                      <option value="SMP">SMP / MTs</option>
                      <option value="SMA/SMK">SMA / SMK / MA</option>
                      <option value="D3">Diploma (D3)</option>
                      <option value="S1">Sarjana (S1)</option>
                      <option value="S2">Magister (S2)</option>
                      <option value="S3">Doktor (S3)</option>
                      <option value="Lainnya">Lainnya</option>
                    </select>
                  </div>
                </div>

                <div className="flex justify-end pt-4">
                  <button
                    type="button"
                    onClick={() => setActiveStep('keluarga')}
                    className="px-6 py-3 bg-slate-900 text-white rounded-xl font-bold text-sm hover:bg-slate-800 transition-all flex items-center gap-2"
                  >
                    <span>Lanjut ke Data Keluarga</span>
                    <ChevronRight className="w-4 h-4" />
                  </button>
                </div>
              </motion.div>
            )}

            {/* STEP 2: KELUARGA & DOMISILI */}
            {activeStep === 'keluarga' && (
              <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="space-y-5">
                <div className="p-4 bg-blue-50/60 rounded-2xl border border-blue-100/80 mb-4">
                  <h3 className="font-black text-blue-900 text-sm mb-0.5">Struktur Keluarga & Tempat Tinggal</h3>
                  <p className="text-xs text-blue-700 font-medium">Tentukan status Kepala Keluarga dan kelompok jamaah Anda</p>
                </div>

                {/* Lokasi Kelompok Masjid */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Lokasi Kelompok Masjid</label>
                  <select
                    name="location"
                    value={selectedLocation}
                    onChange={(e) => {
                      setSelectedLocation(e.target.value as MosqueLocation);
                      setSelectedKKId('');
                      setKkSearchTerm('');
                    }}
                    disabled={isLocationLocked}
                    className={cn(
                      "w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium bg-white",
                      isLocationLocked && "bg-slate-100 cursor-not-allowed opacity-80"
                    )}
                  >
                    {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(l => (
                      <option key={l} value={l}>{l}</option>
                    ))}
                  </select>
                  {isLocationLocked && (
                    <p className="text-[11px] text-slate-500 mt-1">
                      Lokasi kelompok telah ditentukan dari tautan pendaftaran.
                    </p>
                  )}
                </div>

                {/* Status KK Selection */}
                <div className="p-4 bg-slate-50 rounded-2xl border border-slate-200/80 space-y-3">
                  <label className="block text-xs font-bold text-slate-800">Apakah Anda Kepala Keluarga (KK)?</label>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <button
                      type="button"
                      onClick={() => {
                        setIsKK(true);
                        setSelectedKKId('');
                        setKkSearchTerm('');
                      }}
                      className={cn(
                        "p-4 rounded-xl border text-left transition-all",
                        isKK
                          ? "bg-emerald-50 border-emerald-500 text-emerald-900 ring-2 ring-emerald-500"
                          : "bg-white border-slate-200 text-slate-600 hover:bg-slate-100"
                      )}
                    >
                      <p className="font-black text-sm">👨‍👩‍👧‍👦 Ya, Kepala Keluarga</p>
                      <p className="text-[11px] text-slate-500 mt-0.5">Akan dibuatkan nomor induk KK baru</p>
                    </button>

                    <button
                      type="button"
                      onClick={() => setIsKK(false)}
                      className={cn(
                        "p-4 rounded-xl border text-left transition-all",
                        !isKK
                          ? "bg-emerald-50 border-emerald-500 text-emerald-900 ring-2 ring-emerald-500"
                          : "bg-white border-slate-200 text-slate-600 hover:bg-slate-100"
                      )}
                    >
                      <p className="font-black text-sm">👤 Bukan (Anggota Keluarga)</p>
                      <p className="text-[11px] text-slate-500 mt-0.5">Istri, Anak, atau Famili dari KK yang ada</p>
                    </button>
                  </div>
                </div>

                {/* Searchable KK Selector if NOT KK */}
                {!isKK && (
                  <div className="p-4 bg-amber-50/70 border border-amber-200 rounded-2xl space-y-2">
                    <label className="block text-xs font-bold text-amber-900">
                      Pilih Kepala Keluarga Anda ({selectedLocation}) <span className="text-red-500">*</span>
                    </label>

                    <div className="relative">
                      <div className="relative">
                        <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                        <input
                          type="text"
                          value={kkSearchTerm}
                          onChange={(e) => {
                            setKkSearchTerm(e.target.value);
                            setIsKkDropdownOpen(true);
                          }}
                          onFocus={() => setIsKkDropdownOpen(true)}
                          placeholder="Cari Kepala Keluarga (Ketik Nama / ID KK / HP)..."
                          className="w-full pl-9 pr-4 py-2.5 rounded-xl border border-slate-200 bg-white text-xs font-medium focus:ring-2 focus:ring-emerald-500 outline-none"
                        />
                      </div>

                      {/* Dropdown Results */}
                      {isKkDropdownOpen && (
                        <div className="absolute z-20 top-full left-0 right-0 mt-1 max-h-52 overflow-y-auto bg-white border border-slate-200 rounded-xl shadow-xl divide-y divide-slate-100">
                          {filteredKKList.length === 0 ? (
                            <div className="p-3 text-center text-xs text-slate-400">
                              {headsOfFamily.length === 0 
                                ? `Belum ada Kepala Keluarga di kelompok ${selectedLocation}` 
                                : 'Tidak ada Kepala Keluarga yang cocok dengan pencarian'}
                            </div>
                          ) : (
                            filteredKKList.map(k => (
                              <button
                                key={k.id || k.memberId}
                                type="button"
                                onClick={() => {
                                  setSelectedKKId(k.memberId);
                                  setKkSearchTerm(`${k.memberId} - ${k.name} ${k.nickname ? `(${k.nickname})` : ''}`);
                                  setIsKkDropdownOpen(false);
                                }}
                                className={cn(
                                  "w-full px-3.5 py-2.5 text-left text-xs hover:bg-emerald-50 flex items-center justify-between transition-colors",
                                  selectedKKId === k.memberId && "bg-emerald-50 text-emerald-800 font-bold"
                                )}
                              >
                                <div>
                                  <span className="font-bold text-slate-900">{k.name}</span>
                                  {k.nickname && <span className="text-slate-500 ml-1">({k.nickname})</span>}
                                  <p className="text-[10px] text-slate-400">{k.phone || 'Tanpa No. HP'}</p>
                                </div>
                                <span className="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 font-mono text-[10px] font-bold">
                                  {k.memberId}
                                </span>
                              </button>
                            ))
                          )}
                        </div>
                      )}
                    </div>

                    {selectedKKId ? (
                      <p className="text-xs text-emerald-700 font-bold flex items-center gap-1.5 pt-1">
                        <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                        Terhubung dengan KK: <span className="font-mono">{selectedKKId}</span>
                      </p>
                    ) : (
                      <p className="text-[11px] text-amber-700">
                        * Wajib memilih salah satu Kepala Keluarga yang terdaftar.
                      </p>
                    )}
                  </div>
                )}

                {/* Status Pernikahan */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Status Pernikahan</label>
                    <select
                      name="maritalStatus"
                      defaultValue="Belum Menikah"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium bg-white"
                    >
                      <option value="Belum Menikah">Belum Menikah</option>
                      <option value="Menikah">Menikah</option>
                      <option value="Duda">Duda</option>
                      <option value="Janda">Janda</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Nama Suami / Istri (Jika Menikah)</label>
                    <input
                      name="spouseName"
                      placeholder="Nama Pasangan"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium"
                    />
                  </div>

                  <div className="sm:col-span-2">
                    <label className="block text-xs font-bold text-slate-700 mb-1">Alamat Domisili Saat Ini</label>
                    <textarea
                      name="currentAddress"
                      placeholder="Jalan, RT/RW, Kelurahan, Kecamatan, Kota"
                      rows={2}
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium"
                    />
                  </div>

                  <div className="sm:col-span-2">
                    <label className="block text-xs font-bold text-slate-700 mb-1">Alamat Asal / KTP</label>
                    <textarea
                      name="originAddress"
                      placeholder="Alamat asal daerah / sesuai KTP (jika berbeda dengan domisili)"
                      rows={2}
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Nama Ayah</label>
                    <input
                      name="fatherName"
                      placeholder="Nama Ayah Kandung"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Nama Ibu</label>
                    <input
                      name="motherName"
                      placeholder="Nama Ibu Kandung"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium"
                    />
                  </div>
                </div>

                <div className="flex justify-between pt-4">
                  <button
                    type="button"
                    onClick={() => setActiveStep('identitas')}
                    className="px-6 py-3 bg-slate-100 text-slate-700 rounded-xl font-bold text-sm hover:bg-slate-200 transition-all flex items-center gap-2"
                  >
                    <ChevronLeft className="w-4 h-4" />
                    <span>Kembali</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => setActiveStep('keilmuan')}
                    className="px-6 py-3 bg-slate-900 text-white rounded-xl font-bold text-sm hover:bg-slate-800 transition-all flex items-center gap-2"
                  >
                    <span>Lanjut ke Keilmuan & Ibadah</span>
                    <ChevronRight className="w-4 h-4" />
                  </button>
                </div>
              </motion.div>
            )}

            {/* STEP 3: KEILMUAN & IBADAH */}
            {activeStep === 'keilmuan' && (
              <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} className="space-y-5">
                <div className="p-4 bg-purple-50/60 rounded-2xl border border-purple-100/80 mb-4">
                  <h3 className="font-black text-purple-900 text-sm mb-0.5">Keilmuan, Haji & Catatan Kesehatan</h3>
                  <p className="text-xs text-purple-700 font-medium">Informasi pembinaan keagamaan dan riwayat haji</p>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Kategori Jamaah</label>
                    <select
                      name="category"
                      defaultValue="UMUM"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium bg-white"
                    >
                      <option value="UMUM">UMUM (Dewasa / Orang Tua)</option>
                      <option value="GPN">GPN (Generus Muda)</option>
                      <option value="APR">APR (Pra-Remaja / SMP)</option>
                      <option value="ACR">ACR (Cabe Rawit / SD)</option>
                      <option value="DUDA">DUDA</option>
                      <option value="JANDA">JANDA</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Status Mubaligh / Mubalighot</label>
                    <select
                      name="isMubaligh"
                      defaultValue="Bukan"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium bg-white"
                    >
                      <option value="Bukan">Bukan</option>
                      <option value="Mubaligh Kelompok">Mubaligh Kelompok</option>
                      <option value="Mubaligh Desa">Mubaligh Desa</option>
                      <option value="Mubaligh Daerah">Mubaligh Daerah</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Status Riwayat Haji / Umroh</label>
                    <select
                      name="hasHajj"
                      defaultValue="Belum"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium bg-white"
                    >
                      <option value="Belum">Belum Pernah</option>
                      <option value="Sudah">Sudah Pernah Haji</option>
                      <option value="Pernah Umroh">Pernah Umroh</option>
                      <option value="Sudah Daftar Porsi">Sudah Daftar Porsi Haji</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Punya Saham UB</label>
                    <select
                      name="hasUbShares"
                      defaultValue="Tidak"
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium bg-white"
                    >
                      <option value="Tidak">Tidak</option>
                      <option value="Ada">Ada</option>
                    </select>
                  </div>

                  <div className="sm:col-span-2">
                    <label className="block text-xs font-bold text-slate-700 mb-1">Riwayat Penyakit / Catatan Medis (Opsional)</label>
                    <textarea
                      name="medicalHistory"
                      placeholder="Informasi penyakit bawaan, alergi, atau catatan medis jika ada..."
                      rows={2}
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm font-medium"
                    />
                  </div>
                </div>

                {/* Final Submission Buttons */}
                <div className="flex flex-col sm:flex-row gap-3 pt-6 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setActiveStep('keluarga')}
                    disabled={isSaving}
                    className="px-6 py-3.5 bg-slate-100 text-slate-700 rounded-xl font-bold text-sm hover:bg-slate-200 transition-all disabled:opacity-50"
                  >
                    Kembali
                  </button>

                  <button
                    type="submit"
                    disabled={isSaving}
                    className="flex-1 py-3.5 bg-emerald-600 text-white rounded-xl font-bold text-sm hover:bg-emerald-700 transition-all shadow-xl shadow-emerald-200 disabled:opacity-60 disabled:cursor-not-allowed flex items-center justify-center gap-2 active:scale-98"
                  >
                    {isSaving ? (
                      <>
                        <Loader2 className="w-5 h-5 animate-spin" />
                        <span>Menyimpan Data Jamaah...</span>
                      </>
                    ) : (
                      <>
                        <Check className="w-5 h-5" />
                        <span>Kirim & Simpan Data Jamaah</span>
                      </>
                    )}
                  </button>
                </div>
              </motion.div>
            )}
          </form>
        </div>
      </div>
    </div>
  );
}

// --- Modal Generator Link Pendaftaran Mandiri (Admin / Pengurus) ---
function RegistrationLinkModal({
  isOpen,
  onClose,
  profile
}: {
  isOpen: boolean;
  onClose: () => void;
  profile: UserProfile;
}) {
  const { accessToken, spreadsheetId } = useFirebase();
  const { showToast } = useToast();
  const [activeTab, setActiveTab] = useState<'create' | 'history'>('create');
  
  const [targetLocation, setTargetLocation] = useState<MosqueLocation | 'Seluruh Lokasi'>(() => {
    return profile.location || 'Kramat Batu';
  });

  const [durationOption, setDurationOption] = useState<'6h' | '12h' | '24h' | '3d' | '7d' | '14d' | '30d' | 'custom'>('24h');
  const [customDateTime, setCustomDateTime] = useState(() => {
    const d = new Date(Date.now() + 24 * 3600 * 1000);
    return d.toISOString().slice(0, 16);
  });
  const [note, setNote] = useState('');
  
  const [generatedLink, setGeneratedLink] = useState<string | null>(null);
  const [generatedPayload, setGeneratedPayload] = useState<any | null>(null);
  const [copied, setCopied] = useState(false);

  // Saved Links History (from localStorage and query)
  const [historyLinks, setHistoryLinks] = useState<RegistrationLink[]>(() => {
    try {
      const saved = localStorage.getItem('app_registration_links');
      return saved ? JSON.parse(saved) : [];
    } catch (e) {
      return [];
    }
  });

  const calculateExpiresAt = useCallback((): number => {
    const now = Date.now();
    switch (durationOption) {
      case '6h': return now + 6 * 3600 * 1000;
      case '12h': return now + 12 * 3600 * 1000;
      case '24h': return now + 24 * 3600 * 1000;
      case '3d': return now + 3 * 86400 * 1000;
      case '7d': return now + 7 * 86400 * 1000;
      case '14d': return now + 14 * 86400 * 1000;
      case '30d': return now + 30 * 86400 * 1000;
      case 'custom': {
        const parsed = new Date(customDateTime).getTime();
        return isNaN(parsed) ? now + 24 * 3600 * 1000 : parsed;
      }
      default: return now + 24 * 3600 * 1000;
    }
  }, [durationOption, customDateTime]);

  const handleGenerate = () => {
    const expiresAt = calculateExpiresAt();
    const id = `reg_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    const payload = {
      id,
      location: targetLocation,
      expiresAt,
      createdAt: Date.now(),
      createdBy: profile.displayName || profile.role,
      note: note.trim()
    };

    const token = btoa(encodeURIComponent(JSON.stringify(payload)));
    const baseUrl = `${window.location.origin}${window.location.pathname}`;
    const fullUrl = `${baseUrl}?register_jamaah=true&token=${token}`;

    setGeneratedLink(fullUrl);
    setGeneratedPayload(payload);

    // Save to history
    const newRecord: RegistrationLink = {
      id,
      token,
      location: targetLocation,
      expiresAt,
      createdAt: Date.now(),
      createdBy: profile.displayName || profile.role,
      note: note.trim(),
      isActive: true
    };

    const updated = [newRecord, ...historyLinks.filter(l => l.id !== id)];
    setHistoryLinks(updated);
    try {
      localStorage.setItem('app_registration_links', JSON.stringify(updated));
    } catch (e) {}

    // Save asynchronously to spreadsheet registration_links
    saveData(null, accessToken, spreadsheetId, 'registration_links', newRecord, id).catch(() => {});
    showToast('Tautan pendaftaran berhasil dibuat!', 'success');
  };

  const handleCopy = (urlToCopy?: string) => {
    const url = urlToCopy || generatedLink;
    if (!url) return;
    navigator.clipboard.writeText(url);
    setCopied(true);
    showToast('Tautan pendaftaran berhasil disalin ke clipboard!', 'success');
    setTimeout(() => setCopied(false), 2500);
  };

  const handleShareWhatsApp = (url?: string, loc?: string, expTime?: number) => {
    const targetUrl = url || generatedLink;
    const targetLoc = loc || targetLocation;
    const targetExp = expTime || (generatedPayload ? generatedPayload.expiresAt : calculateExpiresAt());
    if (!targetUrl) return;

    const formattedExp = new Date(targetExp).toLocaleString('id-ID', {
      dateStyle: 'full',
      timeStyle: 'short'
    });

    const msg = `Assalamu'alaikum Wr. Wb. Bapak/Ibu/Saudara Jamaah sekalian,\n\nMohon bantuannya untuk mengisi formulir pendaftaran/pemutakhiran data jamaah ${targetLoc !== 'Seluruh Lokasi' ? `Kelompok ${targetLoc}` : 'Desa Gandaria'} melalui tautan resmi berikut:\n\n👉 ${targetUrl}\n\n⚠️ *Perhatian:* Tautan ini aktif sampai dengan *${formattedExp} WIB*.\nMohon segera mengisi sebelum batas waktu berakhir.\n\nAlhamdulillah Jazakumullohu Khoiro.`;
    window.open(`https://wa.me/?text=${encodeURIComponent(msg)}`, '_blank');
  };

  const handleDownloadQR = () => {
    const canvas = document.getElementById('reg-link-qr-canvas') as HTMLCanvasElement;
    if (!canvas) return;
    const url = canvas.toDataURL('image/png');
    const a = document.createElement('a');
    a.href = url;
    a.download = `QR_Pendaftaran_Jamaah_${String(targetLocation).replace(/\s+/g, '_')}.png`;
    a.click();
  };

  const handlePrintPoster = () => {
    const canvas = document.getElementById('reg-link-qr-canvas') as HTMLCanvasElement;
    if (!canvas) return;
    const qrDataUrl = canvas.toDataURL('image/png');
    const expText = generatedPayload?.expiresAt 
      ? new Date(generatedPayload.expiresAt).toLocaleString('id-ID', { dateStyle: 'full', timeStyle: 'short' })
      : '';

    const printWin = window.open('', '_blank');
    if (!printWin) {
      showToast('Popup cetak diblokir oleh browser. Izinkan popup.', 'error');
      return;
    }

    printWin.document.write(`
      <!DOCTYPE html>
      <html>
        <head>
          <title>Poster Pendaftaran Jamaah - ${targetLocation}</title>
          <style>
            @page { size: A4 portrait; margin: 15mm; }
            body {
              font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
              color: #0f172a;
              text-align: center;
              padding: 20px;
              margin: 0;
            }
            .card {
              border: 3px solid #10b981;
              border-radius: 24px;
              padding: 30px;
              max-width: 600px;
              margin: 0 auto;
            }
            h1 { font-size: 26px; margin: 0 0 8px; color: #047857; text-transform: uppercase; }
            h2 { font-size: 18px; margin: 0 0 16px; color: #334155; }
            .location-badge {
              display: inline-block;
              background-color: #ecfdf5;
              color: #065f46;
              padding: 8px 18px;
              border-radius: 9999px;
              font-weight: bold;
              font-size: 16px;
              margin-bottom: 24px;
              border: 1px solid #a7f3d0;
            }
            .qr-wrapper {
              background: #fff;
              padding: 16px;
              border-radius: 16px;
              display: inline-block;
              box-shadow: 0 4px 12px rgba(0,0,0,0.08);
              border: 1px solid #e2e8f0;
              margin-bottom: 20px;
            }
            .qr-img { width: 280px; height: 280px; display: block; }
            .instructions { font-size: 15px; color: #475569; line-height: 1.5; margin: 16px 0; }
            .expiry-box {
              background: #fffbeb;
              border: 1px solid #fef3c7;
              color: #92400e;
              padding: 10px 16px;
              border-radius: 12px;
              font-size: 13px;
              font-weight: bold;
              margin-top: 16px;
            }
          </style>
        </head>
        <body>
          <div class="card">
            <h1>Formulir Pendaftaran Jamaah Mandiri</h1>
            <h2>Sistem Informasi & Manajemen Desa Gandaria</h2>
            <div class="location-badge">Kelompok: ${targetLocation}</div>
            <div class="qr-wrapper">
              <img class="qr-img" src="${qrDataUrl}" alt="QR Code" />
            </div>
            <p class="instructions">
              <strong>Scan QR Code di atas menggunakan Kamera HP / WhatsApp</strong><br/>
              untuk mengisi data jamaah secara mandiri tanpa perlu login.
            </p>
            ${expText ? `<div class="expiry-box">⏳ Berlaku hingga: ${expText} WIB</div>` : ''}
          </div>
          <script>
            window.onload = function() { window.print(); };
          </script>
        </body>
      </html>
    `);
    printWin.document.close();
  };

  const handleDeleteHistory = (id: string) => {
    const updated = historyLinks.filter(l => l.id !== id);
    setHistoryLinks(updated);
    try {
      localStorage.setItem('app_registration_links', JSON.stringify(updated));
    } catch (e) {}
    deleteData(null, accessToken, spreadsheetId, 'registration_links', id).catch(() => {});
    showToast('Tautan berhasil dihapus dari riwayat.', 'info');
  };

  if (!isOpen) return null;

  return (
    <AnimatePresence>
      <motion.div 
        initial={{ opacity: 0 }} 
        animate={{ opacity: 1 }} 
        exit={{ opacity: 0 }} 
        className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50 overflow-y-auto"
      >
        <motion.div 
          initial={{ scale: 0.95, y: 20 }} 
          animate={{ scale: 1, y: 0 }} 
          exit={{ scale: 0.95, y: 20 }} 
          className="bg-white rounded-[2.5rem] p-6 sm:p-8 w-full max-w-2xl shadow-2xl max-h-[92vh] flex flex-col overflow-hidden"
        >
          {/* Header */}
          <div className="flex justify-between items-start mb-6 border-b border-slate-100 pb-4">
            <div className="flex items-center gap-3">
              <div className="w-12 h-12 rounded-2xl bg-emerald-100 text-emerald-700 flex items-center justify-center shadow-inner">
                <Share2 className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-xl font-black text-slate-900">Buat Tautan Pendaftaran Jamaah</h3>
                <p className="text-xs text-slate-500">Generate link pengisian mandiri dengan batas waktu (expired date)</p>
              </div>
            </div>
            <button 
              type="button" 
              onClick={onClose} 
              className="p-2 text-slate-400 hover:text-slate-600 rounded-xl hover:bg-slate-100"
            >
              <X className="w-5 h-5" />
            </button>
          </div>

          {/* Sub Navigation */}
          <div className="flex gap-2 mb-6 border-b border-slate-100 pb-2">
            <button
              type="button"
              onClick={() => setActiveTab('create')}
              className={cn(
                "px-4 py-2 rounded-xl text-xs font-bold transition-all",
                activeTab === 'create'
                  ? "bg-emerald-600 text-white shadow-sm"
                  : "text-slate-500 hover:bg-slate-100"
              )}
            >
              Buat Tautan Baru
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('history')}
              className={cn(
                "px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5",
                activeTab === 'history'
                  ? "bg-emerald-600 text-white shadow-sm"
                  : "text-slate-500 hover:bg-slate-100"
              )}
            >
              <span>Riwayat Tautan</span>
              {historyLinks.length > 0 && (
                <span className="px-1.5 py-0.2 bg-emerald-100 text-emerald-800 rounded-full text-[10px] font-black">
                  {historyLinks.length}
                </span>
              )}
            </button>
          </div>

          {/* Body content */}
          <div className="flex-1 overflow-y-auto pr-1 space-y-6">
            {activeTab === 'create' ? (
              <>
                {/* Form Controls */}
                <div className="space-y-4">
                  {/* Location Selector */}
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1.5">Lokasi Target Jamaah</label>
                    <select
                      value={targetLocation}
                      onChange={(e) => {
                        setTargetLocation(e.target.value as any);
                        setGeneratedLink(null);
                      }}
                      disabled={profile.role === 'pengurus' && Boolean(profile.location)}
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-xs font-bold bg-white focus:ring-2 focus:ring-emerald-500 outline-none"
                    >
                      {profile.role === 'admin' && <option value="Seluruh Lokasi">Seluruh Lokasi (Jamaah Bebas Memilih)</option>}
                      {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(loc => (
                        <option key={loc} value={loc}>{loc}</option>
                      ))}
                    </select>
                  </div>

                  {/* Expiry Presets */}
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1.5">Masa Berlaku Tautan (Expired Date)</label>
                    <div className="grid grid-cols-4 sm:grid-cols-7 gap-2 mb-3">
                      {[
                        { id: '6h', label: '6 Jam' },
                        { id: '12h', label: '12 Jam' },
                        { id: '24h', label: '1 Hari' },
                        { id: '3d', label: '3 Hari' },
                        { id: '7d', label: '7 Hari' },
                        { id: '30d', label: '30 Hari' },
                        { id: 'custom', label: 'Kustom' }
                      ].map(preset => (
                        <button
                          key={preset.id}
                          type="button"
                          onClick={() => {
                            setDurationOption(preset.id as any);
                            setGeneratedLink(null);
                          }}
                          className={cn(
                            "py-2 px-1 text-center rounded-xl text-xs font-bold transition-all border",
                            durationOption === preset.id
                              ? "bg-emerald-50 border-emerald-500 text-emerald-800 ring-1 ring-emerald-500"
                              : "bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100"
                          )}
                        >
                          {preset.label}
                        </button>
                      ))}
                    </div>

                    {durationOption === 'custom' && (
                      <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-1 mb-2">
                        <label className="block text-[11px] font-bold text-slate-600">Pilih Tanggal & Jam Kedaluwarsa:</label>
                        <input
                          type="datetime-local"
                          value={customDateTime}
                          onChange={(e) => {
                            setCustomDateTime(e.target.value);
                            setGeneratedLink(null);
                          }}
                          className="w-full px-3 py-2 rounded-lg border border-slate-300 text-xs font-mono outline-none focus:ring-2 focus:ring-emerald-500 bg-white"
                        />
                      </div>
                    )}

                    {/* Expiry Preview Banner */}
                    <div className="p-3 bg-emerald-50/70 rounded-xl border border-emerald-100 text-xs text-emerald-800 flex items-center gap-2">
                      <Clock className="w-4 h-4 text-emerald-600 flex-shrink-0" />
                      <span>
                        Tautan akan aktif sampai: <strong>{new Date(calculateExpiresAt()).toLocaleString('id-ID', { dateStyle: 'full', timeStyle: 'short' })} WIB</strong>
                      </span>
                    </div>
                  </div>

                  {/* Note / Keperluan */}
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Catatan / Keperluan (Opsional)</label>
                    <input
                      type="text"
                      value={note}
                      onChange={(e) => setNote(e.target.value)}
                      placeholder="Contoh: Pendataan Santri Baru, Sensus Ramadan 2026..."
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-xs font-medium focus:ring-2 focus:ring-emerald-500 outline-none"
                    />
                  </div>

                  {/* Generate Button */}
                  <button
                    type="button"
                    onClick={handleGenerate}
                    className="w-full py-3.5 bg-emerald-600 text-white rounded-2xl font-black text-xs uppercase tracking-wider hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200 flex items-center justify-center gap-2 active:scale-98"
                  >
                    <Sparkles className="w-4 h-4" />
                    <span>{generatedLink ? 'Perbarui Tautan & QR Code' : 'Generate Link & QR Code'}</span>
                  </button>
                </div>

                {/* Generated Link Result Section */}
                {generatedLink && (
                  <motion.div
                    initial={{ opacity: 0, y: 15 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="p-5 bg-gradient-to-br from-slate-900 to-slate-950 text-white rounded-3xl space-y-4 shadow-xl border border-slate-800"
                  >
                    <div className="flex items-center justify-between">
                      <span className="text-[10px] font-black uppercase tracking-widest text-emerald-400 flex items-center gap-1.5">
                        <CheckCheck className="w-3.5 h-3.5" /> Tautan Pendaftaran Siap Dibagikan
                      </span>
                      <span className="text-[11px] text-slate-400 font-mono">
                        Kelompok: {targetLocation}
                      </span>
                    </div>

                    {/* URL Box & Copy */}
                    <div className="flex gap-2">
                      <input
                        type="text"
                        readOnly
                        value={generatedLink}
                        className="flex-1 px-3.5 py-2.5 rounded-xl bg-slate-800/90 border border-slate-700 text-slate-200 text-xs font-mono outline-none select-all"
                      />
                      <button
                        type="button"
                        onClick={() => handleCopy()}
                        className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-xl transition-all shadow-sm flex items-center gap-1.5 flex-shrink-0"
                      >
                        {copied ? (
                          <>
                            <Check className="w-4 h-4" />
                            <span>Tersalin!</span>
                          </>
                        ) : (
                          <>
                            <Copy className="w-4 h-4" />
                            <span>Salin</span>
                          </>
                        )}
                      </button>
                    </div>

                    {/* Quick Action Share Buttons */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                      <button
                        type="button"
                        onClick={() => handleShareWhatsApp()}
                        className="py-3 px-4 bg-[#25D366] hover:bg-[#1EBE5D] text-white rounded-xl text-xs font-bold transition-all flex items-center justify-center gap-2 shadow-md shadow-emerald-950/40"
                      >
                        <Send className="w-4 h-4" />
                        <span>Bagikan ke WhatsApp</span>
                      </button>

                      <a
                        href={generatedLink}
                        target="_blank"
                        rel="noreferrer"
                        className="py-3 px-4 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-xs font-bold transition-all flex items-center justify-center gap-2 border border-slate-700"
                      >
                        <ExternalLink className="w-4 h-4" />
                        <span>Buka Formulir</span>
                      </a>
                    </div>

                    {/* QR Code Container */}
                    <div className="pt-3 border-t border-slate-800 flex flex-col sm:flex-row items-center gap-4">
                      <div className="p-3 bg-white rounded-2xl shadow-inner flex items-center justify-center">
                        <QRCodeCanvas
                          id="reg-link-qr-canvas"
                          value={generatedLink}
                          size={130}
                          level="M"
                          includeMargin={true}
                        />
                      </div>
                      <div className="flex-1 space-y-2 text-center sm:text-left">
                        <p className="text-xs font-bold text-white">QR Code Pendaftaran</p>
                        <p className="text-[11px] text-slate-400">
                          Jamaah dapat memindai QR code ini di masjid atau selebaran untuk langsung mengisi form.
                        </p>
                        <div className="flex flex-wrap gap-2 justify-center sm:justify-start pt-1">
                          <button
                            type="button"
                            onClick={handleDownloadQR}
                            className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 border border-slate-700"
                          >
                            <Download className="w-3.5 h-3.5" />
                            <span>Unduh PNG</span>
                          </button>
                          <button
                            type="button"
                            onClick={handlePrintPoster}
                            className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 border border-slate-700"
                          >
                            <Printer className="w-3.5 h-3.5" />
                            <span>Cetak Poster</span>
                          </button>
                        </div>
                      </div>
                    </div>
                  </motion.div>
                )}
              </>
            ) : (
              /* Riwayat Tautan Tab */
              <div className="space-y-3">
                {historyLinks.length === 0 ? (
                  <div className="p-8 text-center bg-slate-50 rounded-2xl border border-slate-100 text-slate-400 text-xs">
                    Belum ada riwayat tautan pendaftaran yang dibuat.
                  </div>
                ) : (
                  historyLinks.map((item) => {
                    const isItemExpired = Date.now() > item.expiresAt;
                    const fullUrl = `${window.location.origin}${window.location.pathname}?register_jamaah=true&token=${item.token}`;

                    return (
                      <div
                        key={item.id}
                        className="p-4 bg-slate-50 hover:bg-slate-100/80 rounded-2xl border border-slate-200/80 transition-all space-y-2.5 text-xs"
                      >
                        <div className="flex items-center justify-between">
                          <div className="flex items-center gap-2">
                            <span className="font-bold text-slate-900">{item.location || 'Seluruh Lokasi'}</span>
                            <span
                              className={cn(
                                "px-2 py-0.5 rounded-full text-[10px] font-black uppercase",
                                isItemExpired
                                  ? "bg-slate-200 text-slate-600"
                                  : "bg-emerald-100 text-emerald-800"
                              )}
                            >
                              {isItemExpired ? 'Kedaluwarsa' : 'Aktif'}
                            </span>
                          </div>
                          <span className="text-[10px] text-slate-400">
                            Dibuat: {new Date(item.createdAt).toLocaleDateString('id-ID')}
                          </span>
                        </div>

                        {item.note && (
                          <p className="text-slate-600 text-[11px] italic">"{item.note}"</p>
                        )}

                        <div className="flex justify-between items-center text-[11px] text-slate-500 pt-1 border-t border-slate-200/60">
                          <span>
                            Berlaku hingga: <strong>{new Date(item.expiresAt).toLocaleString('id-ID', { dateStyle: 'short', timeStyle: 'short' })}</strong>
                          </span>
                          <div className="flex gap-2">
                            <button
                              type="button"
                              onClick={() => handleCopy(fullUrl)}
                              className="px-2.5 py-1 bg-white border border-slate-200 rounded-lg font-bold text-slate-700 hover:bg-slate-50 flex items-center gap-1"
                            >
                              <Copy className="w-3 h-3" />
                              <span>Salin</span>
                            </button>
                            <button
                              type="button"
                              onClick={() => handleShareWhatsApp(fullUrl, item.location, item.expiresAt)}
                              className="px-2.5 py-1 bg-[#25D366] text-white rounded-lg font-bold hover:bg-[#1EBE5D] flex items-center gap-1"
                            >
                              <Send className="w-3 h-3" />
                              <span>WA</span>
                            </button>
                            <button
                              type="button"
                              onClick={() => handleDeleteHistory(item.id)}
                              className="px-2 py-1 text-red-500 hover:bg-red-50 rounded-lg"
                              title="Hapus dari riwayat"
                            >
                              <X className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>
                      </div>
                    );
                  })
                )}
              </div>
            )}
          </div>
        </motion.div>
      </motion.div>
    </AnimatePresence>
  );
}

function LoginView({ onAttendanceMode, onRegisterJamaahMode }: { onAttendanceMode: () => void, onRegisterJamaahMode?: () => void }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [isRegister, setIsRegister] = useState(false);
  const [displayName, setDisplayName] = useState('');
  const [role, setRole] = useState<UserRole>('pengurus');
  const [location, setLocation] = useState<MosqueLocation>('Kramat Batu');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const { auth, googleSignIn, accessToken, spreadsheetId, syncProfileFromSheet } = useFirebase();

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    if (!auth) {
      setError('Firebase belum diinisialisasi');
      setLoading(false);
      return;
    }

    try {
      if (isRegister) {
        if (!displayName.trim()) {
          setError('Silakan masukkan nama lengkap Anda.');
          setLoading(false);
          return;
        }

        const effectiveSpreadsheetId = spreadsheetId || (import.meta as any).env?.VITE_SPREADSHEET_ID || localStorage.getItem('app_spreadsheet_id') || '';
        const effectiveToken = accessToken || localStorage.getItem('app_access_token') || null;

        const userCredential = await createUserWithEmailAndPassword(auth, email.trim(), password);
        const newUid = userCredential.user.uid;

        const newUserProfile = {
          uid: newUid,
          id: newUid,
          email: email.trim(),
          displayName: displayName.trim(),
          role,
          location: role === 'pengurus' ? location : null,
          isVerified: false, // Akun baru memerlukan verifikasi dari admin!
          createdAt: Date.now(),
          spreadsheetId: effectiveSpreadsheetId || ''
        };

        // 1. Simpan ke local storage agar FirebaseProvider langsung menggunakannya
        localStorage.setItem(`user_profile_${newUid}`, JSON.stringify(newUserProfile));

        // 2. Simpan data lengkap ke Google Spreadsheet (sheet 'users')
        try {
          await saveData(null, effectiveToken, effectiveSpreadsheetId, 'users', newUserProfile, newUid);
        } catch (saveErr) {
          console.warn('[Register] Simpan ke spreadsheet:', saveErr);
        }

        // 3. Sinkronisasi profile
        if (syncProfileFromSheet) {
          await syncProfileFromSheet(newUid).catch(() => {});
        }
      } else {
        await signInWithEmailAndPassword(auth, email.trim(), password);
      }
    } catch (err: any) {
      const msg = err.message || 'Gagal masuk';
      if (msg.includes('email-already-in-use')) {
        setError('Email ini sudah terdaftar. Silakan masuk atau gunakan email lain.');
      } else if (msg.includes('weak-password')) {
        setError('Password terlalu pendek (minimal 6 karakter).');
      } else if (msg.includes('invalid-credential') || msg.includes('user-not-found') || msg.includes('wrong-password')) {
        setError('Email atau password tidak sesuai.');
      } else {
        setError(msg);
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-50 p-4">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        className="bg-white p-8 rounded-3xl shadow-xl w-full max-w-md border border-slate-100"
      >
        <div className="text-center mb-8">
          <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-2xl flex items-center justify-center mx-auto mb-4">
            <LayoutDashboard className="w-8 h-8" />
          </div>
          <h1 className="text-2xl font-bold text-slate-900">Sistem Manajemen Desa GND</h1>
          <p className="text-slate-500 mt-2">Silakan masuk ke akun Anda</p>
        </div>

        <button 
          type="button"
          onClick={googleSignIn}
          className="w-full py-4 bg-white border-2 border-slate-100 rounded-2xl flex items-center justify-center gap-4 hover:border-emerald-500 transition-all group mb-8 shadow-sm"
        >
          <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" className="w-6 h-6" alt="Google" />
          <span className="font-black uppercase tracking-widest text-slate-700 group-hover:text-emerald-600">Masuk dengan Google</span>
        </button>

        <div className="relative flex items-center gap-4 mb-8">
          <div className="flex-grow h-px bg-slate-100"></div>
          <span className="text-[10px] font-black uppercase tracking-widest text-slate-400">Atau Gunakan Email</span>
          <div className="flex-grow h-px bg-slate-100"></div>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4">
          {isRegister && (
            <>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Nama Lengkap</label>
                <input
                  type="text"
                  value={displayName}
                  onChange={(e) => setDisplayName(e.target.value)}
                  className="w-full px-4 py-2.5 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 transition-all outline-none"
                  placeholder="Nama Anda"
                  required
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Role</label>
                  <select
                    value={role}
                    onChange={(e) => setRole(e.target.value as UserRole)}
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none"
                  >
                    <option value="pengurus">Pengurus</option>
                    <option value="admin">Admin</option>
                  </select>
                </div>
                {role === 'pengurus' && (
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">Lokasi</label>
                    <select
                      value={location}
                      onChange={(e) => setLocation(e.target.value as MosqueLocation)}
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none"
                    >
                      <option value="Kramat Batu">Kramat Batu</option>
                      <option value="Karya Utama">Karya Utama</option>
                      <option value="Radio Dalam">Radio Dalam</option>
                      <option value="Cipete">Cipete</option>
                      <option value="Antena">Antena</option>
                    </select>
                  </div>
                )}
              </div>
            </>
          )}

          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Email</label>
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              className="w-full px-4 py-2.5 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 transition-all outline-none"
              placeholder="email@masjid.com"
              required
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Password</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full px-4 py-2.5 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 transition-all outline-none"
              placeholder="••••••••"
              required
            />
          </div>

          {error && <p className="text-red-500 text-sm">{error}</p>}

          <button
            type="submit"
            disabled={loading}
            className="w-full bg-emerald-600 text-white py-3 rounded-xl font-semibold hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200 disabled:opacity-50"
          >
            {loading ? 'Memproses...' : (isRegister ? 'Daftar' : 'Masuk')}
          </button>
        </form>

        <div className="mt-8 pt-6 border-t border-slate-100 space-y-4">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
            <button 
              type="button"
              onClick={onAttendanceMode}
              className="flex items-center justify-center gap-2 w-full py-3.5 rounded-2xl bg-slate-900 text-white font-black uppercase tracking-wider text-[11px] hover:bg-slate-800 transition-all active:scale-95 shadow-md shadow-slate-200"
            >
              <Clock className="w-4 h-4 text-emerald-400" />
              <span>Mode Absensi</span>
            </button>
            {onRegisterJamaahMode && (
              <button 
                type="button"
                onClick={onRegisterJamaahMode}
                className="flex items-center justify-center gap-2 w-full py-3.5 rounded-2xl bg-emerald-50 text-emerald-800 border border-emerald-200 font-black uppercase tracking-wider text-[11px] hover:bg-emerald-100 transition-all active:scale-95 shadow-sm"
              >
                <UserPlus className="w-4 h-4 text-emerald-600" />
                <span>Daftar Mandiri</span>
              </button>
            )}
          </div>
          
          <button
            type="button"
            onClick={() => setIsRegister(!isRegister)}
            className="w-full py-2 text-sm font-bold text-slate-500 hover:text-emerald-600 transition-colors"
          >
            {isRegister ? 'Sudah punya akun? Masuk' : 'Belum punya akun? Daftar Akun Pengurus'}
          </button>
        </div>
      </motion.div>
    </div>
  );
}

// --- Dashboard Sub-Views ---

export function parseImageUrls(input: any): string[] {
  if (!input) return [];
  if (Array.isArray(input)) {
    return input.filter((it): it is string => typeof it === 'string' && it.trim().length > 0);
  }
  if (typeof input === 'string') {
    const trimmed = input.trim();
    if (!trimmed) return [];
    if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
      try {
        const parsed = JSON.parse(trimmed);
        if (Array.isArray(parsed)) {
          return parsed.filter((it): it is string => typeof it === 'string' && it.trim().length > 0);
        }
      } catch (e) {}
    }
    if (trimmed.includes(',')) {
      return trimmed.split(',').map((s) => s.trim()).filter((s) => s.length > 0);
    }
    return [trimmed];
  }
  return [];
}

function ActivityImageSlider({ images }: { images?: any }) {
  const imageList = useMemo(() => parseImageUrls(images), [images]);
  const [index, setIndex] = useState(0);

  useEffect(() => {
    setIndex(0);
  }, [imageList.length]);

  useEffect(() => {
    if (imageList.length <= 1) return;
    const interval = setInterval(() => {
      setIndex((prev) => (prev + 1) % imageList.length);
    }, 4000);
    return () => clearInterval(interval);
  }, [imageList.length]);

  if (imageList.length === 0) {
    return (
      <div className="w-full h-full flex items-center justify-center text-slate-400 bg-slate-100">
        <Calendar className="w-12 h-12" />
      </div>
    );
  }

  const activeSrc = imageList[index] || imageList[0];

  return (
    <div className="relative w-full h-full overflow-hidden group">
      <AnimatePresence initial={false} mode="wait">
        <motion.img
          key={`${activeSrc}-${index}`}
          src={activeSrc}
          initial={{ opacity: 0, scale: 1.1 }}
          animate={{ opacity: 1, scale: 1 }}
          exit={{ opacity: 0 }}
          transition={{ duration: 1 }}
          className="absolute inset-0 w-full h-full object-cover"
        />
      </AnimatePresence>
      {imageList.length > 1 && (
        <div className="absolute bottom-3 left-1/2 -translate-x-1/2 flex gap-1.5 z-10">
          {imageList.map((_, i) => (
            <div 
              key={i} 
              className={cn(
                "w-1.5 h-1.5 rounded-full transition-all duration-300",
                i === index ? "bg-white w-4" : "bg-white/40"
              )}
            />
          ))}
        </div>
      )}
    </div>
  );
}

function Overview({ profile }: { profile: UserProfile }) {
  const { db } = useFirebase();
  const filter = useMemo(() => (profile.role === 'pengurus' && profile.location && (profile.location as string) !== 'Seluruh Lokasi') ? [where('location', '==', profile.location)] : [], [profile.role, profile.location]);
  const { data: jamaah } = useDataQuery<Jamaah>('jamaah', filter);
  const { data: allAssets } = useDataQuery<Asset>('assets', filter);
  const { data: activities } = useDataQuery<Activity>('activities', filter);
  const { data: stats } = useDataQuery<FacilityStat>('facility_stats', []);

  const barangAssets = useMemo(() => allAssets.filter(a => !a.assetType || a.assetType === 'barang'), [allAssets]);
  const tanahAssets = useMemo(() => allAssets.filter(a => a.assetType === 'tanah'), [allAssets]);
  const totalTanahArea = useMemo(() => tanahAssets.reduce((sum, t) => sum + (t.areaSize || 0), 0), [tanahAssets]);
  const totalKK = useMemo(() => jamaah.filter(j => j.isKK).length, [jamaah]);

  const jamaahCategories = useMemo(() => {
    const counts: Record<string, number> = { 'UMUM': 0, 'GPN': 0, 'APR': 0, 'ACR': 0, 'DUDA': 0, 'JANDA': 0 };
    jamaah.forEach(j => {
      const cat = (j.category || 'UMUM').toUpperCase();
      if (counts[cat] !== undefined) {
        counts[cat]++;
      } else {
        counts[cat] = (counts[cat] || 0) + 1;
      }
    });
    return Object.entries(counts)
      .filter(([_, value]) => value > 0 || ['UMUM', 'GPN', 'APR', 'ACR'].includes(_))
      .map(([name, value]) => ({ name, value }));
  }, [jamaah]);

  const COLORS = ['#10b981', '#3b82f6', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899'];

  return (
    <div className="space-y-8">
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 xl:grid-cols-5 gap-6">
        <StatCard title={`Total Jamaah - ${profile.location || 'Desa GND'}`} value={jamaah.length} icon={Users} color="bg-emerald-500" />
        <StatCard title={`Total KK - ${profile.location || 'Desa GND'}`} value={totalKK} icon={Home} color="bg-orange-500" />
        <StatCard title={`Total Luas Tanah - ${profile.location || 'Desa GND'}`} value={`${totalTanahArea} m²`} icon={MapIcon} color="bg-blue-500" />
        {stats.slice(0, 2).map((stat, i) => (
          <StatCard 
            key={stat.id ? `${stat.id}-${i}` : `overview-stat-${i}`} 
            title={stat.name} 
            value={stat.currentUsage} 
            total={stat.capacity} 
            icon={Package} 
            color={COLORS[i % COLORS.length]} 
          />
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
          <h3 className="text-lg font-bold text-slate-900 mb-6">Pertumbuhan Jamaah</h3>
          <div className="h-64">
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={jamaah.sort((a,b) => a.registeredAt - b.registeredAt).map(j => ({ date: new Date(j.registeredAt).toLocaleDateString(), count: 1 }))}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="date" hide />
                <YAxis hide />
                <Tooltip />
                <Line type="monotone" dataKey="count" stroke="#10b981" strokeWidth={3} dot={false} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
          <h3 className="text-lg font-bold text-slate-900 mb-6">Status Inventaris</h3>
          <div className="h-64">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={[
                { name: 'Baik', count: barangAssets.filter(a => a.status === 'baik').length },
                { name: 'Rusak', count: barangAssets.filter(a => a.status === 'rusak').length },
                { name: 'Perbaikan', count: barangAssets.filter(a => a.status === 'perlu perbaikan').length }
              ]}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="name" fontSize={12} />
                <YAxis fontSize={12} />
                <Tooltip />
                <Bar dataKey="count" radius={[4, 4, 0, 0]}>
                  { [0,1,2].map((entry, index) => <Cell key={`cell-${index}`} fill={COLORS[index + 1]} />) }
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
          <h3 className="text-lg font-bold text-slate-900 mb-6">Distribusi Tanah Sabilillah</h3>
          <div className="h-64">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={
                profile.role === 'admin' 
                ? ['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(loc => ({
                    name: loc,
                    value: tanahAssets.filter(t => t.location === loc).reduce((sum, t) => sum + (t.areaSize || 0), 0)
                  }))
                : [
                  { name: 'Wakaf', value: tanahAssets.filter(t => t.status === 'wakaf').length },
                  { name: 'Sertifikasi', value: tanahAssets.filter(t => t.status === 'sertifikasi').length },
                  { name: 'Terjual', value: tanahAssets.filter(t => t.status === 'terjual').length }
                ]
              }>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="name" fontSize={10} />
                <YAxis fontSize={12} />
                <Tooltip />
                <Bar dataKey="value" radius={[4, 4, 0, 0]}>
                  { [0,1,2,3,4].map((entry, index) => <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />) }
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
          <h3 className="text-lg font-bold text-slate-900 mb-6">Kategori Jamaah</h3>
          <div className="h-64">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={jamaahCategories}
                  cx="50%"
                  cy="50%"
                  innerRadius={60}
                  outerRadius={80}
                  paddingAngle={5}
                  dataKey="value"
                >
                  {jamaahCategories.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                  ))}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="space-y-6">
        <div className="flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-900">Kegiatan Terbaru</h3>
          <div className="flex items-center gap-2 text-sm text-slate-400 font-medium bg-slate-50 px-3 py-1 rounded-full">
             <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
             Live Update
          </div>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
          {activities.sort((a,b) => b.date - a.date).slice(0, 3).map((activity, actIdx) => (
            <motion.div 
              key={activity.id ? `${activity.id}-${actIdx}` : `overview-act-${actIdx}`}
              whileHover={{ y: -4 }}
              className="bg-white rounded-[2rem] overflow-hidden border border-slate-100 shadow-sm hover:shadow-xl hover:shadow-slate-100 transition-all group"
            >
              <div className="h-56 relative overflow-hidden">
                <ActivityImageSlider images={activity.imageUrls || []} />
                <div className="absolute top-4 left-4 z-10">
                  <span className={cn(
                    "px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider shadow-lg backdrop-blur-md",
                    activity.type === 'harian' ? "bg-emerald-500/80 text-white" : "bg-blue-500/80 text-white"
                  )}>
                    {activity.type}
                  </span>
                </div>
              </div>
              <div className="p-6">
                <h4 className="font-bold text-slate-900 mb-2 line-clamp-1 group-hover:text-emerald-600 transition-colors">{activity.title}</h4>
                <p className="text-slate-500 text-sm line-clamp-2 mb-6 h-10 leading-relaxed">{activity.description}</p>
                <div className="flex items-center justify-between border-t border-slate-50 pt-4">
                  <div className="flex items-center text-[10px] text-slate-400 gap-4">
                    <span className="flex items-center gap-1.5"><Clock className="w-3.5 h-3.5" /> {new Date(activity.date).toLocaleDateString()}</span>
                    <span className="flex items-center gap-1.5"><MapPin className="w-3.5 h-3.5 text-emerald-500" /> {activity.location}</span>
                  </div>
                  <div className="w-8 h-8 rounded-xl bg-slate-50 flex items-center justify-center group-hover:bg-emerald-50 group-hover:text-emerald-600 transition-all">
                    <Plus className="w-4 h-4" />
                  </div>
                </div>
              </div>
            </motion.div>
          ))}
          {activities.length === 0 && (
            <div className="col-span-full py-20 text-center bg-slate-50 rounded-[2rem] border-2 border-dashed border-slate-200">
               <Calendar className="w-12 h-12 text-slate-200 mx-auto mb-4" />
               <p className="text-slate-400 font-medium">Belum ada kegiatan yang diposting</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

const LOCATION_PREFIXES: Record<string, string> = {
  'Kramat Batu': 'KB',
  'Karya Utama': 'KU',
  'Radio Dalam': 'RD',
  'Cipete': 'CP',
  'Antena': 'AN'
};

const DAPUKAN_OPTIONS = [
  "Ides", "Wides", "Koor. Lupg", "Bosdes", "Mubdes", "Tim Aghniya'", "Ku Des", 
  "Tim Bk", "Tim Bacaan", "Tim Basyiron Wa Nadziron", "Tim Benda Sb", "Tim DhuaFa'", 
  "Tim Faraoid", "Tim Gambuh", "Tim Haji", "Tim Keluarga Bahagia", "Tim Kematian", 
  "Tim Manula", "Tim Mondar Mandir", "Tim Muballigh", "Tim Organisasi", 
  "Tim Pembangunan", "Tim Pkw", "Tim Penyelesaian", "Tim Pramuka", "Tim Sarjana", 
  "Tim Ub", "Tim Zakat", "Tim Cai & Remaja (Karemdes)", "Keputrian Des", 
  "Ikel", "Wikel", "Pjkbm", "Boskel", "Mubkel", "Ku Kel", "Pakar Pendidik", "Ptk", "Keputkel"
];

function JamaahView({ profile, formTrigger, onFormTriggered }: { profile: UserProfile, formTrigger?: string | null, onFormTriggered?: () => void }) {
  const { accessToken, spreadsheetId } = useFirebase();
  const { showToast } = useToast();
  const filter = useMemo(() => (profile.role === 'pengurus' && profile.location && (profile.location as string) !== 'Seluruh Lokasi') ? [where('location', '==', profile.location)] : [], [profile.role, profile.location]);
  const { data: jamaah, loading } = useDataQuery<Jamaah>('jamaah', filter);
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [editingJamaah, setEditingJamaah] = useState<Jamaah | null>(null);
  const [selectedJamaahDetail, setSelectedJamaahDetail] = useState<Jamaah | null>(null);
  const [base64Image, setBase64Image] = useState<string | null>(null);
  const [isKK, setIsKK] = useState(true);
  const [selectedKKId, setSelectedKKId] = useState<string>('');
  const [kkSearchTerm, setKkSearchTerm] = useState<string>('');
  const [isKkDropdownOpen, setIsKkDropdownOpen] = useState<boolean>(false);
  const [selectedDapukan, setSelectedDapukan] = useState<string[]>([]);
  const [dapukanFilter, setDapukanFilter] = useState('');
  const [hasHajjStatus, setHasHajjStatus] = useState<string>('');
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [showRegistrationLinkModal, setShowRegistrationLinkModal] = useState(false);
  const [formTab, setFormTab] = useState<'identitas' | 'keluarga' | 'pendidikan' | 'dapukan' | 'keilmuan'>('identitas');

  useEffect(() => {
    if (formTrigger === 'jamaah') {
      setShowForm(true);
      setEditingJamaah(null);
      setBase64Image(null);
      setIsKK(true);
      setSelectedKKId('');
      setKkSearchTerm('');
      setIsKkDropdownOpen(false);
      setSelectedDapukan([]);
      setHasHajjStatus('');
      setFormTab('identitas');
      onFormTriggered?.();
    }
  }, [formTrigger]);

  const handleImageChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setBase64Image(reader.result as string);
      };
      reader.readAsDataURL(file);
    }
  };

  const filteredJamaah = jamaah.filter(j => 
    j.name.toLowerCase().includes(searchTerm.toLowerCase()) || 
    j.nickname?.toLowerCase().includes(searchTerm.toLowerCase()) ||
    j.phone?.includes(searchTerm) ||
    j.memberId?.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);

  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm]);

  const totalPages = Math.ceil(filteredJamaah.length / pageSize) || 1;
  const paginatedJamaah = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredJamaah.slice(start, start + pageSize);
  }, [filteredJamaah, currentPage, pageSize]);

  const headsOfFamily = jamaah.filter(j => (j.isKK === true || (j.isKK as any) === 'true' || (j.isKK as any) === 'TRUE') && (profile.role === 'admin' || !profile.location || (profile.location as string) === 'Seluruh Lokasi' ? true : j.location === profile.location));

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);

    const nameVal = (formData.get('name') as string || '').trim();
    if (!nameVal) {
      setFormTab('identitas');
      showToast('Nama lengkap wajib diisi', 'error');
      return;
    }

    const phoneVal = (formData.get('phone') as string || '').trim();
    if (!phoneVal) {
      setFormTab('identitas');
      showToast('Nomor telepon/WA wajib diisi', 'error');
      return;
    }

    const location = profile.role === 'pengurus' ? (profile.location as MosqueLocation) : (formData.get('location') as MosqueLocation);
    
    let memberId = editingJamaah?.memberId || '';
    let familyOrder = editingJamaah?.familyOrder || 0;
    let finalKKId = editingJamaah?.kkId || '';

    const isStructureChanged = !editingJamaah || 
                                editingJamaah.isKK !== isKK || 
                                (!isKK && editingJamaah.kkId !== selectedKKId) ||
                                editingJamaah.location !== location;

    if (isStructureChanged) {
      const prefix = LOCATION_PREFIXES[location] || 'JM';
      if (isKK) {
        const kkInLocation = jamaah.filter(j => j.location === location && (j.isKK === true || (j.isKK as any) === 'true' || (j.isKK as any) === 'TRUE'));
        let maxNum = 0;
        kkInLocation.forEach(kk => {
          const numPart = kk.memberId.replace(prefix, '');
          const num = parseInt(numPart);
          if (!isNaN(num) && num > maxNum) maxNum = num;
        });
        const nextNum = maxNum + 1;
        memberId = `${prefix}${nextNum.toString().padStart(3, '0')}`;
        familyOrder = 0;
        finalKKId = memberId;
      } else {
        const kk = jamaah.find(j => j.memberId === selectedKKId);
        if (!kk) {
          showToast('Pilih Kepala Keluarga terlebih dahulu', 'error');
          return;
        }
        const familyMembers = jamaah.filter(j => j.kkId === selectedKKId);
        let maxOrder = 0;
        familyMembers.forEach(m => {
          if (m.familyOrder > maxOrder) maxOrder = m.familyOrder;
        });
        familyOrder = maxOrder + 1;
        memberId = `${selectedKKId}-${familyOrder.toString().padStart(2, '0')}`;
        finalKKId = selectedKKId;
      }
    }

    const originAddress = (formData.get('originAddress') as string) || '';
    const currentAddress = (formData.get('currentAddress') as string) || '';

    const data: Jamaah = {
      id: editingJamaah?.id || '',
      memberId,
      name: (formData.get('name') as string) || '',
      nickname: (formData.get('nickname') as string) || '',
      gender: (formData.get('gender') as string) || '',
      originAddress,
      currentAddress,
      address: currentAddress || originAddress,
      phone: (formData.get('phone') as string) || '',
      location,
      category: (formData.get('category') as JamaahCategory) || 'UMUM',
      isKK,
      kkId: finalKKId,
      familyOrder,
      dapukan: selectedDapukan,
      positions: selectedDapukan,
      photoUrl: base64Image || editingJamaah?.photoUrl || undefined,
      registeredAt: editingJamaah ? editingJamaah.registeredAt : Date.now(),
      placeOfBirth: (formData.get('placeOfBirth') as string) || '',
      dateOfBirth: (formData.get('dateOfBirth') as string) || '',
      fatherName: (formData.get('fatherName') as string) || '',
      motherName: (formData.get('motherName') as string) || '',
      parentPhone: (formData.get('parentPhone') as string) || '',
      lastEducation: (formData.get('lastEducation') as string) || '',
      majorOrClass: (formData.get('majorOrClass') as string) || '',
      schoolOrUniversity: (formData.get('schoolOrUniversity') as string) || '',
      currentJob: (formData.get('currentJob') as string) || '',
      workplaceAddress: (formData.get('workplaceAddress') as string) || '',
      maritalStatus: (formData.get('maritalStatus') as string) || '',
      marriageYear: (formData.get('marriageYear') as string) || '',
      spouseName: (formData.get('spouseName') as string) || '',
      hasJurusKeras: (formData.get('hasJurusKeras') as string) || '',
      hasJurusHalus: (formData.get('hasJurusHalus') as string) || '',
      bloodType: (formData.get('bloodType') as string) || '',
      hasUbShares: (formData.get('hasUbShares') as string) || '',
      previousDapukan: (formData.get('previousDapukan') as string) || '',
      isMubaligh: (formData.get('isMubaligh') as string) || '',
      hasHajj: (formData.get('hasHajj') as string) || '',
      hajjPortionNumber: (formData.get('hajjPortionNumber') as string) || '',
      plannedHajjYear: (formData.get('plannedHajjYear') as string) || '',
      hajjName: (formData.get('hajjName') as string) || '',
      hajjYear: (formData.get('hajjYear') as string) || '',
      medicalHistory: (formData.get('medicalHistory') as string) || ''
    };

    setIsSaving(true);
    try {
      await saveData(null, accessToken, spreadsheetId, 'jamaah', data, editingJamaah?.id);
      showToast(`Data jamaah ${data.name} berhasil disimpan!`, 'success');
      closeForm();
    } catch (err: any) {
      showToast(`Gagal menyimpan data: ${err.message}`, 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const closeForm = () => {
    setShowForm(false);
    setEditingJamaah(null);
    setBase64Image(null);
    setIsKK(true);
    setSelectedKKId('');
    setKkSearchTerm('');
    setIsKkDropdownOpen(false);
    setSelectedDapukan([]);
    setHasHajjStatus('');
    setFormTab('identitas');
  };

  const openEdit = (j: Jamaah) => {
    setEditingJamaah(j);
    setBase64Image(j.photoUrl || null);
    const kkStatus = j.isKK === true || (j.isKK as any) === 'true' || (j.isKK as any) === 'TRUE';
    setIsKK(kkStatus);
    setSelectedKKId(j.kkId || '');
    if (!kkStatus && j.kkId) {
      const match = headsOfFamily.find(k => k.memberId === j.kkId);
      setKkSearchTerm(match ? `${match.memberId} - ${match.name}` : j.kkId);
    } else {
      setKkSearchTerm('');
    }
    setIsKkDropdownOpen(false);
    const currentDapukan = Array.isArray(j.dapukan) && j.dapukan.length > 0 
      ? j.dapukan 
      : (Array.isArray(j.positions) ? j.positions : []);
    setSelectedDapukan(currentDapukan);
    setHasHajjStatus(j.hasHajj || '');
    setFormTab('identitas');
    setShowForm(true);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-slate-900">Data Jamaah</h2>
          <p className="text-slate-500">Kelola data jamaah {profile.role === 'pengurus' ? `di ${profile.location}` : 'seluruh lokasi'}</p>
        </div>
        <div className="flex flex-wrap items-center gap-2.5">
          <button 
            type="button"
            onClick={() => setShowRegistrationLinkModal(true)}
            className="bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 px-4 py-2.5 rounded-xl font-bold text-xs sm:text-sm flex items-center justify-center gap-2 transition-all shadow-sm active:scale-95"
            title="Generate Link Pengisian Jamaah Mandiri"
          >
            <Share2 className="w-4 h-4 text-emerald-600" />
            <span>Bagikan Link Pengisian</span>
          </button>
          <button 
            type="button"
            onClick={() => {
              closeForm();
              setShowForm(true);
            }}
            className="bg-emerald-600 text-white px-5 py-2.5 rounded-xl font-semibold flex items-center justify-center gap-2 hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200 active:scale-95 text-xs sm:text-sm"
          >
            <Plus className="w-5 h-5" />
            <span>Tambah Jamaah</span>
          </button>
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
        <div className="p-4 border-b border-slate-100 bg-slate-50/50">
          <div className="relative max-w-md">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
            <input 
              type="text" 
              placeholder="Cari jamaah (Nama, Panggilan, ID, Telepon)..." 
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-10 pr-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none bg-white"
            />
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead className="bg-slate-50 text-slate-500 text-sm font-medium">
              <tr>
                <th className="px-6 py-4 text-center">Foto</th>
                <th className="px-6 py-4">ID / Nama</th>
                <th className="px-6 py-4">Dapukan</th>
                <th className="px-6 py-4">Kategori</th>
                <th className="px-6 py-4">Telepon</th>
                <th className="px-6 py-4">Lokasi</th>
                <th className="px-6 py-4 text-center">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                <tr>
                  <td colSpan={7} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center justify-center gap-3">
                      <Loader2 className="w-8 h-8 text-emerald-600 animate-spin" />
                      <p className="text-sm font-semibold text-slate-700">Memuat data jamaah dari Google Sheets...</p>
                      <p className="text-xs text-slate-400">Harap tunggu sebentar</p>
                    </div>
                  </td>
                </tr>
              ) : filteredJamaah.length === 0 ? (
                <tr><td colSpan={7} className="px-6 py-8 text-center text-slate-400">Tidak ada data ditemukan</td></tr>
              ) : paginatedJamaah.map((j, jIdx) => {
                const dapukanList = j.dapukan || j.positions || [];
                return (
                  <tr key={j.id ? `${j.id}-${jIdx}` : `jamaah-${jIdx}`} className="hover:bg-slate-50 transition-colors">
                    <td className="px-6 py-4">
                      <div className="w-10 h-10 rounded-full bg-slate-100 overflow-hidden flex items-center justify-center mx-auto">
                        {j.photoUrl ? (
                          <img src={j.photoUrl} alt={j.name} className="w-full h-full object-cover" />
                        ) : (
                          <Users className="w-5 h-5 text-slate-400" />
                        )}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex flex-col">
                        <span className="text-xs font-bold text-emerald-600">{j.memberId}</span>
                        <span className="font-semibold text-slate-900">
                          {j.name} {j.nickname ? <span className="text-xs font-normal text-slate-500">({j.nickname})</span> : null}
                        </span>
                        {(j.isKK === true || (j.isKK as any) === 'true' || (j.isKK as any) === 'TRUE') && <span className="text-[10px] text-emerald-600 font-medium">Kepala Keluarga</span>}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex flex-wrap gap-1 max-w-xs">
                        {dapukanList.length > 0 ? (
                          dapukanList.map((p, pIdx) => (
                            <span key={`${p}-${pIdx}`} className="px-1.5 py-0.5 rounded bg-slate-100 text-slate-700 text-[9px] font-bold uppercase whitespace-nowrap">
                              {p}
                            </span>
                          ))
                        ) : (
                          <span className="text-[10px] text-slate-300 italic">-</span>
                        )}
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <span className={cn(
                        "px-2 py-0.5 rounded text-[10px] font-bold uppercase",
                        j.category === 'UMUM' ? "bg-slate-100 text-slate-600" :
                        j.category === 'ACR' ? "bg-emerald-100 text-emerald-700" :
                        j.category === 'APR' ? "bg-blue-100 text-blue-700" :
                        "bg-purple-100 text-purple-700"
                      )}>
                        {j.category}
                      </span>
                    </td>
                    <td className="px-6 py-4 text-slate-600 text-sm">{j.phone || '-'}</td>
                    <td className="px-6 py-4">
                      <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-700 text-xs font-semibold">
                        <MapPin className="w-3 h-3" /> {j.location}
                      </span>
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex justify-center gap-2">
                        <button 
                          onClick={() => setSelectedJamaahDetail(j)}
                          className="p-1.5 rounded-lg text-slate-600 hover:bg-slate-100 hover:text-slate-900 transition-colors"
                          title="Lihat Detail"
                        >
                          <Eye className="w-4 h-4" />
                        </button>
                        <button 
                          onClick={() => openEdit(j)}
                          className="text-emerald-600 hover:text-emerald-700 font-medium text-sm"
                        >
                          Edit
                        </button>
                        <button 
                          onClick={() => setDeleteId(j.id)}
                          className="text-red-500 hover:text-red-700 font-medium text-sm"
                        >
                          Hapus
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <Pagination
          currentPage={currentPage}
          totalPages={totalPages}
          totalItems={filteredJamaah.length}
          pageSize={pageSize}
          onPageChange={setCurrentPage}
          onPageSizeChange={setPageSize}
          className="px-6 py-4 bg-white"
        />
      </div>

      <DeleteConfirmation 
        isOpen={!!deleteId}
        onClose={() => setDeleteId(null)}
        onConfirm={async () => {
          if (!deleteId) return;
          setIsDeleting(true);
          try {
            await deleteData(null, accessToken, spreadsheetId, 'jamaah', deleteId);
            showToast('Data jamaah berhasil dihapus!', 'success');
            setDeleteId(null);
          } catch (err: any) {
            showToast(`Gagal menghapus: ${err.message}`, 'error');
          } finally {
            setIsDeleting(false);
          }
        }}
        loading={isDeleting}
        title="Hapus Data Jamaah"
        message="Apakah Anda yakin ingin menghapus data jamaah ini? Data yang dihapus tidak dapat dikembalikan."
      />

      {/* Detail Jamaah Modal */}
      <AnimatePresence>
        {selectedJamaahDetail && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50"
          >
            <motion.div 
              initial={{ scale: 0.95, y: 20 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.95, y: 20 }}
              className="bg-white rounded-3xl p-6 sm:p-8 w-full max-w-2xl shadow-2xl max-h-[90vh] overflow-y-auto"
            >
              <div className="flex justify-between items-start mb-6 border-b pb-4">
                <div className="flex items-center gap-4">
                  <div className="w-16 h-16 rounded-2xl bg-slate-100 overflow-hidden flex items-center justify-center border border-slate-200">
                    {selectedJamaahDetail.photoUrl ? (
                      <img src={selectedJamaahDetail.photoUrl} className="w-full h-full object-cover" />
                    ) : (
                      <Users className="w-8 h-8 text-slate-400" />
                    )}
                  </div>
                  <div>
                    <span className="text-xs font-bold text-emerald-600 uppercase tracking-wider">{selectedJamaahDetail.memberId}</span>
                    <h3 className="text-xl font-bold text-slate-900">{selectedJamaahDetail.name} {selectedJamaahDetail.nickname ? `(${selectedJamaahDetail.nickname})` : ''}</h3>
                    <div className="flex items-center gap-2 mt-1">
                      <span className="text-xs px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 font-semibold">{selectedJamaahDetail.location}</span>
                      <span className="text-xs px-2 py-0.5 rounded bg-slate-100 text-slate-700 font-semibold">{selectedJamaahDetail.category}</span>
                    </div>
                  </div>
                </div>
                <button 
                  onClick={() => setSelectedJamaahDetail(null)}
                  className="p-2 text-slate-400 hover:text-slate-600 rounded-xl hover:bg-slate-100"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="space-y-6 text-sm text-slate-700">
                {/* Identitas Diri */}
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 space-y-2">
                  <h4 className="font-bold text-slate-900 flex items-center gap-2 text-emerald-700">
                    <Users className="w-4 h-4" /> Identitas Diri & Kontak
                  </h4>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2">
                    <div><span className="text-slate-400 block text-xs">Jenis Kelamin</span> <span className="font-semibold">{selectedJamaahDetail.gender || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">Tempat, Tgl Lahir</span> <span className="font-semibold">{selectedJamaahDetail.placeOfBirth || '-'}, {selectedJamaahDetail.dateOfBirth || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">Golongan Darah</span> <span className="font-semibold">{selectedJamaahDetail.bloodType || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">No. Telepon / WA</span> <span className="font-semibold">{selectedJamaahDetail.phone || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">No. Ortu yang Bisa Dihubungi</span> <span className="font-semibold">{selectedJamaahDetail.parentPhone || '-'}</span></div>
                    <div className="sm:col-span-2"><span className="text-slate-400 block text-xs">Alamat Asal</span> <span className="font-semibold">{selectedJamaahDetail.originAddress || '-'}</span></div>
                    <div className="sm:col-span-2"><span className="text-slate-400 block text-xs">Alamat Saat Ini</span> <span className="font-semibold">{selectedJamaahDetail.currentAddress || selectedJamaahDetail.address || '-'}</span></div>
                  </div>
                </div>

                {/* Keluarga */}
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 space-y-2">
                  <h4 className="font-bold text-slate-900 flex items-center gap-2 text-emerald-700">
                    <Home className="w-4 h-4" /> Informasi Keluarga
                  </h4>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2">
                    <div><span className="text-slate-400 block text-xs">Nama Ayah</span> <span className="font-semibold">{selectedJamaahDetail.fatherName || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">Nama Ibu</span> <span className="font-semibold">{selectedJamaahDetail.motherName || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">Status Pernikahan</span> <span className="font-semibold">{selectedJamaahDetail.maritalStatus || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">Tahun Nikah</span> <span className="font-semibold">{selectedJamaahDetail.marriageYear || '-'}</span></div>
                    <div className="sm:col-span-2"><span className="text-slate-400 block text-xs">Nama Suami/Istri</span> <span className="font-semibold">{selectedJamaahDetail.spouseName || '-'}</span></div>
                  </div>
                </div>

                {/* Pendidikan & Kerja */}
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 space-y-2">
                  <h4 className="font-bold text-slate-900 flex items-center gap-2 text-emerald-700">
                    <GraduationCap className="w-4 h-4" /> Pendidikan & Pekerjaan
                  </h4>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2">
                    <div><span className="text-slate-400 block text-xs">Pendidikan Terakhir</span> <span className="font-semibold">{selectedJamaahDetail.lastEducation || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">Kelas / Jurusan</span> <span className="font-semibold">{selectedJamaahDetail.majorOrClass || '-'}</span></div>
                    <div className="sm:col-span-2"><span className="text-slate-400 block text-xs">Sekolah / Universitas</span> <span className="font-semibold">{selectedJamaahDetail.schoolOrUniversity || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">Pekerjaan Saat Ini</span> <span className="font-semibold">{selectedJamaahDetail.currentJob || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">Alamat Tempat Kerja</span> <span className="font-semibold">{selectedJamaahDetail.workplaceAddress || '-'}</span></div>
                  </div>
                </div>

                {/* Dapukan & Keorganisasian */}
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 space-y-2">
                  <h4 className="font-bold text-slate-900 flex items-center gap-2 text-emerald-700">
                    <Briefcase className="w-4 h-4" /> Dapukan & Keorganisasian
                  </h4>
                  <div className="space-y-2 pt-2">
                    <div>
                      <span className="text-slate-400 block text-xs mb-1">Dapukan Saat Ini</span>
                      <div className="flex flex-wrap gap-1">
                        {(selectedJamaahDetail.dapukan || selectedJamaahDetail.positions || []).length > 0 ? (
                          (selectedJamaahDetail.dapukan || selectedJamaahDetail.positions || []).map((d, dIdx) => (
                            <span key={`${d}-${dIdx}`} className="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold text-xs">{d}</span>
                          ))
                        ) : (
                          <span className="font-semibold italic text-slate-400">Belum ada dapukan</span>
                        )}
                      </div>
                    </div>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-2">
                      <div><span className="text-slate-400 block text-xs">Dapukan Sebelumnya</span> <span className="font-semibold">{selectedJamaahDetail.previousDapukan || '-'}</span></div>
                      <div><span className="text-slate-400 block text-xs">Status Mubaligh</span> <span className="font-semibold">{selectedJamaahDetail.isMubaligh || '-'}</span></div>
                      <div><span className="text-slate-400 block text-xs">Punya Saham UB</span> <span className="font-semibold">{selectedJamaahDetail.hasUbShares || '-'}</span></div>
                    </div>
                  </div>
                </div>

                {/* Keilmuan, Haji & Kesehatan */}
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 space-y-2">
                  <h4 className="font-bold text-slate-900 flex items-center gap-2 text-emerald-700">
                    <HeartPulse className="w-4 h-4" /> Keilmuan, Haji & Kesehatan
                  </h4>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2">
                    <div><span className="text-slate-400 block text-xs">Sudah Jurus Keras</span> <span className="font-semibold">{selectedJamaahDetail.hasJurusKeras || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">Sudah Jurus Halus</span> <span className="font-semibold">{selectedJamaahDetail.hasJurusHalus || '-'}</span></div>
                    <div><span className="text-slate-400 block text-xs">Sudah Haji</span> <span className="font-semibold">{selectedJamaahDetail.hasHajj || '-'}</span></div>
                    {selectedJamaahDetail.hasHajj === 'Sudah' ? (
                      <>
                        <div><span className="text-slate-400 block text-xs">Nama Haji</span> <span className="font-semibold">{selectedJamaahDetail.hajjName || '-'}</span></div>
                        <div><span className="text-slate-400 block text-xs">Tahun Haji</span> <span className="font-semibold">{selectedJamaahDetail.hajjYear || '-'}</span></div>
                      </>
                    ) : (
                      <>
                        <div><span className="text-slate-400 block text-xs">No. Porsi Haji</span> <span className="font-semibold">{selectedJamaahDetail.hajjPortionNumber || '-'}</span></div>
                        <div><span className="text-slate-400 block text-xs">Rencana Tahun Berangkat</span> <span className="font-semibold">{selectedJamaahDetail.plannedHajjYear || '-'}</span></div>
                      </>
                    )}
                    <div className="sm:col-span-2"><span className="text-slate-400 block text-xs">Riwayat Penyakit</span> <span className="font-semibold">{selectedJamaahDetail.medicalHistory || '-'}</span></div>
                  </div>
                </div>
              </div>

              <div className="mt-6 flex justify-end">
                <button 
                  onClick={() => setSelectedJamaahDetail(null)}
                  className="px-6 py-2.5 rounded-xl bg-slate-100 font-semibold text-slate-700 hover:bg-slate-200 transition-colors"
                >
                  Tutup
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Form Tambah/Edit Jamaah Modal */}
      <AnimatePresence>
        {showForm && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50"
          >
            <motion.div 
              initial={{ scale: 0.95, y: 20 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.95, y: 20 }}
              className="bg-white rounded-3xl p-6 sm:p-8 w-full max-w-3xl shadow-2xl max-h-[92vh] flex flex-col overflow-hidden"
            >
              <div className="flex justify-between items-center mb-4">
                <h3 className="text-xl font-bold">{editingJamaah ? 'Edit Data Jamaah' : 'Tambah Jamaah Baru'}</h3>
                <button type="button" onClick={closeForm} className="p-2 text-slate-400 hover:text-slate-600 rounded-xl hover:bg-slate-100">
                  <X className="w-5 h-5" />
                </button>
              </div>

              {!editingJamaah && (
                <div className="mb-4 p-3 bg-gradient-to-r from-emerald-50 to-teal-50 rounded-2xl border border-emerald-100 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
                  <div className="flex items-center gap-2.5">
                    <div className="w-8 h-8 rounded-xl bg-emerald-600 text-white flex items-center justify-center shadow-sm flex-shrink-0">
                      <Share2 className="w-4 h-4" />
                    </div>
                    <div>
                      <p className="text-xs font-bold text-slate-800">Ingin jamaah mengisi datanya sendiri?</p>
                      <p className="text-[11px] text-slate-500">Buat link pendaftaran mandiri berbatas waktu (expired date) & QR Code.</p>
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      closeForm();
                      setShowRegistrationLinkModal(true);
                    }}
                    className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold transition-all shadow-sm flex items-center justify-center gap-1.5 flex-shrink-0 active:scale-95"
                  >
                    <LinkIcon className="w-3.5 h-3.5" />
                    <span>Buat Link Pengisian</span>
                  </button>
                </div>
              )}

              {/* Form Navigation Tabs */}
              <div className="flex border-b border-slate-100 gap-1 overflow-x-auto pb-2 mb-4 scrollbar-none">
                {[
                  { id: 'identitas', label: '1. Identitas & Kontak' },
                  { id: 'keluarga', label: '2. Orang Tua & Keluarga' },
                  { id: 'pendidikan', label: '3. Pendidikan & Pekerjaan' },
                  { id: 'dapukan', label: '4. Dapukan (Checklist)' },
                  { id: 'keilmuan', label: '5. Keilmuan, Haji & Kesehatan' },
                ].map(tab => (
                  <button
                    key={tab.id}
                    type="button"
                    onClick={() => setFormTab(tab.id as any)}
                    className={cn(
                      "px-3.5 py-2 rounded-xl text-xs font-bold whitespace-nowrap transition-all",
                      formTab === tab.id
                        ? "bg-emerald-600 text-white shadow-md shadow-emerald-200"
                        : "bg-slate-50 text-slate-600 hover:bg-slate-100"
                    )}
                  >
                    {tab.label}
                  </button>
                ))}
              </div>

              <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto space-y-4 pr-1">
                {/* TAB 1: IDENTITAS & KONTAK */}
                <div className={cn("space-y-4", formTab === 'identitas' ? 'block' : 'hidden')}>
                  <div className="flex justify-center mb-4">
                    <label className="relative group cursor-pointer">
                      <div className="w-24 h-24 rounded-full bg-slate-100 border-2 border-dashed border-slate-300 flex items-center justify-center overflow-hidden transition-all group-hover:border-emerald-500">
                        {base64Image ? (
                          <img src={base64Image} className="w-full h-full object-cover" />
                        ) : (
                          <Plus className="w-8 h-8 text-slate-400 group-hover:text-emerald-500" />
                        )}
                      </div>
                      <input type="file" accept="image/*" onChange={handleImageChange} className="hidden" />
                      <div className="absolute -bottom-1 -right-1 bg-emerald-600 text-white p-1.5 rounded-full shadow-lg">
                        <Plus className="w-3 h-3" />
                      </div>
                    </label>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Nama Lengkap *</label>
                      <input name="name" defaultValue={editingJamaah?.name} className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Nama Panggilan</label>
                      <input name="nickname" defaultValue={editingJamaah?.nickname} placeholder="Contoh: Budi" className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Tempat Lahir</label>
                      <input name="placeOfBirth" defaultValue={editingJamaah?.placeOfBirth} placeholder="Contoh: Jakarta" className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Tanggal Lahir</label>
                      <input type="date" name="dateOfBirth" defaultValue={editingJamaah?.dateOfBirth} className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Jenis Kelamin</label>
                      <select name="gender" defaultValue={editingJamaah?.gender || ''} className="w-full px-4 py-2.5 rounded-xl border outline-none text-sm bg-white">
                        <option value="">Pilih Jenis Kelamin...</option>
                        <option value="Laki-laki">Laki-laki</option>
                        <option value="Perempuan">Perempuan</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Golongan Darah</label>
                      <select name="bloodType" defaultValue={editingJamaah?.bloodType || ''} className="w-full px-4 py-2.5 rounded-xl border outline-none text-sm bg-white">
                        <option value="">Pilih Golongan Darah...</option>
                        <option value="A">A</option>
                        <option value="B">B</option>
                        <option value="AB">AB</option>
                        <option value="O">O</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Nomor Telepon / WA *</label>
                      <input type="tel" name="phone" defaultValue={editingJamaah?.phone} className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                    </div>
                  </div>

                  <div className="flex items-center gap-3 p-3 bg-slate-50 rounded-xl border">
                    <input 
                      type="checkbox" 
                      id="isKK" 
                      checked={isKK} 
                      onChange={(e) => setIsKK(e.target.checked)}
                      className="w-4 h-4 text-emerald-600 rounded"
                    />
                    <label htmlFor="isKK" className="text-sm font-semibold text-slate-700">Kepala Keluarga (KK)</label>
                  </div>

                  {!isKK && (
                    <div className="relative">
                      <label className="block text-xs font-bold text-slate-700 mb-1">
                        Pilih Kepala Keluarga (KK) <span className="text-rose-500">*</span>
                      </label>
                      <div className="relative">
                        <div className="relative flex items-center">
                          <Search className="absolute left-3.5 w-4 h-4 text-slate-400 pointer-events-none" />
                          <input 
                            type="text"
                            placeholder="Cari KK berdasarkan Nama, ID, atau Telepon..."
                            value={kkSearchTerm}
                            onChange={(e) => {
                              setKkSearchTerm(e.target.value);
                              setIsKkDropdownOpen(true);
                            }}
                            onFocus={() => setIsKkDropdownOpen(true)}
                            className="w-full pl-10 pr-10 py-2.5 rounded-xl border border-slate-200 focus:ring-2 focus:ring-emerald-500 outline-none text-sm bg-white font-medium"
                          />
                          {selectedKKId ? (
                            <button
                              type="button"
                              onClick={() => {
                                setSelectedKKId('');
                                setKkSearchTerm('');
                              }}
                              className="absolute right-3 p-1 text-slate-400 hover:text-rose-500 rounded-md transition-colors"
                              title="Hapus pilihan KK"
                            >
                              <X className="w-4 h-4" />
                            </button>
                          ) : (
                            <ChevronDown 
                              className="absolute right-3 w-4 h-4 text-slate-400 cursor-pointer"
                              onClick={() => setIsKkDropdownOpen(prev => !prev)}
                            />
                          )}
                        </div>

                        {selectedKKId && (
                          <div className="mt-1.5 p-2.5 bg-emerald-50 border border-emerald-200 rounded-xl flex items-center justify-between">
                            {(() => {
                              const currKK = headsOfFamily.find(k => k.memberId === selectedKKId);
                              return (
                                <div className="flex items-center gap-2 overflow-hidden">
                                  <span className="px-2 py-0.5 rounded bg-emerald-600 text-white font-mono font-black text-xs shrink-0">{selectedKKId}</span>
                                  <span className="text-xs font-bold text-emerald-950 truncate">{currKK?.name || 'Kepala Keluarga Terpilih'}</span>
                                  {currKK?.location && (
                                    <span className="text-[10px] text-emerald-700 bg-emerald-100/80 px-1.5 py-0.5 rounded font-bold shrink-0">{currKK.location}</span>
                                  )}
                                </div>
                              );
                            })()}
                            <span className="text-[10px] text-emerald-700 font-bold uppercase tracking-wider shrink-0 ml-2">KK Aktif</span>
                          </div>
                        )}

                        {isKkDropdownOpen && (
                          <>
                            <div 
                              className="fixed inset-0 z-40" 
                              onClick={() => setIsKkDropdownOpen(false)}
                            />
                            <div className="absolute z-50 w-full mt-1.5 bg-white rounded-2xl border border-slate-200 shadow-2xl overflow-hidden max-h-60 overflow-y-auto divide-y divide-slate-100">
                              {(() => {
                                const eligibleKKs = headsOfFamily
                                  .filter(kk => kk.id !== editingJamaah?.id)
                                  .filter(kk => {
                                    if (!kkSearchTerm.trim()) return true;
                                    const term = kkSearchTerm.toLowerCase();
                                    return (
                                      kk.name.toLowerCase().includes(term) ||
                                      kk.memberId.toLowerCase().includes(term) ||
                                      (kk.nickname && kk.nickname.toLowerCase().includes(term)) ||
                                      (kk.phone && kk.phone.includes(term)) ||
                                      (kk.location && kk.location.toLowerCase().includes(term))
                                    );
                                  });

                                if (eligibleKKs.length === 0) {
                                  return (
                                    <div className="p-4 text-center text-xs text-slate-400 italic">
                                      Tidak ada Kepala Keluarga yang cocok dengan pencarian "{kkSearchTerm}"
                                    </div>
                                  );
                                }

                                return eligibleKKs.map((kk, kkIdx) => (
                                  <button
                                    key={kk.id ? `${kk.id}-${kkIdx}` : `kk-search-${kkIdx}`}
                                    type="button"
                                    onClick={() => {
                                      setSelectedKKId(kk.memberId);
                                      setKkSearchTerm(`${kk.memberId} - ${kk.name}`);
                                      setIsKkDropdownOpen(false);
                                    }}
                                    className={cn(
                                      "w-full px-4 py-2.5 text-left hover:bg-slate-50 transition-colors flex items-center justify-between gap-2",
                                      selectedKKId === kk.memberId ? "bg-emerald-50/70 text-emerald-950 font-bold" : "text-slate-800"
                                    )}
                                  >
                                    <div className="flex items-center gap-2.5 min-w-0">
                                      <div className="w-8 h-8 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center font-black text-xs shrink-0">
                                        {kk.name.charAt(0).toUpperCase()}
                                      </div>
                                      <div className="min-w-0">
                                        <div className="font-bold text-xs truncate flex items-center gap-1.5">
                                          <span>{kk.name}</span>
                                          {kk.nickname && <span className="text-slate-400 font-normal">({kk.nickname})</span>}
                                        </div>
                                        <div className="text-[10px] text-slate-500 font-mono font-medium">{kk.memberId} {kk.phone ? `• ${kk.phone}` : ''}</div>
                                      </div>
                                    </div>
                                    <span className="text-[10px] px-2 py-0.5 rounded-md bg-slate-100 font-bold text-slate-600 shrink-0">
                                      {kk.location}
                                    </span>
                                  </button>
                                ));
                              })()}
                            </div>
                          </>
                        )}
                      </div>
                    </div>
                  )}

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Kategori Jamaah</label>
                      <select name="category" defaultValue={editingJamaah?.category || 'UMUM'} className="w-full px-4 py-2.5 rounded-xl border outline-none text-sm">
                        <option value="UMUM">UMUM</option>
                        <option value="GPN">GPN (Generus Pra Nikah)</option>
                        <option value="APR">APR (Anak Pra Remaja)</option>
                        <option value="ACR">ACR (Anak Caberawit)</option>
                        <option value="DUDA">DUDA</option>
                        <option value="JANDA">JANDA</option>
                      </select>
                    </div>
                    {profile.role === 'admin' && (
                      <div>
                        <label className="block text-xs font-bold text-slate-700 mb-1">Lokasi Masjid</label>
                        <select name="location" defaultValue={editingJamaah?.location || 'Kramat Batu'} className="w-full px-4 py-2.5 rounded-xl border outline-none text-sm">
                          {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(l => (
                            <option key={l} value={l}>{l}</option>
                          ))}
                        </select>
                      </div>
                    )}
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Alamat Asal</label>
                    <textarea name="originAddress" defaultValue={editingJamaah?.originAddress} rows={2} placeholder="Alamat asal/daerah..." className="w-full px-4 py-2 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Alamat Saat Ini</label>
                    <textarea name="currentAddress" defaultValue={editingJamaah?.currentAddress || editingJamaah?.address} rows={2} placeholder="Alamat tinggal sekarang..." className="w-full px-4 py-2 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                  </div>

                  <div className="pt-2 flex justify-end">
                    <button type="button" onClick={() => setFormTab('keluarga')} className="px-6 py-2 rounded-xl bg-slate-900 text-white font-semibold text-xs">
                      Lanjut: Orang Tua & Keluarga &rarr;
                    </button>
                  </div>
                </div>

                {/* TAB 2: ORANG TUA & KELUARGA */}
                <div className={cn("space-y-4", formTab === 'keluarga' ? 'block' : 'hidden')}>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Nama Ayah</label>
                      <input name="fatherName" defaultValue={editingJamaah?.fatherName} placeholder="Nama Ayah Kandung" className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Nama Ibu</label>
                      <input name="motherName" defaultValue={editingJamaah?.motherName} placeholder="Nama Ibu Kandung" className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Nomor Ortu yang Bisa Dihubungi</label>
                    <input type="tel" name="parentPhone" defaultValue={editingJamaah?.parentPhone} placeholder="No. Telepon / WA Orang Tua" className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Status Pernikahan</label>
                      <select name="maritalStatus" defaultValue={editingJamaah?.maritalStatus || ''} className="w-full px-4 py-2.5 rounded-xl border outline-none text-sm">
                        <option value="">Pilih Status...</option>
                        <option value="Muda Mudi">Muda Mudi</option>
                        <option value="Nikah">Nikah</option>
                        <option value="Janda">Janda</option>
                        <option value="Duda">Duda</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Tahun Nikah</label>
                      <input name="marriageYear" defaultValue={editingJamaah?.marriageYear} placeholder="Contoh: 2018" className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Nama Suami / Istri</label>
                    <input name="spouseName" defaultValue={editingJamaah?.spouseName} placeholder="Nama Suami atau Istri" className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                  </div>

                  <div className="pt-2 flex justify-between">
                    <button type="button" onClick={() => setFormTab('identitas')} className="px-4 py-2 rounded-xl border text-slate-600 font-semibold text-xs">
                      &larr; Kembali
                    </button>
                    <button type="button" onClick={() => setFormTab('pendidikan')} className="px-6 py-2 rounded-xl bg-slate-900 text-white font-semibold text-xs">
                      Lanjut: Pendidikan & Kerja &rarr;
                    </button>
                  </div>
                </div>

                {/* TAB 3: PENDIDIKAN & PEKERJAAN */}
                <div className={cn("space-y-4", formTab === 'pendidikan' ? 'block' : 'hidden')}>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Pendidikan Terakhir</label>
                      <select name="lastEducation" defaultValue={editingJamaah?.lastEducation || ''} className="w-full px-4 py-2.5 rounded-xl border outline-none text-sm">
                        <option value="">Pilih Pendidikan...</option>
                        <option value="SD">SD</option>
                        <option value="SMP">SMP</option>
                        <option value="SMA">SMA</option>
                        <option value="D1">D1</option>
                        <option value="D2">D2</option>
                        <option value="D3">D3</option>
                        <option value="S1">S1</option>
                        <option value="S2">S2</option>
                        <option value="S3">S3</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Kelas / Jurusan</label>
                      <input name="majorOrClass" defaultValue={editingJamaah?.majorOrClass} placeholder="Contoh: Teknik Informatika / XII IPA" className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Nama Sekolah / Universitas</label>
                    <input name="schoolOrUniversity" defaultValue={editingJamaah?.schoolOrUniversity} placeholder="Contoh: Universitas Indonesia" className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Pekerjaan Saat Ini</label>
                    <input name="currentJob" defaultValue={editingJamaah?.currentJob} placeholder="Contoh: Karyawan Swasta / PNS" className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Alamat Tempat Kerja</label>
                    <textarea name="workplaceAddress" defaultValue={editingJamaah?.workplaceAddress} rows={2} placeholder="Alamat kantor/perusahaan..." className="w-full px-4 py-2 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                  </div>

                  <div className="pt-2 flex justify-between">
                    <button type="button" onClick={() => setFormTab('keluarga')} className="px-4 py-2 rounded-xl border text-slate-600 font-semibold text-xs">
                      &larr; Kembali
                    </button>
                    <button type="button" onClick={() => setFormTab('dapukan')} className="px-6 py-2 rounded-xl bg-slate-900 text-white font-semibold text-xs">
                      Lanjut: Dapukan (Checklist) &rarr;
                    </button>
                  </div>
                </div>

                {/* TAB 4: DAPUKAN (CHECKLIST MULTI-SELECT) */}
                <div className={cn("space-y-4", formTab === 'dapukan' ? 'block' : 'hidden')}>
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-emerald-50 p-3 rounded-2xl border border-emerald-100">
                    <div>
                      <h4 className="font-bold text-emerald-800 text-sm">Pilih Dapukan (Tugas Sabilillah)</h4>
                      <p className="text-xs text-emerald-600">Centang dapukan yang diampu oleh jamaah. ({selectedDapukan.length} dipilih)</p>
                    </div>
                    <input 
                      type="text"
                      placeholder="Cari dapukan..."
                      value={dapukanFilter}
                      onChange={(e) => setDapukanFilter(e.target.value)}
                      className="px-3 py-1.5 rounded-xl border border-emerald-200 text-xs outline-none bg-white"
                    />
                  </div>

                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5 p-3 bg-slate-50 rounded-2xl border border-slate-100 max-h-64 overflow-y-auto">
                    {DAPUKAN_OPTIONS.filter(d => d.toLowerCase().includes(dapukanFilter.toLowerCase())).map((pos, posIdx) => {
                      const isChecked = selectedDapukan.includes(pos);
                      return (
                        <label key={`${pos}-${posIdx}`} className={cn(
                          "flex items-center gap-2.5 p-2 rounded-xl border transition-all cursor-pointer text-xs font-medium",
                          isChecked 
                            ? "bg-emerald-50 border-emerald-500 text-emerald-900 font-bold shadow-sm" 
                            : "bg-white border-slate-200 text-slate-700 hover:border-slate-300"
                        )}>
                          <input 
                            type="checkbox" 
                            checked={isChecked}
                            onChange={(e) => {
                              if (e.target.checked) {
                                setSelectedDapukan([...selectedDapukan, pos]);
                              } else {
                                setSelectedDapukan(selectedDapukan.filter(p => p !== pos));
                              }
                            }}
                            className="w-4 h-4 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500"
                          />
                          <span className="truncate">{pos}</span>
                        </label>
                      );
                    })}
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Dapukan di Tempat Sebelumnya</label>
                    <input name="previousDapukan" defaultValue={editingJamaah?.previousDapukan} placeholder="Dapukan/Tugas di sambungan atau kelompok lama" className="w-full px-4 py-2.5 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Status Mubaligh</label>
                      <select name="isMubaligh" defaultValue={editingJamaah?.isMubaligh || ''} className="w-full px-4 py-2.5 rounded-xl border outline-none text-sm">
                        <option value="">Pilih Status Mubaligh...</option>
                        <option value="Ya">Ya</option>
                        <option value="Tidak">Tidak</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Punya Saham UB</label>
                      <select name="hasUbShares" defaultValue={editingJamaah?.hasUbShares || ''} className="w-full px-4 py-2.5 rounded-xl border outline-none text-sm">
                        <option value="">Pilih Saham UB...</option>
                        <option value="Sudah">Sudah</option>
                        <option value="Belum">Belum</option>
                      </select>
                    </div>
                  </div>

                  <div className="pt-2 flex justify-between">
                    <button type="button" onClick={() => setFormTab('pendidikan')} className="px-4 py-2 rounded-xl border text-slate-600 font-semibold text-xs">
                      &larr; Kembali
                    </button>
                    <button type="button" onClick={() => setFormTab('keilmuan')} className="px-6 py-2 rounded-xl bg-slate-900 text-white font-semibold text-xs">
                      Lanjut: Keilmuan & Haji &rarr;
                    </button>
                  </div>
                </div>

                {/* TAB 5: KEILMUAN, HAJI & KESEHATAN */}
                <div className={cn("space-y-4", formTab === 'keilmuan' ? 'block' : 'hidden')}>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Sudah Jurus Keras</label>
                      <select name="hasJurusKeras" defaultValue={editingJamaah?.hasJurusKeras || ''} className="w-full px-4 py-2.5 rounded-xl border outline-none text-sm">
                        <option value="">Pilih Status...</option>
                        <option value="Sudah">Sudah</option>
                        <option value="Belum">Belum</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Sudah Jurus Halus</label>
                      <select name="hasJurusHalus" defaultValue={editingJamaah?.hasJurusHalus || ''} className="w-full px-4 py-2.5 rounded-xl border outline-none text-sm">
                        <option value="">Pilih Status...</option>
                        <option value="Sudah">Sudah</option>
                        <option value="Belum">Belum</option>
                      </select>
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Sudah Haji</label>
                    <select 
                      name="hasHajj" 
                      value={hasHajjStatus} 
                      onChange={(e) => setHasHajjStatus(e.target.value)}
                      className="w-full px-4 py-2.5 rounded-xl border outline-none text-sm bg-white"
                    >
                      <option value="">Pilih Status Haji...</option>
                      <option value="Sudah">Sudah</option>
                      <option value="Belum">Belum</option>
                    </select>
                  </div>

                  {hasHajjStatus === 'Belum' && (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 p-3 bg-amber-50 rounded-2xl border border-amber-100">
                      <div>
                        <label className="block text-xs font-bold text-amber-900 mb-1">Nomor Porsi Haji (Bila Belum)</label>
                        <input name="hajjPortionNumber" defaultValue={editingJamaah?.hajjPortionNumber} placeholder="No. Porsi pendaftaran" className="w-full px-4 py-2 rounded-xl border focus:ring-2 focus:ring-amber-500 outline-none text-sm bg-white" />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-amber-900 mb-1">Rencana Tahun Berangkat (Bila Belum)</label>
                        <input name="plannedHajjYear" defaultValue={editingJamaah?.plannedHajjYear} placeholder="Contoh: 2028" className="w-full px-4 py-2 rounded-xl border focus:ring-2 focus:ring-amber-500 outline-none text-sm bg-white" />
                      </div>
                    </div>
                  )}

                  {hasHajjStatus === 'Sudah' && (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 p-3 bg-emerald-50 rounded-2xl border border-emerald-100">
                      <div>
                        <label className="block text-xs font-bold text-emerald-900 mb-1">Nama Haji (Bila Sudah)</label>
                        <input name="hajjName" defaultValue={editingJamaah?.hajjName} placeholder="Gelar / Nama Haji" className="w-full px-4 py-2 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm bg-white" />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-emerald-900 mb-1">Tahun Haji (Bila Sudah)</label>
                        <input name="hajjYear" defaultValue={editingJamaah?.hajjYear} placeholder="Contoh: 2022" className="w-full px-4 py-2 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm bg-white" />
                      </div>
                    </div>
                  )}

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Riwayat Penyakit</label>
                    <textarea name="medicalHistory" defaultValue={editingJamaah?.medicalHistory} rows={2} placeholder="Catatan riwayat kesehatan/penyakit..." className="w-full px-4 py-2 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none text-sm" />
                  </div>

                  <div className="pt-2 flex justify-between">
                    <button type="button" onClick={() => setFormTab('dapukan')} className="px-4 py-2 rounded-xl border text-slate-600 font-semibold text-xs">
                      &larr; Kembali
                    </button>
                  </div>
                </div>

                {/* Form Action Buttons */}
                <div className="flex gap-3 pt-6 border-t mt-4">
                  <button 
                    type="button" 
                    onClick={closeForm} 
                    disabled={isSaving}
                    className="flex-1 px-6 py-2.5 rounded-xl border hover:bg-slate-50 transition-all font-semibold text-sm text-slate-700 disabled:opacity-50"
                  >
                    Batal
                  </button>
                  <button 
                    type="submit" 
                    disabled={isSaving}
                    className="flex-1 px-6 py-2.5 rounded-xl bg-emerald-600 text-white hover:bg-emerald-700 transition-all font-semibold text-sm shadow-lg shadow-emerald-200 disabled:opacity-60 disabled:cursor-not-allowed flex items-center justify-center gap-2"
                  >
                    {isSaving ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        <span>Menyimpan...</span>
                      </>
                    ) : (
                      editingJamaah ? 'Simpan Perubahan' : 'Simpan Data Jamaah'
                    )}
                  </button>
                </div>
              </form>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Modal Generate Link Pendaftaran Mandiri */}
      <RegistrationLinkModal
        isOpen={showRegistrationLinkModal}
        onClose={() => setShowRegistrationLinkModal(false)}
        profile={profile}
      />
    </div>
  );
}

function InventarisView({ profile, formTrigger, onFormTriggered, assetType = 'barang' }: { profile: UserProfile, formTrigger?: string | null, onFormTriggered?: () => void, assetType?: 'barang' | 'tanah' }) {
  const { accessToken, spreadsheetId } = useFirebase();
  const { showToast } = useToast();
  const baseFilter = useMemo(() => (profile.role === 'pengurus' && profile.location && (profile.location as string) !== 'Seluruh Lokasi') ? [where('location', '==', profile.location)] : [], [profile.role, profile.location]);
  const { data: allAssets, loading } = useDataQuery<Asset>('assets', baseFilter);
  const assets = useMemo(() => allAssets.filter(a => {
    if (assetType === 'barang') return !a.assetType || a.assetType === 'barang';
    return a.assetType === 'tanah';
  }), [allAssets, assetType]);

  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(9);

  useEffect(() => {
    setCurrentPage(1);
  }, [assetType]);

  const totalPages = Math.ceil(assets.length / pageSize) || 1;
  const paginatedAssets = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return assets.slice(start, start + pageSize);
  }, [assets, currentPage, pageSize]);

  const [showForm, setShowForm] = useState(false);
  const [editingAsset, setEditingAsset] = useState<Asset | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [base64Image, setBase64Image] = useState<string | null>(null);

  useEffect(() => {
    if (formTrigger === 'inventaris' && assetType === 'barang') {
      setShowForm(true);
      setEditingAsset(null);
      setBase64Image(null);
      onFormTriggered?.();
    } else if (formTrigger === 'tanah' && assetType === 'tanah') {
      setShowForm(true);
      setEditingAsset(null);
      setBase64Image(null);
      onFormTriggered?.();
    }
  }, [formTrigger, assetType]);

  const handleImageChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setBase64Image(reader.result as string);
      };
      reader.readAsDataURL(file);
    }
  };

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const data = {
      name: formData.get('name') as string,
      category: formData.get('category') as string,
      quantity: assetType === 'barang' ? Number(formData.get('quantity')) : 1,
      areaSize: assetType === 'tanah' ? Number(formData.get('areaSize')) : null,
      status: formData.get('status') as Asset['status'],
      location: (formData.get('location') as MosqueLocation) || (profile.location as MosqueLocation) || 'Utama',
      lastChecked: Date.now(),
      assetType,
      photoUrl: base64Image || editingAsset?.photoUrl || null,
      description: formData.get('description') as string
    };

    setIsSaving(true);
    try {
      await saveData(null, accessToken, spreadsheetId, 'assets', data, editingAsset?.id);
      showToast(`Data ${isTanah ? 'tanah' : 'barang'} ${data.name} berhasil disimpan!`, 'success');
      closeForm();
    } catch (err: any) {
      showToast(`Gagal menyimpan data: ${err.message}`, 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const closeForm = () => {
    setShowForm(false);
    setEditingAsset(null);
    setBase64Image(null);
  };

  const openEdit = (a: Asset) => {
    setEditingAsset(a);
    setBase64Image(a.photoUrl || null);
    setShowForm(true);
  };

  const isTanah = assetType === 'tanah';

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-bold text-slate-900">{isTanah ? 'Tanah Sabilillah' : 'Inventaris Sabilillah'}</h2>
          <p className="text-slate-500">{isTanah ? 'Kelola Tanah Wakaf dan Sabilillah' : 'Kelola Inventaris Sabilillah'}</p>
        </div>
        <button 
          onClick={() => setShowForm(true)}
          className="bg-emerald-600 text-white px-6 py-2.5 rounded-xl font-semibold flex items-center gap-2 hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200"
        >
          <Plus className="w-5 h-5" />
          {isTanah ? 'Tambah Tanah' : 'Tambah Barang'}
        </button>
      </div>

      {loading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {[1, 2, 3].map((i) => (
            <div key={i} className="bg-white rounded-2xl border border-slate-100 p-6 shadow-sm animate-pulse space-y-4">
              <div className="h-32 bg-slate-100 rounded-xl w-full flex items-center justify-center">
                <Loader2 className="w-8 h-8 text-emerald-600 animate-spin" />
              </div>
              <p className="text-center text-xs font-semibold text-slate-500">Memuat data {isTanah ? 'tanah' : 'barang'}...</p>
            </div>
          ))}
        </div>
      ) : assets.length === 0 ? (
        <div className="bg-white rounded-2xl border border-slate-100 p-12 text-center text-slate-400 font-medium">
          Tidak ada data {isTanah ? 'tanah' : 'barang'} ditemukan
        </div>
      ) : (
        <div className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {paginatedAssets.map((asset, assetIdx) => (
              <motion.div 
                key={asset.id ? `${asset.id}-${assetIdx}` : `asset-${assetIdx}`}
                layout
                className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden flex flex-col"
              >
                {asset.photoUrl && (
                  <div className="h-40 overflow-hidden">
                    <img src={asset.photoUrl} className="w-full h-full object-cover" />
                  </div>
                )}
                <div className="p-6">
                  <div className="flex justify-between items-start mb-4">
                    <div className="p-3 bg-emerald-50 rounded-xl">
                      {isTanah ? <MapIcon className="w-6 h-6 text-emerald-600" /> : <Package className="w-6 h-6 text-emerald-600" />}
                    </div>
                    <span className={cn(
                      "px-2 py-1 rounded-lg text-[10px] font-bold uppercase tracking-wider",
                      asset.status === 'baik' || asset.status === 'wakaf' ? "bg-emerald-100 text-emerald-700" : 
                      asset.status === 'rusak' ? "bg-red-100 text-red-700" : "bg-amber-100 text-amber-700"
                    )}>
                      {asset.status}
                    </span>
                  </div>
                  <h3 className="font-bold text-slate-900 text-lg mb-1">{asset.name}</h3>
                  <p className="text-slate-500 text-sm mb-4">{asset.category}</p>
                  
                  <div className="space-y-3">
                    {isTanah ? (
                      <div className="flex items-center justify-between py-2 border-t border-slate-50">
                        <span className="text-sm text-slate-500">Luas Tanah</span>
                        <span className="font-bold text-slate-900">{asset.areaSize} m²</span>
                      </div>
                    ) : (
                      <div className="flex items-center justify-between py-2 border-t border-slate-50">
                        <span className="text-sm text-slate-500">Jumlah</span>
                        <span className="font-bold text-slate-900">{asset.quantity} Unit</span>
                      </div>
                    )}
                    <div className="flex items-center justify-between py-2 border-t border-slate-50">
                      <span className="text-sm text-slate-500">Lokasi</span>
                      <span className="text-sm font-medium text-slate-700">{asset.location}</span>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-3 mt-6">
                    <button 
                      onClick={() => openEdit(asset)}
                      className="py-2 rounded-xl text-xs font-semibold text-blue-600 border border-blue-100 hover:bg-blue-50 transition-all"
                    >
                      Edit
                    </button>
                    <button 
                      onClick={() => setDeleteId(asset.id)}
                      className="py-2 rounded-xl text-xs font-semibold text-red-500 border border-red-50 hover:bg-red-50 transition-all"
                    >
                      Hapus
                    </button>
                  </div>
                </div>
              </motion.div>
            ))}
          </div>

          <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
            <Pagination
              currentPage={currentPage}
              totalPages={totalPages}
              totalItems={assets.length}
              pageSize={pageSize}
              onPageChange={setCurrentPage}
              onPageSizeChange={setPageSize}
              pageSizeOptions={[6, 9, 15, 30]}
              className="px-6 py-4"
            />
          </div>
        </div>
      )}

      <DeleteConfirmation 
        isOpen={!!deleteId}
        onClose={() => setDeleteId(null)}
        onConfirm={async () => {
          if (!deleteId) return;
          setIsDeleting(true);
          try {
            await deleteData(null, accessToken, spreadsheetId, 'assets', deleteId);
            showToast(`Data ${isTanah ? 'tanah' : 'barang'} berhasil dihapus!`, 'success');
            setDeleteId(null);
          } catch (err: any) {
            showToast(`Gagal menghapus: ${err.message}`, 'error');
          } finally {
            setIsDeleting(false);
          }
        }}
        loading={isDeleting}
        title={isTanah ? "Hapus Data Tanah" : "Hapus Barang Inventaris"}
        message={isTanah ? "Hapus data tanah sabilillah ini? Tindakan ini permanen." : "Apakah Anda yakin ingin menghapus barang ini dari inventaris?"}
      />

      <AnimatePresence>
        {showForm && (
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
            <motion.div initial={{ scale: 0.95 }} animate={{ scale: 1 }} className="bg-white rounded-3xl p-8 w-full max-w-lg shadow-2xl max-h-[90vh] overflow-y-auto my-8">
              <h3 className="text-xl font-bold mb-6">{editingAsset ? `Edit ${isTanah ? 'Tanah' : 'Barang'}` : `Tambah ${isTanah ? 'Tanah Baru' : 'Barang Baru'}`}</h3>
              <form onSubmit={handleSubmit} className="space-y-4">
                <div className="flex justify-center mb-6">
                  <label className="relative group cursor-pointer">
                    <div className="w-32 h-24 rounded-2xl bg-slate-100 border-2 border-dashed border-slate-300 flex items-center justify-center overflow-hidden transition-all group-hover:border-emerald-500">
                      {base64Image ? (
                        <img src={base64Image} className="w-full h-full object-cover" />
                      ) : (
                        <div className="text-center">
                          <ImageIcon className="w-8 h-8 text-slate-400 mx-auto mb-1 group-hover:text-emerald-500" />
                          <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Foto</span>
                        </div>
                      )}
                    </div>
                    <input type="file" accept="image/*" onChange={handleImageChange} className="hidden" />
                  </label>
                </div>

                <div>
                  <label className="block text-sm font-medium mb-1">Nama {isTanah ? 'Lahan / Tanah' : 'Barang'}</label>
                  <input name="name" defaultValue={editingAsset?.name} required className="w-full px-4 py-2 rounded-xl border outline-none focus:ring-2 focus:ring-emerald-500" />
                </div>
                
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium mb-1">Kategori</label>
                    <input name="category" defaultValue={editingAsset?.category} required className="w-full px-4 py-2 rounded-xl border outline-none" placeholder={isTanah ? 'Sabilillah / Wakaf' : 'Elektronik / Alat'} />
                  </div>
                  <div>
                    {isTanah ? (
                      <>
                        <label className="block text-sm font-medium mb-1">Luas (m²)</label>
                        <input type="number" name="areaSize" defaultValue={editingAsset?.areaSize} required className="w-full px-4 py-2 rounded-xl border outline-none" />
                      </>
                    ) : (
                      <>
                        <label className="block text-sm font-medium mb-1">Jumlah</label>
                        <input type="number" name="quantity" defaultValue={editingAsset?.quantity} required className="w-full px-4 py-2 rounded-xl border outline-none" />
                      </>
                    )}
                  </div>
                </div>

                <div>
                  <label className="block text-sm font-medium mb-1">Status</label>
                  <select name="status" defaultValue={editingAsset?.status} className="w-full px-4 py-2 rounded-xl border outline-none">
                    {isTanah ? (
                      <>
                        <option value="wakaf">Wakaf</option>
                        <option value="sertifikasi">Proses Sertifikasi</option>
                        <option value="terjual">Sudah Terjual</option>
                      </>
                    ) : (
                      <>
                        <option value="baik">Baik</option>
                        <option value="perlu perbaikan">Perlu Perbaikan</option>
                        <option value="rusak">Rusak</option>
                      </>
                    )}
                  </select>
                </div>

                {profile.role === 'admin' && (
                  <div>
                    <label className="block text-sm font-medium mb-1">Lokasi</label>
                    <select name="location" defaultValue={editingAsset?.location || profile.location} className="w-full px-4 py-2 rounded-xl border outline-none">
                      {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(l => (
                        <option key={l} value={l}>{l}</option>
                      ))}
                    </select>
                  </div>
                )}

                <div>
                  <label className="block text-sm font-medium mb-1">Keterangan</label>
                  <textarea name="description" defaultValue={editingAsset?.description} className="w-full px-4 py-2 rounded-xl border outline-none min-h-[80px]" placeholder="Informasi tambahan..." />
                </div>

                <div className="flex gap-3 pt-4">
                  <button 
                    type="button" 
                    onClick={closeForm} 
                    disabled={isSaving}
                    className="flex-1 px-6 py-2.5 rounded-xl border font-semibold text-slate-600 hover:bg-slate-50 transition-all disabled:opacity-50"
                  >
                    Batal
                  </button>
                  <button 
                    type="submit" 
                    disabled={isSaving}
                    className="flex-1 px-6 py-2.5 rounded-xl bg-emerald-600 text-white font-semibold hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-100 disabled:opacity-60 disabled:cursor-not-allowed flex items-center justify-center gap-2"
                  >
                    {isSaving ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        <span>Menyimpan...</span>
                      </>
                    ) : (
                      'Simpan'
                    )}
                  </button>
                </div>
              </form>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

function ActivitiesView({ profile }: { profile: UserProfile }) {
  const { accessToken, spreadsheetId } = useFirebase();
  const { showToast } = useToast();
  const filter = useMemo(() => (profile.role === 'pengurus' && profile.location && (profile.location as string) !== 'Seluruh Lokasi') ? [where('location', '==', profile.location)] : [], [profile.role, profile.location]);
  const { data: activities, loading } = useDataQuery<Activity>('activities', filter);
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(6);

  const totalPages = Math.ceil(activities.length / pageSize) || 1;
  const paginatedActivities = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return activities.slice(start, start + pageSize);
  }, [activities, currentPage, pageSize]);

  const [showForm, setShowForm] = useState(false);
  const [editingActivity, setEditingActivity] = useState<Activity | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [base64Images, setBase64Images] = useState<string[]>([]);

  const handleImagesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files) return;
    const files = Array.from(e.target.files);
    files.forEach((file: File) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        setBase64Images(prev => [...prev, reader.result as string]);
      };
      reader.readAsDataURL(file);
    });
  };

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const data = {
      title: formData.get('title') as string,
      description: formData.get('description') as string,
      type: formData.get('type') as Activity['type'],
      date: editingActivity ? editingActivity.date : Date.now(),
      imageUrls: base64Images.length > 0 ? base64Images : editingActivity?.imageUrls || [],
      location: profile.role === 'pengurus' ? profile.location : formData.get('location') as MosqueLocation,
      createdBy: profile.uid
    };

    setIsSaving(true);
    try {
      await saveData(null, accessToken, spreadsheetId, 'activities', data, editingActivity?.id);
      showToast('Data kegiatan berhasil disimpan!', 'success');
      closeForm();
    } catch (err: any) {
      showToast(`Gagal menyimpan data: ${err.message}`, 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const closeForm = () => {
    setShowForm(false);
    setEditingActivity(null);
    setBase64Images([]);
  };

  const openEdit = (a: Activity) => {
    setEditingActivity(a);
    setBase64Images(parseImageUrls(a.imageUrls));
    setShowForm(true);
  };

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-bold text-slate-900">Kegiatan Desa Gandaria</h2>
          <p className="text-slate-500">Kelola Kegiatan Desa Gandaria</p>
        </div>
        <button 
          onClick={() => setShowForm(true)}
          className="bg-emerald-600 text-white px-6 py-2.5 rounded-xl font-semibold flex items-center gap-2 hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200"
        >
          <Plus className="w-5 h-5" />
          Tambah Kegiatan
        </button>
      </div>

      {loading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
          {[1, 2, 3].map((i) => (
            <div key={i} className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm animate-pulse space-y-4">
              <div className="h-48 bg-slate-100 rounded-2xl w-full flex items-center justify-center">
                <Loader2 className="w-8 h-8 text-emerald-600 animate-spin" />
              </div>
              <p className="text-center text-xs font-semibold text-slate-500">Memuat data kegiatan...</p>
            </div>
          ))}
        </div>
      ) : activities.length === 0 ? (
        <div className="bg-white rounded-3xl border border-slate-100 p-12 text-center text-slate-400 font-medium">
          Belum ada kegiatan terdaftar
        </div>
      ) : (
        <div className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
            {paginatedActivities.map((activity, actIdx) => {
              const actImages = parseImageUrls(activity.imageUrls);
              return (
                <div key={activity.id ? `${activity.id}-${actIdx}` : `act-${actIdx}`} className="bg-white rounded-3xl overflow-hidden border border-slate-100 shadow-sm flex flex-col">
                  <div className="h-56 relative group">
                    <div className="w-full h-full flex overflow-x-auto snap-x snap-mandatory scrollbar-hide">
                      {actImages.length > 0 ? (
                        actImages.map((url, imgIdx) => (
                          <img key={`act-img-${actIdx}-${imgIdx}`} src={url} className="w-full h-full object-cover flex-shrink-0 snap-center" />
                        ))
                      ) : (
                        <div className="w-full h-full bg-slate-100 flex items-center justify-center text-slate-400">
                          <Calendar className="w-12 h-12" />
                        </div>
                      )}
                    </div>
                  <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center gap-3">
                     <button 
                      onClick={() => openEdit(activity)}
                      className="bg-white/20 backdrop-blur-md text-white p-3 rounded-full hover:bg-emerald-500 transition-colors"
                    >
                      <Plus className="w-6 h-6" />
                    </button>
                     <button 
                      onClick={() => setDeleteId(activity.id)}
                      className="bg-white/20 backdrop-blur-md text-white p-3 rounded-full hover:bg-red-500 transition-colors"
                    >
                      <Plus className="w-6 h-6 rotate-45" />
                    </button>
                  </div>
                  <div className="absolute top-4 left-4">
                    <span className={cn(
                      "px-3 py-1 rounded-full text-xs font-bold text-white uppercase",
                      activity.type === 'harian' ? "bg-emerald-500" : "bg-blue-500"
                    )}>
                      {activity.type}
                    </span>
                  </div>
                </div>
                <div className="p-6">
                  <h3 className="text-lg font-bold text-slate-900 mb-2">{activity.title}</h3>
                  <p className="text-slate-500 text-sm mb-4 flex-grow">{activity.description}</p>
                  <div className="flex items-center gap-2 text-xs text-slate-400">
                    <Clock className="w-4 h-4" />
                    {new Date(activity.date).toLocaleString()}
                  </div>
                </div>
              </div>
              );
            })}
          </div>

          <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden">
            <Pagination
              currentPage={currentPage}
              totalPages={totalPages}
              totalItems={activities.length}
              pageSize={pageSize}
              onPageChange={setCurrentPage}
              onPageSizeChange={setPageSize}
              pageSizeOptions={[6, 12, 24]}
              className="px-6 py-4"
            />
          </div>
        </div>
      )}

      <AnimatePresence>
        {showForm && (
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
            <motion.div initial={{ scale: 0.95 }} animate={{ scale: 1 }} className="bg-white rounded-3xl p-8 w-full max-w-lg shadow-2xl max-h-[90vh] overflow-y-auto">
              <h3 className="text-xl font-bold mb-6">{editingActivity ? 'Edit Kegiatan' : 'Tambah Kegiatan'}</h3>
              <form onSubmit={handleSubmit} className="space-y-4">
                <div>
                  <label className="block text-sm font-medium mb-1">Judul Kegiatan</label>
                  <input name="title" defaultValue={editingActivity?.title} required className="w-full px-4 py-2 rounded-xl border outline-none" />
                </div>
                <div>
                  <label className="block text-sm font-medium mb-1">Deskripsi</label>
                  <textarea name="description" defaultValue={editingActivity?.description} required className="w-full px-4 py-2 rounded-xl border outline-none min-h-[100px]" />
                </div>
                <div className="grid grid-cols-1 gap-4">
                  <div>
                    <label className="block text-sm font-medium mb-1">Upload Foto (Bisa banyak)</label>
                    <div className="flex flex-wrap gap-2 mb-2">
                      {base64Images.map((img, i) => (
                        <div key={i} className="relative w-16 h-16 group">
                          <img src={img} className="w-full h-full object-cover rounded-lg" />
                          <button 
                            type="button" 
                            onClick={() => setBase64Images(prev => prev.filter((_, idx) => idx !== i))}
                            className="absolute -top-1 -right-1 bg-red-500 text-white rounded-full p-0.5 opacity-0 group-hover:opacity-100 transition-opacity"
                          >
                            <Plus className="w-3 h-3 rotate-45" />
                          </button>
                        </div>
                      ))}
                    </div>
                    <div className="flex items-center gap-4">
                      <label className="flex-grow cursor-pointer bg-slate-50 border-2 border-dashed border-slate-200 rounded-xl p-4 hover:border-emerald-500 transition-all text-center">
                        <span className="text-xs text-slate-400">Klik untuk upload foto-foto</span>
                        <input type="file" accept="image/*" multiple onChange={handleImagesChange} className="hidden" />
                      </label>
                    </div>
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-sm font-medium mb-1">Tipe</label>
                      <select name="type" defaultValue={editingActivity?.type} className="w-full px-4 py-2 rounded-xl border outline-none">
                        <option value="harian">Harian</option>
                        <option value="mingguan">Mingguan</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-sm font-medium mb-1">Lokasi</label>
                      <select name="location" defaultValue={editingActivity?.location || profile.location} className="w-full px-4 py-2 rounded-xl border outline-none" disabled={profile.role === 'pengurus'}>
                        {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(l => (
                          <option key={l} value={l}>{l}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                </div>
                <div className="flex gap-3 pt-4">
                  <button 
                    type="button" 
                    onClick={closeForm} 
                    disabled={isSaving}
                    className="flex-1 px-6 py-2.5 rounded-xl border font-semibold hover:bg-slate-50 transition-all disabled:opacity-50"
                  >
                    Batal
                  </button>
                  <button 
                    type="submit" 
                    disabled={isSaving}
                    className="flex-1 px-6 py-2.5 rounded-xl bg-emerald-600 text-white font-semibold hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-100 disabled:opacity-60 disabled:cursor-not-allowed flex items-center justify-center gap-2"
                  >
                    {isSaving ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        <span>Menyimpan...</span>
                      </>
                    ) : (
                      'Simpan'
                    )}
                  </button>
                </div>
              </form>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <DeleteConfirmation 
        isOpen={!!deleteId}
        onClose={() => setDeleteId(null)}
        onConfirm={async () => {
          if (!deleteId) return;
          setIsDeleting(true);
          try {
            await deleteData(null, accessToken, spreadsheetId, 'activities', deleteId);
            showToast('Kegiatan berhasil dihapus!', 'success');
            setDeleteId(null);
          } catch (err: any) {
            showToast(`Gagal menghapus: ${err.message}`, 'error');
          } finally {
            setIsDeleting(false);
          }
        }}
        loading={isDeleting}
        title="Hapus Kegiatan"
        message="Hapus postingan kegiatan ini? Foto dan deskripsi akan dihapus secara permanen."
      />
    </div>
  );
}

function FacilityView({ profile }: { profile: UserProfile }) {
  const { accessToken, spreadsheetId } = useFirebase();
  const { showToast } = useToast();
  const { data: stats, loading } = useDataQuery<FacilityStat>('facility_stats', []);
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(4);

  const totalPages = Math.ceil(stats.length / pageSize) || 1;
  const paginatedStats = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return stats.slice(start, start + pageSize);
  }, [stats, currentPage, pageSize]);

  const [showForm, setShowForm] = useState(false);
  const [editingStat, setEditingStat] = useState<FacilityStat | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  const updateUsage = async (id: string, current: number, capacity: number, delta: number) => {
    const newVal = Math.max(0, Math.min(capacity, current + delta));
    try {
      await saveData(null, accessToken, spreadsheetId, 'facility_stats', { 
        currentUsage: newVal,
        updatedAt: Date.now()
      }, id);
      showToast('Penggunaan fasilitas diperbarui!', 'success');
    } catch (err: any) {
      showToast(`Gagal mengupdate status: ${err.message}`, 'error');
    }
  };

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const data = {
      name: formData.get('name') as string,
      capacity: Number(formData.get('capacity')),
      currentUsage: editingStat ? editingStat.currentUsage : 0,
      updatedAt: Date.now()
    };

    setIsSaving(true);
    try {
      await saveData(null, accessToken, spreadsheetId, 'facility_stats', data, editingStat?.id);
      showToast('Data fasilitas berhasil disimpan!', 'success');
      closeForm();
    } catch (err: any) {
      showToast(`Gagal menyimpan data: ${err.message}`, 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const closeForm = () => {
    setShowForm(false);
    setEditingStat(null);
  };

  const openEdit = (s: FacilityStat) => {
    setEditingStat(s);
    setShowForm(true);
  };

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-bold text-slate-900">Penggunaan Fasilitas</h2>
          <p className="text-slate-500">Monitor penggunaan fasilitas secara real-time</p>
        </div>
        {profile.role === 'admin' && (
          <button 
            onClick={() => setShowForm(true)}
            className="bg-emerald-600 text-white px-6 py-2.5 rounded-xl font-semibold flex items-center gap-2 hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200"
          >
            <Plus className="w-5 h-5" />
            Tambah Fasilitas
          </button>
        )}
      </div>

      {loading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
          {[1, 2].map((i) => (
            <div key={i} className="bg-white p-8 rounded-3xl border border-slate-100 shadow-sm animate-pulse space-y-4">
              <div className="h-8 bg-slate-100 rounded-md w-1/2 mx-auto" />
              <div className="h-20 bg-slate-100 rounded-2xl w-full flex items-center justify-center">
                <Loader2 className="w-8 h-8 text-emerald-600 animate-spin" />
              </div>
              <p className="text-center text-xs font-semibold text-slate-500">Memuat data fasilitas...</p>
            </div>
          ))}
        </div>
      ) : stats.length === 0 ? (
        <div className="bg-white rounded-3xl border border-slate-100 p-12 text-center text-slate-400 font-medium">
          Tidak ada data fasilitas
        </div>
      ) : (
        <div className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
            {paginatedStats.map((stat, statIdx) => (
              <div key={stat.id ? `${stat.id}-${statIdx}` : `stat-${statIdx}`} className="bg-white p-8 rounded-3xl border border-slate-100 shadow-sm">
                <div className="flex items-center justify-between mb-8">
                  <h3 className="text-xl font-bold text-slate-900">{stat.name}</h3>
                  <div className="flex items-center gap-3">
                    {profile.role === 'admin' && (
                      <div className="flex gap-2 mr-2">
                        <button onClick={() => openEdit(stat)} className="text-slate-400 hover:text-emerald-500"><Plus className="w-4 h-4" /></button>
                        <button onClick={() => setDeleteId(stat.id)} className="text-slate-400 hover:text-red-500"><Plus className="w-4 h-4 rotate-45" /></button>
                      </div>
                    )}
                    <span className={cn(
                      "px-3 py-1 rounded-full text-sm font-bold",
                      (stat.currentUsage / stat.capacity) > 0.8 ? "bg-red-100 text-red-700" : "bg-emerald-100 text-emerald-700"
                    )}>
                      {Math.round((stat.currentUsage / stat.capacity) * 100)}% Terisi
                    </span>
                  </div>
                </div>

                <div className="relative h-4 bg-slate-100 rounded-full overflow-hidden mb-8">
                  <motion.div 
                    initial={{ width: 0 }}
                    animate={{ width: `${(stat.currentUsage / stat.capacity) * 100}%` }}
                    className={cn(
                      "h-full transition-colors duration-500",
                      (stat.currentUsage / stat.capacity) > 0.8 ? "bg-red-500" : "bg-emerald-500"
                    )}
                  />
                </div>

                <div className="flex items-center justify-between">
                  <div className="text-center">
                    <p className="text-3xl font-black text-slate-900">{stat.currentUsage}</p>
                    <p className="text-xs text-slate-400 uppercase font-bold tracking-widest mt-1">Sekarang</p>
                  </div>
                  <div className="flex gap-4">
                    <button 
                      onClick={() => updateUsage(stat.id, stat.currentUsage, stat.capacity, -1)}
                      className="w-12 h-12 rounded-2xl bg-slate-100 flex items-center justify-center text-2xl font-bold text-slate-600 hover:bg-slate-200 transition-all"
                    >
                      -
                    </button>
                    <button 
                      onClick={() => updateUsage(stat.id, stat.currentUsage, stat.capacity, 1)}
                      className="w-12 h-12 rounded-2xl bg-emerald-600 flex items-center justify-center text-2xl font-bold text-white hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-100"
                    >
                      +
                    </button>
                  </div>
                  <div className="text-center">
                    <p className="text-3xl font-black text-slate-200">{stat.capacity}</p>
                    <p className="text-xs text-slate-400 uppercase font-bold tracking-widest mt-1">Kapasitas</p>
                  </div>
                </div>
              </div>
            ))}
          </div>

          <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden">
            <Pagination
              currentPage={currentPage}
              totalPages={totalPages}
              totalItems={stats.length}
              pageSize={pageSize}
              onPageChange={setCurrentPage}
              onPageSizeChange={setPageSize}
              pageSizeOptions={[4, 6, 12]}
              className="px-6 py-4"
            />
          </div>
        </div>
      )}

      <AnimatePresence>
        {showForm && (
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
            <motion.div initial={{ scale: 0.95 }} animate={{ scale: 1 }} className="bg-white rounded-3xl p-8 w-full max-w-lg shadow-2xl">
              <h3 className="text-xl font-bold mb-6">{editingStat ? 'Edit Fasilitas' : 'Tambah Fasilitas'}</h3>
              <form onSubmit={handleSubmit} className="space-y-4">
                <div>
                  <label className="block text-sm font-medium mb-1">Nama Fasilitas</label>
                  <input name="name" defaultValue={editingStat?.name} required className="w-full px-4 py-2 rounded-xl border outline-none" />
                </div>
                <div>
                  <label className="block text-sm font-medium mb-1">Kapasitas Maksimal</label>
                  <input type="number" name="capacity" defaultValue={editingStat?.capacity} required className="w-full px-4 py-2 rounded-xl border outline-none" />
                </div>
                <div className="flex gap-3 pt-4">
                  <button 
                    type="button" 
                    onClick={closeForm} 
                    disabled={isSaving}
                    className="flex-1 px-6 py-2.5 rounded-xl border font-semibold hover:bg-slate-50 transition-all disabled:opacity-50"
                  >
                    Batal
                  </button>
                  <button 
                    type="submit" 
                    disabled={isSaving}
                    className="flex-1 px-6 py-2.5 rounded-xl bg-emerald-600 text-white font-semibold hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-100 disabled:opacity-60 disabled:cursor-not-allowed flex items-center justify-center gap-2"
                  >
                    {isSaving ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        <span>Menyimpan...</span>
                      </>
                    ) : (
                      'Simpan'
                    )}
                  </button>
                </div>
              </form>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <DeleteConfirmation 
        isOpen={!!deleteId}
        onClose={() => setDeleteId(null)}
        onConfirm={async () => {
          if (!deleteId) return;
          setIsDeleting(true);
          try {
            await deleteData(null, accessToken, spreadsheetId, 'facility_stats', deleteId);
            showToast('Fasilitas berhasil dihapus!', 'success');
            setDeleteId(null);
          } catch (err: any) {
            showToast(`Gagal menghapus: ${err.message}`, 'error');
          } finally {
            setIsDeleting(false);
          }
        }}
        loading={isDeleting}
        title="Hapus Fasilitas"
        message="Hapus fasilitas ini dari sistem pemantauan?"
      />
    </div>
  );
}

function AttendanceReportView({ profile }: { profile: UserProfile }) {
  const { accessToken, spreadsheetId } = useFirebase();
  const { showToast } = useToast();
  const [selectedDay, setSelectedDay] = useState<number | 'all'>('all');
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth());
  const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());
  const [editingAttendance, setEditingAttendance] = useState<Attendance | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const [isUpdating, setIsUpdating] = useState(false);
  
  const filter = useMemo(() => {
    const constraints = [];
    if (profile.role === 'pengurus') {
      constraints.push(where('location', '==', profile.location));
    }
    return constraints;
  }, [profile.role, profile.location]);

  const { data: attendanceData, loading } = useDataQuery<Attendance>('attendance', filter);

  const months = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];

  const filteredData = useMemo(() => {
    return attendanceData.filter(a => {
      const date = new Date(a.date);
      const matchMonth = date.getMonth() === selectedMonth;
      const matchYear = date.getFullYear() === selectedYear;
      const matchDay = selectedDay === 'all' || date.getDate() === selectedDay;
      return matchMonth && matchYear && matchDay;
    }).sort((a, b) => b.date - a.date);
  }, [attendanceData, selectedDay, selectedMonth, selectedYear]);

  const daysInMonth = useMemo(() => {
    return new Date(selectedYear, selectedMonth + 1, 0).getDate();
  }, [selectedMonth, selectedYear]);

  const years = useMemo(() => {
    const currentYear = new Date().getFullYear();
    const result = [];
    for (let i = currentYear - 2; i <= currentYear + 1; i++) {
      result.push(i);
    }
    return result;
  }, []);

  const stats = useMemo(() => {
    const total = filteredData.length;
    const hadir = filteredData.filter(a => a.status === 'hadir' || !a.status).length;
    const izin = filteredData.filter(a => a.status === 'izin').length;
    const confirmed = filteredData.filter(a => a.isConfirmed).length;
    const byCategory = filteredData.reduce((acc, curr) => {
      acc[curr.category] = (acc[curr.category] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
    const bySession = filteredData.reduce((acc, curr) => {
      acc[curr.sessionType] = (acc[curr.sessionType] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
    return { total, hadir, izin, confirmed, byCategory, bySession };
  }, [filteredData]);

  const handleExportExcel = () => {
    const data = filteredData.map(a => ({
      'Hari': a.day,
      'Tanggal': new Date(a.date).toLocaleDateString('id-ID'),
      'Jam': new Date(a.date).toLocaleTimeString('id-ID'),
      'Nama Jamaah': a.jamaahName,
      'Kategori': a.category,
      'Sesi': a.sessionType,
      'Status': a.status || 'hadir',
      'Keterangan': a.reason || '-',
      'Lokasi': a.location,
      'Terkonfirmasi': a.isConfirmed ? 'Ya' : 'Belum'
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Laporan Absensi");
    XLSX.writeFile(wb, `Laporan_Absensi_${months[selectedMonth]}_${selectedYear}.xlsx`);
  };

  const handleConfirm = async (id: string, current: boolean) => {
    try {
      await saveData(null, accessToken, spreadsheetId, 'attendance', {
        isConfirmed: !current
      }, id);
      showToast(!current ? 'Kehadiran dikonfirmasi!' : 'Konfirmasi dibatalkan', 'success');
    } catch (err: any) {
      showToast(`Gagal mengonfirmasi: ${err.message}`, 'error');
    }
  };

  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(20);

  useEffect(() => {
    setCurrentPage(1);
  }, [selectedDay, selectedMonth, selectedYear]);

  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredData.slice(start, start + pageSize);
  }, [filteredData, currentPage, pageSize]);

  const groupedData = useMemo(() => {
    const groups: Record<string, Attendance[]> = {};
    paginatedData.forEach(a => {
      const dateKey = new Date(a.date).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
      if (!groups[dateKey]) groups[dateKey] = [];
      groups[dateKey].push(a);
    });
    return groups;
  }, [paginatedData]);

  return (
    <div className="space-y-8">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
        <div>
          <h2 className="text-2xl font-black text-slate-900 uppercase tracking-tight">Laporan Absensi</h2>
          <p className="text-slate-500 mt-1">Data kehadiran jamaah per bulan</p>
        </div>
        
        <div className="flex flex-wrap items-center gap-3">
          <div className="flex items-center gap-2 bg-white p-2 rounded-2xl border border-slate-100 shadow-sm">
            <div className="flex items-center gap-1">
              <select 
                value={selectedDay} 
                onChange={(e) => setSelectedDay(e.target.value === 'all' ? 'all' : parseInt(e.target.value))}
                className="bg-transparent border-none outline-none font-bold text-slate-700 text-sm px-2 cursor-pointer focus:ring-0"
              >
                <option value="all">Semua Tgl</option>
                {Array.from({ length: daysInMonth }, (_, i) => i + 1).map(d => (
                  <option key={d} value={d}>{d}</option>
                ))}
              </select>
              <div className="w-px h-4 bg-slate-100 mx-1" />
              <select 
                value={selectedMonth} 
                onChange={(e) => setSelectedMonth(parseInt(e.target.value))}
                className="bg-transparent border-none outline-none font-bold text-slate-700 text-sm px-2 cursor-pointer focus:ring-0"
              >
                {months.map((m, i) => (
                  <option key={m} value={i}>{m}</option>
                ))}
              </select>
              <div className="w-px h-4 bg-slate-100 mx-1" />
              <select 
                value={selectedYear} 
                onChange={(e) => setSelectedYear(parseInt(e.target.value))}
                className="bg-transparent border-none outline-none font-bold text-slate-700 text-sm px-2 cursor-pointer focus:ring-0"
              >
                {years.map(y => (
                  <option key={y} value={y}>{y}</option>
                ))}
              </select>
            </div>
          </div>
          
          <button 
            onClick={handleExportExcel}
            className="px-4 py-2.5 bg-emerald-600 text-white rounded-2xl text-[10px] font-black uppercase tracking-widest flex items-center gap-2 hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-100"
          >
            <ClipboardList className="w-4 h-4" />
            Excel
          </button>
        </div>
    </div>

      <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
        <div className="bg-emerald-600 p-6 rounded-3xl text-white shadow-xl shadow-emerald-100">
          <p className="text-emerald-100 text-xs font-black uppercase tracking-widest mb-1">Total Hadir</p>
          <h3 className="text-4xl font-black">{stats.hadir}</h3>
          <p className="text-emerald-200/60 text-[10px] mt-4 font-bold uppercase tracking-wider italic">
            {selectedDay === 'all' ? `Bulan ${months[selectedMonth]}` : `${selectedDay} ${months[selectedMonth]} ${selectedYear}`}
          </p>
        </div>

        <div className="bg-amber-500 p-6 rounded-3xl text-white shadow-xl shadow-amber-100">
          <p className="text-amber-100 text-xs font-black uppercase tracking-widest mb-1">Total Izin</p>
          <h3 className="text-4xl font-black">{stats.izin}</h3>
          <p className="text-amber-200/60 text-[10px] mt-4 font-bold uppercase tracking-wider italic">
            {selectedDay === 'all' ? `Bulan ${months[selectedMonth]}` : `${selectedDay} ${months[selectedMonth]} ${selectedYear}`}
          </p>
        </div>
        
        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
          <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-4">Berdasarkan Kategori</p>
          <div className="flex flex-wrap gap-2">
            {Object.entries(stats.byCategory).map(([cat, count], catIdx) => (
              <div key={`stat-cat-${cat}-${catIdx}`} className="px-3 py-2 bg-slate-50 rounded-xl flex items-center gap-3">
                <span className="text-xs font-black text-slate-900">{cat}</span>
                <span className="w-6 h-6 rounded-lg bg-emerald-100 text-emerald-700 text-[10px] flex items-center justify-center font-black">{count}</span>
              </div>
            ))}
            {Object.keys(stats.byCategory).length === 0 && <p className="text-slate-300 text-xs italic">Tidak ada data</p>}
          </div>
        </div>

        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
          <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mb-4">Berdasarkan Sesi</p>
          <div className="flex flex-wrap gap-2">
            {Object.entries(stats.bySession).map(([sess, count], sessIdx) => (
              <div key={`stat-sess-${sess}-${sessIdx}`} className="px-3 py-2 bg-slate-50 rounded-xl flex items-center gap-3">
                <span className="text-xs font-black text-slate-900">{sess}</span>
                <span className="w-6 h-6 rounded-lg bg-blue-100 text-blue-700 text-[10px] flex items-center justify-center font-black">{count}</span>
              </div>
            ))}
            {Object.keys(stats.bySession).length === 0 && <p className="text-slate-300 text-xs italic">Tidak ada data</p>}
          </div>
        </div>
      </div>

      <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50/50">
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Tanggal</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Nama Jamaah</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Kategori</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Sesi</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Status</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Keterangan</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Konfirmasi</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Lokasi</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {loading ? (
                <tr>
                  <td colSpan={10} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center justify-center gap-2">
                      <Loader2 className="w-8 h-8 text-emerald-600 animate-spin" />
                      <p className="text-xs font-semibold text-slate-500">Memuat laporan absensi...</p>
                    </div>
                  </td>
                </tr>
              ) : Object.keys(groupedData).length > 0 ? (
                Object.entries(groupedData).map(([date, items]: [string, Attendance[]], grpIdx) => (
                  <React.Fragment key={`group-${date}-${grpIdx}`}>
                    <tr className="bg-slate-50/30">
                      <td colSpan={10} className="px-6 py-3 text-[10px] font-black text-emerald-600 uppercase tracking-widest bg-emerald-50/30">
                        {date} ({items.length} Absensi)
                      </td>
                    </tr>
                    {items.map((a, aIdx) => (
                      <tr key={a.id ? `${a.id}-${aIdx}` : `att-row-${grpIdx}-${aIdx}`} className="hover:bg-slate-50 transition-colors group">
                        <td className="px-6 py-4">
                          <div className="flex items-center gap-3">
                            <div className="w-10 h-10 rounded-xl bg-slate-100 flex flex-col items-center justify-center group-hover:bg-white transition-colors">
                              <span className="text-[10px] font-black text-slate-400 uppercase leading-none">{new Date(a.date).toLocaleDateString('id-ID', { weekday: 'short' })}</span>
                              <span className="text-sm font-black text-slate-900 leading-none mt-0.5">{new Date(a.date).getDate()}</span>
                            </div>
                            <div className="text-[10px] text-slate-400 font-bold">
                              {new Date(a.date).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })}
                            </div>
                          </div>
                        </td>
                        <td className="px-6 py-4">
                          <p className="font-black text-slate-900 text-sm">{a.jamaahName}</p>
                        </td>
                        <td className="px-6 py-4">
                          <span className="px-2 py-1 rounded-lg bg-emerald-50 text-emerald-700 text-[9px] font-black uppercase tracking-wider">{a.category}</span>
                        </td>
                        <td className="px-6 py-4">
                          <span className={cn(
                            "px-2 py-1 rounded-lg text-[9px] font-black uppercase tracking-wider",
                            a.sessionType === 'Kelompok' ? "bg-blue-50 text-blue-700" : 
                            a.sessionType === 'Desa' ? "bg-purple-50 text-purple-700" : "bg-amber-50 text-amber-700"
                          )}>{a.sessionType}</span>
                        </td>
                        <td className="px-6 py-4">
                          <span className={cn(
                            "px-2 py-1 rounded-lg text-[9px] font-black uppercase tracking-wider",
                            a.status === 'izin' ? "bg-amber-100 text-amber-700" : "bg-emerald-100 text-emerald-700"
                          )}>{a.status || 'hadir'}</span>
                        </td>
                        <td className="px-6 py-4">
                          <p className="text-xs text-slate-500 font-medium max-w-[200px] truncate" title={a.reason}>{a.reason || '-'}</p>
                        </td>
                        <td className="px-6 py-4">
                          <button 
                            onClick={() => handleConfirm(a.id, !!a.isConfirmed)}
                            disabled={profile.role !== 'pengurus' && profile.role !== 'admin'}
                            className={cn(
                              "px-2 py-1 rounded-lg text-[9px] font-black uppercase tracking-wider transition-all",
                              a.isConfirmed 
                                ? "bg-emerald-100 text-emerald-700" 
                                : "bg-slate-100 text-slate-400 hover:bg-emerald-50 hover:text-emerald-600"
                            )}
                          >
                            {a.isConfirmed ? 'Datang' : 'Konfirmasi'}
                          </button>
                        </td>
                        <td className="px-6 py-4">
                          <div className="flex items-center gap-2 text-slate-500 text-xs font-bold">
                            <MapPin className="w-3 h-3 text-slate-300" />
                            {a.location}
                          </div>
                        </td>
                        <td className="px-6 py-4 text-right">
                          <div className="flex gap-3 justify-end">
                            <button onClick={() => setEditingAttendance(a)} className="text-emerald-600 hover:text-emerald-700 font-bold text-xs uppercase tracking-wider">Edit</button>
                            <button onClick={() => setDeleteId(a.id)} className="text-red-500 hover:text-red-600 font-bold text-xs uppercase tracking-wider">Hapus</button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </React.Fragment>
                ))
              ) : (
                <tr>
                  <td colSpan={10} className="px-6 py-20 text-center">
                    <div className="max-w-[200px] mx-auto opacity-20 mb-4 grayscale">
                      <ClipboardList className="w-12 h-12 mx-auto text-slate-900" />
                    </div>
                    <p className="text-sm font-bold text-slate-400">Belum ada data absensi untuk periode ini</p>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        <Pagination
          currentPage={currentPage}
          totalPages={totalPages}
          totalItems={filteredData.length}
          pageSize={pageSize}
          onPageChange={setCurrentPage}
          onPageSizeChange={setPageSize}
          pageSizeOptions={[10, 20, 50, 100]}
          className="px-6 py-4 bg-white border-t border-slate-100"
        />
      </div>

      <AnimatePresence>
        {editingAttendance && (
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
            <motion.div initial={{ scale: 0.95 }} animate={{ scale: 1 }} className="bg-white rounded-3xl p-8 w-full max-w-lg shadow-2xl">
              <h3 className="text-xl font-bold mb-6">Edit Data Absensi</h3>
              <form 
                onSubmit={async (e) => {
                  e.preventDefault();
                  if (!editingAttendance) return;
                  setIsUpdating(true);
                  const formData = new FormData(e.currentTarget);
                  try {
                    await saveData(null, accessToken, spreadsheetId, 'attendance', {
                      ...editingAttendance,
                      sessionType: formData.get('sessionType'),
                      status: formData.get('status'),
                      reason: formData.get('reason'),
                      updatedAt: Date.now()
                    }, editingAttendance.id);
                    showToast('Data absensi berhasil diperbarui!', 'success');
                    setEditingAttendance(null);
                  } catch (err: any) {
                    showToast(`Gagal mengupdate: ${err.message}`, 'error');
                  } finally {
                    setIsUpdating(false);
                  }
                }} 
                className="space-y-4"
              >
                <div>
                  <label className="block text-sm font-medium mb-1 uppercase text-[10px] tracking-widest text-slate-400">Nama Jamaah</label>
                  <p className="font-black text-slate-900">{editingAttendance.jamaahName}</p>
                </div>
                <div>
                  <label className="block text-sm font-medium mb-1 uppercase text-[10px] tracking-widest text-slate-400">Sesi</label>
                  <select name="sessionType" defaultValue={editingAttendance.sessionType} className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 transition-all font-bold text-slate-700">
                    <option value="Kelompok">Kelompok</option>
                    <option value="Desa">Desa</option>
                    <option value="Acara">Acara</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-medium mb-1 uppercase text-[10px] tracking-widest text-slate-400">Status</label>
                  <select name="status" defaultValue={editingAttendance.status || 'hadir'} className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 transition-all font-bold text-slate-700">
                    <option value="hadir">Hadir</option>
                    <option value="izin">Izin</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-medium mb-1 uppercase text-[10px] tracking-widest text-slate-400">Keterangan / Alasan</label>
                  <textarea name="reason" defaultValue={editingAttendance.reason} className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 transition-all font-bold text-slate-700 min-h-[100px]" />
                </div>
                <div className="flex gap-3 pt-4">
                  <button type="button" onClick={() => setEditingAttendance(null)} className="flex-1 px-6 py-3 rounded-xl border border-slate-200 font-bold uppercase tracking-widest text-xs text-slate-500">Batal</button>
                  <button type="submit" disabled={isUpdating} className="flex-1 px-6 py-3 rounded-xl bg-emerald-600 text-white font-bold uppercase tracking-widest text-xs shadow-lg shadow-emerald-100 disabled:opacity-50">
                    {isUpdating ? 'Menyimpan...' : 'Simpan Perubahan'}
                  </button>
                </div>
              </form>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <DeleteConfirmation 
        isOpen={!!deleteId}
        onClose={() => setDeleteId(null)}
        onConfirm={async () => {
          if (!deleteId) return;
          setIsDeleting(true);
          try {
            await deleteData(null, accessToken, spreadsheetId, 'attendance', deleteId);
            showToast('Data absensi berhasil dihapus!', 'success');
            setDeleteId(null);
          } catch (err: any) {
            showToast(`Gagal menghapus: ${err.message}`, 'error');
          } finally {
            setIsDeleting(false);
          }
        }}
        loading={isDeleting}
        title="Hapus Data Absensi"
        message="Hapus catatan kehadiran ini secara permanen?"
      />
    </div>
  );
}

function UBShoppingView({ profile }: { profile: UserProfile }) {
  const { accessToken, spreadsheetId } = useFirebase();
  const { showToast } = useToast();
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth());
  const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());
  const [selectedLocation, setSelectedLocation] = useState<MosqueLocation | 'all'>(profile.role === 'pengurus' ? profile.location! : 'all');
  
  const shoppingFilter = useMemo(() => {
    const filters = [
      where('month', '==', selectedMonth),
      where('year', '==', selectedYear)
    ];
    if (selectedLocation !== 'all') {
      filters.push(where('location', '==', selectedLocation));
    }
    return filters;
  }, [selectedMonth, selectedYear, selectedLocation]);

  const { data: shoppingRecords, loading: loadingShopping } = useDataQuery<UBShopping>('ub_shopping', shoppingFilter);
  
  const jamaahFilter = useMemo(() => {
    if (selectedLocation !== 'all') {
      return [where('location', '==', selectedLocation)];
    }
    return [];
  }, [selectedLocation]);
  
  const { data: allJamaah, loading: loadingJamaah } = useDataQuery<Jamaah>('jamaah', jamaahFilter);
  
  const [showForm, setShowForm] = useState(false);
  const [editingShopping, setEditingShopping] = useState<UBShopping | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isUpdating, setIsUpdating] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);

  // Searchable Select State
  const [jamaahSearch, setJamaahSearch] = useState('');
  const [selectedJamaahId, setSelectedJamaahId] = useState('');
  const [isSearchOpen, setIsSearchOpen] = useState(false);

  const filteredJamaah = useMemo(() => {
    if (!jamaahSearch) return allJamaah.slice(0, 10);
    const lowerSearch = jamaahSearch.toLowerCase();
    return allJamaah.filter(j => 
      j.name.toLowerCase().includes(lowerSearch) || 
      j.memberId.toLowerCase().includes(lowerSearch)
    ).slice(0, 10);
  }, [allJamaah, jamaahSearch]);

  const months = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];

  const years = Array.from({ length: 5 }, (_, i) => new Date().getFullYear() - 2 + i);

  const shoppingByKK = useMemo(() => {
    const map: Record<string, number> = {};
    const kkDocIdMap: Record<string, string> = {}; // memberId -> docId
    
    allJamaah.forEach(j => {
      if (j.isKK) kkDocIdMap[j.memberId] = j.id;
    });

    shoppingRecords.forEach(rec => {
      const jamaah = allJamaah.find(j => j.id === rec.jamaahId);
      if (jamaah) {
        let targetKKDocId = jamaah.id;
        if (!jamaah.isKK && jamaah.kkId) {
          targetKKDocId = kkDocIdMap[jamaah.kkId] || jamaah.id;
        }
        map[targetKKDocId] = (map[targetKKDocId] || 0) + rec.amount;
      }
    });
    return map;
  }, [shoppingRecords, allJamaah]);

  const stats = useMemo(() => {
    const kks = allJamaah.filter(j => j.isKK);
    const totalKK = kks.length;
    const metTarget = kks.filter(j => (shoppingByKK[j.id] || 0) >= 100000).length;
    const totalAmount = shoppingRecords.reduce((sum, rec) => sum + rec.amount, 0);
    
    return {
      totalJamaah: totalKK,
      metTarget,
      notMetTarget: totalKK - metTarget,
      totalAmount
    };
  }, [allJamaah, shoppingByKK, shoppingRecords]);

  const kkList = useMemo(() => allJamaah.filter(j => j.isKK), [allJamaah]);
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);

  useEffect(() => {
    setCurrentPage(1);
  }, [selectedMonth, selectedYear, selectedLocation]);

  const totalPages = Math.ceil(kkList.length / pageSize) || 1;
  const paginatedKKList = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return kkList.slice(start, start + pageSize);
  }, [kkList, currentPage, pageSize]);

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    setIsUpdating(true);
    
    const formData = new FormData(e.currentTarget);
    const jamaahId = selectedJamaahId;
    const jamaah = allJamaah.find(j => j.id === jamaahId);
    
    if (!jamaah) return;

    const amount = Number(formData.get('amount'));
    const date = new Date(formData.get('date') as string).getTime();
    const d = new Date(date);

    const data: Partial<UBShopping> = {
      jamaahId: jamaah.id,
      jamaahName: jamaah.name,
      location: jamaah.location,
      amount,
      date,
      month: d.getMonth(),
      year: d.getFullYear(),
      note: formData.get('note') as string,
      createdAt: editingShopping ? editingShopping.createdAt : Date.now()
    };

    try {
      await saveData(null, accessToken, spreadsheetId, 'ub_shopping', data, editingShopping?.id);
      showToast('Data belanja UB berhasil disimpan!', 'success');
      setShowForm(false);
      setEditingShopping(null);
    } catch (error: any) {
      showToast(`Gagal menyimpan data: ${error.message}`, 'error');
    } finally {
      setIsUpdating(false);
    }
  };

  const handleExportExcel = () => {
    const kks = allJamaah.filter(j => j.isKK);
    const data = kks.map(j => ({
      'ID Jamaah (KK)': j.memberId,
      'Nama Kepala Keluarga': j.name,
      'Lokasi': j.location,
      'Total Belanja Keluarga': shoppingByKK[j.id] || 0,
      'Status': (shoppingByKK[j.id] || 0) >= 100000 ? 'TERPENUHI' : 'BELUM TERPENUHI',
      'Kurang': Math.max(0, 100000 - (shoppingByKK[j.id] || 0))
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Laporan Belanja UB");
    XLSX.writeFile(wb, `Laporan_Belanja_UB_${months[selectedMonth]}_${selectedYear}.xlsx`);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6">
        <div>
          <h2 className="text-3xl font-black text-slate-900 tracking-tight">Laporan Belanja UB</h2>
          <p className="text-slate-400 text-xs font-bold uppercase tracking-[0.2em] mt-1">Target Minimal: Rp 100.000 / Bulan</p>
        </div>

        <div className="flex flex-wrap items-center gap-3 bg-white p-2 rounded-[2rem] border border-slate-100 shadow-sm">
          <div className="flex items-center gap-2 px-4 py-2 bg-slate-50 rounded-2xl border border-slate-100">
            <Calendar className="w-4 h-4 text-emerald-600" />
            <select 
              value={selectedMonth} 
              onChange={(e) => setSelectedMonth(Number(e.target.value))}
              className="bg-transparent text-[10px] font-black uppercase tracking-widest outline-none border-none cursor-pointer text-slate-700"
            >
              {months.map((m, i) => <option key={m} value={i}>{m}</option>)}
            </select>
            <select 
              value={selectedYear} 
              onChange={(e) => setSelectedYear(Number(e.target.value))}
              className="bg-transparent text-[10px] font-black uppercase tracking-widest outline-none border-none cursor-pointer text-slate-700"
            >
              {years.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
          </div>

          {profile.role === 'admin' && (
            <div className="flex items-center gap-2 px-4 py-2 bg-slate-50 rounded-2xl border border-slate-100">
              <MapPin className="w-4 h-4 text-emerald-600" />
              <select 
                value={selectedLocation} 
                onChange={(e) => setSelectedLocation(e.target.value as any)}
                className="bg-transparent text-[10px] font-black uppercase tracking-widest outline-none border-none cursor-pointer text-slate-700"
              >
                <option value="all">Semua Lokasi</option>
                {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(l => <option key={l} value={l}>{l}</option>)}
              </select>
            </div>
          )}

          <button 
            onClick={handleExportExcel}
            className="p-2.5 bg-emerald-100 text-emerald-700 rounded-2xl hover:bg-emerald-200 transition-all"
            title="Export Excel"
          >
            <ClipboardList className="w-5 h-5" />
          </button>

          <button 
            onClick={() => { 
              setEditingShopping(null); 
              setSelectedJamaahId('');
              setJamaahSearch('');
              setShowForm(true); 
            }}
            className="px-6 py-2.5 bg-emerald-600 text-white rounded-2xl text-[10px] font-black uppercase tracking-widest flex items-center gap-2 hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-100"
          >
            <Plus className="w-4 h-4" />
            Input Belanja
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
        <div className="bg-emerald-600 p-6 rounded-3xl text-white shadow-xl shadow-emerald-100">
          <p className="text-emerald-100 text-xs font-black uppercase tracking-widest mb-1">Total Belanja</p>
          <h3 className="text-3xl font-black italic">Rp {stats.totalAmount.toLocaleString('id-ID')}</h3>
          <p className="text-emerald-200/60 text-[10px] mt-4 font-bold uppercase tracking-wider">
            Bulan {months[selectedMonth]} {selectedYear}
          </p>
        </div>

        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
          <p className="text-slate-400 text-xs font-black uppercase tracking-widest mb-1 text-[10px]">Terpenuhi (≥ 100k)</p>
          <h3 className="text-3xl font-black text-emerald-600">{stats.metTarget}</h3>
          <div className="w-full bg-slate-50 h-2 rounded-full mt-4 overflow-hidden border border-slate-100">
            <div 
              className="h-full bg-emerald-500 rounded-full" 
              style={{ width: `${(stats.metTarget / (stats.totalJamaah || 1)) * 100}%` }} 
            />
          </div>
        </div>

        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
          <p className="text-slate-400 text-xs font-black uppercase tracking-widest mb-1 text-[10px]">Belum Terpenuhi</p>
          <h3 className="text-3xl font-black text-rose-500">{stats.notMetTarget}</h3>
          <div className="w-full bg-slate-50 h-2 rounded-full mt-4 overflow-hidden border border-slate-100">
            <div 
              className="h-full bg-rose-500 rounded-full" 
              style={{ width: `${(stats.notMetTarget / (stats.totalJamaah || 1)) * 100}%` }} 
            />
          </div>
        </div>

        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
          <p className="text-slate-400 text-xs font-black uppercase tracking-widest mb-1 text-[10px]">Total Kepala Keluarga</p>
          <h3 className="text-3xl font-black text-slate-900">{stats.totalJamaah}</h3>
          <p className="text-slate-300 text-[10px] mt-4 font-bold uppercase tracking-wider italic">
            {selectedLocation === 'all' ? 'Seluruh Lokasi' : selectedLocation}
          </p>
        </div>
      </div>

      <div className="bg-white rounded-[2rem] border border-slate-100 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50/50">
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">ID / Nama Jamaah</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Lokasi</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Total Belanja</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Status</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest">Kurang</th>
                <th className="px-6 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right">Rincian</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {loadingJamaah || loadingShopping ? (
                <tr>
                  <td colSpan={6} className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center justify-center gap-2">
                      <Loader2 className="w-8 h-8 text-emerald-600 animate-spin" />
                      <p className="text-xs font-semibold text-slate-500">Memuat data belanja UB...</p>
                    </div>
                  </td>
                </tr>
              ) : kkList.length === 0 ? (
                <tr>
                  <td colSpan={6} className="px-6 py-12 text-center text-slate-400 font-medium">
                    Tidak ada data kepala keluarga untuk lokasi ini
                  </td>
                </tr>
              ) : paginatedKKList.map((j, jIdx) => {
                const total = shoppingByKK[j.id] || 0;
                const isMet = total >= 100000;
                const remaining = Math.max(0, 100000 - total);
                
                // Get all family members' IDs including KK
                const familyIds = allJamaah.filter(member => member.id === j.id || member.kkId === j.memberId).map(m => m.id);
                const familyRecords = shoppingRecords.filter(r => familyIds.includes(r.jamaahId));
                
                return (
                  <tr key={j.id ? `${j.id}-${jIdx}` : `kk-row-${jIdx}`} className="hover:bg-slate-50/50 transition-colors">
                    <td className="px-6 py-4">
                      <div className="flex flex-col">
                        <span className="text-[10px] font-black text-emerald-600 uppercase tracking-widest">{j.memberId}</span>
                        <span className="text-sm font-black text-slate-900">{j.name}</span>
                        <span className="text-[10px] text-slate-400 font-medium">Kepala Keluarga</span>
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{j.location}</span>
                    </td>
                    <td className="px-6 py-4">
                      <span className={cn("text-sm font-black italic", isMet ? "text-emerald-600" : "text-slate-900")}>
                        Rp {total.toLocaleString('id-ID')}
                      </span>
                    </td>
                    <td className="px-6 py-4">
                      <span className={cn(
                        "px-2.5 py-1 rounded-lg text-[9px] font-black uppercase tracking-widest",
                        isMet ? "bg-emerald-100 text-emerald-700" : "bg-rose-100 text-rose-700"
                      )}>
                        {isMet ? 'TERPENUHI' : 'BELUM TERPENUHI'}
                      </span>
                    </td>
                    <td className="px-6 py-4">
                      <span className={cn("text-xs font-bold", remaining > 0 ? "text-rose-500" : "text-slate-300")}>
                        {remaining > 0 ? `Rp ${remaining.toLocaleString('id-ID')}` : '-'}
                      </span>
                    </td>
                    <td className="px-6 py-4 text-right">
                      <div className="flex justify-end gap-2 flex-wrap max-w-[150px]">
                        {familyRecords.map((r, rIdx) => {
                          const recordJamaah = allJamaah.find(jm => jm.id === r.jamaahId);
                          return (
                            <button 
                              key={r.id ? `${r.id}-${rIdx}` : `fam-rec-${jIdx}-${rIdx}`}
                              onClick={() => { 
                                setEditingShopping(r); 
                                setSelectedJamaahId(r.jamaahId);
                                setJamaahSearch(`${recordJamaah?.memberId} - ${recordJamaah?.name}`);
                                setShowForm(true); 
                              }}
                              className="p-1.5 hover:bg-slate-100 rounded-lg transition-all"
                              title={`${recordJamaah?.name}: Rp ${r.amount.toLocaleString()}`}
                            >
                              <div className={cn(
                                "w-2 h-2 rounded-full",
                                recordJamaah?.isKK ? "bg-emerald-500" : "bg-blue-400"
                              )} />
                            </button>
                          );
                        })}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <Pagination
          currentPage={currentPage}
          totalPages={totalPages}
          totalItems={kkList.length}
          pageSize={pageSize}
          onPageChange={setCurrentPage}
          onPageSizeChange={setPageSize}
          pageSizeOptions={[10, 20, 50]}
          className="px-6 py-4 bg-white border-t border-slate-100"
        />
      </div>

      <AnimatePresence>
        {showForm && (
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
            <motion.div initial={{ scale: 0.95 }} animate={{ scale: 1 }} className="bg-white rounded-[2.5rem] p-8 w-full max-w-md shadow-2xl max-h-[90vh] overflow-y-auto">
              <div className="flex items-center justify-between mb-8">
                <div>
                  <h3 className="text-xl font-black text-slate-900">{editingShopping ? 'Edit Data Belanja' : 'Input Data Belanja'}</h3>
                  <p className="text-slate-400 text-[10px] font-black uppercase tracking-widest mt-1">Usaha Bersama (UB)</p>
                </div>
                <button onClick={() => setShowForm(false)} className="p-2 hover:bg-slate-100 rounded-2xl transition-all">
                  <X className="w-5 h-5 text-slate-400" />
                </button>
              </div>

              <form onSubmit={handleSubmit} className="space-y-6">
                <div>
                  <label className="block text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2">Pilih Jamaah</label>
                  <div className="relative">
                    <div className="relative">
                      <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                      <input 
                        type="text"
                        placeholder="Cari Nama atau ID Jamaah..."
                        value={jamaahSearch}
                        onChange={(e) => {
                          setJamaahSearch(e.target.value);
                          if (!e.target.value) setSelectedJamaahId('');
                          setIsSearchOpen(true);
                        }}
                        onFocus={() => setIsSearchOpen(true)}
                        onBlur={() => setTimeout(() => setIsSearchOpen(false), 200)}
                        className="w-full pl-11 pr-5 py-4 rounded-[1.5rem] border border-slate-100 bg-slate-50 focus:bg-white outline-none focus:ring-2 focus:ring-emerald-500 transition-all font-bold text-sm text-slate-700"
                      />
                    </div>
                    
                    {isSearchOpen && (
                      <div className="absolute z-50 w-full mt-2 bg-white rounded-2xl border border-slate-100 shadow-xl overflow-hidden max-h-60 overflow-y-auto">
                        {filteredJamaah.length > 0 ? (
                          filteredJamaah.map((j, jIdx) => (
                            <button
                              key={j.id ? `${j.id}-${jIdx}` : `ub-search-j-${jIdx}`}
                              type="button"
                              onClick={() => {
                                setSelectedJamaahId(j.id);
                                setJamaahSearch(`${j.memberId} - ${j.name}`);
                                setIsSearchOpen(false);
                              }}
                              className={cn(
                                "w-full px-5 py-3 text-left hover:bg-slate-50 transition-colors flex flex-col",
                                selectedJamaahId === j.id ? "bg-emerald-50" : ""
                              )}
                            >
                              <span className="text-[10px] font-black text-emerald-600 uppercase tracking-widest">{j.memberId}</span>
                              <span className="text-sm font-bold text-slate-700">{j.name}</span>
                            </button>
                          ))
                        ) : (
                          <div className="px-5 py-4 text-center text-slate-400 text-xs italic">Jamaah tidak ditemukan</div>
                        )}
                      </div>
                    )}
                  </div>
                  {/* Invisible input to keep form valid if needed, but we use state in handleSubmit */}
                  <input type="hidden" name="jamaahId" value={selectedJamaahId} required />
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2">Nominal (Rp)</label>
                    <input 
                      type="number" 
                      name="amount" 
                      defaultValue={editingShopping?.amount} 
                      required 
                      className="w-full px-5 py-4 rounded-[1.5rem] border border-slate-100 bg-slate-50 focus:bg-white outline-none focus:ring-2 focus:ring-emerald-500 transition-all font-black italic text-sm text-slate-900" 
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2">Tanggal</label>
                    <input 
                      type="date" 
                      name="date" 
                      defaultValue={editingShopping ? new Date(editingShopping.date).toISOString().split('T')[0] : new Date().toISOString().split('T')[0]} 
                      required 
                      className="w-full px-5 py-4 rounded-[1.5rem] border border-slate-100 bg-slate-50 focus:bg-white outline-none focus:ring-2 focus:ring-emerald-500 transition-all font-bold text-xs text-slate-700" 
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2">Catatan (Opsional)</label>
                  <textarea 
                    name="note" 
                    defaultValue={editingShopping?.note} 
                    className="w-full px-5 py-4 rounded-[1.5rem] border border-slate-100 bg-slate-50 focus:bg-white outline-none focus:ring-2 focus:ring-emerald-500 transition-all font-bold text-sm text-slate-700 min-h-[100px]" 
                  />
                </div>

                <div className="flex gap-4 pt-4">
                  {editingShopping && (
                    <button 
                      type="button"
                      onClick={async () => {
                        if (!editingShopping) return;
                        setIsDeleting(true);
                        try {
                          await deleteData(null, accessToken, spreadsheetId, 'ub_shopping', editingShopping.id);
                          showToast('Data belanja UB berhasil dihapus!', 'success');
                          setShowForm(false);
                          setEditingShopping(null);
                        } catch (err: any) {
                          showToast(`Gagal menghapus: ${err.message}`, 'error');
                        } finally {
                          setIsDeleting(false);
                        }
                      }}
                      className="p-4 rounded-[1.5rem] border border-rose-100 text-rose-500 hover:bg-rose-50 transition-all"
                    >
                      <Plus className="w-5 h-5 rotate-45" />
                    </button>
                  )}
                  <button 
                    type="submit" 
                    disabled={isUpdating}
                    className="flex-1 px-8 py-4 rounded-[1.5rem] bg-emerald-600 text-white font-black uppercase tracking-widest text-xs shadow-xl shadow-emerald-100 disabled:opacity-50 hover:bg-emerald-700 transition-all"
                  >
                    {isUpdating ? 'Menyimpan...' : editingShopping ? 'Update Data' : 'Simpan Data'}
                  </button>
                </div>
              </form>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

function UsersView({ profile }: { profile?: any }) {
  const { accessToken, spreadsheetId } = useFirebase();
  const { showToast } = useToast();
  
  const effectiveSpreadsheetId = profile?.spreadsheetId || spreadsheetId || (import.meta as any).env?.VITE_SPREADSHEET_ID || localStorage.getItem('app_spreadsheet_id') || '';
  const activeToken = accessToken || localStorage.getItem('app_access_token') || null;

  const { data: users, loading } = useDataQuery<UserProfile>('users', [], effectiveSpreadsheetId);
  const [searchTerm, setSearchTerm] = useState('');
  const [roleFilter, setRoleFilter] = useState<'all' | 'admin' | 'pengurus'>('all');
  const [statusFilter, setStatusFilter] = useState<'all' | 'verified' | 'pending'>('all');
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [isSyncingSelf, setIsSyncingSelf] = useState(false);

  const [editingUser, setEditingUser] = useState<UserProfile | null>(null);
  const [showAddForm, setShowAddForm] = useState(false);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  const handleRefresh = async () => {
    setIsRefreshing(true);
    try {
      clearSheetMemoryCache('users');
      window.dispatchEvent(new CustomEvent('data_updated', { detail: { collectionName: 'users' } }));
      showToast('Memuat ulang data pengguna langsung dari Google Sheets...', 'info');
    } finally {
      setTimeout(() => setIsRefreshing(false), 600);
    }
  };

  const isCurrentProfileInUsers = useMemo(() => {
    if (!profile?.uid && !profile?.email) return true;
    return users.some(u => 
      (profile.uid && (u.uid === profile.uid || u.id === profile.uid)) ||
      (profile.email && u.email && u.email.toLowerCase() === profile.email.toLowerCase())
    );
  }, [users, profile]);

  const handleSyncMyProfile = async () => {
    if (!profile) return;
    setIsSyncingSelf(true);
    try {
      const myId = profile.uid || 'usr_' + Math.random().toString(36).substring(2, 9);
      const userObj = {
        uid: myId,
        id: myId,
        displayName: profile.displayName || 'Admin',
        email: profile.email || '',
        role: profile.role || 'admin',
        location: profile.location || null,
        isVerified: true,
        createdAt: profile.createdAt || Date.now(),
        spreadsheetId: effectiveSpreadsheetId
      };
      await saveData(null, activeToken, effectiveSpreadsheetId, 'users', userObj, myId);
      clearSheetMemoryCache('users');
      window.dispatchEvent(new CustomEvent('data_updated', { detail: { collectionName: 'users' } }));
      showToast('Akun Anda berhasil disimpan ke sheet "users" di Google Spreadsheet!', 'success');
    } catch (err: any) {
      showToast(`Gagal menyelaraskan akun: ${err.message}`, 'error');
    } finally {
      setIsSyncingSelf(false);
    }
  };

  const filteredUsers = useMemo(() => {
    return users.filter(u => {
      const isVerified = u.isVerified === true || (u.isVerified as any) === 'true' || (u.isVerified as any) === 'TRUE';
      if (statusFilter === 'verified' && !isVerified) return false;
      if (statusFilter === 'pending' && isVerified) return false;
      if (roleFilter !== 'all' && u.role !== roleFilter) return false;
      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase();
        const nameMatch = (u.displayName || '').toLowerCase().includes(q);
        const emailMatch = (u.email || '').toLowerCase().includes(q);
        const locMatch = (u.location || '').toLowerCase().includes(q);
        const roleMatch = (u.role || '').toLowerCase().includes(q);
        return nameMatch || emailMatch || locMatch || roleMatch;
      }
      return true;
    });
  }, [users, statusFilter, roleFilter, searchTerm]);

  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);

  useEffect(() => {
    setCurrentPage(1);
  }, [statusFilter, roleFilter, searchTerm]);

  const totalPages = Math.ceil(filteredUsers.length / pageSize) || 1;
  const paginatedUsers = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredUsers.slice(start, start + pageSize);
  }, [filteredUsers, currentPage, pageSize]);

  const stats = useMemo(() => {
    const total = users.length;
    const verified = users.filter(u => u.isVerified === true || (u.isVerified as any) === 'true' || (u.isVerified as any) === 'TRUE').length;
    const pending = total - verified;
    const admins = users.filter(u => u.role === 'admin').length;
    const pengurus = users.filter(u => u.role === 'pengurus').length;
    return { total, verified, pending, admins, pengurus };
  }, [users]);

  const toggleVerify = async (userId: string, current: boolean) => {
    const targetUser = users.find(u => (u.uid || u.id) === userId);
    if (!targetUser) return;

    try {
      const targetId = userId || targetUser.uid || targetUser.id;
      const newStatus = !current;
      const updated = { 
        ...targetUser, 
        isVerified: newStatus, 
        uid: targetId, 
        id: targetId,
        displayName: targetUser.displayName || '',
        email: targetUser.email || '',
        role: targetUser.role || 'pengurus',
        location: targetUser.location || '',
        spreadsheetId: effectiveSpreadsheetId
      };

      await saveData(null, activeToken, effectiveSpreadsheetId, 'users', updated, targetId);

      // Sinkronkan ke local cache profil jika akun ini yang sedang login
      const cacheKey = `user_profile_${targetId}`;
      const cached = localStorage.getItem(cacheKey);
      if (cached) {
        try {
          const parsed = JSON.parse(cached);
          parsed.isVerified = newStatus;
          localStorage.setItem(cacheKey, JSON.stringify(parsed));
        } catch (e) {}
      }

      showToast(!current ? 'Status verifikasi pengguna berhasil diaktifkan!' : 'Verifikasi pengguna dinonaktifkan', 'success');
    } catch (err: any) {
      showToast(`Gagal mengubah status di spreadsheet: ${err.message}`, 'error');
    }
  };

  const handleUpdate = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!editingUser) return;

    setIsSaving(true);
    const formData = new FormData(e.currentTarget);
    const userId = editingUser.uid || editingUser.id;
    const updated = {
      ...editingUser,
      uid: userId,
      id: userId,
      displayName: String(formData.get('displayName') || editingUser.displayName),
      role: String(formData.get('role') || editingUser.role) as UserRole,
      location: formData.get('role') === 'pengurus' ? String(formData.get('location') || '') as MosqueLocation : null,
      spreadsheetId: effectiveSpreadsheetId
    };
    try {
      await saveData(null, activeToken, effectiveSpreadsheetId, 'users', updated, userId);
      showToast('Profil pengguna berhasil diperbarui di Google Sheets!', 'success');
      setEditingUser(null);
    } catch (err: any) {
      showToast(`Gagal mengupdate profil: ${err.message}`, 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleAddUser = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();

    setIsSaving(true);
    const formData = new FormData(e.currentTarget);
    const newUid = 'usr_' + Math.random().toString(36).substring(2, 11);
    const role = String(formData.get('role') || 'pengurus') as UserRole;
    const newUser = {
      uid: newUid,
      id: newUid,
      displayName: String(formData.get('displayName')),
      email: String(formData.get('email')),
      role,
      location: role === 'pengurus' ? String(formData.get('location') || '') as MosqueLocation : null,
      isVerified: formData.get('isVerified') === 'true',
      createdAt: Date.now(),
      spreadsheetId: effectiveSpreadsheetId
    };

    try {
      await saveData(null, activeToken, effectiveSpreadsheetId, 'users', newUser, newUid);
      showToast('Pengguna baru berhasil ditambahkan ke Google Sheets!', 'success');
      setShowAddForm(false);
    } catch (err: any) {
      showToast(`Gagal menambahkan pengguna ke spreadsheet: ${err.message}`, 'error');
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <div className="flex items-center gap-3">
            <h2 className="text-2xl font-bold text-slate-900">Manajemen Pengguna</h2>
            <span className="px-3 py-1 bg-emerald-100 text-emerald-800 text-xs font-bold rounded-full flex items-center gap-1.5 shadow-sm">
              <Database className="w-3.5 h-3.5" />
              Google Sheets (Sheet: users)
            </span>
          </div>
          <p className="text-slate-500 text-sm mt-1">Kelola data pengguna, peranan (Admin/Pengurus), dan hak akses yang tersimpan di Google Spreadsheet</p>
        </div>
        <div className="flex items-center gap-3">
          <button
            onClick={handleRefresh}
            disabled={isRefreshing}
            className="p-2.5 rounded-xl border border-slate-200 text-slate-600 hover:text-emerald-700 hover:bg-emerald-50 transition-all shadow-sm flex items-center gap-2 text-sm font-semibold disabled:opacity-50"
            title="Muat ulang dari Google Sheets"
          >
            <RefreshCw className={cn("w-4 h-4", isRefreshing && "animate-spin text-emerald-600")} />
            <span className="hidden sm:inline">Segarkan Data</span>
          </button>
          <button 
            onClick={() => setShowAddForm(true)}
            className="bg-emerald-600 text-white px-5 py-2.5 rounded-xl font-semibold flex items-center gap-2 hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200 text-sm"
          >
            <Plus className="w-4 h-4" />
            Tambah Pengguna
          </button>
        </div>
      </div>

      {/* Spreadsheet Status & Info Banner */}
      <div className="bg-gradient-to-r from-emerald-50 via-teal-50 to-slate-50 border border-emerald-200/80 rounded-2xl p-4 flex flex-col md:flex-row md:items-center justify-between gap-4 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-emerald-600 text-white flex items-center justify-center font-bold text-sm shadow-md shadow-emerald-200">
            <Database className="w-5 h-5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h4 className="font-bold text-slate-900 text-sm">Database Terhubung: Google Sheets</h4>
              <span className="text-[11px] font-semibold text-emerald-700 bg-emerald-100/80 px-2 py-0.5 rounded-md">
                Sheet: users
              </span>
            </div>
            <p className="text-xs text-slate-600 mt-0.5">
              {effectiveSpreadsheetId ? (
                <>ID: <code className="font-mono bg-white/80 px-1.5 py-0.5 rounded border border-emerald-200 text-slate-700">{effectiveSpreadsheetId}</code></>
              ) : (
                <span className="text-amber-700 font-medium">ID Spreadsheet belum dikonfigurasi. Silakan cek menu Migrasi / Pengaturan.</span>
              )}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          {effectiveSpreadsheetId && (
            <a
              href={`https://docs.google.com/spreadsheets/d/${effectiveSpreadsheetId}/edit#gid=0`}
              target="_blank"
              rel="noreferrer"
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl bg-white border border-slate-200 text-xs font-bold text-slate-700 hover:text-emerald-700 hover:border-emerald-300 transition-all shadow-sm"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              Buka di Google Drive
            </a>
          )}
        </div>
      </div>

      {/* Sync Self Banner if current admin is not yet in users sheet */}
      {!loading && !isCurrentProfileInUsers && profile && (
        <div className="bg-amber-50 border border-amber-200 rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-sm">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-amber-100 text-amber-800 rounded-xl">
              <AlertCircle className="w-5 h-5" />
            </div>
            <div>
              <p className="text-sm font-bold text-amber-900">Akun Anda ({profile.displayName || profile.email}) belum tersinkronisasi di sheet 'users'</p>
              <p className="text-xs text-amber-700">Daftarkan akun ini agar terdata secara resmi sebagai Admin di Google Spreadsheet.</p>
            </div>
          </div>
          <button
            onClick={handleSyncMyProfile}
            disabled={isSyncingSelf}
            className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold rounded-xl transition-all shadow-sm whitespace-nowrap disabled:opacity-50 flex items-center justify-center gap-2"
          >
            {isSyncingSelf ? <Loader2 className="w-4 h-4 animate-spin" /> : <ShieldCheck className="w-4 h-4" />}
            Sinkronkan Akun Saya
          </button>
        </div>
      )}

      {/* Stat Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 sm:gap-4">
        <div className="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm">
          <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Total Pengguna</p>
          <p className="text-2xl font-black text-slate-900 mt-1">{stats.total}</p>
        </div>
        <div className="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm">
          <p className="text-xs font-bold text-emerald-600 uppercase tracking-wider">Terverifikasi</p>
          <p className="text-2xl font-black text-emerald-700 mt-1">{stats.verified}</p>
        </div>
        <div className="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm">
          <p className="text-xs font-bold text-amber-600 uppercase tracking-wider">Menunggu</p>
          <p className="text-2xl font-black text-amber-700 mt-1">{stats.pending}</p>
        </div>
        <div className="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm">
          <p className="text-xs font-bold text-purple-600 uppercase tracking-wider">Admin</p>
          <p className="text-2xl font-black text-purple-700 mt-1">{stats.admins}</p>
        </div>
        <div className="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm col-span-2 sm:col-span-1">
          <p className="text-xs font-bold text-blue-600 uppercase tracking-wider">Pengurus</p>
          <p className="text-2xl font-black text-blue-700 mt-1">{stats.pengurus}</p>
        </div>
      </div>

      {/* Filter and Search Bar */}
      <div className="bg-white rounded-2xl p-4 border border-slate-100 shadow-sm flex flex-col md:flex-row gap-3 items-stretch md:items-center justify-between">
        <div className="relative flex-1">
          <Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            type="text"
            placeholder="Cari berdasarkan nama, email, role, atau lokasi..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm outline-none focus:ring-2 focus:ring-emerald-500 focus:bg-white transition-all"
          />
        </div>

        <div className="flex items-center gap-2 overflow-x-auto pb-1 md:pb-0">
          <select
            value={roleFilter}
            onChange={(e) => setRoleFilter(e.target.value as any)}
            className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-700 outline-none focus:ring-2 focus:ring-emerald-500"
          >
            <option value="all">Semua Peran</option>
            <option value="admin">Admin</option>
            <option value="pengurus">Pengurus</option>
          </select>

          <div className="flex bg-slate-100 p-1 rounded-xl">
            <button
              onClick={() => setStatusFilter('all')}
              className={cn(
                "px-3 py-1.5 rounded-lg text-xs font-bold transition-all",
                statusFilter === 'all' ? "bg-white text-slate-900 shadow-sm" : "text-slate-500 hover:text-slate-700"
              )}
            >
              Semua
            </button>
            <button
              onClick={() => setStatusFilter('verified')}
              className={cn(
                "px-3 py-1.5 rounded-lg text-xs font-bold transition-all",
                statusFilter === 'verified' ? "bg-white text-emerald-700 shadow-sm" : "text-slate-500 hover:text-emerald-700"
              )}
            >
              Terverifikasi
            </button>
            <button
              onClick={() => setStatusFilter('pending')}
              className={cn(
                "px-3 py-1.5 rounded-lg text-xs font-bold transition-all",
                statusFilter === 'pending' ? "bg-white text-amber-700 shadow-sm" : "text-slate-500 hover:text-amber-700"
              )}
            >
              Menunggu
            </button>
          </div>
        </div>
      </div>

      {/* User Table */}
      <div className="bg-white rounded-3xl overflow-hidden border border-slate-100 shadow-sm">
        <table className="w-full text-left">
          <thead className="bg-slate-50 border-b border-slate-100">
            <tr>
              <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Nama & Email</th>
              <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Role</th>
              <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Lokasi Tugas</th>
              <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Status Akses</th>
              <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase tracking-wider text-right">Aksi</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {loading ? (
              <tr>
                <td colSpan={5} className="px-6 py-12 text-center">
                  <div className="flex flex-col items-center justify-center gap-2">
                    <Loader2 className="w-8 h-8 text-emerald-600 animate-spin" />
                    <p className="text-xs font-semibold text-slate-500">Membaca data pengguna dari Google Sheets...</p>
                  </div>
                </td>
              </tr>
            ) : filteredUsers.length === 0 ? (
              <tr>
                <td colSpan={5} className="px-6 py-12 text-center">
                  <div className="max-w-md mx-auto space-y-3">
                    <p className="text-slate-400 font-medium text-sm">
                      {searchTerm || roleFilter !== 'all' || statusFilter !== 'all' 
                        ? 'Tidak ada pengguna yang cocok dengan kriteria pencarian.' 
                        : 'Belum ada data pengguna yang tersimpan di Google Sheets.'}
                    </p>
                    {!users.length && (
                      <div className="flex justify-center gap-2">
                        {profile && (
                          <button
                            onClick={handleSyncMyProfile}
                            className="px-4 py-2 bg-emerald-600 text-white text-xs font-bold rounded-xl hover:bg-emerald-700 transition-all shadow-sm"
                          >
                            Daftarkan Akun Saya
                          </button>
                        )}
                        <button
                          onClick={() => setShowAddForm(true)}
                          className="px-4 py-2 bg-slate-900 text-white text-xs font-bold rounded-xl hover:bg-slate-800 transition-all shadow-sm"
                        >
                          Tambah Pengguna Baru
                        </button>
                      </div>
                    )}
                  </div>
                </td>
              </tr>
            ) : (
              paginatedUsers.map((u, uIdx) => {
                const userId = u.uid || u.id || u.email || `user-${uIdx}`;
                const isVerified = u.isVerified === true || (u.isVerified as any) === 'true' || (u.isVerified as any) === 'TRUE';
                const isCurrentUser = profile && (profile.uid === userId || (profile.email && u.email && profile.email.toLowerCase() === u.email.toLowerCase()));

                return (
                  <tr key={`${userId}-${uIdx}`} className="hover:bg-slate-50 transition-colors">
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-3">
                        <div className={cn(
                          "w-9 h-9 rounded-xl flex items-center justify-center font-bold text-xs shadow-sm",
                          u.role === 'admin' ? "bg-purple-100 text-purple-700" : "bg-blue-100 text-blue-700"
                        )}>
                          {(u.displayName || u.email || 'U').charAt(0).toUpperCase()}
                        </div>
                        <div>
                          <div className="flex items-center gap-1.5">
                            <span className="font-bold text-slate-900 text-sm">{u.displayName || '-'}</span>
                            {isCurrentUser && (
                              <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 text-[10px] font-bold rounded-full">
                                Anda
                              </span>
                            )}
                          </div>
                          <span className="text-xs text-slate-500 block font-mono">{u.email}</span>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <span className={cn(
                        "px-2.5 py-1 rounded-full text-xs font-bold uppercase tracking-wide",
                        u.role === 'admin' ? "bg-purple-100 text-purple-700" : "bg-blue-100 text-blue-700"
                      )}>
                        {u.role}
                      </span>
                    </td>
                    <td className="px-6 py-4">
                      <span className="text-sm font-medium text-slate-700">
                        {u.location ? (
                          <span className="inline-flex items-center gap-1">
                            <MapPin className="w-3.5 h-3.5 text-slate-400" />
                            {u.location}
                          </span>
                        ) : (
                          <span className="text-slate-400 text-xs italic">Semua Lokasi (Global)</span>
                        )}
                      </span>
                    </td>
                    <td className="px-6 py-4">
                      <button 
                        onClick={() => toggleVerify(userId, isVerified)}
                        title="Klik untuk mengubah status verifikasi di Google Sheets"
                        className={cn(
                          "px-3 py-1 rounded-full text-[11px] font-bold uppercase transition-all flex items-center gap-1.5 cursor-pointer shadow-sm",
                          isVerified 
                            ? "bg-emerald-100 text-emerald-800 hover:bg-amber-100 hover:text-amber-800" 
                            : "bg-amber-100 text-amber-800 hover:bg-emerald-100 hover:text-emerald-800"
                        )}
                      >
                        {isVerified ? (
                          <>
                            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                            Terverifikasi
                          </>
                        ) : (
                          <>
                            <Clock className="w-3.5 h-3.5 text-amber-600" />
                            Menunggu Verifikasi
                          </>
                        )}
                      </button>
                    </td>
                    <td className="px-6 py-4 text-right">
                      <div className="flex items-center justify-end gap-2">
                        <button 
                          onClick={() => setEditingUser(u)} 
                          className="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-600 rounded-lg font-bold text-xs transition-colors"
                        >
                          Edit
                        </button>
                        <button 
                          onClick={() => setDeleteId(userId)} 
                          className="px-3 py-1.5 bg-slate-100 hover:bg-red-50 hover:text-red-600 text-slate-500 rounded-lg font-bold text-xs transition-colors"
                        >
                          Hapus
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
        <Pagination
          currentPage={currentPage}
          totalPages={totalPages}
          totalItems={filteredUsers.length}
          pageSize={pageSize}
          onPageChange={setCurrentPage}
          onPageSizeChange={setPageSize}
          pageSizeOptions={[10, 20, 50]}
          className="px-6 py-4 bg-white border-t border-slate-100"
        />
      </div>

      {/* Modal Tambah Pengguna */}
      <AnimatePresence>
        {showAddForm && (
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
            <motion.div initial={{ scale: 0.95 }} animate={{ scale: 1 }} className="bg-white rounded-3xl p-8 w-full max-w-lg shadow-2xl">
              <div className="flex items-center justify-between mb-6">
                <div>
                  <h3 className="text-xl font-bold text-slate-900">Tambah Pengguna Baru</h3>
                  <p className="text-xs text-slate-500">Data akan langsung disimpan sebagai baris baru di Google Sheets</p>
                </div>
                <button onClick={() => setShowAddForm(false)} className="text-slate-400 hover:text-slate-600">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <form onSubmit={handleAddUser} className="space-y-4">
                <div>
                  <label className="block text-sm font-semibold text-slate-700 mb-1">Nama Tampilan</label>
                  <input name="displayName" placeholder="Contoh: Ustadz Ahmad Hidayat" required className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm" />
                </div>
                <div>
                  <label className="block text-sm font-semibold text-slate-700 mb-1">Email</label>
                  <input type="email" name="email" placeholder="nama@gmail.com" required className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm" />
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-sm font-semibold text-slate-700 mb-1">Peran (Role)</label>
                    <select name="role" defaultValue="pengurus" className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm bg-white">
                      <option value="admin">Admin</option>
                      <option value="pengurus">Pengurus</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-semibold text-slate-700 mb-1">Status Verifikasi</label>
                    <select name="isVerified" defaultValue="true" className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm bg-white">
                      <option value="true">Langsung Aktif</option>
                      <option value="false">Menunggu Verifikasi</option>
                    </select>
                  </div>
                </div>
                <div>
                  <label className="block text-sm font-semibold text-slate-700 mb-1">Lokasi Tugas (Khusus Pengurus)</label>
                  <select name="location" defaultValue="" className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm bg-white">
                    <option value="">Semua Lokasi / Tidak Terbatas</option>
                    {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(l => (
                      <option key={l} value={l}>{l}</option>
                    ))}
                  </select>
                </div>

                <div className="p-3 bg-emerald-50 rounded-xl border border-emerald-100 text-xs text-emerald-800 flex items-center gap-2">
                  <Database className="w-4 h-4 text-emerald-600 shrink-0" />
                  <span>Data pengguna akan langsung tersimpan di Google Spreadsheet (Sheet <strong>users</strong>).</span>
                </div>

                <div className="flex gap-3 pt-4">
                  <button type="button" onClick={() => setShowAddForm(false)} className="flex-1 px-6 py-2.5 rounded-xl border border-slate-200 font-semibold text-slate-600 text-sm">Batal</button>
                  <button type="submit" disabled={isSaving} className="flex-1 px-6 py-2.5 rounded-xl bg-emerald-600 text-white font-semibold text-sm disabled:opacity-50 hover:bg-emerald-700 transition-all flex items-center justify-center gap-2 shadow-lg shadow-emerald-200">
                    {isSaving && <Loader2 className="w-4 h-4 animate-spin" />}
                    {isSaving ? 'Menyimpan...' : 'Tambah ke Spreadsheet'}
                  </button>
                </div>
              </form>
            </motion.div>
          </motion.div>
        )}

        {/* Modal Edit Pengguna */}
        {editingUser && (
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
            <motion.div initial={{ scale: 0.95 }} animate={{ scale: 1 }} className="bg-white rounded-3xl p-8 w-full max-w-lg shadow-2xl">
              <div className="flex items-center justify-between mb-6">
                <div>
                  <h3 className="text-xl font-bold text-slate-900">Edit Profil Pengguna</h3>
                  <p className="text-xs text-slate-500">Perubahan akan langsung diperbarui pada Google Sheets</p>
                </div>
                <button onClick={() => setEditingUser(null)} className="text-slate-400 hover:text-slate-600">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <form onSubmit={handleUpdate} className="space-y-4">
                <div>
                  <label className="block text-sm font-semibold text-slate-700 mb-1">Nama Tampilan</label>
                  <input name="displayName" defaultValue={editingUser.displayName} required className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm" />
                </div>
                <div>
                  <label className="block text-sm font-semibold text-slate-700 mb-1">Email (Hanya Baca)</label>
                  <input value={editingUser.email} disabled className="w-full px-4 py-2.5 rounded-xl border border-slate-200 bg-slate-50 text-slate-500 text-sm font-mono cursor-not-allowed" />
                </div>
                <div>
                  <label className="block text-sm font-semibold text-slate-700 mb-1">Peran (Role)</label>
                  <select name="role" defaultValue={editingUser.role} className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm bg-white">
                    <option value="admin">Admin</option>
                    <option value="pengurus">Pengurus</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-semibold text-slate-700 mb-1">Lokasi Tugas (Khusus Pengurus)</label>
                  <select name="location" defaultValue={editingUser.location || ''} className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm bg-white">
                    <option value="">Semua Lokasi / Tidak Terbatas</option>
                    {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(l => (
                      <option key={l} value={l}>{l}</option>
                    ))}
                  </select>
                </div>

                <div className="flex gap-3 pt-4">
                  <button type="button" onClick={() => setEditingUser(null)} className="flex-1 px-6 py-2.5 rounded-xl border border-slate-200 font-semibold text-slate-600 text-sm">Batal</button>
                  <button type="submit" disabled={isSaving} className="flex-1 px-6 py-2.5 rounded-xl bg-emerald-600 text-white font-semibold text-sm disabled:opacity-50 hover:bg-emerald-700 transition-all flex items-center justify-center gap-2 shadow-lg shadow-emerald-200">
                    {isSaving && <Loader2 className="w-4 h-4 animate-spin" />}
                    {isSaving ? 'Menyimpan...' : 'Simpan Perubahan'}
                  </button>
                </div>
              </form>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <DeleteConfirmation 
        isOpen={!!deleteId}
        onClose={() => setDeleteId(null)}
        onConfirm={async () => {
          if (!deleteId) return;
          setIsDeleting(true);
          try {
            await deleteData(null, activeToken, effectiveSpreadsheetId, 'users', deleteId);
            clearSheetMemoryCache('users');
            showToast('Pengguna berhasil dihapus dari Google Sheets!', 'success');
            setDeleteId(null);
          } catch (err: any) {
            showToast(`Gagal menghapus pengguna dari spreadsheet: ${err.message}`, 'error');
          } finally {
            setIsDeleting(false);
          }
        }}
        loading={isDeleting}
        title="Hapus Pengguna dari Spreadsheet"
        message="Hapus akses pengguna ini dari Google Sheets? Pengguna tidak akan dapat mengakses dashboard ini lagi."
      />
    </div>
  );
}

// --- Main Layout ---

  const MigrationView = ({ profile }: { profile: UserProfile }) => {
    const { accessToken, googleSignIn, setSpreadsheetId, spreadsheetId } = useFirebase();
    const [status, setStatus] = useState<'idle' | 'running' | 'success' | 'error'>('idle');
    const [logs, setLogs] = useState<string[]>([]);
    const [inputSpreadsheetId, setInputSpreadsheetId] = useState(spreadsheetId || '');
    const envSpreadsheetId = (import.meta as any).env?.VITE_SPREADSHEET_ID || '';
    const [saveIdSuccess, setSaveIdSuccess] = useState(false);

    const envScriptUrl = (import.meta as any).env?.VITE_APPS_SCRIPT_URL || '';
    const [inputScriptUrl, setInputScriptUrl] = useState(localStorage.getItem('app_script_url') || envScriptUrl);
    const [saveScriptSuccess, setSaveScriptSuccess] = useState(false);
    const [copiedScript, setCopiedScript] = useState(false);

    useEffect(() => {
      if (spreadsheetId) {
        setInputSpreadsheetId(spreadsheetId);
      }
    }, [spreadsheetId]);

    const addLog = (msg: string) => setLogs(prev => [...prev, `[${new Date().toLocaleTimeString()}] ${msg}`]);

    const handleSaveManualId = (e: React.FormEvent) => {
      e.preventDefault();
      const trimmed = inputSpreadsheetId.trim();
      setSpreadsheetId(trimmed || null);
      setSaveIdSuccess(true);
      setTimeout(() => setSaveIdSuccess(false), 3000);
    };

    const handleSaveScriptUrl = (e: React.FormEvent) => {
      e.preventDefault();
      const trimmed = inputScriptUrl.trim();
      if (trimmed) {
        localStorage.setItem('app_script_url', trimmed);
      } else {
        localStorage.removeItem('app_script_url');
      }
      setSaveScriptSuccess(true);
      setTimeout(() => setSaveScriptSuccess(false), 3000);
    };

    const appsScriptCode = `// KODE GOOGLE APPS SCRIPT (UNTUK AKSES BEBAS LOGIN & BEBAS TOKEN)
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = data.collection || 'data';
    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    if (data.action === 'save') {
      var item = data.item;
      var rows = sheet.getDataRange().getValues();
      var headers = rows.length > 0 && rows[0][0] !== '' ? rows[0] : [];
      
      var itemKeys = Object.keys(item);
      var newHeaders = headers.slice();
      itemKeys.forEach(function(k) {
        if (newHeaders.indexOf(k) === -1) newHeaders.push(k);
      });
      
      if (newHeaders.length > headers.length || headers.length === 0) {
        headers = newHeaders;
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
      
      var idIndex = headers.indexOf('id');
      if (idIndex === -1) idIndex = headers.indexOf('uid');
      var existingRow = -1;
      if (idIndex !== -1 && rows.length > 1) {
        for (var i = 1; i < rows.length; i++) {
          if (String(rows[i][idIndex]) === String(data.id)) {
            existingRow = i + 1;
            break;
          }
        }
      }
      
      var rowData = headers.map(function(h) {
        var val = item[h];
        if (val === undefined || val === null) return '';
        if (typeof val === 'object') return JSON.stringify(val);
        return val;
      });
      
      if (existingRow !== -1) {
        sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', id: data.id })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (data.action === 'delete') {
      var rows = sheet.getDataRange().getValues();
      var headers = rows.length > 0 ? rows[0] : [];
      var idIndex = headers.indexOf('id');
      if (idIndex === -1) idIndex = headers.indexOf('uid');
      if (idIndex !== -1) {
        for (var i = 1; i < rows.length; i++) {
          if (String(rows[i][idIndex]) === String(data.id)) {
            sheet.deleteRow(i + 1);
            break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'unknown_action' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = (e && e.parameter && (e.parameter.collection || e.parameter.sheet)) || 'jamaah';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ values: [] })).setMimeType(ContentService.MimeType.JSON);
    }
    var values = sheet.getDataRange().getValues();
    return ContentService.createTextOutput(JSON.stringify({ values: values })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ values: [], error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}`;

    const handleCopyCode = () => {
      navigator.clipboard.writeText(appsScriptCode);
      setCopiedScript(true);
      setTimeout(() => setCopiedScript(false), 3000);
    };

    const runMigration = async () => {
      if (!accessToken) return;
      setStatus('running');
      setLogs([]);
      
      try {
        addLog('Memulai pembuatan database Google Sheets baru...');
        
        // 1. Create Spreadsheet
        addLog('Membuat file Spreadsheet di Google Drive...');
        const spreadsheet = await createSpreadsheet(accessToken, 'Database Mosque Management');
        const newSpreadsheetId = spreadsheet.spreadsheetId;
        addLog(`Spreadsheet berhasil dibuat! ID: ${newSpreadsheetId}`);

        // 2. Setup standard tables
        const collections = [
          { name: 'jamaah', headers: ['id', 'memberId', 'name', 'phone', 'location', 'category', 'gender', 'createdAt'] },
          { name: 'attendance', headers: ['id', 'jamaahId', 'jamaahName', 'location', 'category', 'date', 'sessionType', 'day', 'status', 'reason'] },
          { name: 'assets', headers: ['id', 'name', 'code', 'category', 'condition', 'location', 'quantity', 'purchaseDate', 'value'] },
          { name: 'activities', headers: ['id', 'title', 'description', 'date', 'time', 'location', 'speaker', 'category'] },
          { name: 'facility_stats', headers: ['id', 'facilityName', 'capacity', 'status', 'lastCleaned'] },
          { name: 'ub_shopping', headers: ['id', 'itemName', 'quantity', 'estimatedCost', 'requestedBy', 'status', 'createdAt'] },
          { name: 'users', headers: ['uid', 'id', 'email', 'displayName', 'role', 'location', 'isVerified', 'createdAt', 'spreadsheetId'] }
        ];

        for (const col of collections) {
          addLog(`Inisialisasi tabel '${col.name}'...`);
          await updateSheetValues(accessToken, newSpreadsheetId, `${col.name}!A1`, [col.headers]).catch(() => {});
        }

        if (profile) {
          const userRow = { ...profile, spreadsheetId: newSpreadsheetId, id: profile.uid };
          const headers = ['uid', 'id', 'email', 'displayName', 'role', 'location', 'isVerified', 'createdAt', 'spreadsheetId'];
          const values = [headers, headers.map(h => (userRow as any)[h] ?? '')];
          await updateSheetValues(accessToken, newSpreadsheetId, `users!A1`, values).catch(() => {});
        }

        setSpreadsheetId(newSpreadsheetId);
        addLog('Setup Selesai! Firestore kini dinonaktifkan sepenuhnya dan data dikelola 100% via Google Sheets.');
        setStatus('success');
      } catch (error: any) {
        console.error(error);
        addLog(`Error: ${error.message}`);
        setStatus('error');
      }
    };

    return (
      <div className="max-w-2xl mx-auto py-12">
        <div className="bg-white rounded-[2.5rem] p-10 shadow-xl border border-slate-100">
          <div className="flex items-center gap-6 mb-10">
            <div className="p-4 bg-emerald-100 rounded-3xl">
              <CloudCog className="w-10 h-10 text-emerald-600" />
            </div>
            <div>
              <h2 className="text-3xl font-black text-slate-900 leading-tight">Migrasi Database</h2>
              <p className="text-slate-500 font-medium">Pindahkan seluruh data dari Firestore ke Google Sheets</p>
            </div>
          </div>

          <div className="space-y-8">
            {!accessToken ? (
              <div className="p-8 bg-slate-50 rounded-[2rem] border border-slate-100">
                <p className="text-slate-600 font-bold mb-6 text-center">Hubungkan akun Google Anda untuk memulai migrasi.</p>
                <button 
                  onClick={googleSignIn}
                  className="w-full py-4 bg-white border-2 border-slate-100 rounded-2xl flex items-center justify-center gap-4 hover:border-emerald-500 transition-all group"
                >
                  <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" className="w-6 h-6" alt="Google" />
                  <span className="font-black uppercase tracking-widest text-slate-700 group-hover:text-emerald-600">Hubungkan Google Sheets</span>
                </button>
              </div>
            ) : (
              <div className="space-y-6">
                <div className="flex items-center gap-4 p-4 bg-emerald-50 rounded-2xl text-emerald-700 border border-emerald-100">
                  <CheckCircle2 className="w-5 h-5" />
                  <p className="text-sm font-black uppercase tracking-widest">Akun Google Terhubung</p>
                </div>

                <button 
                  onClick={runMigration}
                  disabled={status === 'running'}
                  className={cn(
                    "w-full py-6 rounded-[1.8rem] font-black uppercase tracking-widest flex items-center justify-center gap-4 transition-all shadow-xl",
                    status === 'running' ? "bg-slate-100 text-slate-400 cursor-not-allowed" : "bg-emerald-600 text-white hover:bg-emerald-700 shadow-emerald-200"
                  )}
                >
                  {status === 'running' ? (
                    <motion.div animate={{ rotate: 360 }} transition={{ repeat: Infinity, duration: 1, ease: "linear" }} className="w-6 h-6 border-4 border-slate-300 border-t-slate-500 rounded-full" />
                  ) : (
                    <>
                      <Database className="w-6 h-6" />
                      Mulai Migrasi Data
                    </>
                  )}
                </button>
              </div>
            )}

            {(logs.length > 0 || status === 'running') && (
              <div className="bg-slate-900 rounded-[2rem] p-6 font-mono text-xs text-emerald-400 space-y-2 max-h-60 overflow-y-auto shadow-inner">
                {logs.map((log, i) => (
                  <div key={i} className="opacity-80 leading-relaxed">{log}</div>
                ))}
                {status === 'running' && <div className="animate-pulse">_</div>}
              </div>
            )}

            {status === 'success' && (
              <div className="p-6 bg-emerald-50 rounded-2xl border border-emerald-100 text-center">
                <p className="text-emerald-800 font-black text-sm uppercase tracking-widest mb-2">🎉 Migrasi Berhasil!</p>
                <p className="text-emerald-600 text-xs font-medium">Data Anda sekarang tersimpan di Google Sheets. Sistem akan otomatis beralih menggunakan spreadsheet.</p>
              </div>
            )}

            {/* Pengaturan ID Spreadsheet & Environment Variable */}
            <div className="p-6 bg-slate-50 rounded-[2rem] border border-slate-200 space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold text-slate-800 text-sm">Google Spreadsheet ID</h4>
                  <p className="text-xs text-slate-500">ID spreadsheet yang digunakan sebagai database utama aplikasi</p>
                </div>
                {envSpreadsheetId ? (
                  <span className="px-3 py-1 bg-emerald-100 text-emerald-800 text-[11px] font-bold rounded-full">
                    Tersimpan di .env
                  </span>
                ) : (
                  <span className="px-3 py-1 bg-slate-200 text-slate-700 text-[11px] font-medium rounded-full">
                    Penyimpanan Browser
                  </span>
                )}
              </div>

              <form onSubmit={handleSaveManualId} className="space-y-3">
                <div className="flex gap-2">
                  <input 
                    type="text"
                    value={inputSpreadsheetId}
                    onChange={(e) => setInputSpreadsheetId(e.target.value)}
                    placeholder="Contoh: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
                    className="flex-1 px-4 py-2.5 rounded-xl border border-slate-300 text-xs font-mono outline-none focus:ring-2 focus:ring-emerald-500 bg-white"
                  />
                  <button
                    type="submit"
                    className="px-5 py-2.5 bg-slate-900 text-white text-xs font-bold rounded-xl hover:bg-slate-800 transition-all shadow-sm"
                  >
                    Simpan ID
                  </button>
                </div>

                {saveIdSuccess && (
                  <p className="text-xs text-emerald-600 font-bold flex items-center gap-1.5">
                    <CheckCircle2 className="w-3.5 h-3.5" /> ID Spreadsheet berhasil diperbarui!
                  </p>
                )}
              </form>

              <div className="p-3.5 bg-blue-50/70 rounded-xl border border-blue-100 text-xs text-blue-800 space-y-1">
                <p className="font-bold">💡 Tips Permanen (Bebas Token):</p>
                <p className="text-slate-600 leading-relaxed">
                  Tambahkan <code className="bg-blue-100 px-1.5 py-0.5 rounded font-mono text-blue-900">VITE_SPREADSHEET_ID=ID_SPREADSHEET_ANDA</code> pada file <strong>.env</strong>. Dengan cara ini, semua pengguna dan role (Admin, Pengurus, dsb) akan otomatis terhubung ke database yang sama tanpa perlu login ulang Google atau bergantung pada token sesi!
                </p>
              </div>
            </div>

            {/* Pengaturan Google Apps Script (Bebas Login & Bebas Token 100%) */}
            <div className="p-6 bg-gradient-to-br from-emerald-50/70 to-teal-50/40 rounded-[2rem] border border-emerald-200/80 space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold text-slate-900 text-sm flex items-center gap-2">
                    <span>⚡ Mode Bebas Token (Apps Script Web App)</span>
                  </h4>
                  <p className="text-xs text-slate-600">Simpan & baca data Google Sheets 100% tanpa perlu login akun Google berulang kali</p>
                </div>
                {inputScriptUrl ? (
                  <span className="px-3 py-1 bg-emerald-600 text-white text-[11px] font-black rounded-full flex items-center gap-1 shadow-sm">
                    <CheckCircle2 className="w-3 h-3" /> Bebas Token Aktif
                  </span>
                ) : (
                  <span className="px-3 py-1 bg-slate-200 text-slate-700 text-[11px] font-medium rounded-full">
                    Opsional
                  </span>
                )}
              </div>

              <form onSubmit={handleSaveScriptUrl} className="space-y-3">
                <div className="flex gap-2">
                  <input 
                    type="url"
                    value={inputScriptUrl}
                    onChange={(e) => setInputScriptUrl(e.target.value)}
                    placeholder="https://script.google.com/macros/s/.../exec"
                    className="flex-1 px-4 py-2.5 rounded-xl border border-emerald-300 text-xs font-mono outline-none focus:ring-2 focus:ring-emerald-500 bg-white"
                  />
                  <button
                    type="submit"
                    className="px-5 py-2.5 bg-emerald-700 text-white text-xs font-bold rounded-xl hover:bg-emerald-800 transition-all shadow-sm flex items-center gap-1.5"
                  >
                    Simpan URL
                  </button>
                </div>

                {saveScriptSuccess && (
                  <p className="text-xs text-emerald-700 font-bold flex items-center gap-1.5">
                    <CheckCircle2 className="w-3.5 h-3.5" /> URL Web App berhasil disimpan! Sekarang data tersimpan tanpa token.
                  </p>
                )}
              </form>

              <div className="p-4 bg-white/90 rounded-xl border border-emerald-100 text-xs space-y-3">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-800">Langkah Memasang Skrip (Hanya 1 Menit):</span>
                  <button
                    type="button"
                    onClick={handleCopyCode}
                    className="px-3 py-1.5 bg-slate-900 text-white hover:bg-slate-800 rounded-lg text-[11px] font-bold flex items-center gap-1.5 transition-all shadow-sm"
                  >
                    {copiedScript ? (
                      <>
                        <Check className="w-3.5 h-3.5 text-emerald-400" />
                        <span>Tersalin!</span>
                      </>
                    ) : (
                      <>
                        <Copy className="w-3.5 h-3.5" />
                        <span>Salin Kode Skrip</span>
                      </>
                    )}
                  </button>
                </div>
                <ol className="list-decimal list-inside space-y-1.5 text-slate-600 leading-relaxed text-[11px]">
                  <li>Buka spreadsheet Anda di Google Drive.</li>
                  <li>Klik menu <strong>Ekstensi (Extensions)</strong> &gt; <strong>Apps Script</strong>.</li>
                  <li>Hapus kode yang ada, lalu tempel (Paste) kode yang disalin dari tombol di atas.</li>
                  <li>Klik tombol <strong>Terapkan (Deploy)</strong> di pojok kanan atas &gt; <strong>Penerapan Baru (New deployment)</strong>.</li>
                  <li>Pilih jenis <strong>Aplikasi Web (Web App)</strong>. Pada bagian <em>Akses (Who has access)</em>, pilih <strong>Siapa saja (Anyone)</strong>.</li>
                  <li>Klik <strong>Terapkan (Deploy)</strong>, lalu salin URL Web App (akhiran <code className="bg-slate-100 px-1 font-mono text-slate-800">/exec</code>) dan tempel pada input di atas!</li>
                </ol>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  };

function UnverifiedAccountView({ 
  profile, 
  onCheckStatus, 
  onSignOut 
}: { 
  profile: any; 
  onCheckStatus: () => Promise<void>; 
  onSignOut: () => void; 
}) {
  const [checking, setChecking] = useState(false);
  const [feedback, setFeedback] = useState<string | null>(null);

  const checkRef = React.useRef(onCheckStatus);
  useEffect(() => {
    checkRef.current = onCheckStatus;
  }, [onCheckStatus]);

  // Auto-polling verification status from Google Sheets every 6 seconds
  useEffect(() => {
    const interval = setInterval(() => {
      checkRef.current?.().catch(() => {});
    }, 6000);
    return () => clearInterval(interval);
  }, []);

  const handleManualCheck = async () => {
    setChecking(true);
    setFeedback(null);
    try {
      await onCheckStatus();
      setFeedback('Status berhasil disinkronkan dengan Google Sheets');
      setTimeout(() => setFeedback(null), 3000);
    } catch (err: any) {
      setFeedback('Gagal memeriksa status: ' + (err.message || 'Coba lagi'));
    } finally {
      setChecking(false);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-50 p-4">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        className="bg-white p-8 md:p-10 rounded-3xl shadow-xl max-w-md w-full border border-slate-100 text-center"
      >
        <div className="w-16 h-16 bg-amber-100 text-amber-600 rounded-2xl flex items-center justify-center mx-auto mb-5">
          <AlertCircle className="w-8 h-8" />
        </div>
        <h2 className="text-2xl font-bold text-slate-900 mb-2 tracking-tight">Akun Menunggu Verifikasi</h2>
        <p className="text-slate-500 mb-6 font-medium text-sm leading-relaxed">
          Pendaftaran berhasil! Akun Anda telah dicatat di database dan sedang menunggu verifikasi dari admin sebelum dapat mengakses sistem.
        </p>

        {/* Ringkasan Data Akun Terdaftar */}
        <div className="bg-slate-50 border border-slate-200/80 rounded-2xl p-4 mb-6 text-left space-y-2.5 text-xs">
          <div className="flex justify-between items-center">
            <span className="text-slate-500 font-medium">Nama Lengkap</span>
            <span className="font-bold text-slate-900">{profile?.displayName || '-'}</span>
          </div>
          <div className="flex justify-between items-center">
            <span className="text-slate-500 font-medium">Email</span>
            <span className="font-bold text-slate-900">{profile?.email || '-'}</span>
          </div>
          <div className="flex justify-between items-center">
            <span className="text-slate-500 font-medium">Role</span>
            <span className="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-blue-100 text-blue-700 capitalize">
              {profile?.role || 'pengurus'}
            </span>
          </div>
          {profile?.role === 'pengurus' && (
            <div className="flex justify-between items-center">
              <span className="text-slate-500 font-medium">Lokasi Tugas</span>
              <span className="font-bold text-slate-900">{profile?.location || '-'}</span>
            </div>
          )}
          <div className="flex justify-between items-center pt-2 border-t border-slate-200">
            <span className="text-slate-500 font-medium">Status Verifikasi</span>
            <span className="px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase bg-amber-100 text-amber-700">
              Belum Terverifikasi
            </span>
          </div>
        </div>

        {feedback && (
          <p className="text-xs font-semibold text-emerald-600 mb-4 animate-pulse">{feedback}</p>
        )}

        <div className="space-y-3">
          <button 
            type="button"
            onClick={handleManualCheck}
            disabled={checking}
            className="w-full py-3.5 bg-emerald-600 text-white rounded-2xl font-bold text-sm hover:bg-emerald-700 transition-all flex items-center justify-center gap-2 shadow-lg shadow-emerald-200 disabled:opacity-50"
          >
            {checking ? (
              <>
                <Loader2 className="w-4 h-4 animate-spin" />
                <span>Memeriksa Status...</span>
              </>
            ) : (
              <>
                <RefreshCw className="w-4 h-4" />
                <span>Cek Status Verifikasi</span>
              </>
            )}
          </button>

          <button 
            type="button"
            onClick={onSignOut}
            className="w-full py-3.5 bg-slate-100 text-slate-700 rounded-2xl font-bold text-sm hover:bg-slate-200 transition-all"
          >
            Keluar
          </button>
        </div>
      </motion.div>
    </div>
  );
}

function DashboardContent() {
  const { user, db, auth, profile, spreadsheetId, syncProfileFromSheet, loading: authLoading } = useFirebase();
  const [activeTab, setActiveTab] = useState('overview');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [formTrigger, setFormTrigger] = useState<string | null>(null);
  const [isAttendanceMode, setIsAttendanceMode] = useState(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      return params.get('attendance') === 'true' || Boolean(params.get('location'));
    }
    return false;
  });

  const handleCheckStatus = useCallback(async () => {
    if (syncProfileFromSheet && user?.uid) {
      await syncProfileFromSheet(user.uid);
    }
  }, [syncProfileFromSheet, user?.uid]);

  if (authLoading) return (
    <div className="min-h-screen flex items-center justify-center">
      <motion.div animate={{ rotate: 360 }} transition={{ repeat: Infinity, duration: 1, ease: "linear" }} className="w-8 h-8 border-4 border-emerald-500 border-t-transparent rounded-full" />
    </div>
  );

  if (isAttendanceMode) {
    return (
      <PublicAttendanceView 
        onBack={() => {
          setIsAttendanceMode(false);
          if (typeof window !== 'undefined' && window.history) {
            window.history.replaceState({}, '', window.location.pathname);
          }
        }} 
        spreadsheetId={spreadsheetId || undefined} 
      />
    );
  }

  if (!user || !profile) return <LoginView onAttendanceMode={() => setIsAttendanceMode(true)} />;

  const isUserVerified = profile.isVerified === true || (profile.isVerified as any) === 'true' || (profile.isVerified as any) === 'TRUE';

  if (!isUserVerified) {
    return (
      <UnverifiedAccountView 
        profile={profile} 
        onCheckStatus={handleCheckStatus}
        onSignOut={() => auth && signOut(auth)} 
      />
    );
  }

  const renderContent = () => {
    const profileWithSheet = { ...profile, spreadsheetId };

    switch (activeTab) {
      case 'overview': return <Overview profile={profileWithSheet} />;
      case 'jamaah': return <JamaahView profile={profileWithSheet} formTrigger={formTrigger} onFormTriggered={() => setFormTrigger(null)} />;
      case 'cacah_jiwa': return <CacahJiwaView profile={profileWithSheet} />;
      case 'inventaris': return <InventarisView profile={profileWithSheet} formTrigger={formTrigger} onFormTriggered={() => setFormTrigger(null)} assetType="barang" />;
      case 'tanah': return <InventarisView profile={profileWithSheet} formTrigger={formTrigger} onFormTriggered={() => setFormTrigger(null)} assetType="tanah" />;
      case 'activities': return <ActivitiesView profile={profileWithSheet} />;
      case 'facilities': return <FacilityView profile={profileWithSheet} />;
      case 'attendance_report': return <AttendanceReportView profile={profileWithSheet} />;
      case 'ub_shopping': return <UBShoppingView profile={profileWithSheet} />;
      case 'users': return <UsersView profile={profileWithSheet} />;
      case 'migration': return <MigrationView profile={profile} />;
      default: return <Overview profile={profileWithSheet} />;
    }
  };

  return (
    <div className="flex min-h-screen bg-slate-50 text-slate-900 font-sans">
      {/* Sidebar - Desktop */}
      <aside className="w-72 bg-white border-r border-slate-100 hidden lg:flex flex-col sticky top-0 h-screen">
        <div className="p-8 border-b border-slate-50">
          <div className="flex items-center gap-3 text-emerald-600">
            <div className="p-2 bg-emerald-100 rounded-lg">
              <LayoutDashboard className="w-6 h-6" />
            </div>
            <h1 className="text-xl font-black tracking-tight leading-none uppercase">Mosque<br/><span className="text-slate-400 text-[10px] font-bold tracking-[0.2em]">Management</span></h1>
          </div>
        </div>

        <nav className="flex-grow p-6 space-y-2 overflow-y-auto">
          <SidebarItem icon={LayoutDashboard} label="Dashboard" active={activeTab === 'overview'} onClick={() => setActiveTab('overview')} />
          <SidebarItem icon={Users} label="Data Jamaah" active={activeTab === 'jamaah'} onClick={() => setActiveTab('jamaah')} />
          <SidebarItem icon={FileSpreadsheet} label="Cacah Jiwa" active={activeTab === 'cacah_jiwa'} onClick={() => setActiveTab('cacah_jiwa')} />
          <SidebarItem icon={Package} label="Inventaris" active={activeTab === 'inventaris'} onClick={() => setActiveTab('inventaris')} />
          <SidebarItem icon={MapIcon} label="Tanah Sabilillah" active={activeTab === 'tanah'} onClick={() => setActiveTab('tanah')} />
          <SidebarItem icon={Calendar} label="Kegiatan" active={activeTab === 'activities'} onClick={() => setActiveTab('activities')} />
          <SidebarItem icon={Clock} label="Fasilitas" active={activeTab === 'facilities'} onClick={() => setActiveTab('facilities')} />
          <SidebarItem icon={ClipboardList} label="Laporan Absensi" active={activeTab === 'attendance_report'} onClick={() => setActiveTab('attendance_report')} />
          <SidebarItem icon={ShoppingBag} label="Belanja UB" active={activeTab === 'ub_shopping'} onClick={() => setActiveTab('ub_shopping')} />
          
          <div className="pt-3 pb-1">
            <div className="h-[1px] bg-slate-100 mb-2" />
            <span className="px-3 text-[10px] font-black uppercase tracking-wider text-slate-400">Sistem & Akses</span>
          </div>

          <SidebarItem icon={UserCheck} label="Pengguna (User)" active={activeTab === 'users'} onClick={() => setActiveTab('users')} />
          <SidebarItem icon={Database} label="Migrasi Database" active={activeTab === 'migration'} onClick={() => setActiveTab('migration')} />
        </nav>
        
        <div className="p-6 border-t border-slate-50">
          <div className="flex items-center gap-3 mb-6 p-2">
            <div className="w-10 h-10 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-700 font-bold">
              {profile.displayName.charAt(0)}
            </div>
            <div className="flex-grow overflow-hidden">
              <p className="font-bold text-sm truncate">{profile.displayName}</p>
              <p className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">{profile.role} {profile.location ? `• ${profile.location}` : ''}</p>
            </div>
          </div>
          <button 
            onClick={() => auth && signOut(auth)}
            className="w-full flex items-center gap-3 px-4 py-3 rounded-xl text-red-500 hover:bg-red-50 transition-all font-medium"
          >
            <LogOut className="w-5 h-5" />
            Keluar
          </button>
        </div>
      </aside>

      {/* Mobile Menu Overlay */}
      <AnimatePresence>
        {isMobileMenuOpen && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] lg:hidden"
            onClick={() => setIsMobileMenuOpen(false)}
          >
            <motion.div 
              initial={{ x: '-100%' }}
              animate={{ x: 0 }}
              exit={{ x: '-100%' }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="w-[280px] h-full bg-white flex flex-col"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="p-6 border-b border-slate-50 flex items-center justify-between">
                <div className="flex items-center gap-2 text-emerald-600">
                  <LayoutDashboard className="w-6 h-6" />
                  <span className="font-black uppercase tracking-tight">Mosque</span>
                </div>
                <button onClick={() => setIsMobileMenuOpen(false)} className="p-2 text-slate-400">
                  <X className="w-6 h-6" />
                </button>
              </div>

              <nav className="flex-grow p-4 space-y-2 overflow-y-auto">
                <SidebarItem icon={LayoutDashboard} label="Dashboard" active={activeTab === 'overview'} onClick={() => { setActiveTab('overview'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={Users} label="Data Jamaah" active={activeTab === 'jamaah'} onClick={() => { setActiveTab('jamaah'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={FileSpreadsheet} label="Cacah Jiwa" active={activeTab === 'cacah_jiwa'} onClick={() => { setActiveTab('cacah_jiwa'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={Package} label="Inventaris" active={activeTab === 'inventaris'} onClick={() => { setActiveTab('inventaris'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={MapIcon} label="Tanah Sabilillah" active={activeTab === 'tanah'} onClick={() => { setActiveTab('tanah'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={Calendar} label="Kegiatan" active={activeTab === 'activities'} onClick={() => { setActiveTab('activities'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={Clock} label="Fasilitas" active={activeTab === 'facilities'} onClick={() => { setActiveTab('facilities'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={ClipboardList} label="Laporan Absensi" active={activeTab === 'attendance_report'} onClick={() => { setActiveTab('attendance_report'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={ShoppingBag} label="Belanja UB" active={activeTab === 'ub_shopping'} onClick={() => { setActiveTab('ub_shopping'); setIsMobileMenuOpen(false); }} />
                
                <div className="pt-3 pb-1">
                  <div className="h-[1px] bg-slate-100 mb-2" />
                  <span className="px-3 text-[10px] font-black uppercase tracking-wider text-slate-400">Sistem & Akses</span>
                </div>

                <SidebarItem icon={UserCheck} label="Pengguna (User)" active={activeTab === 'users'} onClick={() => { setActiveTab('users'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={Database} label="Migrasi Database" active={activeTab === 'migration'} onClick={() => { setActiveTab('migration'); setIsMobileMenuOpen(false); }} />
              </nav>

              <div className="p-6 border-t border-slate-50">
                <button 
                  onClick={() => auth && signOut(auth)}
                  className="w-full flex items-center gap-3 px-4 py-3 rounded-xl text-red-500 hover:bg-red-50 transition-all font-medium"
                >
                  <LogOut className="w-5 h-5" />
                  Keluar
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Main Content */}
      <main className="flex-grow p-4 lg:p-12 max-w-7xl mx-auto w-full">
        <header className="mb-12 flex flex-col sm:flex-row sm:items-center justify-between gap-6">
          <div className="flex items-center justify-between lg:block">
            <div>
              <h1 className="text-2xl lg:text-3xl font-black text-slate-900">Selamat Datang, {profile.displayName.split(' ')[0]}!</h1>
              <p className="text-slate-500 mt-1 hidden sm:block">Sistem Manajemen Desa Gandaria</p>
            </div>
            <button 
              onClick={() => setIsMobileMenuOpen(true)}
              className="lg:hidden p-3 bg-white rounded-2xl border border-slate-100 shadow-sm text-slate-600"
            >
              <Menu className="w-6 h-6" />
            </button>
          </div>
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-3 px-4 py-2 bg-white rounded-2xl border border-slate-100 shadow-sm text-sm font-medium text-slate-600">
               <MapPin className="w-4 h-4 text-emerald-500" />
               {profile.location || 'Seluruh Lokasi'}
            </div>
            <button 
              onClick={() => auth && signOut(auth)}
              className="lg:hidden p-3 bg-white rounded-2xl border border-slate-100 shadow-sm text-red-500 hover:bg-red-50 transition-all"
              title="Keluar"
            >
              <LogOut className="w-5 h-5" />
            </button>
          </div>
        </header>

        {!(spreadsheetId) && activeTab !== 'migration' && (
          <motion.div 
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            className="mb-12 p-6 bg-amber-50 rounded-[2rem] border border-amber-200 flex flex-col md:flex-row items-center gap-6"
          >
            <div className="p-4 bg-amber-100 rounded-2xl text-amber-600">
              <AlertCircle className="w-8 h-8" />
            </div>
            <div className="flex-grow text-center md:text-left">
              <h3 className="text-lg font-black text-amber-900">Google Sheets Belum Terhubung</h3>
              <p className="text-amber-700 font-medium">Data bisnis saat ini dikunci. Silakan lakukan migrasi database untuk mengaktifkan penyimpanan di Google Sheets.</p>
            </div>
            <button 
              onClick={() => setActiveTab('migration')}
              className="px-8 py-3 bg-amber-600 text-white rounded-xl font-black uppercase tracking-wider hover:bg-amber-700 transition-all shadow-lg shadow-amber-200 whitespace-nowrap"
            >
              Ke Menu Migrasi
            </button>
          </motion.div>
        )}

        <motion.div
          key={activeTab}
          initial={{ opacity: 0, x: 20 }}
          animate={{ opacity: 1, x: 0 }}
          transition={{ duration: 0.3 }}
        >
          {renderContent()}
        </motion.div>
      </main>
    </div>
  );
}

export default function App() {
  return (
    <FirebaseProvider>
      <ToastProvider>
        <DashboardContent />
      </ToastProvider>
    </FirebaseProvider>
  );
}
