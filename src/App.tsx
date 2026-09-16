import * as XLSX from 'xlsx';
import React, { useState, useEffect, useMemo } from 'react';
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
  ChevronRight
} from 'lucide-react';
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
import { 
  doc, 
  getDoc, 
  setDoc, 
  collection, 
  addDoc, 
  updateDoc, 
  deleteDoc,
  where,
  query,
  getDocs,
  Timestamp
} from 'firebase/firestore';
import { 
  getAuth, 
  signInWithEmailAndPassword, 
  createUserWithEmailAndPassword, 
  signOut 
} from 'firebase/auth';
import { useFirestoreQuery } from './hooks/useFirestore';
import { UserProfile, UserRole, MosqueLocation, Jamaah, Asset, Activity, FacilityStat, JamaahCategory, Attendance } from './types';
import { cn } from './lib/utils';

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

function PublicAttendanceView({ onBack }: { onBack: () => void }) {
  const { db } = useFirebase();
  const [selectedLocation, setSelectedLocation] = useState<MosqueLocation | ''>('');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedJamaah, setSelectedJamaah] = useState<Jamaah | null>(null);
  const [sessionType, setSessionType] = useState<'Kelompok' | 'Desa' | 'Acara'>('Kelompok');
  const [status, setStatus] = useState<'hadir' | 'izin'>('hadir');
  const [reason, setReason] = useState('');
  const [loading, setLoading] = useState(false);
  const [success, setSuccess] = useState(false);
  const [showAlreadyAttended, setShowAlreadyAttended] = useState(false);

  const { data: jamaahList } = useFirestoreQuery<Jamaah>(db, 'jamaah', 
    selectedLocation ? [where('location', '==', selectedLocation)] : []
  );

  const filteredJamaah = useMemo(() => {
    if (!searchTerm || selectedJamaah) return [];
    return jamaahList.filter(j => 
      j.name.toLowerCase().includes(searchTerm.toLowerCase()) || 
      j.memberId.toLowerCase().includes(searchTerm.toLowerCase())
    ).slice(0, 5);
  }, [jamaahList, searchTerm, selectedJamaah]);

  const handleSubmit = async () => {
    if (!selectedJamaah || !db) return;
    setLoading(true);
    try {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const startOfDay = today.getTime();
      
      const nextDay = new Date(today);
      nextDay.setDate(today.getDate() + 1);
      const endOfDay = nextDay.getTime();

      const attendanceRef = collection(db, 'attendance');
      const q = query(
        attendanceRef, 
        where('jamaahId', '==', selectedJamaah.id),
        where('date', '>=', startOfDay),
        where('date', '<', endOfDay)
      );
      
      const querySnapshot = await getDocs(q);
      if (!querySnapshot.empty) {
        setShowAlreadyAttended(true);
        setLoading(false);
        return;
      }

      const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
      const dayName = days[new Date().getDay()];

      await addDoc(collection(db, 'attendance'), {
        jamaahId: selectedJamaah.id,
        jamaahName: selectedJamaah.name,
        location: selectedJamaah.location,
        category: selectedJamaah.category,
        date: Date.now(),
        sessionType,
        day: dayName,
        status,
        reason: status === 'izin' ? reason : ''
      });
      setSuccess(true);
      setSelectedJamaah(null);
      setSearchTerm('');
      setStatus('hadir');
      setReason('');
    } catch (error) {
      console.error(error);
      alert('Gagal mengirim absensi');
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
        
        <div className="flex items-center justify-between mb-10">
          <button onClick={onBack} className="p-3 bg-slate-50 rounded-2xl text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition-all"><X className="w-6 h-6" /></button>
          <div className="text-center">
            <h2 className="text-lg font-black uppercase tracking-widest text-emerald-600 leading-none">Absensi</h2>
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-[0.3em] mt-1">Sambung Jamaah</p>
          </div>
          <div className="w-12"></div>
        </div>

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
                      {filteredJamaah.map(j => (
                        <button
                          key={j.id}
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

function LoginView({ onAttendanceMode }: { onAttendanceMode: () => void }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [isRegister, setIsRegister] = useState(false);
  const [displayName, setDisplayName] = useState('');
  const [role, setRole] = useState<UserRole>('pengurus');
  const [location, setLocation] = useState<MosqueLocation>('Kramat Batu');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const { db, auth } = useFirebase();

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    if (!auth) {
      setError('Firebase not initialized');
      setLoading(false);
      return;
    }

    try {
      if (isRegister) {
        const userCredential = await createUserWithEmailAndPassword(auth, email, password);
        if (db) {
          await setDoc(doc(db, 'users', userCredential.user.uid), {
            uid: userCredential.user.uid,
            email,
            displayName,
            role,
            location: role === 'pengurus' ? location : null,
            isVerified: false,
            createdAt: Date.now()
          });
          setError('Pendaftaran berhasil! Akun Anda sedang menunggu verifikasi dari admin.');
          setIsRegister(false);
        }
      } else {
        const userCredential = await signInWithEmailAndPassword(auth, email, password);
        if (db) {
          const userDoc = await getDoc(doc(db, 'users', userCredential.user.uid));
          const userData = userDoc.data() as UserProfile;
          if (!userData?.isVerified) {
            await signOut(auth);
            setError('Akun Anda belum diverifikasi oleh admin. Silakan hubungi admin untuk aktivasi.');
            return;
          }
        }
      }
    } catch (err: any) {
      setError(err.message || 'Gagal masuk');
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
          <h1 className="text-2xl font-bold text-slate-900">Sistem Manajemen Masjid</h1>
          <p className="text-slate-500 mt-2">Silakan masuk ke akun Anda</p>
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
          <button 
            onClick={onAttendanceMode}
            className="flex items-center justify-center gap-3 w-full py-4 rounded-2xl bg-slate-900 text-white font-black uppercase tracking-wider text-xs hover:bg-slate-800 transition-all active:scale-95 shadow-xl shadow-slate-200"
          >
            <Clock className="w-5 h-5 text-emerald-400" />
            Mode Absensi Jamaah
          </button>
          
          <button
            onClick={() => setIsRegister(!isRegister)}
            className="w-full py-2 text-sm font-bold text-slate-500 hover:text-emerald-600 transition-colors"
          >
            {isRegister ? 'Sudah punya akun? Masuk' : 'Belum punya akun? Daftar'}
          </button>
        </div>
      </motion.div>
    </div>
  );
}

// --- Dashboard Sub-Views ---

function ActivityImageSlider({ images }: { images: string[] }) {
  const [index, setIndex] = useState(0);

  useEffect(() => {
    if (!images || images.length <= 1) return;
    const interval = setInterval(() => {
      setIndex((prev) => (prev + 1) % images.length);
    }, 4000);
    return () => clearInterval(interval);
  }, [images]);

  if (!images || images.length === 0) {
    return (
      <div className="w-full h-full flex items-center justify-center text-slate-400 bg-slate-100">
        <Calendar className="w-12 h-12" />
      </div>
    );
  }

  return (
    <div className="relative w-full h-full overflow-hidden group">
      <AnimatePresence initial={false} mode="wait">
        <motion.img
          key={index}
          src={images[index]}
          initial={{ opacity: 0, scale: 1.1 }}
          animate={{ opacity: 1, scale: 1 }}
          exit={{ opacity: 0 }}
          transition={{ duration: 1 }}
          className="absolute inset-0 w-full h-full object-cover"
        />
      </AnimatePresence>
      {images.length > 1 && (
        <div className="absolute bottom-3 left-1/2 -translate-x-1/2 flex gap-1.5 z-10">
          {images.map((_, i) => (
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
  const filter = useMemo(() => profile.role === 'pengurus' ? [where('location', '==', profile.location)] : [], [profile.role, profile.location]);
  const { data: jamaah } = useFirestoreQuery<Jamaah>(db, 'jamaah', filter);
  const { data: allAssets } = useFirestoreQuery<Asset>(db, 'assets', filter);
  const { data: activities } = useFirestoreQuery<Activity>(db, 'activities', filter);
  const { data: stats } = useFirestoreQuery<FacilityStat>(db, 'facility_stats');

  const barangAssets = useMemo(() => allAssets.filter(a => !a.assetType || a.assetType === 'barang'), [allAssets]);
  const tanahAssets = useMemo(() => allAssets.filter(a => a.assetType === 'tanah'), [allAssets]);
  const totalTanahArea = useMemo(() => tanahAssets.reduce((sum, t) => sum + (t.areaSize || 0), 0), [tanahAssets]);
  const totalKK = useMemo(() => jamaah.filter(j => j.isKK).length, [jamaah]);

  const jamaahCategories = useMemo(() => {
    const counts: Record<string, number> = { 'UMUM': 0, 'ACR': 0, 'APR': 0, 'GPN': 0 };
    jamaah.forEach(j => {
      if (counts[j.category] !== undefined) counts[j.category]++;
    });
    return Object.entries(counts).map(([name, value]) => ({ name, value }));
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
            key={stat.id} 
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
          {activities.sort((a,b) => b.date - a.date).slice(0, 3).map((activity) => (
            <motion.div 
              key={activity.id}
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

function JamaahView({ profile, formTrigger, onFormTriggered }: { profile: UserProfile, formTrigger?: string | null, onFormTriggered?: () => void }) {
  const { db } = useFirebase();
  const filter = useMemo(() => profile.role === 'pengurus' ? [where('location', '==', profile.location)] : [], [profile.role, profile.location]);
  const { data: jamaah, loading } = useFirestoreQuery<Jamaah>(db, 'jamaah', filter);
  const [searchTerm, setSearchTerm] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [editingJamaah, setEditingJamaah] = useState<Jamaah | null>(null);
  const [base64Image, setBase64Image] = useState<string | null>(null);
  const [isKK, setIsKK] = useState(true);
  const [selectedKKId, setSelectedKKId] = useState<string>('');
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);

  useEffect(() => {
    if (formTrigger === 'jamaah') {
      setShowForm(true);
      setEditingJamaah(null);
      setBase64Image(null);
      setIsKK(true);
      setSelectedKKId('');
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
    j.phone.includes(searchTerm) ||
    j.memberId?.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const headsOfFamily = jamaah.filter(j => j.isKK && (profile.role === 'admin' ? true : j.location === profile.location));

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!db) return;
    const formData = new FormData(e.currentTarget);
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
        const kkInLocation = jamaah.filter(j => j.location === location && j.isKK);
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
          alert('Pilih Kepala Keluarga terlebih dahulu');
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

    const data = {
      memberId,
      name: formData.get('name') as string,
      phone: formData.get('phone') as string,
      address: formData.get('address') as string,
      location,
      category: formData.get('category') as JamaahCategory,
      isKK: isKK,
      kkId: finalKKId,
      familyOrder,
      photoUrl: base64Image || editingJamaah?.photoUrl || null,
      registeredAt: editingJamaah ? editingJamaah.registeredAt : Date.now()
    };

    if (editingJamaah) {
      await updateDoc(doc(db, 'jamaah', editingJamaah.id), data as any);
    } else {
      await addDoc(collection(db, 'jamaah'), data);
    }
    closeForm();
  };

  const closeForm = () => {
    setShowForm(false);
    setEditingJamaah(null);
    setBase64Image(null);
    setIsKK(true);
    setSelectedKKId('');
  };

  const openEdit = (j: Jamaah) => {
    setEditingJamaah(j);
    setBase64Image(j.photoUrl || null);
    setIsKK(j.isKK);
    setSelectedKKId(j.kkId || '');
    setShowForm(true);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-slate-900">Data Jamaah</h2>
          <p className="text-slate-500">Kelola data jamaah {profile.role === 'pengurus' ? `di ${profile.location}` : 'seluruh lokasi'}</p>
        </div>
        <button 
          onClick={() => setShowForm(true)}
          className="bg-emerald-600 text-white px-6 py-2.5 rounded-xl font-semibold flex items-center justify-center gap-2 hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200"
        >
          <Plus className="w-5 h-5" />
          Tambah Jamaah
        </button>
      </div>

      <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
        <div className="p-4 border-b border-slate-100 bg-slate-50/50">
          <div className="relative max-w-md">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
            <input 
              type="text" 
              placeholder="Cari jamaah..." 
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
                <th className="px-6 py-4">Kategori</th>
                <th className="px-6 py-4">Telepon</th>
                <th className="px-6 py-4">Lokasi</th>
                <th className="px-6 py-4 text-center">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                <tr><td colSpan={6} className="px-6 py-8 text-center text-slate-400">Memuat data...</td></tr>
              ) : filteredJamaah.length === 0 ? (
                <tr><td colSpan={6} className="px-6 py-8 text-center text-slate-400">Tidak ada data ditemukan</td></tr>
              ) : filteredJamaah.map((j) => (
                <tr key={j.id} className="hover:bg-slate-50 transition-colors">
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
                      <span className="font-semibold text-slate-900">{j.name}</span>
                      {j.isKK && <span className="text-[10px] text-slate-400 font-medium">Kepala Keluarga</span>}
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
                  <td className="px-6 py-4 text-slate-600 text-sm">{j.phone}</td>
                  <td className="px-6 py-4">
                    <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-700 text-xs font-semibold">
                      <MapPin className="w-3 h-3" /> {j.location}
                    </span>
                  </td>
                  <td className="px-6 py-4">
                    <div className="flex gap-3">
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
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <DeleteConfirmation 
        isOpen={!!deleteId}
        onClose={() => setDeleteId(null)}
        onConfirm={async () => {
          if (!deleteId || !db) return;
          setIsDeleting(true);
          try {
            await deleteDoc(doc(db, 'jamaah', deleteId));
            setDeleteId(null);
          } finally {
            setIsDeleting(false);
          }
        }}
        loading={isDeleting}
        title="Hapus Data Jamaah"
        message="Apakah Anda yakin ingin menghapus data jamaah ini? Data yang dihapus tidak dapat dikembalikan."
      />

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
              className="bg-white rounded-3xl p-8 w-full max-w-lg shadow-2xl"
            >
              <h3 className="text-xl font-bold mb-6">{editingJamaah ? 'Edit Data Jamaah' : 'Tambah Jamaah Baru'}</h3>
              <form onSubmit={handleSubmit} className="space-y-4">
                <div className="flex justify-center mb-6">
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
                <div>
                  <label className="block text-sm font-medium mb-1">Nama</label>
                  <input name="name" defaultValue={editingJamaah?.name} required className="w-full px-4 py-2 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none" />
                </div>
                <div className="flex items-center gap-3 p-3 bg-slate-50 rounded-xl">
                  <input 
                    type="checkbox" 
                    id="isKK" 
                    checked={isKK} 
                    onChange={(e) => setIsKK(e.target.checked)}
                    className="w-4 h-4 text-emerald-600"
                  />
                  <label htmlFor="isKK" className="text-sm font-semibold text-slate-700">Kepala Keluarga (KK)</label>
                </div>
                {!isKK && (
                  <div>
                    <label className="block text-sm font-medium mb-1">Pilih Kepala Keluarga</label>
                    <select 
                      value={selectedKKId} 
                      onChange={(e) => setSelectedKKId(e.target.value)}
                      required={!isKK}
                      className="w-full px-4 py-2 rounded-xl border outline-none"
                    >
                      <option value="">Pilih KK...</option>
                      {headsOfFamily.filter(kk => kk.id !== editingJamaah?.id).map(kk => (
                        <option key={kk.id} value={kk.memberId}>{kk.memberId} - {kk.name}</option>
                      ))}
                    </select>
                  </div>
                )}
                <div>
                  <label className="block text-sm font-medium mb-1">Kategori</label>
                  <select name="category" defaultValue={editingJamaah?.category || 'UMUM'} className="w-full px-4 py-2 rounded-xl border outline-none">
                    <option value="UMUM">UMUM</option>
                    <option value="ACR">ACR (Anak Caberawit)</option>
                    <option value="APR">APR (Anak Pra Remaja)</option>
                    <option value="GPN">GPN (Generus Pra Nikah)</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-medium mb-1">Telepon</label>
                  <input name="phone" defaultValue={editingJamaah?.phone} required className="w-full px-4 py-2 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none" />
                </div>
                <div>
                  <label className="block text-sm font-medium mb-1">Alamat</label>
                  <textarea name="address" defaultValue={editingJamaah?.address} required className="w-full px-4 py-2 rounded-xl border focus:ring-2 focus:ring-emerald-500 outline-none" />
                </div>
                {profile.role === 'admin' && (
                  <div>
                    <label className="block text-sm font-medium mb-1">Lokasi</label>
                    <select name="location" defaultValue={editingJamaah?.location} className="w-full px-4 py-2 rounded-xl border outline-none">
                      {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(l => (
                        <option key={l} value={l}>{l}</option>
                      ))}
                    </select>
                  </div>
                )}
                <div className="flex gap-3 pt-4">
                  <button type="button" onClick={closeForm} className="flex-1 px-6 py-2.5 rounded-xl border hover:bg-slate-50 transition-all font-semibold">Batal</button>
                  <button type="submit" className="flex-1 px-6 py-2.5 rounded-xl bg-emerald-600 text-white hover:bg-emerald-700 transition-all font-semibold shadow-lg shadow-emerald-200">Simpan</button>
                </div>
              </form>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

function InventarisView({ profile, formTrigger, onFormTriggered, assetType = 'barang' }: { profile: UserProfile, formTrigger?: string | null, onFormTriggered?: () => void, assetType?: 'barang' | 'tanah' }) {
  const { db } = useFirebase();
  const baseFilter = useMemo(() => profile.role === 'pengurus' ? [where('location', '==', profile.location)] : [], [profile.role, profile.location]);
  const { data: allAssets, loading } = useFirestoreQuery<Asset>(db, 'assets', baseFilter);
  const assets = useMemo(() => allAssets.filter(a => {
    if (assetType === 'barang') return !a.assetType || a.assetType === 'barang';
    return a.assetType === 'tanah';
  }), [allAssets, assetType]);
  const [showForm, setShowForm] = useState(false);
  const [editingAsset, setEditingAsset] = useState<Asset | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
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
    if (!db) return;
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

    if (editingAsset) {
      await updateDoc(doc(db, 'assets', editingAsset.id), data as any);
    } else {
      await addDoc(collection(db, 'assets'), data);
    }
    closeForm();
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

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {assets.map((asset) => (
          <motion.div 
            key={asset.id}
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

      <DeleteConfirmation 
        isOpen={!!deleteId}
        onClose={() => setDeleteId(null)}
        onConfirm={async () => {
          if (!deleteId || !db) return;
          setIsDeleting(true);
          try {
            await deleteDoc(doc(db, 'assets', deleteId));
            setDeleteId(null);
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
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50 overflow-y-auto">
            <motion.div initial={{ scale: 0.95 }} animate={{ scale: 1 }} className="bg-white rounded-3xl p-8 w-full max-w-lg shadow-2xl my-8">
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
                  <button type="button" onClick={closeForm} className="flex-1 px-6 py-2.5 rounded-xl border font-semibold text-slate-600 hover:bg-slate-50 transition-all">Batal</button>
                  <button type="submit" className="flex-1 px-6 py-2.5 rounded-xl bg-emerald-600 text-white font-semibold hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-100">Simpan</button>
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
  const { db } = useFirebase();
  const filter = useMemo(() => profile.role === 'pengurus' ? [where('location', '==', profile.location)] : [], [profile.role, profile.location]);
  const { data: activities } = useFirestoreQuery<Activity>(db, 'activities', filter);
  const [showForm, setShowForm] = useState(false);
  const [editingActivity, setEditingActivity] = useState<Activity | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
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
    if (!db) return;
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

    if (editingActivity) {
      await updateDoc(doc(db, 'activities', editingActivity.id), data as any);
    } else {
      await addDoc(collection(db, 'activities'), data);
    }
    closeForm();
  };

  const closeForm = () => {
    setShowForm(false);
    setEditingActivity(null);
    setBase64Images([]);
  };

  const openEdit = (a: Activity) => {
    setEditingActivity(a);
    setBase64Images(a.imageUrls || []);
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

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
        {activities.map((activity) => (
          <div key={activity.id} className="bg-white rounded-3xl overflow-hidden border border-slate-100 shadow-sm flex flex-col">
            <div className="h-56 relative group">
              <div className="w-full h-full flex overflow-x-auto snap-x snap-mandatory scrollbar-hide">
                {activity.imageUrls && activity.imageUrls.length > 0 ? (
                  activity.imageUrls.map((url, idx) => (
                    <img key={idx} src={url} className="w-full h-full object-cover flex-shrink-0 snap-center" />
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
        ))}
      </div>

      <AnimatePresence>
        {showForm && (
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
            <motion.div initial={{ scale: 0.95 }} animate={{ scale: 1 }} className="bg-white rounded-3xl p-8 w-full max-w-lg shadow-2xl">
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
                  <button type="button" onClick={closeForm} className="flex-1 px-6 py-2.5 rounded-xl border font-semibold">Batal</button>
                  <button type="submit" className="flex-1 px-6 py-2.5 rounded-xl bg-emerald-600 text-white font-semibold">Simpan</button>
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
          if (!deleteId || !db) return;
          setIsDeleting(true);
          try {
            await deleteDoc(doc(db, 'activities', deleteId));
            setDeleteId(null);
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
  const { db } = useFirebase();
  const { data: stats } = useFirestoreQuery<FacilityStat>(db, 'facility_stats');
  const [showForm, setShowForm] = useState(false);
  const [editingStat, setEditingStat] = useState<FacilityStat | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);

  const updateUsage = async (id: string, current: number, capacity: number, delta: number) => {
    if (!db) return;
    const newVal = Math.max(0, Math.min(capacity, current + delta));
    await updateDoc(doc(db, 'facility_stats', id), {
      currentUsage: newVal,
      updatedAt: Date.now()
    });
  };

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!db) return;
    const formData = new FormData(e.currentTarget);
    const data = {
      name: formData.get('name') as string,
      capacity: Number(formData.get('capacity')),
      currentUsage: editingStat ? editingStat.currentUsage : 0,
      updatedAt: Date.now()
    };

    if (editingStat) {
      await updateDoc(doc(db, 'facility_stats', editingStat.id), data);
    } else {
      await addDoc(collection(db, 'facility_stats'), data);
    }
    closeForm();
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

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
        {stats.map((stat) => (
          <div key={stat.id} className="bg-white p-8 rounded-3xl border border-slate-100 shadow-sm">
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
                  <button type="button" onClick={closeForm} className="flex-1 px-6 py-2.5 rounded-xl border font-semibold">Batal</button>
                  <button type="submit" className="flex-1 px-6 py-2.5 rounded-xl bg-emerald-600 text-white font-semibold">Simpan</button>
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
          if (!deleteId || !db) return;
          setIsDeleting(true);
          try {
            await deleteDoc(doc(db, 'facility_stats', deleteId));
            setDeleteId(null);
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
  const { db } = useFirebase();
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

  const { data: attendanceData, loading } = useFirestoreQuery<Attendance>(db, 'attendance', filter);

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
    if (!db) return;
    await updateDoc(doc(db, 'attendance', id), {
      isConfirmed: !current
    });
  };

  const groupedData = useMemo(() => {
    const groups: Record<string, Attendance[]> = {};
    filteredData.forEach(a => {
      const dateKey = new Date(a.date).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
      if (!groups[dateKey]) groups[dateKey] = [];
      groups[dateKey].push(a);
    });
    return groups;
  }, [filteredData]);

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
            {Object.entries(stats.byCategory).map(([cat, count]) => (
              <div key={cat} className="px-3 py-2 bg-slate-50 rounded-xl flex items-center gap-3">
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
            {Object.entries(stats.bySession).map(([sess, count]) => (
              <div key={sess} className="px-3 py-2 bg-slate-50 rounded-xl flex items-center gap-3">
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
                    <motion.div animate={{ rotate: 360 }} transition={{ repeat: Infinity, duration: 1, ease: "linear" }} className="w-6 h-6 border-2 border-emerald-500 border-t-transparent rounded-full mx-auto" />
                  </td>
                </tr>
              ) : Object.keys(groupedData).length > 0 ? (
                Object.entries(groupedData).map(([date, items]: [string, Attendance[]]) => (
                  <React.Fragment key={date}>
                    <tr className="bg-slate-50/30">
                      <td colSpan={10} className="px-6 py-3 text-[10px] font-black text-emerald-600 uppercase tracking-widest bg-emerald-50/30">
                        {date} ({items.length} Absensi)
                      </td>
                    </tr>
                    {items.map((a) => (
                      <tr key={a.id} className="hover:bg-slate-50 transition-colors group">
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
      </div>

      <AnimatePresence>
        {editingAttendance && (
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
            <motion.div initial={{ scale: 0.95 }} animate={{ scale: 1 }} className="bg-white rounded-3xl p-8 w-full max-w-lg shadow-2xl">
              <h3 className="text-xl font-bold mb-6">Edit Data Absensi</h3>
              <form 
                onSubmit={async (e) => {
                  e.preventDefault();
                  if (!db || !editingAttendance) return;
                  setIsUpdating(true);
                  const formData = new FormData(e.currentTarget);
                  try {
                    await updateDoc(doc(db, 'attendance', editingAttendance.id), {
                      sessionType: formData.get('sessionType'),
                      status: formData.get('status'),
                      reason: formData.get('reason')
                    });
                    setEditingAttendance(null);
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
          if (!deleteId || !db) return;
          setIsDeleting(true);
          try {
            await deleteDoc(doc(db, 'attendance', deleteId));
            setDeleteId(null);
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

function UsersView() {
  const { db } = useFirebase();
  const { data: users, loading } = useFirestoreQuery<UserProfile>(db, 'users');
  const [editingUser, setEditingUser] = useState<UserProfile | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);

  const toggleVerify = async (uid: string, current: boolean) => {
    if (!db) return;
    await updateDoc(doc(db, 'users', uid), {
      isVerified: !current
    });
  };

  const handleUpdate = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!db || !editingUser) return;
    const formData = new FormData(e.currentTarget);
    await updateDoc(doc(db, 'users', editingUser.uid), {
      displayName: formData.get('displayName'),
      role: formData.get('role'),
      location: formData.get('role') === 'pengurus' ? formData.get('location') : null
    });
    setEditingUser(null);
  };

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-bold text-slate-900">Manajemen Pengguna</h2>
          <p className="text-slate-500">Kelola hak akses admin dan pengurus</p>
        </div>
      </div>

      <div className="bg-white rounded-3xl overflow-hidden border border-slate-100 shadow-sm">
        <table className="w-full text-left">
          <thead className="bg-slate-50 border-b border-slate-100">
            <tr>
              <th className="px-6 py-4 text-sm font-bold text-slate-500 uppercase tracking-wider">Nama</th>
              <th className="px-6 py-4 text-sm font-bold text-slate-500 uppercase tracking-wider">Email</th>
              <th className="px-6 py-4 text-sm font-bold text-slate-500 uppercase tracking-wider">Role</th>
              <th className="px-6 py-4 text-sm font-bold text-slate-500 uppercase tracking-wider">Lokasi</th>
              <th className="px-6 py-4 text-sm font-bold text-slate-500 uppercase tracking-wider">Verifikasi</th>
              <th className="px-6 py-4 text-sm font-bold text-slate-500 uppercase tracking-wider">Aksi</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-50">
            {users.map((u) => (
              <tr key={u.uid} className="hover:bg-slate-50 transition-colors">
                <td className="px-6 py-4 font-bold text-slate-900">{u.displayName}</td>
                <td className="px-6 py-4 text-slate-600">{u.email}</td>
                <td className="px-6 py-4">
                  <span className={cn(
                    "px-3 py-1 rounded-full text-xs font-bold uppercase",
                    u.role === 'admin' ? "bg-purple-100 text-purple-700" : "bg-blue-100 text-blue-700"
                  )}>
                    {u.role}
                  </span>
                </td>
                <td className="px-6 py-4 text-slate-600 font-medium">{u.location || '-'}</td>
                <td className="px-6 py-4">
                  <button 
                    onClick={() => toggleVerify(u.uid, u.isVerified)}
                    className={cn(
                      "px-3 py-1 rounded-full text-[10px] font-black uppercase transition-all",
                      u.isVerified 
                        ? "bg-emerald-100 text-emerald-700" 
                        : "bg-amber-100 text-amber-700 hover:bg-emerald-100"
                    )}
                  >
                    {u.isVerified ? 'Terverifikasi' : 'Belum'}
                  </button>
                </td>
                <td className="px-6 py-4">
                  <div className="flex gap-3">
                    <button onClick={() => setEditingUser(u)} className="text-emerald-600 hover:text-emerald-700 font-bold text-sm">Edit</button>
                    <button onClick={() => setDeleteId(u.uid)} className="text-red-500 hover:text-red-600 font-bold text-sm">Hapus</button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <AnimatePresence>
        {editingUser && (
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
            <motion.div initial={{ scale: 0.95 }} animate={{ scale: 1 }} className="bg-white rounded-3xl p-8 w-full max-w-lg shadow-2xl">
              <h3 className="text-xl font-bold mb-6">Edit Profil Pengguna</h3>
              <form onSubmit={handleUpdate} className="space-y-4">
                <div>
                  <label className="block text-sm font-medium mb-1">Nama Tampilan</label>
                  <input name="displayName" defaultValue={editingUser.displayName} required className="w-full px-4 py-2 rounded-xl border outline-none" />
                </div>
                <div>
                  <label className="block text-sm font-medium mb-1">Role</label>
                  <select name="role" defaultValue={editingUser.role} className="w-full px-4 py-2 rounded-xl border outline-none">
                    <option value="admin">Admin</option>
                    <option value="pengurus">Pengurus</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-medium mb-1">Lokasi (Khusus Pengurus)</label>
                  <select name="location" defaultValue={editingUser.location} className="w-full px-4 py-2 rounded-xl border outline-none">
                    <option value="">Pilih Lokasi</option>
                    {['Kramat Batu', 'Karya Utama', 'Radio Dalam', 'Cipete', 'Antena'].map(l => (
                      <option key={l} value={l}>{l}</option>
                    ))}
                  </select>
                </div>
                <div className="flex gap-3 pt-4">
                  <button type="button" onClick={() => setEditingUser(null)} className="flex-1 px-6 py-2.5 rounded-xl border font-semibold">Batal</button>
                  <button type="submit" className="flex-1 px-6 py-2.5 rounded-xl bg-emerald-600 text-white font-semibold">Simpan</button>
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
          if (!deleteId || !db) return;
          setIsDeleting(true);
          try {
            await deleteDoc(doc(db, 'users', deleteId));
            setDeleteId(null);
          } finally {
            setIsDeleting(false);
          }
        }}
        loading={isDeleting}
        title="Hapus Pengguna"
        message="Hapus akses pengguna ini? Pengguna tidak akan dapat mengakses dashboard ini lagi."
      />
    </div>
  );
}

// --- Main Layout ---

function DashboardContent() {
  const { user, db, auth } = useFirebase();
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [activeTab, setActiveTab] = useState('overview');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [formTrigger, setFormTrigger] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [isAttendanceMode, setIsAttendanceMode] = useState(false);

  useEffect(() => {
    if (user && db) {
      const fetchProfile = async () => {
        const docRef = doc(db, 'users', user.uid);
        const docSnap = await getDoc(docRef);
        if (docSnap.exists()) {
          setProfile(docSnap.data() as UserProfile);
        }
        setLoading(false);
      };
      fetchProfile();
    } else if (!user) {
      setLoading(false);
    }
  }, [user, db]);

  if (loading) return (
    <div className="min-h-screen flex items-center justify-center">
      <motion.div animate={{ rotate: 360 }} transition={{ repeat: Infinity, duration: 1, ease: "linear" }} className="w-8 h-8 border-4 border-emerald-500 border-t-transparent rounded-full" />
    </div>
  );

  if (isAttendanceMode) {
    return <PublicAttendanceView onBack={() => setIsAttendanceMode(false)} />;
  }

  if (!user || !profile) return <LoginView onAttendanceMode={() => setIsAttendanceMode(true)} />;

  if (!profile.isVerified) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50 p-4 text-center">
        <div className="bg-white p-10 rounded-3xl shadow-xl max-w-sm border border-slate-100">
          <AlertCircle className="w-16 h-16 text-amber-500 mx-auto mb-6" />
          <h2 className="text-2xl font-bold text-slate-900 mb-4 tracking-tight">Akun Belum Aktif</h2>
          <p className="text-slate-500 mb-8 font-medium leading-relaxed">Akun Anda sedang menunggu verifikasi dari admin. Silakan hubungi admin untuk aktivasi agar dapat mengakses dashboard.</p>
          <button 
            onClick={() => auth && signOut(auth)} 
            className="w-full py-4 bg-slate-900 text-white rounded-2xl font-black uppercase tracking-wider hover:bg-slate-800 transition-all active:scale-95 shadow-xl shadow-slate-200"
          >
            Keluar
          </button>
        </div>
      </div>
    );
  }

  const renderContent = () => {
    switch (activeTab) {
      case 'overview': return <Overview profile={profile} />;
      case 'jamaah': return <JamaahView profile={profile} formTrigger={formTrigger} onFormTriggered={() => setFormTrigger(null)} />;
      case 'inventaris': return <InventarisView profile={profile} formTrigger={formTrigger} onFormTriggered={() => setFormTrigger(null)} assetType="barang" />;
      case 'tanah': return <InventarisView profile={profile} formTrigger={formTrigger} onFormTriggered={() => setFormTrigger(null)} assetType="tanah" />;
      case 'activities': return <ActivitiesView profile={profile} />;
      case 'facilities': return <FacilityView profile={profile} />;
      case 'attendance_report': return <AttendanceReportView profile={profile} />;
      case 'users': return <UsersView />;
      default: return <Overview profile={profile} />;
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
          <SidebarItem icon={Package} label="Inventaris" active={activeTab === 'inventaris'} onClick={() => setActiveTab('inventaris')} />
          <SidebarItem icon={MapIcon} label="Tanah Sabilillah" active={activeTab === 'tanah'} onClick={() => setActiveTab('tanah')} />
          <SidebarItem icon={Calendar} label="Kegiatan" active={activeTab === 'activities'} onClick={() => setActiveTab('activities')} />
          <SidebarItem icon={Clock} label="Fasilitas" active={activeTab === 'facilities'} onClick={() => setActiveTab('facilities')} />
          <SidebarItem icon={ClipboardList} label="Laporan Absensi" active={activeTab === 'attendance_report'} onClick={() => setActiveTab('attendance_report')} />
          {profile.role === 'admin' && (
            <SidebarItem icon={Plus} label="Pengguna" active={activeTab === 'users'} onClick={() => setActiveTab('users')} />
          )}
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
                <SidebarItem icon={Package} label="Inventaris" active={activeTab === 'inventaris'} onClick={() => { setActiveTab('inventaris'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={MapIcon} label="Tanah Sabilillah" active={activeTab === 'tanah'} onClick={() => { setActiveTab('tanah'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={Calendar} label="Kegiatan" active={activeTab === 'activities'} onClick={() => { setActiveTab('activities'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={Clock} label="Fasilitas" active={activeTab === 'facilities'} onClick={() => { setActiveTab('facilities'); setIsMobileMenuOpen(false); }} />
                <SidebarItem icon={ClipboardList} label="Laporan Absensi" active={activeTab === 'attendance_report'} onClick={() => { setActiveTab('attendance_report'); setIsMobileMenuOpen(false); }} />
                {profile.role === 'admin' && (
                  <SidebarItem icon={Plus} label="Pengguna" active={activeTab === 'users'} onClick={() => { setActiveTab('users'); setIsMobileMenuOpen(false); }} />
                )}
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
      <DashboardContent />
    </FirebaseProvider>
  );
}
