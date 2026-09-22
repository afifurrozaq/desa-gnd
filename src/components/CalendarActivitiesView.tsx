import React, { useState, useMemo, useCallback } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Calendar as CalendarIcon, 
  ChevronLeft, 
  ChevronRight, 
  Plus, 
  Clock, 
  MapPin, 
  User, 
  Search, 
  Filter, 
  Trash2, 
  Edit, 
  Eye, 
  Image as ImageIcon, 
  Layers, 
  List, 
  Grid, 
  CalendarDays, 
  X, 
  Loader2, 
  CheckCircle2, 
  Sparkles,
  Tag,
  ArrowRight,
  ExternalLink
} from 'lucide-react';
import { UserProfile, MosqueLocation, Activity } from '../types';
import { useFirebase } from './FirebaseProvider';
import { useToast } from './ToastContext';
import { useDataQuery, where } from '../hooks/useDataQuery';
import { saveData, deleteData } from '../lib/dataService';
import { DeleteConfirmation } from './DeleteConfirmation';
import { cn } from '../lib/utils';
import { parseImageUrls } from '../App';
import { 
  MONTH_NAMES, 
  DAY_NAMES, 
  DAY_FULL_NAMES, 
  MOSQUE_LOCATIONS, 
  ACTIVITY_CATEGORIES 
} from '../constants';

function getCategoryMeta(catName?: string) {
  if (!catName) return ACTIVITY_CATEGORIES[ACTIVITY_CATEGORIES.length - 1];
  const found = ACTIVITY_CATEGORIES.find(c => c.id.toLowerCase() === catName.toLowerCase() || c.label.toLowerCase() === catName.toLowerCase() || catName.toLowerCase().includes(c.id.toLowerCase()));
  return found || {
    id: catName,
    label: catName,
    bg: 'bg-slate-50',
    text: 'text-slate-700',
    border: 'border-slate-200',
    dot: 'bg-slate-500',
    badge: 'bg-slate-600 text-white',
    width: 'col-span-1 md:col-span-1',
    height: 'min-h-[180px]',
    cardWidth: 'w-full',
    cardHeight: 'h-auto'
  };
}

export function CalendarActivitiesView({ profile }: { profile: UserProfile }) {
  const { accessToken, spreadsheetId } = useFirebase();
  const { showToast } = useToast();

  // Current selected month & year for calendar navigation
  const today = new Date();
  const [currentYear, setCurrentYear] = useState<number>(today.getFullYear());
  const [currentMonth, setCurrentMonth] = useState<number>(today.getMonth()); // 0-indexed: 0 = Jan, 8 = Sep

  // View modes: 'calendar' | 'agenda' | 'cards'
  const [viewMode, setViewMode] = useState<'calendar' | 'agenda' | 'cards'>('calendar');

  // Filters
  const [selectedLocation, setSelectedLocation] = useState<string>(
    profile.role === 'pengurus' && profile.location ? profile.location : 'Semua Lokasi'
  );
  const [selectedCategory, setSelectedCategory] = useState<string>('Semua Kategori');
  const [searchTerm, setSearchTerm] = useState<string>('');

  // Selected date popover / day detail
  const [selectedDayEvents, setSelectedDayEvents] = useState<{ date: Date; events: Activity[] } | null>(null);

  // Form & Modals state
  const [showForm, setShowForm] = useState(false);
  const [editingActivity, setEditingActivity] = useState<Activity | null>(null);
  const [formPreselectedDate, setFormPreselectedDate] = useState<string>(''); // YYYY-MM-DD
  const [selectedDetailActivity, setSelectedDetailActivity] = useState<Activity | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [base64Images, setBase64Images] = useState<string[]>([]);

  // Query filter by location if pengurus
  const filter = useMemo(() => 
    (profile.role === 'pengurus' && profile.location && (profile.location as string) !== 'Seluruh Lokasi') 
      ? [where('location', '==', profile.location)] 
      : [], 
    [profile.role, profile.location]
  );

  const { data: activities, loading, refresh } = useDataQuery<Activity>('activities', filter);

  // Parse activity date helper
  const getActivityDate = useCallback((act: Activity): Date => {
    if (!act.date) return new Date();
    if (typeof act.date === 'number') {
      // Check if unix in seconds or ms
      const d = act.date < 10000000000 ? act.date * 1000 : act.date;
      return new Date(d);
    }
    const parsed = new Date(act.date);
    return isNaN(parsed.getTime()) ? new Date() : parsed;
  }, []);

  // Filter activities
  const filteredActivities = useMemo(() => {
    return activities.filter(act => {
      // Filter by location
      if (selectedLocation !== 'Semua Lokasi' && act.location !== selectedLocation) {
        return false;
      }
      // Filter by category
      if (selectedCategory !== 'Semua Kategori') {
        const cat = act.category || act.type || 'Lainnya';
        if (cat.toLowerCase() !== selectedCategory.toLowerCase()) {
          return false;
        }
      }
      // Filter by search
      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase();
        const matchTitle = String(act.title || '').toLowerCase().includes(q);
        const matchDesc = String(act.description || '').toLowerCase().includes(q);
        const matchLoc = String(act.location || '').toLowerCase().includes(q);
        const matchSpeaker = String(act.speaker || '').toLowerCase().includes(q);
        const matchCat = String(act.category || '').toLowerCase().includes(q);
        if (!matchTitle && !matchDesc && !matchLoc && !matchSpeaker && !matchCat) {
          return false;
        }
      }
      return true;
    });
  }, [activities, selectedLocation, selectedCategory, searchTerm]);

  // Activities strictly in the selected month & year
  const currentMonthActivities = useMemo(() => {
    return filteredActivities.filter(act => {
      const d = getActivityDate(act);
      return d.getFullYear() === currentYear && d.getMonth() === currentMonth;
    }).sort((a, b) => getActivityDate(a).getTime() - getActivityDate(b).getTime());
  }, [filteredActivities, currentYear, currentMonth, getActivityDate]);

  // Group activities in current month by date string 'YYYY-MM-DD'
  const activitiesByDateMap = useMemo(() => {
    const map = new Map<string, Activity[]>();
    currentMonthActivities.forEach(act => {
      const d = getActivityDate(act);
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      if (!map.has(key)) {
        map.set(key, []);
      }
      map.get(key)!.push(act);
    });
    return map;
  }, [currentMonthActivities, getActivityDate]);

  // Calendar matrix calculation for current month
  const calendarGrid = useMemo(() => {
    const firstDayOfMonth = new Date(currentYear, currentMonth, 1);
    const lastDayOfMonth = new Date(currentYear, currentMonth + 1, 0);
    const daysInMonth = lastDayOfMonth.getDate();

    // Day of week: 0 = Sunday, 1 = Monday, ... 6 = Saturday
    // We want Monday (1) as column 0, Sunday (0) as column 6
    let startDayOfWeek = firstDayOfMonth.getDay() - 1;
    if (startDayOfWeek === -1) startDayOfWeek = 6;

    const days = [];

    // Prev month padding
    const prevMonthLastDay = new Date(currentYear, currentMonth, 0).getDate();
    for (let i = startDayOfWeek - 1; i >= 0; i--) {
      const dayNum = prevMonthLastDay - i;
      const prevDate = new Date(currentYear, currentMonth - 1, dayNum);
      const dateKey = `${prevDate.getFullYear()}-${String(prevDate.getMonth() + 1).padStart(2, '0')}-${String(dayNum).padStart(2, '0')}`;
      days.push({
        dayNumber: dayNum,
        date: prevDate,
        dateKey,
        isCurrentMonth: false,
        events: [] as Activity[]
      });
    }

    // Current month days
    for (let dayNum = 1; dayNum <= daysInMonth; dayNum++) {
      const date = new Date(currentYear, currentMonth, dayNum);
      const dateKey = `${currentYear}-${String(currentMonth + 1).padStart(2, '0')}-${String(dayNum).padStart(2, '0')}`;
      const events = activitiesByDateMap.get(dateKey) || [];
      days.push({
        dayNumber: dayNum,
        date,
        dateKey,
        isCurrentMonth: true,
        events
      });
    }

    // Next month padding to fill grid (multiple of 7)
    const remaining = (7 - (days.length % 7)) % 7;
    for (let dayNum = 1; dayNum <= remaining; dayNum++) {
      const nextDate = new Date(currentYear, currentMonth + 1, dayNum);
      const dateKey = `${nextDate.getFullYear()}-${String(nextDate.getMonth() + 1).padStart(2, '0')}-${String(dayNum).padStart(2, '0')}`;
      days.push({
        dayNumber: dayNum,
        date: nextDate,
        dateKey,
        isCurrentMonth: false,
        events: [] as Activity[]
      });
    }

    return days;
  }, [currentYear, currentMonth, activitiesByDateMap]);

  // Quick stats for the month
  const monthStats = useMemo(() => {
    const total = currentMonthActivities.length;
    const nowTime = new Date().getTime();
    let upcoming = 0;
    let completed = 0;

    currentMonthActivities.forEach(act => {
      const d = getActivityDate(act);
      if (d.getTime() >= nowTime) {
        upcoming++;
      } else {
        completed++;
      }
    });

    return { total, upcoming, completed };
  }, [currentMonthActivities, getActivityDate]);

  // Navigation helpers
  const handlePrevMonth = () => {
    if (currentMonth === 0) {
      setCurrentMonth(11);
      setCurrentYear(prev => prev - 1);
    } else {
      setCurrentMonth(prev => prev - 1);
    }
  };

  const handleNextMonth = () => {
    if (currentMonth === 11) {
      setCurrentMonth(0);
      setCurrentYear(prev => prev + 1);
    } else {
      setCurrentMonth(prev => prev + 1);
    }
  };

  const handleToday = () => {
    const now = new Date();
    setCurrentYear(now.getFullYear());
    setCurrentMonth(now.getMonth());
  };

  // Open add activity form for a specific date
  const openAddFormForDate = (dateStr?: string) => {
    let initialDate = dateStr;
    if (!initialDate) {
      const y = currentYear;
      const m = String(currentMonth + 1).padStart(2, '0');
      const d = String(today.getDate()).padStart(2, '0');
      initialDate = `${y}-${m}-${d}`;
    }
    setFormPreselectedDate(initialDate);
    setEditingActivity(null);
    setBase64Images([]);
    setShowForm(true);
  };

  const openEditForm = (act: Activity) => {
    const d = getActivityDate(act);
    const dateStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    setFormPreselectedDate(dateStr);
    setEditingActivity(act);
    setBase64Images(parseImageUrls(act.imageUrls));
    setShowForm(true);
    setSelectedDetailActivity(null);
    setSelectedDayEvents(null);
  };

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

  // Handle submit form (Save / Update activity)
  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    
    const dateInput = formData.get('activityDate') as string;
    const timeInput = formData.get('time') as string;
    
    // Parse timestamp
    let activityTimestamp = Date.now();
    if (dateInput) {
      if (timeInput && timeInput.includes(':')) {
        const [hh, mm] = timeInput.split(':').map(Number);
        const [year, month, day] = dateInput.split('-').map(Number);
        activityTimestamp = new Date(year, month - 1, day, hh || 0, mm || 0).getTime();
      } else {
        const [year, month, day] = dateInput.split('-').map(Number);
        activityTimestamp = new Date(year, month - 1, day, 9, 0).getTime();
      }
    }

    const data: Partial<Activity> = {
      title: (formData.get('title') as string).trim(),
      description: (formData.get('description') as string || '').trim(),
      category: (formData.get('category') as string) || 'Pengajian Rutin',
      type: (formData.get('type') as string) || 'harian',
      time: (formData.get('time') as string || '').trim(),
      speaker: (formData.get('speaker') as string || '').trim(),
      date: activityTimestamp,
      imageUrls: base64Images.length > 0 ? base64Images : editingActivity?.imageUrls || [],
      location: profile.role === 'pengurus' && profile.location ? profile.location : (formData.get('location') as string || 'Kramat Batu'),
      createdBy: profile.uid || 'admin'
    };

    setIsSaving(true);
    try {
      await saveData(null, accessToken, spreadsheetId, 'activities', data, editingActivity?.id);
      showToast(editingActivity ? 'Kegiatan berhasil diperbarui!' : 'Kegiatan baru berhasil ditambahkan ke kalender!', 'success');
      setShowForm(false);
      setEditingActivity(null);
      setBase64Images([]);
      if (refresh) refresh();
    } catch (err: any) {
      showToast(`Gagal menyimpan kegiatan: ${err.message}`, 'error');
    } finally {
      setIsSaving(false);
    }
  };

  // Check if a date cell is today
  const isDateToday = (d: Date) => {
    const n = new Date();
    return d.getDate() === n.getDate() && d.getMonth() === n.getMonth() && d.getFullYear() === n.getFullYear();
  };

  return (
    <div className="space-y-6">
      {/* Header Banner & Month Navigator */}
      <div className="bg-white rounded-3xl p-6 lg:p-8 border border-slate-100 shadow-sm">
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-6">
          <div className="flex items-start gap-4">
            <div className="p-3 bg-emerald-100/80 rounded-2xl text-emerald-700 shadow-sm mt-0.5">
              <CalendarDays className="w-8 h-8" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-2xl lg:text-3xl font-black text-slate-900 tracking-tight">
                  Kalender Kegiatan Bulanan
                </h2>
                <span className="hidden sm:inline-flex px-3 py-0.5 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
                  Desa Gandaria
                </span>
              </div>
              <p className="text-slate-500 text-sm mt-1">
                Jadwal & rencana kegiatan sabilillah, pengajian, musyawarah, dan agenda pembinaan per bulan.
              </p>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-3">
            <button
              onClick={() => openAddFormForDate()}
              className="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2.5 rounded-2xl font-bold flex items-center gap-2 shadow-lg shadow-emerald-200 transition-all hover:scale-[1.02] active:scale-[0.98] text-sm"
            >
              <Plus className="w-5 h-5" />
              <span>Tambah Kegiatan</span>
            </button>
          </div>
        </div>

        {/* Month Selector Bar & Quick Stats */}
        <div className="mt-8 pt-6 border-t border-slate-100 flex flex-col md:flex-row items-center justify-between gap-4">
          {/* Month Navigator Controls */}
          <div className="flex items-center gap-3 w-full md:w-auto justify-between md:justify-start">
            <div className="flex items-center bg-slate-100/80 p-1 rounded-2xl border border-slate-200/60">
              <button
                onClick={handlePrevMonth}
                className="p-2 rounded-xl text-slate-600 hover:text-slate-900 hover:bg-white transition-all"
                title="Bulan Sebelumnya"
              >
                <ChevronLeft className="w-5 h-5" />
              </button>

              <div className="flex items-center gap-2 px-3">
                <select
                  value={currentMonth}
                  onChange={(e) => setCurrentMonth(Number(e.target.value))}
                  className="bg-transparent font-bold text-slate-900 text-sm lg:text-base outline-none cursor-pointer hover:text-emerald-700"
                >
                  {MONTH_NAMES.map((m, idx) => (
                    <option key={m} value={idx}>{m}</option>
                  ))}
                </select>

                <select
                  value={currentYear}
                  onChange={(e) => setCurrentYear(Number(e.target.value))}
                  className="bg-transparent font-bold text-slate-900 text-sm lg:text-base outline-none cursor-pointer hover:text-emerald-700"
                >
                  {[2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030].map((y) => (
                    <option key={y} value={y}>{y}</option>
                  ))}
                </select>
              </div>

              <button
                onClick={handleNextMonth}
                className="p-2 rounded-xl text-slate-600 hover:text-slate-900 hover:bg-white transition-all"
                title="Bulan Berikutnya"
              >
                <ChevronRight className="w-5 h-5" />
              </button>
            </div>

            <button
              onClick={handleToday}
              className="px-3.5 py-2 rounded-xl text-xs font-bold text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 transition-colors"
            >
              Hari Ini
            </button>
          </div>

          {/* Month Stats Badges */}
          <div className="flex items-center gap-2 lg:gap-3 w-full md:w-auto justify-end overflow-x-auto">
            <div className="px-3 py-1.5 bg-slate-50 border border-slate-200/80 rounded-xl flex items-center gap-2">
              <span className="text-xs text-slate-500 font-medium">Bulan Ini:</span>
              <span className="font-bold text-slate-900 text-sm">{monthStats.total} Kegiatan</span>
            </div>
            <div className="px-3 py-1.5 bg-emerald-50 border border-emerald-200/80 rounded-xl flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
              <span className="text-xs text-emerald-700 font-medium">Akan Datang:</span>
              <span className="font-bold text-emerald-800 text-sm">{monthStats.upcoming}</span>
            </div>
            <div className="px-3 py-1.5 bg-blue-50 border border-blue-200/80 rounded-xl flex items-center gap-2">
              <span className="text-xs text-blue-700 font-medium">Terlaksana:</span>
              <span className="font-bold text-blue-800 text-sm">{monthStats.completed}</span>
            </div>
          </div>
        </div>
      </div>

      {/* Filter & View Switcher Bar */}
      <div className="bg-white rounded-2xl p-4 border border-slate-100 shadow-sm flex flex-col md:flex-row items-center justify-between gap-4">
        {/* Search & Select Filters */}
        <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
          {/* Search box */}
          <div className="relative flex-grow sm:flex-grow-0 sm:w-64">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="Cari kegiatan / narasumber..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium outline-none focus:ring-2 focus:ring-emerald-500 transition-all"
            />
            {searchTerm && (
              <button onClick={() => setSearchTerm('')} className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600">
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          {/* Location filter */}
          {profile.role !== 'pengurus' && (
            <select
              value={selectedLocation}
              onChange={(e) => setSelectedLocation(e.target.value)}
              className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-700 outline-none focus:ring-2 focus:ring-emerald-500"
            >
              <option value="Semua Lokasi">Semua Lokasi / Kelompok</option>
              {MOSQUE_LOCATIONS.map(loc => (
                <option key={loc} value={loc}>{loc}</option>
              ))}
            </select>
          )}

          {/* Category filter */}
          <select
            value={selectedCategory}
            onChange={(e) => setSelectedCategory(e.target.value)}
            className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-700 outline-none focus:ring-2 focus:ring-emerald-500"
          >
            <option value="Semua Kategori">Semua Kategori Kegiatan</option>
            {ACTIVITY_CATEGORIES.map(cat => (
              <option key={cat.id} value={cat.id}>{cat.label}</option>
            ))}
          </select>
        </div>

        {/* View Mode Switcher */}
        <div className="flex items-center bg-slate-100 p-1 rounded-xl border border-slate-200/60 self-end md:self-auto">
          <button
            onClick={() => setViewMode('calendar')}
            className={cn(
              "flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all",
              viewMode === 'calendar' ? "bg-white text-emerald-700 shadow-sm" : "text-slate-500 hover:text-slate-800"
            )}
            title="Tampilan Kalender Bulanan"
          >
            <CalendarIcon className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Kalender</span>
          </button>
          <button
            onClick={() => setViewMode('agenda')}
            className={cn(
              "flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all",
              viewMode === 'agenda' ? "bg-white text-emerald-700 shadow-sm" : "text-slate-500 hover:text-slate-800"
            )}
            title="Tampilan Daftar Agenda"
          >
            <List className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Agenda List</span>
          </button>
          <button
            onClick={() => setViewMode('cards')}
            className={cn(
              "flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all",
              viewMode === 'cards' ? "bg-white text-emerald-700 shadow-sm" : "text-slate-500 hover:text-slate-800"
            )}
            title="Tampilan Kartu / Galeri"
          >
            <Grid className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Kartu</span>
          </button>
        </div>
      </div>

      {/* Main Content Area */}
      {loading ? (
        <div className="bg-white rounded-3xl p-16 border border-slate-100 shadow-sm text-center flex flex-col items-center justify-center">
          <Loader2 className="w-10 h-10 text-emerald-600 animate-spin mb-4" />
          <p className="text-slate-500 font-semibold text-sm">Memuat kalender & kegiatan Desa Gandaria...</p>
        </div>
      ) : (
        <div>
          {/* 1. CALENDAR VIEW */}
          {viewMode === 'calendar' && (
            <div className="bg-white rounded-3xl border border-slate-200/80 shadow-sm overflow-hidden">
              {/* Day Headers (Mon - Sun) */}
              <div className="grid grid-cols-7 bg-slate-50/80 border-b border-slate-200 text-center py-3">
                {DAY_NAMES.map((name, i) => (
                  <div key={name} className="text-xs font-black uppercase tracking-wider text-slate-500">
                    <span className="hidden md:inline">{DAY_FULL_NAMES[i]}</span>
                    <span className="md:hidden">{name}</span>
                  </div>
                ))}
              </div>

              {/* Day Cells Grid */}
              <div className="grid grid-cols-7 divide-x divide-y divide-slate-100 bg-slate-100/30">
                {calendarGrid.map((cell, idx) => {
                  const isToday = isDateToday(cell.date);
                  const hasEvents = cell.events.length > 0;
                  const dateStr = `${cell.date.getFullYear()}-${String(cell.date.getMonth() + 1).padStart(2, '0')}-${String(cell.dayNumber).padStart(2, '0')}`;

                  return (
                    <div
                      key={`grid-${idx}-${cell.dateKey}`}
                      className={cn(
                        "min-h-[110px] md:min-h-[135px] p-2 bg-white flex flex-col justify-between transition-colors relative group",
                        !cell.isCurrentMonth && "bg-slate-50/60 text-slate-400",
                        isToday && "bg-emerald-50/30 ring-1 ring-inset ring-emerald-500/40"
                      )}
                    >
                      {/* Top row of day cell */}
                      <div className="flex items-center justify-between mb-1">
                        <span
                          className={cn(
                            "w-6 h-6 flex items-center justify-center rounded-full text-xs font-bold",
                            isToday
                              ? "bg-emerald-600 text-white shadow-sm"
                              : cell.isCurrentMonth
                              ? "text-slate-800"
                              : "text-slate-400"
                          )}
                        >
                          {cell.dayNumber}
                        </span>

                        {/* Quick Add Button on cell hover */}
                        {cell.isCurrentMonth && (
                          <button
                            onClick={() => openAddFormForDate(dateStr)}
                            className="opacity-0 group-hover:opacity-100 p-1 rounded-md text-emerald-600 hover:bg-emerald-50 transition-all"
                            title={`Tambah kegiatan pada ${cell.dayNumber} ${MONTH_NAMES[currentMonth]}`}
                          >
                            <Plus className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>

                      {/* Event Chips List */}
                      <div className="space-y-1 flex-grow overflow-hidden">
                        {cell.events.slice(0, 3).map((event, eIdx) => {
                          const meta = getCategoryMeta(event.category || event.type);
                          return (
                            <div
                              key={event.id || `ev-${eIdx}`}
                              onClick={() => setSelectedDetailActivity(event)}
                              className={cn(
                                "px-1.5 py-0.5 rounded-md border text-[10px] font-semibold cursor-pointer truncate transition-transform hover:scale-[1.02] flex items-center gap-1 shadow-2xs",
                                meta.bg, meta.text, meta.border
                              )}
                              title={`${event.title} (${event.time || 'Waktu belum diatur'})`}
                            >
                              <span className={cn("w-1.5 h-1.5 rounded-full shrink-0", meta.dot)}></span>
                              {event.time && <span className="opacity-75 shrink-0">{event.time}</span>}
                              <span className="truncate font-medium">{event.title}</span>
                            </div>
                          );
                        })}

                        {/* If more than 3 events */}
                        {cell.events.length > 3 && (
                          <button
                            onClick={() => setSelectedDayEvents({ date: cell.date, events: cell.events })}
                            className="w-full text-left text-[10px] font-bold text-emerald-600 hover:text-emerald-800 hover:underline px-1"
                          >
                            +{cell.events.length - 3} lainnya...
                          </button>
                        )}
                      </div>

                      {/* Empty cell indicator on click */}
                      {cell.isCurrentMonth && cell.events.length === 0 && (
                        <div
                          onClick={() => openAddFormForDate(dateStr)}
                          className="w-full h-8 cursor-pointer rounded-lg hover:bg-slate-50 transition-colors flex items-center justify-center text-slate-300 opacity-0 group-hover:opacity-100"
                        >
                          <span className="text-[10px] font-medium">+ Tambah</span>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* 2. AGENDA LIST VIEW */}
          {viewMode === 'agenda' && (
            <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-6 space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-4">
                <h3 className="font-bold text-slate-800 text-lg">
                  Daftar Agenda Kegiatan: {MONTH_NAMES[currentMonth]} {currentYear}
                </h3>
                <span className="text-xs font-semibold text-slate-500">
                  {currentMonthActivities.length} Kegiatan Terjadwal
                </span>
              </div>

              {currentMonthActivities.length === 0 ? (
                <div className="p-12 text-center text-slate-400">
                  <CalendarDays className="w-12 h-12 mx-auto text-slate-300 mb-3" />
                  <p className="font-semibold">Tidak ada kegiatan di bulan {MONTH_NAMES[currentMonth]} {currentYear}</p>
                  <p className="text-xs text-slate-400 mt-1">Gunakan tombol 'Tambah Kegiatan' untuk merencanakan agenda.</p>
                </div>
              ) : (
                <div className="divide-y divide-slate-100">
                  {currentMonthActivities.map((act) => {
                    const actDate = getActivityDate(act);
                    const meta = getCategoryMeta(act.category || act.type);
                    const actImages = parseImageUrls(act.imageUrls);
                    const isUpcoming = actDate.getTime() >= new Date().getTime();

                    return (
                      <div key={act.id} className="py-4.5 flex flex-col md:flex-row items-start md:items-center justify-between gap-4 hover:bg-slate-50/80 px-3 rounded-2xl transition-colors">
                        {/* Date badge */}
                        <div className="flex items-center gap-4">
                          <div className="w-14 h-14 rounded-2xl bg-emerald-50 border border-emerald-200/80 flex flex-col items-center justify-center text-emerald-800 shrink-0">
                            <span className="text-[10px] font-bold uppercase tracking-wider">{DAY_NAMES[actDate.getDay() === 0 ? 6 : actDate.getDay() - 1]}</span>
                            <span className="text-xl font-black leading-none">{actDate.getDate()}</span>
                          </div>

                          {/* Event details */}
                          <div className="space-y-1">
                            <div className="flex flex-wrap items-center gap-2">
                              <span className={cn("px-2.5 py-0.5 rounded-full text-[10px] font-bold border", meta.bg, meta.text, meta.border)}>
                                {act.category || act.type || 'Kegiatan'}
                              </span>
                              <span className="text-xs font-bold text-slate-800">{act.location}</span>
                              {isUpcoming ? (
                                <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800">
                                  Akan Datang
                                </span>
                              ) : (
                                <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-600">
                                  Selesai
                                </span>
                              )}
                            </div>

                            <h4 
                              onClick={() => setSelectedDetailActivity(act)}
                              className="font-bold text-slate-900 text-base hover:text-emerald-700 cursor-pointer"
                            >
                              {act.title}
                            </h4>

                            <div className="flex flex-wrap items-center gap-4 text-xs text-slate-500">
                              {act.time && (
                                <span className="flex items-center gap-1">
                                  <Clock className="w-3.5 h-3.5 text-slate-400" />
                                  {act.time}
                                </span>
                              )}
                              {act.speaker && (
                                <span className="flex items-center gap-1">
                                  <User className="w-3.5 h-3.5 text-slate-400" />
                                  {act.speaker}
                                </span>
                              )}
                              {actImages.length > 0 && (
                                <span className="flex items-center gap-1 text-emerald-600 font-medium">
                                  <ImageIcon className="w-3.5 h-3.5" />
                                  {actImages.length} Foto
                                </span>
                              )}
                            </div>
                          </div>
                        </div>

                        {/* Action buttons */}
                        <div className="flex items-center gap-2 self-end md:self-auto">
                          <button
                            onClick={() => setSelectedDetailActivity(act)}
                            className="p-2 rounded-xl text-slate-500 hover:text-slate-900 hover:bg-slate-200/60 transition-colors"
                            title="Lihat Detail"
                          >
                            <Eye className="w-4 h-4" />
                          </button>
                          <button
                            onClick={() => openEditForm(act)}
                            className="p-2 rounded-xl text-emerald-600 hover:bg-emerald-50 transition-colors"
                            title="Edit Kegiatan"
                          >
                            <Edit className="w-4 h-4" />
                          </button>
                          <button
                            onClick={() => setDeleteId(act.id)}
                            className="p-2 rounded-xl text-red-500 hover:bg-red-50 transition-colors"
                            title="Hapus Kegiatan"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          )}

          {/* 3. CARD / GALLERY VIEW */}
          {viewMode === 'cards' && (
            <div>
              {currentMonthActivities.length === 0 ? (
                <div className="bg-white rounded-3xl border border-slate-100 p-12 text-center text-slate-400 font-medium">
                  Belum ada kegiatan di bulan {MONTH_NAMES[currentMonth]} {currentYear}
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  {currentMonthActivities.map((act) => {
                    const actDate = getActivityDate(act);
                    const meta = getCategoryMeta(act.category || act.type);
                    const actImages = parseImageUrls(act.imageUrls);

                    return (
                      <div 
                        key={act.id} 
                        className={cn(
                          "bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden flex flex-col hover:shadow-md transition-all",
                          meta.width,
                          meta.height,
                          meta.cardWidth,
                          meta.cardHeight
                        )}
                      >
                        {/* Image banner */}
                        <div className="h-48 relative bg-slate-100 group overflow-hidden">
                          {actImages.length > 0 ? (
                            <img src={actImages[0]} alt={act.title} className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300" />
                          ) : (
                            <div className="w-full h-full flex flex-col items-center justify-center text-slate-400">
                              <CalendarIcon className="w-10 h-10 mb-1 opacity-50" />
                              <span className="text-xs font-medium">Dokumentasi Kegiatan</span>
                            </div>
                          )}

                          <div className="absolute top-3 left-3 flex gap-2">
                            <span className={cn("px-3 py-1 rounded-full text-xs font-bold shadow-xs", meta.badge)}>
                              {act.category || act.type || 'Kegiatan'}
                            </span>
                          </div>

                          <div className="absolute top-3 right-3 bg-white/90 backdrop-blur-md px-2.5 py-1 rounded-xl text-xs font-bold text-slate-800 shadow-xs">
                            {actDate.getDate()} {MONTH_NAMES[actDate.getMonth()]}
                          </div>
                        </div>

                        {/* Content */}
                        <div className="p-6 flex flex-col flex-grow justify-between space-y-4">
                          <div className="space-y-2">
                            <h4 className="font-bold text-slate-900 text-lg leading-snug hover:text-emerald-700 cursor-pointer" onClick={() => setSelectedDetailActivity(act)}>
                              {act.title}
                            </h4>
                            <p className="text-slate-500 text-xs line-clamp-2">{act.description || 'Tidak ada deskripsi rincian.'}</p>
                          </div>

                          <div className="space-y-2 pt-3 border-t border-slate-100 text-xs text-slate-500">
                            <div className="flex items-center justify-between">
                              <span className="flex items-center gap-1.5">
                                <MapPin className="w-3.5 h-3.5 text-emerald-600" />
                                <span className="font-medium text-slate-700">{act.location}</span>
                              </span>
                              {act.time && (
                                <span className="flex items-center gap-1 text-slate-500">
                                  <Clock className="w-3.5 h-3.5" />
                                  {act.time}
                                </span>
                              )}
                            </div>

                            {act.speaker && (
                              <div className="flex items-center gap-1.5 text-slate-600">
                                <User className="w-3.5 h-3.5 text-slate-400" />
                                <span className="truncate">Pemateri: {act.speaker}</span>
                              </div>
                            )}
                          </div>

                          {/* Card action footer */}
                          <div className="flex items-center justify-between pt-3 border-t border-slate-100">
                            <button
                              onClick={() => setSelectedDetailActivity(act)}
                              className="text-xs font-bold text-emerald-600 hover:text-emerald-800 flex items-center gap-1"
                            >
                              Detail Acara <ArrowRight className="w-3 h-3" />
                            </button>
                            <div className="flex items-center gap-1">
                              <button
                                onClick={() => openEditForm(act)}
                                className="p-1.5 rounded-lg text-slate-600 hover:bg-slate-100"
                                title="Edit"
                              >
                                <Edit className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => setDeleteId(act.id)}
                                className="p-1.5 rounded-lg text-red-500 hover:bg-red-50"
                                title="Hapus"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </div>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {/* MODAL: Tambah / Edit Kegiatan */}
      <AnimatePresence>
        {showForm && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50 overflow-y-auto"
            onClick={() => setShowForm(false)}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-3xl p-6 lg:p-8 w-full max-w-2xl shadow-2xl my-8 max-h-[90vh] overflow-y-auto"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="flex items-center justify-between pb-4 border-b border-slate-100 mb-6">
                <div>
                  <h3 className="text-xl font-black text-slate-900">
                    {editingActivity ? 'Edit Kegiatan / Acara' : 'Tambah Kegiatan Baru'}
                  </h3>
                  <p className="text-xs text-slate-500">
                    Isi rincian jadwal kegiatan untuk bulan {MONTH_NAMES[currentMonth]} {currentYear}
                  </p>
                </div>
                <button
                  onClick={() => setShowForm(false)}
                  className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-full transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <form onSubmit={handleSubmit} className="space-y-4">
                {/* Judul Kegiatan */}
                <div>
                  <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">
                    Judul Kegiatan <span className="text-red-500">*</span>
                  </label>
                  <input
                    name="title"
                    defaultValue={editingActivity?.title}
                    required
                    placeholder="Contoh: Pengajian Rutin Malam Kamis, Musyawarah Desa, dll."
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm font-medium"
                  />
                </div>

                {/* Kategori & Lokasi */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">
                      Kategori Kegiatan
                    </label>
                    <select
                      name="category"
                      defaultValue={editingActivity?.category || 'Pengajian Rutin'}
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm font-medium"
                    >
                      {ACTIVITY_CATEGORIES.map(cat => (
                        <option key={cat.id} value={cat.id}>{cat.label}</option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">
                      Lokasi / Kelompok
                    </label>
                    <select
                      name="location"
                      defaultValue={editingActivity?.location || (profile.role === 'pengurus' ? profile.location : 'Kramat Batu')}
                      disabled={profile.role === 'pengurus'}
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm font-medium"
                    >
                      {MOSQUE_LOCATIONS.map(l => (
                        <option key={l} value={l}>{l}</option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* Tanggal & Waktu */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">
                      Tanggal Kegiatan <span className="text-red-500">*</span>
                    </label>
                    <input
                      type="date"
                      name="activityDate"
                      required
                      defaultValue={formPreselectedDate}
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm font-medium"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">
                      Waktu / Jam
                    </label>
                    <input
                      type="text"
                      name="time"
                      defaultValue={editingActivity?.time}
                      placeholder="Contoh: 19:30 - 21:00 WIB / Ba'da Maghrib"
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm font-medium"
                    />
                  </div>
                </div>

                {/* Pemateri & Tipe Frekuensi */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">
                      Pemateri / Penanggung Jawab
                    </label>
                    <input
                      name="speaker"
                      defaultValue={editingActivity?.speaker}
                      placeholder="Contoh: Ust. H. Ahmad / Tim Mubaligh Desa"
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm font-medium"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">
                      Tipe Pelaksanaan
                    </label>
                    <select
                      name="type"
                      defaultValue={editingActivity?.type || 'harian'}
                      className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm font-medium"
                    >
                      <option value="harian">Sekali Jalan (Harian)</option>
                      <option value="mingguan">Rutin Mingguan</option>
                      <option value="bulanan">Rutin Bulanan</option>
                      <option value="khusus">Khusus / Sambung Desa</option>
                    </select>
                  </div>
                </div>

                {/* Deskripsi Kegiatan */}
                <div>
                  <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">
                    Deskripsi / Rincian Acara
                  </label>
                  <textarea
                    name="description"
                    defaultValue={editingActivity?.description}
                    rows={3}
                    placeholder="Rincian agenda, materi kajian, atau instruksi sabilillah untuk jamaah..."
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 outline-none focus:ring-2 focus:ring-emerald-500 text-sm font-medium"
                  />
                </div>

                {/* Upload Foto Dokumentasi */}
                <div>
                  <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">
                    Foto Dokumentasi Kegiatan (Bisa banyak)
                  </label>
                  
                  {base64Images.length > 0 && (
                    <div className="flex flex-wrap gap-2.5 mb-3 p-3 bg-slate-50 rounded-2xl border border-slate-100">
                      {base64Images.map((img, i) => (
                        <div key={i} className="relative w-16 h-16 rounded-xl overflow-hidden group shadow-xs">
                          <img src={img} className="w-full h-full object-cover" alt={`foto-${i}`} />
                          <button
                            type="button"
                            onClick={() => setBase64Images(prev => prev.filter((_, idx) => idx !== i))}
                            className="absolute inset-0 bg-black/50 text-white flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity"
                          >
                            <Trash2 className="w-4 h-4 text-red-400" />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}

                  <label className="flex items-center justify-center gap-2 cursor-pointer bg-slate-50 hover:bg-slate-100 border-2 border-dashed border-slate-200 rounded-2xl p-4 transition-all text-slate-500 text-xs font-semibold">
                    <ImageIcon className="w-4 h-4 text-emerald-600" />
                    <span>Klik untuk memilih foto dokumentasi</span>
                    <input type="file" accept="image/*" multiple onChange={handleImagesChange} className="hidden" />
                  </label>
                </div>

                {/* Buttons */}
                <div className="flex items-center gap-3 pt-4 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setShowForm(false)}
                    disabled={isSaving}
                    className="flex-1 py-3 px-5 rounded-xl border border-slate-200 text-slate-700 font-bold hover:bg-slate-50 text-sm transition-all"
                  >
                    Batal
                  </button>
                  <button
                    type="submit"
                    disabled={isSaving}
                    className="flex-1 py-3 px-5 rounded-xl bg-emerald-600 text-white font-bold hover:bg-emerald-700 shadow-lg shadow-emerald-200 text-sm transition-all flex items-center justify-center gap-2 disabled:opacity-60"
                  >
                    {isSaving ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        <span>Menyimpan Kegiatan...</span>
                      </>
                    ) : (
                      <span>{editingActivity ? 'Simpan Perubahan' : 'Jadwalkan Kegiatan'}</span>
                    )}
                  </button>
                </div>
              </form>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* MODAL: Detail Acara & Dokumentasi */}
      <AnimatePresence>
        {selectedDetailActivity && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50"
            onClick={() => setSelectedDetailActivity(null)}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-3xl p-6 lg:p-8 w-full max-w-lg shadow-2xl max-h-[90vh] overflow-y-auto space-y-6"
              onClick={(e) => e.stopPropagation()}
            >
              {/* Header */}
              <div className="flex items-start justify-between">
                <div>
                  <span className={cn(
                    "px-3 py-1 rounded-full text-xs font-bold inline-block mb-2",
                    getCategoryMeta(selectedDetailActivity.category || selectedDetailActivity.type).badge
                  )}>
                    {selectedDetailActivity.category || selectedDetailActivity.type || 'Kegiatan'}
                  </span>
                  <h3 className="text-xl font-black text-slate-900 leading-snug">
                    {selectedDetailActivity.title}
                  </h3>
                </div>
                <button
                  onClick={() => setSelectedDetailActivity(null)}
                  className="p-1.5 rounded-full text-slate-400 hover:text-slate-600 hover:bg-slate-100"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Meta Grid */}
              <div className="grid grid-cols-2 gap-3 p-4 bg-slate-50 rounded-2xl text-xs">
                <div>
                  <span className="text-slate-400 block font-medium">Tanggal</span>
                  <span className="font-bold text-slate-800">
                    {new Date(selectedDetailActivity.date).toLocaleDateString('id-ID', {
                      weekday: 'long',
                      day: 'numeric',
                      month: 'long',
                      year: 'numeric'
                    })}
                  </span>
                </div>
                <div>
                  <span className="text-slate-400 block font-medium">Waktu</span>
                  <span className="font-bold text-slate-800">{selectedDetailActivity.time || 'Waktu Fleksibel'}</span>
                </div>
                <div>
                  <span className="text-slate-400 block font-medium">Lokasi</span>
                  <span className="font-bold text-slate-800">{selectedDetailActivity.location}</span>
                </div>
                <div>
                  <span className="text-slate-400 block font-medium">Pemateri / Narasumber</span>
                  <span className="font-bold text-slate-800">{selectedDetailActivity.speaker || '-'}</span>
                </div>
              </div>

              {/* Description */}
              <div>
                <h5 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Rincian Agenda</h5>
                <p className="text-sm text-slate-700 whitespace-pre-line leading-relaxed">
                  {selectedDetailActivity.description || 'Tidak ada deskripsi tambahan.'}
                </p>
              </div>

              {/* Photos */}
              {parseImageUrls(selectedDetailActivity.imageUrls).length > 0 && (
                <div>
                  <h5 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Foto Dokumentasi</h5>
                  <div className="grid grid-cols-2 gap-2">
                    {parseImageUrls(selectedDetailActivity.imageUrls).map((url, imgIdx) => (
                      <a key={imgIdx} href={url} target="_blank" rel="noreferrer" className="block rounded-xl overflow-hidden group relative h-32 border border-slate-200">
                        <img src={url} alt={`Dokumentasi ${imgIdx + 1}`} className="w-full h-full object-cover group-hover:scale-105 transition-transform" />
                        <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center text-white text-xs font-medium">
                          Lihat Penuh
                        </div>
                      </a>
                    ))}
                  </div>
                </div>
              )}

              {/* Action buttons */}
              <div className="flex items-center gap-3 pt-4 border-t border-slate-100">
                <button
                  onClick={() => openEditForm(selectedDetailActivity)}
                  className="flex-1 py-2.5 rounded-xl bg-emerald-50 text-emerald-700 font-bold hover:bg-emerald-100 transition-colors flex items-center justify-center gap-2 text-sm"
                >
                  <Edit className="w-4 h-4" />
                  <span>Edit Data</span>
                </button>
                <button
                  onClick={() => {
                    setDeleteId(selectedDetailActivity.id);
                    setSelectedDetailActivity(null);
                  }}
                  className="p-2.5 rounded-xl text-red-600 hover:bg-red-50 transition-colors border border-red-200"
                  title="Hapus Kegiatan"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* MODAL: Day Events Popover (when >3 events in day cell) */}
      <AnimatePresence>
        {selectedDayEvents && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50"
            onClick={() => setSelectedDayEvents(null)}
          >
            <motion.div
              initial={{ scale: 0.95 }}
              animate={{ scale: 1 }}
              exit={{ scale: 0.95 }}
              className="bg-white rounded-3xl p-6 w-full max-w-md shadow-2xl space-y-4 max-h-[80vh] overflow-y-auto"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div>
                  <h4 className="font-bold text-slate-900 text-base">
                    Agenda: {selectedDayEvents.date.toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}
                  </h4>
                  <span className="text-xs text-slate-500">{selectedDayEvents.events.length} Kegiatan Terjadwal</span>
                </div>
                <button onClick={() => setSelectedDayEvents(null)} className="p-1.5 text-slate-400 hover:text-slate-600">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="space-y-2">
                {selectedDayEvents.events.map(ev => {
                  const meta = getCategoryMeta(ev.category || ev.type);
                  return (
                    <div
                      key={ev.id}
                      onClick={() => {
                        setSelectedDetailActivity(ev);
                        setSelectedDayEvents(null);
                      }}
                      className={cn(
                        "p-3 rounded-2xl border cursor-pointer hover:shadow-xs transition-all flex items-center justify-between",
                        meta.bg, meta.border
                      )}
                    >
                      <div className="space-y-0.5">
                        <span className={cn("text-[10px] font-bold uppercase", meta.text)}>
                          {ev.category || ev.type || 'Kegiatan'}
                        </span>
                        <h5 className="font-bold text-slate-900 text-sm">{ev.title}</h5>
                        <div className="flex items-center gap-3 text-[11px] text-slate-500">
                          {ev.time && <span>{ev.time}</span>}
                          <span>• {ev.location}</span>
                        </div>
                      </div>
                      <Eye className="w-4 h-4 text-slate-400" />
                    </div>
                  );
                })}
              </div>

              <button
                onClick={() => {
                  const d = selectedDayEvents.date;
                  const dateStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
                  setSelectedDayEvents(null);
                  openAddFormForDate(dateStr);
                }}
                className="w-full py-2.5 rounded-xl bg-emerald-600 text-white font-bold text-xs hover:bg-emerald-700 transition-colors flex items-center justify-center gap-2"
              >
                <Plus className="w-4 h-4" />
                <span>Tambah Kegiatan di Tanggal Ini</span>
              </button>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Delete Confirmation */}
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
            if (refresh) refresh();
          } catch (err: any) {
            showToast(`Gagal menghapus kegiatan: ${err.message}`, 'error');
          } finally {
            setIsDeleting(false);
          }
        }}
        loading={isDeleting}
        title="Hapus Kegiatan"
        message="Apakah Anda yakin ingin menghapus data kegiatan ini dari kalender? Tindakan ini tidak dapat dibatalkan."
      />
    </div>
  );
}

export default CalendarActivitiesView;