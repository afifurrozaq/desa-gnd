import React from 'react';
import { ChevronLeft, ChevronRight } from 'lucide-react';
import { cn } from '../lib/utils';

interface PaginationProps {
  currentPage: number;
  totalPages: number;
  totalItems: number;
  pageSize: number;
  onPageChange: (page: number) => void;
  onPageSizeChange?: (size: number) => void;
  pageSizeOptions?: number[];
  className?: string;
}

export function Pagination({
  currentPage,
  totalPages,
  totalItems,
  pageSize,
  onPageChange,
  onPageSizeChange,
  pageSizeOptions = [10, 20, 50, 100],
  className
}: PaginationProps) {
  if (totalItems === 0) return null;

  const from = Math.min((currentPage - 1) * pageSize + 1, totalItems);
  const to = Math.min(currentPage * pageSize, totalItems);

  // Generate page numbers with ellipses
  const getPageNumbers = () => {
    const pages: (number | string)[] = [];
    if (totalPages <= 7) {
      for (let i = 1; i <= totalPages; i++) {
        pages.push(i);
      }
    } else {
      pages.push(1);
      if (currentPage > 3) {
        pages.push('...');
      }
      
      const start = Math.max(2, currentPage - 1);
      const end = Math.min(totalPages - 1, currentPage + 1);
      for (let i = start; i <= end; i++) {
        pages.push(i);
      }
      
      if (currentPage < totalPages - 2) {
        pages.push('...');
      }
      pages.push(totalPages);
    }
    return pages;
  };

  return (
    <div className={cn("flex flex-col sm:flex-row items-center justify-between gap-4 py-4 px-2 border-t border-slate-100 text-sm text-slate-600", className)}>
      <div className="flex flex-wrap items-center gap-3">
        <span>
          Menampilkan <strong className="text-slate-900 font-semibold">{from}</strong> - <strong className="text-slate-900 font-semibold">{to}</strong> dari <strong className="text-slate-900 font-semibold">{totalItems}</strong> data
        </span>
        {onPageSizeChange && (
          <div className="flex items-center gap-1.5 text-xs text-slate-500">
            <span>Per halaman:</span>
            <select
              value={pageSize}
              onChange={(e) => {
                onPageSizeChange(Number(e.target.value));
                onPageChange(1);
              }}
              className="px-2 py-1 bg-white border border-slate-200 rounded-lg text-xs font-semibold text-slate-700 outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer shadow-2xs"
            >
            {pageSizeOptions.map((opt, optIdx) => (
              <option key={`page-size-${opt}-${optIdx}`} value={opt}>
                {opt}
              </option>
            ))}
            </select>
          </div>
        )}
      </div>

      <div className="flex items-center gap-1">
        <button
          onClick={() => onPageChange(currentPage - 1)}
          disabled={currentPage <= 1}
          className="p-2 rounded-xl border border-slate-200 text-slate-600 hover:bg-slate-50 hover:text-emerald-700 transition-colors disabled:opacity-30 disabled:cursor-not-allowed shadow-2xs"
          title="Halaman Sebelumnya"
          aria-label="Previous Page"
        >
          <ChevronLeft className="w-4 h-4" />
        </button>

        <div className="flex items-center gap-1">
          {getPageNumbers().map((page, idx) => {
            if (page === '...') {
              return (
                <span key={`ellipsis-${idx}`} className="px-2 py-1 text-slate-400 text-xs">
                  •••
                </span>
              );
            }
            const isCurrent = page === currentPage;
            return (
              <button
                key={`page-btn-${page}-${idx}`}
                onClick={() => onPageChange(Number(page))}
                className={cn(
                  "min-w-8 h-8 px-2.5 rounded-xl text-xs font-bold transition-all shadow-2xs",
                  isCurrent
                    ? "bg-emerald-600 text-white shadow-emerald-200"
                    : "border border-slate-200 text-slate-600 hover:bg-slate-50 hover:text-emerald-700"
                )}
              >
                {page}
              </button>
            );
          })}
        </div>

        <button
          onClick={() => onPageChange(currentPage + 1)}
          disabled={currentPage >= totalPages}
          className="p-2 rounded-xl border border-slate-200 text-slate-600 hover:bg-slate-50 hover:text-emerald-700 transition-colors disabled:opacity-30 disabled:cursor-not-allowed shadow-2xs"
          title="Halaman Selanjutnya"
          aria-label="Next Page"
        >
          <ChevronRight className="w-4 h-4" />
        </button>
      </div>
    </div>
  );
}
