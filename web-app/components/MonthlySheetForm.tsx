'use client';

import { CaseData } from '@/lib/api';

interface MonthlySheetFormProps {
  formData: CaseData;
  clients: string[];
  statuses: string[];
  months: string[];
  loading: boolean;
  onChange: (field: keyof CaseData, value: string | number | undefined) => void;
}

export default function MonthlySheetForm({
  formData,
  clients,
  statuses,
  months,
  loading,
  onChange,
}: MonthlySheetFormProps) {
  return (
    <div className="space-y-6">
      {/* 月選択 */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          月 <span className="text-red-500">*</span>
        </label>
        <select
          value={formData.month}
          onChange={(e) => onChange('month', e.target.value)}
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          required
          disabled={loading}
        >
          <option value="">選択してください</option>
          {months.map((month) => (
            <option key={month} value={month}>
              {month}
            </option>
          ))}
        </select>
      </div>

      {/* 登録日 */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          登録日 <span className="text-red-500">*</span>
        </label>
        <input
          type="date"
          value={formData.registrationDate}
          onChange={(e) => onChange('registrationDate', e.target.value)}
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          required
        />
        <p className="mt-1 text-sm text-gray-500">未入力の場合は当日が設定されます</p>
      </div>

      {/* クライアント名 */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          クライアント名 <span className="text-red-500">*</span>
        </label>
        <select
          value={formData.clientName}
          onChange={(e) => onChange('clientName', e.target.value)}
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          required
          disabled={loading}
        >
          <option value="">選択してください</option>
          {clients.map((client) => (
            <option key={client} value={client}>
              {client}
            </option>
          ))}
        </select>
      </div>

      {/* 業種 */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          業種
        </label>
        <input
          type="text"
          value={formData.industry || ''}
          onChange={(e) => onChange('industry', e.target.value)}
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          placeholder="業種を入力してください"
        />
      </div>

      {/* 経費（円） */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          経費（円）
        </label>
        <input
          type="number"
          value={formData.expense || ''}
          onChange={(e) => onChange('expense', e.target.value ? parseFloat(e.target.value) : undefined)}
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          placeholder="0"
          min="0"
        />
      </div>

      {/* 経費（ドル） */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          経費（ドル）
        </label>
        <input
          type="number"
          value={formData.expenseUsd || ''}
          onChange={(e) => onChange('expenseUsd', e.target.value ? parseFloat(e.target.value) : undefined)}
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          placeholder="0"
          min="0"
          step="0.01"
        />
      </div>

      {/* 売上 */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          売上 <span className="text-red-500">*</span>
        </label>
        <input
          type="number"
          value={formData.revenue || ''}
          onChange={(e) => onChange('revenue', parseFloat(e.target.value) || 0)}
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          placeholder="0"
          required
          min="0"
        />
      </div>

      {/* ステータス */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          ステータス <span className="text-red-500">*</span>
        </label>
        <select
          value={formData.status}
          onChange={(e) => onChange('status', e.target.value)}
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          required
          disabled={loading}
        >
          <option value="">選択してください</option>
          {statuses.map((status) => (
            <option key={status} value={status}>
              {status}
            </option>
          ))}
        </select>
      </div>

      {/* 備考 */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          備考 <span className="text-red-500">*</span>
        </label>
        <textarea
          value={formData.notes}
          onChange={(e) => onChange('notes', e.target.value)}
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          rows={4}
          placeholder="備考を入力してください"
          required
        />
      </div>
    </div>
  );
}

