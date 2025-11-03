'use client';

import { useState, useEffect } from 'react';
import { CaseData, fetchClients, fetchStatuses, fetchMonths, registerCase } from '@/lib/api';

interface CaseFormProps {
  onSuccess: () => void;
}

export default function CaseForm({ onSuccess }: CaseFormProps) {
  const [loading, setLoading] = useState(false);
  const [clients, setClients] = useState<string[]>([]);
  const [statuses, setStatuses] = useState<string[]>([]);
  const [months, setMonths] = useState<string[]>([]);
  const [formData, setFormData] = useState<CaseData>({
    sheetType: '月別シート',
    month: '',
    registrationDate: new Date().toISOString().split('T')[0],
    clientName: '',
    industry: '',
    expense: undefined,
    expenseUsd: undefined,
    revenue: 0,
    orderDate: undefined,
    status: '',
    notes: '',
  });
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    loadMasterData();
  }, []);

  const loadMasterData = async () => {
    try {
      setLoading(true);
      setError(null);
      const [clientsData, statusesData, monthsData] = await Promise.all([
        fetchClients(),
        fetchStatuses(),
        fetchMonths(),
      ]);
      setClients(clientsData);
      setStatuses(statusesData);
      setMonths(monthsData);
      
      // デフォルトの月を設定（最初の月）
      if (monthsData.length > 0 && !formData.month) {
        setFormData(prev => ({ ...prev, month: monthsData[0] }));
      }
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'マスターデータの読み込みに失敗しました';
      setError(errorMessage);
      console.error('Failed to load master data:', err);
    } finally {
      setLoading(false);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setLoading(true);

    try {
      await registerCase(formData);
      onSuccess();
    } catch (err) {
      setError(err instanceof Error ? err.message : '案件登録に失敗しました');
    } finally {
      setLoading(false);
    }
  };

  const handleChange = (field: keyof CaseData, value: string | number | undefined) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

  return (
    <div className="max-w-2xl mx-auto p-4">
      <h1 className="text-2xl font-bold mb-6 text-gray-800">案件登録</h1>
      
      {error && (
        <div className="mb-4 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700">
          {error}
        </div>
      )}

      <form onSubmit={handleSubmit} className="space-y-6">
        {/* シートタイプ選択 */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            登録先シート <span className="text-red-500">*</span>
          </label>
          <select
            value={formData.sheetType}
            onChange={(e) => handleChange('sheetType', e.target.value as '月別シート' | '月末請求シート')}
            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
            required
          >
            <option value="月別シート">月別シート</option>
            <option value="月末請求シート">月末請求シート</option>
          </select>
        </div>

        {/* 月選択 */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            月 <span className="text-red-500">*</span>
          </label>
          <select
            value={formData.month}
            onChange={(e) => handleChange('month', e.target.value)}
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
            onChange={(e) => handleChange('registrationDate', e.target.value)}
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
            onChange={(e) => handleChange('clientName', e.target.value)}
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
            onChange={(e) => handleChange('industry', e.target.value)}
            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
            placeholder="業種を入力してください"
          />
        </div>

        {/* 経費（月別シートの場合のみ表示） */}
        {formData.sheetType === '月別シート' && (
          <>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                経費（円）
              </label>
              <input
                type="number"
                value={formData.expense || ''}
                onChange={(e) => handleChange('expense', e.target.value ? parseFloat(e.target.value) : undefined)}
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                placeholder="0"
                min="0"
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                経費（ドル）
              </label>
              <input
                type="number"
                value={formData.expenseUsd || ''}
                onChange={(e) => handleChange('expenseUsd', e.target.value ? parseFloat(e.target.value) : undefined)}
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                placeholder="0"
                min="0"
                step="0.01"
              />
            </div>
          </>
        )}

        {/* 受注日（月末請求シートの場合のみ表示） */}
        {formData.sheetType === '月末請求シート' && (
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              受注日
            </label>
            <input
              type="date"
              value={formData.orderDate || formData.registrationDate}
              onChange={(e) => handleChange('orderDate', e.target.value)}
              className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
        )}

        {/* 経費（月末請求シートの場合） */}
        {formData.sheetType === '月末請求シート' && (
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              経費（円）
            </label>
            <input
              type="number"
              value={formData.expense || ''}
              onChange={(e) => handleChange('expense', e.target.value ? parseFloat(e.target.value) : undefined)}
              className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              placeholder="0"
              min="0"
            />
          </div>
        )}

        {/* 売上 */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            売上 <span className="text-red-500">*</span>
          </label>
          <input
            type="number"
            value={formData.revenue || ''}
            onChange={(e) => handleChange('revenue', parseFloat(e.target.value) || 0)}
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
            onChange={(e) => handleChange('status', e.target.value)}
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
            onChange={(e) => handleChange('notes', e.target.value)}
            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
            rows={4}
            placeholder="備考を入力してください"
            required
          />
        </div>

        {/* 送信ボタン */}
        <div className="flex justify-end space-x-4">
          <button
            type="submit"
            disabled={loading}
            className="px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 disabled:bg-gray-400 disabled:cursor-not-allowed font-medium"
          >
            {loading ? '登録中...' : '登録'}
          </button>
        </div>
      </form>
    </div>
  );
}

