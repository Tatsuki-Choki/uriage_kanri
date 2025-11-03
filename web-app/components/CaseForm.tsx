'use client';

import { useState, useEffect } from 'react';
import { CaseData, fetchClients, fetchStatuses, fetchMonths, registerCase } from '@/lib/api';
import MonthlySheetForm from './MonthlySheetForm';
import EndOfMonthSheetForm from './EndOfMonthSheetForm';

interface CaseFormProps {
  onSuccess: () => void;
}

type SheetType = '月別シート' | '月末請求シート';

export default function CaseForm({ onSuccess }: CaseFormProps) {
  const [activeTab, setActiveTab] = useState<SheetType>('月別シート');
  const [loading, setLoading] = useState(false);
  const [clients, setClients] = useState<string[]>([]);
  const [statuses, setStatuses] = useState<string[]>([]);
  const [months, setMonths] = useState<string[]>([]);
  const [monthlyFormData, setMonthlyFormData] = useState<CaseData>({
    sheetType: '月別シート',
    month: '',
    registrationDate: new Date().toISOString().split('T')[0],
    clientName: '',
    industry: '',
    expense: undefined,
    expenseUsd: undefined,
    revenue: 0,
    status: '',
    notes: '',
  });
  const [endOfMonthFormData, setEndOfMonthFormData] = useState<CaseData>({
    sheetType: '月末請求シート',
    month: '',
    registrationDate: new Date().toISOString().split('T')[0],
    clientName: '',
    expense: undefined,
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
      if (monthsData.length > 0) {
        if (!monthlyFormData.month) {
          setMonthlyFormData(prev => ({ ...prev, month: monthsData[0] }));
        }
        if (!endOfMonthFormData.month) {
          setEndOfMonthFormData(prev => ({ ...prev, month: monthsData[0] }));
        }
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
      const formData = activeTab === '月別シート' ? monthlyFormData : endOfMonthFormData;
      await registerCase(formData);
      onSuccess();
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : '案件登録に失敗しました';
      setError(errorMessage);
      console.error('案件登録エラー:', err);
    } finally {
      setLoading(false);
    }
  };

  const handleMonthlyChange = (field: keyof CaseData, value: string | number | undefined) => {
    setMonthlyFormData(prev => ({ ...prev, [field]: value }));
  };

  const handleEndOfMonthChange = (field: keyof CaseData, value: string | number | undefined) => {
    setEndOfMonthFormData(prev => ({ ...prev, [field]: value }));
  };

  const handleTabChange = (tab: SheetType) => {
    setActiveTab(tab);
    setError(null);
  };

  return (
    <div className="max-w-2xl mx-auto p-4">
      <h1 className="text-2xl font-bold mb-6 text-gray-800">案件登録</h1>
      
      {error && (
        <div className="mb-4 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700">
          {error}
        </div>
      )}

      {/* タブ */}
      <div className="mb-6 border-b border-gray-200">
        <nav className="-mb-px flex space-x-8">
          <button
            type="button"
            onClick={() => handleTabChange('月別シート')}
            className={`
              py-4 px-1 border-b-2 font-medium text-sm
              ${activeTab === '月別シート'
                ? 'border-blue-500 text-blue-600'
                : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
              }
            `}
          >
            月別シート
          </button>
          <button
            type="button"
            onClick={() => handleTabChange('月末請求シート')}
            className={`
              py-4 px-1 border-b-2 font-medium text-sm
              ${activeTab === '月末請求シート'
                ? 'border-blue-500 text-blue-600'
                : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
              }
            `}
          >
            月末請求シート
          </button>
        </nav>
      </div>

      {/* フォーム */}
      <form onSubmit={handleSubmit} className="space-y-6">
        {activeTab === '月別シート' ? (
          <MonthlySheetForm
            formData={monthlyFormData}
            clients={clients}
            statuses={statuses}
            months={months}
            loading={loading}
            onChange={handleMonthlyChange}
          />
        ) : (
          <EndOfMonthSheetForm
            formData={endOfMonthFormData}
            clients={clients}
            statuses={statuses}
            months={months}
            loading={loading}
            onChange={handleEndOfMonthChange}
          />
        )}

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
