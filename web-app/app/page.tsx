'use client';

import { useState, useEffect } from 'react';
import Login from '@/components/Login';
import CaseForm from '@/components/CaseForm';
import Success from '@/components/Success';
import { User } from '@/lib/auth';

type AppState = 'login' | 'form' | 'success';

const USER_STORAGE_KEY = 'case_registration_user';
const STATE_STORAGE_KEY = 'case_registration_state';

export default function Home() {
  const [state, setState] = useState<AppState>('login');
  const [user, setUser] = useState<User | null>(null);
  const [isLoading, setIsLoading] = useState(true);

  // ページ読み込み時にログイン状態を復元
  useEffect(() => {
    const savedUser = localStorage.getItem(USER_STORAGE_KEY);
    const savedState = localStorage.getItem(STATE_STORAGE_KEY) as AppState | null;

    if (savedUser) {
      try {
        const parsedUser = JSON.parse(savedUser);
        setUser(parsedUser);
        // ログイン済みの場合はフォーム画面へ
        if (savedState === 'form' || savedState === 'success') {
          setState(savedState);
        } else {
          setState('form');
        }
      } catch (error) {
        console.error('Failed to restore user state:', error);
        localStorage.removeItem(USER_STORAGE_KEY);
        localStorage.removeItem(STATE_STORAGE_KEY);
      }
    }
    setIsLoading(false);
  }, []);

  const handleLoginSuccess = (loggedInUser: User) => {
    setUser(loggedInUser);
    setState('form');
    // localStorageに保存
    localStorage.setItem(USER_STORAGE_KEY, JSON.stringify(loggedInUser));
    localStorage.setItem(STATE_STORAGE_KEY, 'form');
  };

  const handleRegisterSuccess = () => {
    setState('success');
    localStorage.setItem(STATE_STORAGE_KEY, 'success');
  };

  const handleBackToForm = () => {
    setState('form');
    localStorage.setItem(STATE_STORAGE_KEY, 'form');
  };

  const handleLogout = () => {
    setUser(null);
    setState('login');
    // localStorageから削除
    localStorage.removeItem(USER_STORAGE_KEY);
    localStorage.removeItem(STATE_STORAGE_KEY);
  };

  // ローディング中は何も表示しない
  if (isLoading) {
    return (
      <div className="flex items-center justify-center min-h-screen">
        <div className="text-gray-600">読み込み中...</div>
      </div>
    );
  }

  return (
    <>
      {state === 'login' && <Login onLoginSuccess={handleLoginSuccess} />}
      {state === 'form' && (
        <div className="min-h-screen bg-gray-50">
          <div className="bg-white shadow-sm border-b">
            <div className="max-w-7xl mx-auto px-4 py-4 flex justify-between items-center">
              <h1 className="text-xl font-semibold text-gray-800">案件登録システム</h1>
              <div className="flex items-center space-x-4">
                {user && (
                  <span className="text-sm text-gray-600">{user.name || user.email}</span>
                )}
                <button
                  onClick={handleLogout}
                  className="px-4 py-2 text-sm text-gray-600 hover:text-gray-800 border border-gray-300 rounded-lg hover:bg-gray-50"
                >
                  ログアウト
                </button>
              </div>
            </div>
          </div>
          <CaseForm onSuccess={handleRegisterSuccess} />
        </div>
      )}
      {state === 'success' && <Success onBack={handleBackToForm} />}
    </>
  );
}
