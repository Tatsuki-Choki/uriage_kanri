'use client';

import { GoogleLogin, CredentialResponse } from '@react-oauth/google';
import { handleGoogleLoginSuccess, User } from '@/lib/auth';

interface LoginProps {
  onLoginSuccess: (user: User) => void;
}

export default function Login({ onLoginSuccess }: LoginProps) {
  const handleSuccess = async (response: CredentialResponse) => {
    try {
      const user = await handleGoogleLoginSuccess(response);
      onLoginSuccess(user);
    } catch (error) {
      console.error('ログインエラー:', error);
      alert('ログインに失敗しました。もう一度お試しください。');
    }
  };

  const handleError = () => {
    alert('ログインに失敗しました。もう一度お試しください。');
  };

  return (
    <div className="flex flex-col items-center justify-center min-h-screen bg-gray-50 px-4">
      <div className="w-full max-w-md bg-white rounded-lg shadow-md p-8">
        <h1 className="text-2xl font-bold text-center mb-6 text-gray-800">
          案件登録システム
        </h1>
        <p className="text-center text-gray-600 mb-8">
          Googleアカウントでログインしてください
        </p>
        <div className="flex justify-center">
          <GoogleLogin
            onSuccess={handleSuccess}
            onError={handleError}
            useOneTap={false}
          />
        </div>
      </div>
    </div>
  );
}

