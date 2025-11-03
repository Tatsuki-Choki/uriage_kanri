'use client';

import { GoogleLogin, CredentialResponse } from '@react-oauth/google';

export interface User {
  name: string;
  email: string;
  picture?: string;
}

export function handleGoogleLoginSuccess(response: CredentialResponse): Promise<User> {
  return new Promise((resolve, reject) => {
    if (!response.credential) {
      reject(new Error('認証情報が取得できませんでした'));
      return;
    }

    // Google ID Tokenをデコードしてユーザー情報を取得
    // 実際の実装では、バックエンドでID Tokenを検証することを推奨
    // ここでは簡易的にJWTをデコード（実際のプロダクションではサーバー側で検証）
    try {
      const base64Url = response.credential.split('.')[1];
      const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
      const jsonPayload = decodeURIComponent(
        atob(base64)
          .split('')
          .map((c) => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
          .join('')
      );
      const payload = JSON.parse(jsonPayload);
      
      resolve({
        name: payload.name || '',
        email: payload.email || '',
        picture: payload.picture || '',
      });
    } catch (error) {
      reject(error);
    }
  });
}

