const GAS_API_URL = process.env.NEXT_PUBLIC_GAS_API_URL || '';

export interface Client {
  name: string;
}

export interface Status {
  name: string;
}

export interface CaseData {
  sheetType: '月別シート' | '月末請求シート';
  month: string;
  registrationDate?: string;
  clientName: string;
  industry?: string;
  expense?: number;
  expenseUsd?: number;
  revenue: number;
  orderDate?: string;
  status: string;
  notes: string;
}

export interface ApiResponse<T> {
  success: boolean;
  error?: string;
  data?: T;
}

async function checkResponse(response: Response): Promise<any> {
  const contentType = response.headers.get('content-type');
  
  // HTMLが返ってきた場合（Google Apps ScriptがWebアプリとして公開されていない）
  if (contentType && contentType.includes('text/html')) {
    throw new Error('Google Apps ScriptがWebアプリとして公開されていません。Google Apps ScriptエディタでWebアプリとして公開してください。');
  }

  if (!response.ok) {
    throw new Error(`APIリクエストに失敗しました: ${response.status} ${response.statusText}`);
  }

  try {
    const data = await response.json();
    return data;
  } catch (error) {
    throw new Error('APIレスポンスの解析に失敗しました。JSON形式のデータが返されていません。');
  }
}

export async function fetchClients(): Promise<string[]> {
  if (!GAS_API_URL) {
    throw new Error('GAS_API_URLが設定されていません');
  }

  try {
    const url = `${GAS_API_URL}?action=clients`;
    const response = await fetch(url);
    const data = await checkResponse(response);
    
    if (!data.success) {
      throw new Error(data.error || 'クライアント一覧の取得に失敗しました');
    }

    return data.clients || [];
  } catch (error) {
    console.error('fetchClients error:', error);
    throw error instanceof Error ? error : new Error('クライアント一覧の取得に失敗しました');
  }
}

export async function fetchStatuses(): Promise<string[]> {
  if (!GAS_API_URL) {
    throw new Error('GAS_API_URLが設定されていません');
  }

  try {
    const url = `${GAS_API_URL}?action=statuses`;
    const response = await fetch(url);
    const data = await checkResponse(response);
    
    if (!data.success) {
      throw new Error(data.error || 'ステータス一覧の取得に失敗しました');
    }

    return data.statuses || [];
  } catch (error) {
    console.error('fetchStatuses error:', error);
    throw error instanceof Error ? error : new Error('ステータス一覧の取得に失敗しました');
  }
}

export async function fetchMonths(): Promise<string[]> {
  if (!GAS_API_URL) {
    throw new Error('GAS_API_URLが設定されていません');
  }

  try {
    const url = `${GAS_API_URL}?action=months`;
    const response = await fetch(url);
    const data = await checkResponse(response);
    
    if (!data.success) {
      throw new Error(data.error || '月一覧の取得に失敗しました');
    }

    return data.months || [];
  } catch (error) {
    console.error('fetchMonths error:', error);
    throw error instanceof Error ? error : new Error('月一覧の取得に失敗しました');
  }
}

export async function registerCase(caseData: CaseData): Promise<{ success: boolean; message?: string; error?: string }> {
  if (!GAS_API_URL) {
    throw new Error('GAS_API_URLが設定されていません');
  }

  try {
    // POSTリクエストをGETリクエストに変更（CORSエラー回避のため）
    const params = new URLSearchParams({
      action: 'registerCase',
      ...Object.fromEntries(
        Object.entries(caseData).map(([key, value]) => [
          key,
          value !== undefined && value !== null ? String(value) : ''
        ])
      ),
    });

    const url = `${GAS_API_URL}?${params.toString()}`;
    const response = await fetch(url, {
      method: 'GET',
    });

    const data = await checkResponse(response);
    
    if (!data.success) {
      throw new Error(data.error || '案件登録に失敗しました');
    }

    return data;
  } catch (error) {
    console.error('registerCase error:', error);
    throw error instanceof Error ? error : new Error('案件登録に失敗しました');
  }
}

