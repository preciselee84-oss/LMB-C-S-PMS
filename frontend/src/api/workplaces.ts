import { apiClient } from './client';

export interface Workplace {
  id: number;
  workplace_name: string;
  business_number?: string | null;
  business_alias?: string | null;
  regular_payment_day: number;
  manager_name?: string | null;
  manager_contact?: string | null;
  memo?: string | null;
  created_at: string;
  updated_at: string;
}

export interface BankAccount {
  id: number;
  account_type: string;
  account_name: string;
  bank_name: string;
  account_number: string;
  holder_name?: string | null;
  linked_workplace_id?: number | null;
  balance: number;
  memo?: string | null;
  balance_updated_at?: string | null;
  created_at: string;
  updated_at: string;
}

export interface AdvanceRequest {
  id: number;
  workplace_id: number;
  workplace_name: string;
  request_amount: number;
  requested_by: string;
  request_reason?: string | null;
  status: string;
  reject_reason?: string | null;
  approved_at?: string | null;
  processed_by?: string | null;
  transfer_generated_at?: string | null;
  paid_at?: string | null;
  requested_at: string;
}

export interface WorkplaceForecast {
  workplace_id: number;
  workplace_name: string;
  average_monthly_amount: number;
  suggested_amount: number;
  guide: string;
}

export interface WorkplaceSummary {
  workplace_count: number;
  request_count: number;
  pending_count: number;
  approved_count: number;
  paid_count: number;
  paid_amount: number;
  month_request_count: number;
  accounts_balance: number;
  forecasts: WorkplaceForecast[];
}

export interface TransferRow {
  request_id: number;
  workplace_name: string;
  bank_name: string;
  account_number: string;
  holder_name: string;
  amount: number;
  memo: string;
}

export async function fetchWorkplaceDashboard() {
  const [summary, workplaces, accounts, requests] = await Promise.all([
    apiClient.get<WorkplaceSummary>('/workplaces/summary'),
    apiClient.get<Workplace[]>('/workplaces'),
    apiClient.get<BankAccount[]>('/workplaces/accounts'),
    apiClient.get<AdvanceRequest[]>('/workplaces/requests'),
  ]);

  return {
    summary: summary.data,
    workplaces: workplaces.data,
    accounts: accounts.data,
    requests: requests.data,
  };
}

export async function createWorkplace(payload: Partial<Workplace>) {
  const response = await apiClient.post<Workplace>('/workplaces', payload);
  return response.data;
}

export async function createBankAccount(payload: Partial<BankAccount>) {
  const response = await apiClient.post<BankAccount>('/workplaces/accounts', payload);
  return response.data;
}

export async function createAdvanceRequest(payload: {
  workplace_id: number;
  request_amount: number;
  requested_by: string;
  request_reason?: string;
}) {
  const response = await apiClient.post<AdvanceRequest>('/workplaces/requests', payload);
  return response.data;
}

export async function approveAdvanceRequest(id: number, processed_by: string) {
  const response = await apiClient.post<AdvanceRequest>(`/workplaces/requests/${id}/approve`, { processed_by });
  return response.data;
}

export async function rejectAdvanceRequest(id: number, processed_by: string, reject_reason?: string) {
  const response = await apiClient.post<AdvanceRequest>(`/workplaces/requests/${id}/reject`, {
    processed_by,
    reject_reason,
  });
  return response.data;
}

export async function generateTransferRow(id: number) {
  const response = await apiClient.post<TransferRow>(`/workplaces/requests/${id}/transfer`);
  return response.data;
}

export async function markRequestPaid(id: number) {
  const response = await apiClient.post<AdvanceRequest>(`/workplaces/requests/${id}/paid`);
  return response.data;
}
