import { apiClient } from './client';

export type SalesLead = {
  id: number;
  customer_name: string;
  owner_name: string;
  owner_contact?: string | null;
  meeting_note?: string | null;
  expected_amount: number;
  status: string;
  claimed_at: string;
};

export type PaymentMatch = {
  id: number;
  sales_lead_id: number;
  customer_name: string;
  owner_name: string;
  depositor_name: string;
  amount: number;
  matched_rule: string;
  confidence: number;
  created_at: string;
};

export type PipelineSummary = {
  total_leads: number;
  waiting_payment: number;
  paid: number;
  total_expected_amount: number;
  total_paid_amount: number;
  overdue_risk: SalesLead[];
};

export type CreateLeadPayload = {
  customer_name: string;
  owner_name: string;
  owner_contact?: string;
  meeting_note?: string;
  expected_amount: number;
};

export type MatchTransactionPayload = {
  depositor_name: string;
  amount: number;
};

export async function fetchSalesDashboard() {
  const [summary, leads, matches] = await Promise.all([
    apiClient.get<PipelineSummary>('/sales/summary'),
    apiClient.get<SalesLead[]>('/sales/leads'),
    apiClient.get<PaymentMatch[]>('/sales/matches'),
  ]);

  return {
    summary: summary.data,
    leads: leads.data,
    matches: matches.data,
  };
}

export async function createSalesLead(payload: CreateLeadPayload) {
  const response = await apiClient.post<SalesLead>('/sales/leads', payload);
  return response.data;
}

export async function matchTransaction(payload: MatchTransactionPayload) {
  const response = await apiClient.post<PaymentMatch | null>('/sales/transactions/match', payload);
  return response.data;
}
