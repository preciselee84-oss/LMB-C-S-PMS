import { apiClient } from './client';

export type BillingPreviewRow = {
  source_type: string;
  sequence: string;
  customer_number: string;
  business_number: string;
  company_name: string;
  manager_name: string;
  base_date: string;
  first_login: string;
  latest_login: string;
  login_count: number;
  billing_company_name: string;
  bank_company_name: string;
  match_status: string;
  note: string;
};

export type BillingPreviewSummary = {
  total_count: number;
  matched_count: number;
  name_mismatch_count: number;
  missing_count: number;
  open_count: number;
  erp_count: number;
};

export type BillingPreview = {
  spreadsheet_title: string;
  spreadsheet_url: string;
  generated_from: string[];
  rows: BillingPreviewRow[];
  summary: BillingPreviewSummary;
};

export async function fetchBillingPreview() {
  const response = await apiClient.get<BillingPreview>('/billing/preview');
  return response.data;
}

export async function uploadBillingLoginFile(file: File) {
  const formData = new FormData();
  formData.append('login_file', file);
  const response = await apiClient.post<BillingPreview>('/billing/preview', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
  });
  return response.data;
}
