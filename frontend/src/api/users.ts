import { apiClient } from './client';

export interface AdminUser {
  id: number;
  username: string;
  name: string;
  email?: string | null;
  role: string;
  is_active: boolean;
  created_at: string;
  updated_at: string;
}

export interface AdminUserCreate {
  username: string;
  password: string;
  name: string;
  email?: string;
  role: string;
  is_active: boolean;
}

export interface AdminUserUpdate {
  password?: string;
  name?: string;
  email?: string;
  role?: string;
  is_active?: boolean;
}

export async function fetchUsers() {
  const response = await apiClient.get<AdminUser[]>('/users');
  return response.data;
}

export async function createUser(payload: AdminUserCreate) {
  const response = await apiClient.post<AdminUser>('/users', payload);
  return response.data;
}

export async function updateUser(id: number, payload: AdminUserUpdate) {
  const response = await apiClient.patch<AdminUser>(`/users/${id}`, payload);
  return response.data;
}

export async function deleteUser(id: number) {
  await apiClient.delete(`/users/${id}`);
}
