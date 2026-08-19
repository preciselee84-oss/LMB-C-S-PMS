import { apiClient } from './client';

export type VisitVocPayload = {
  company_name: string;
  visit_date: string;
  visitor_name: string;
  contact_name?: string;
  channel: string;
  sentiment: string;
  product_area?: string;
  voc_text: string;
  next_action?: string;
};

export type VisitVoc = VisitVocPayload & {
  id: string;
  created_at: string;
  status: string;
};

export type MeetingMinutes = {
  title: string;
  company_name?: string | null;
  meeting_date?: string | null;
  participants: string[];
  source_file_name?: string | null;
  summary: string;
  key_topics: string[];
  decisions: string[];
  action_items: string[];
  risks: string[];
  original_transcript: string;
};

export async function createVisitVoc(payload: VisitVocPayload) {
  const response = await apiClient.post<VisitVoc>('/voc/entries', payload);
  return response.data;
}

export async function createMinutesFromUpload(payload: {
  title: string;
  company_name?: string;
  meeting_date?: string;
  participants?: string;
  transcript_text?: string;
  recording_file?: File;
}) {
  const formData = new FormData();
  formData.append('title', payload.title);
  if (payload.company_name) formData.append('company_name', payload.company_name);
  if (payload.meeting_date) formData.append('meeting_date', payload.meeting_date);
  if (payload.participants) formData.append('participants', payload.participants);
  if (payload.transcript_text) formData.append('transcript_text', payload.transcript_text);
  if (payload.recording_file) formData.append('recording_file', payload.recording_file);

  const response = await apiClient.post<MeetingMinutes>('/voc/minutes/upload', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
  });
  return response.data;
}
