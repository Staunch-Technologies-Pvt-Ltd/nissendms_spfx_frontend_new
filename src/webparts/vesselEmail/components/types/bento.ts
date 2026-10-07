// Bento email log type extracted from VesselEmail.tsx

export interface BentoEmailLog {
  id: number;
  datasource_tag_requested?: string;
  datasource_tag_used: string;
  tag_label?: string;
  tag_was_valid: boolean;
  vessel_name?: string;
  subject: string;
  body?: string;
  recipient: string;
  status: 'pending' | 'completed' | 'failed' | string;
  display_status?: string;
  error_message?: string;
  created_at?: string;
  sent_at?: string;
  attachments_count: number;
  attachment_names: string[];
  /** Compose email: mailbox it was sent from, and documents it carried. */
  sender?: string | null;
  documents?: Array<{ name: string; kind: 'file' | 'link' | 'folder' | string }>;
}

