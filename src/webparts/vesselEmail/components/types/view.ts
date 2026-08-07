// View/navigation & modal-mode types extracted from VesselEmail.tsx

export type ModalMode = 'none' | 'create' | 'edit' | 'delete' | 'user' | 'template' | 'folder';

export type AppView = 
  | 'dashboard'
  | 'list'
  | 'vessels'
  | 'templates'
  | 'approvals'
  | 'notifications'
  | 'reports'
  | 'users'
  | 'settings'
  | 'bento_email'
  | 'email_notify'
  | 'recycle'
  | 'archive';
