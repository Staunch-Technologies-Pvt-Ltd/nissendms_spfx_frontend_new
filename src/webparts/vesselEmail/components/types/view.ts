// View/navigation & modal-mode types extracted from VesselEmail.tsx

export type ModalMode = 'none' | 'create' | 'edit' | 'delete' | 'user' | 'template' | 'folder';

export type AppView = 
  | 'dashboard'
  | 'list'
  | 'sites'
  | 'vessels'
  | 'templates'
  | 'approvals'
  | 'reports'
  | 'users'
  | 'settings'
  | 'bento_email'
  | 'email_notify'
  | 'recycle'
  | 'archive'
  | 'alerts';
