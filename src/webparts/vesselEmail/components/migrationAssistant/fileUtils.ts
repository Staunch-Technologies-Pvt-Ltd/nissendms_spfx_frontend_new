// Small file-display helpers for the Migration Assistant module — ported
// from the standalone project's fileUtils.ts, using Fluent UI icon names
// (this module renders inside SPFx/Fluent UI 8, not lucide-react).

export function fileIconName(ext?: string): { iconName: string; color: string; label: string; previewable: boolean } {
  const e = (ext ?? '').toLowerCase();
  if (e === 'pdf') return { iconName: 'PDF', color: '#e11d48', label: 'PDF', previewable: true };
  if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tif', 'tiff'].includes(e))
    return { iconName: 'FileImage', color: '#7c3aed', label: e.toUpperCase(), previewable: true };
  if (['xls', 'xlsx', 'csv'].includes(e))
    return { iconName: 'ExcelDocument', color: '#059669', label: e.toUpperCase(), previewable: false };
  if (['doc', 'docx'].includes(e))
    return { iconName: 'WordDocument', color: '#2563eb', label: e.toUpperCase(), previewable: false };
  if (['zip', 'rar', '7z'].includes(e))
    return { iconName: 'ZipFolder', color: '#d97706', label: e.toUpperCase(), previewable: false };
  return { iconName: 'Page', color: '#64748b', label: e ? e.toUpperCase() : 'FILE', previewable: false };
}

export function formatSize(bytes?: number | null): string {
  if (bytes === null || bytes === undefined) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
