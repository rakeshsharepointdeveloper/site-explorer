const OFFICE_EXT = /\.(docx?|xlsx?|pptx?|one|vsdx?|docm|xlsm|pptm)$/i;

/** Fluent UI icon name for a file, by extension. */
export function fileIcon(name: string): string {
  const ext = (name.split('.').pop() ?? '').toLowerCase();
  if (ext.indexOf('doc') === 0) return 'WordDocument';
  if (ext.indexOf('xls') === 0 || ext === 'csv') return 'ExcelDocument';
  if (ext.indexOf('ppt') === 0) return 'PowerPointDocument';
  if (ext === 'pdf') return 'PDF';
  if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg', 'webp'].indexOf(ext) >= 0) return 'FileImage';
  if (['json', 'js', 'ts', 'xml', 'html', 'css'].indexOf(ext) >= 0) return 'FileCode';
  return 'Page';
}

/** Office files open in Office for the web; everything else opens as-is. */
export function officeUrl(name: string, url: string): string {
  if (!OFFICE_EXT.test(name)) return url;
  return `${url}${url.indexOf('?') >= 0 ? '&' : '?'}web=1`;
}
