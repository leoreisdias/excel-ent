export const convertType = (
  type: 'string' | 'number' | 'boolean' | 'date' | 'blank' | 'error',
) => {
  switch (type) {
    case 'string':
      return 's';
    case 'number':
      return 'n';
    case 'boolean':
      return 'b';
    case 'date':
      return 'd';
    case 'blank':
      return 'z';
    case 'error':
      return 'e';
    default:
      return 's';
  }
};

export function downloadFile(output: string, fileName: string): void {
  const link = document.createElement('a');
  document.body.appendChild(link);
  link.download = fileName;
  link.href = output;
  link.click();
}

export function objectToSemicolons(data: any[]) {
  const colsHead = Object.keys(data[0])
    .map(key => [key])
    .join(';');
  const colsData = data
    .map(obj => [
      Object.keys(obj)
        .map(col => [obj[col]])
        .join(';'),
    ])
    .join('\n');

  return `${colsHead}\n${colsData}`;
}
