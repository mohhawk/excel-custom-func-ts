export function getBaseUrl(): string {
  if (process.env.NODE_ENV === 'development' || window.location.hostname === 'localhost') {
    return 'http://localhost:3001';
  } else {
    return 'https://github-jirventures-cube-olap-excel-view-32764122184.us-central1.run.app';
  }
}

export function getApiUrl(endpoint: string): string {
  return `${getBaseUrl()}${endpoint}`;
}
