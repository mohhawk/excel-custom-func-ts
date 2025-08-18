import { getApiUrl } from './url';
import { getEPMSettings } from '../taskpane/taskpane';

export async function exportDataSlice(cubeName: string, payload: any): Promise<any> {
  const settings = getEPMSettings();
  const url = getApiUrl('/api/exportDataSlice');

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      cubeName,
      payload,
      settings,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
  }

  return response.json();
}
