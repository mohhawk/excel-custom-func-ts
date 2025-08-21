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
    const errorResponse = await response.json().catch(() => response.text());
    console.error('Full error response:', errorResponse);

    let errorMessage = `HTTP error! status: ${response.status}`;
    if (typeof errorResponse === 'object' && errorResponse !== null && errorResponse.message) {
      errorMessage = errorResponse.message;
    } else if (typeof errorResponse === 'string') {
      try {
        const parsed = JSON.parse(errorResponse);
        if (parsed.message) {
          errorMessage = parsed.message;
        }
      } catch (e) {
        // Not a JSON string, use the raw text if it's not too long
        errorMessage = errorResponse.substring(0, 100);
      }
    }
    throw new Error(errorMessage);
  }

  return response.json();
}
