import type { ViewerAPI } from '../electron/types';

export const DESKTOP_UPDATE_MESSAGE = 'The desktop component is from an older session. Copy any unsaved data before closing, then restart Clear Sheet (or stop and rerun npm run dev). The save dialog requires the updated desktop component.';

export function requireSaveBridge(api: ViewerAPI): void {
  if (typeof api?.chooseSavePath !== 'function' || typeof api.save !== 'function') throw new Error(DESKTOP_UPDATE_MESSAGE);
}
