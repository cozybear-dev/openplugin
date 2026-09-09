export const PIN_THRESHOLD_PX = 64;

export type ScrollMetrics = {
  scrollTop: number;
  scrollHeight: number;
  clientHeight: number;
};

export function isPinnedToBottom(el: ScrollMetrics, thresholdPx = PIN_THRESHOLD_PX): boolean {
  return el.scrollHeight - el.scrollTop - el.clientHeight <= thresholdPx;
}

export function scrollToBottom(el: { scrollTop: number; scrollHeight: number }): void {
  el.scrollTop = el.scrollHeight;
}
