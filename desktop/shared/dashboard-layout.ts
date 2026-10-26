// Copyright (c) 2026 AbeLin · MIT License
export const DASHBOARD_CANVAS = { width: 1920, height: 1080 } as const;
export const CANVAS_BOUNDS = { left: 24, top: 120, right: 1896, bottom: 1056 } as const;
export const MIN_CARD = { width: 240, height: 168 } as const;
export interface ChartLayout { x: number; y: number; width: number; height: number }
export type ResizeHandle = 'n' | 's' | 'e' | 'w' | 'ne' | 'nw' | 'se' | 'sw';
const clamp = (n: number, low: number, high: number) => Math.min(high, Math.max(low, Math.round(n)));

export function validLayout(value: unknown): value is ChartLayout {
  if (!value || typeof value !== 'object') return false;
  const r = value as ChartLayout;
  return [r.x, r.y, r.width, r.height].every(Number.isInteger)
    && r.x >= CANVAS_BOUNDS.left && r.y >= CANVAS_BOUNDS.top
    && r.width >= MIN_CARD.width && r.height >= MIN_CARD.height
    && r.x + r.width <= CANVAS_BOUNDS.right && r.y + r.height <= CANVAS_BOUNDS.bottom;
}

// 拖动和数字输入共用边界约束；方案导入另外严格校验，不悄悄改写坏数据。
export function constrainLayout(r: ChartLayout): ChartLayout {
  const width = clamp(r.width, MIN_CARD.width, CANVAS_BOUNDS.right - CANVAS_BOUNDS.left);
  const height = clamp(r.height, MIN_CARD.height, CANVAS_BOUNDS.bottom - CANVAS_BOUNDS.top);
  return { x: clamp(r.x, CANVAS_BOUNDS.left, CANVAS_BOUNDS.right - width),
    y: clamp(r.y, CANVAS_BOUNDS.top, CANVAS_BOUNDS.bottom - height), width, height };
}

export function arrangeLayouts(count: number): ChartLayout[] {
  if (count < 1) return [];
  const cols = count === 1 ? 1 : count <= 4 ? 2 : count <= 9 ? 3 : 4;
  const rows = Math.ceil(count / cols), gap = 24;
  const width = Math.floor((1872 - gap * (cols - 1)) / cols);
  const height = Math.floor((936 - gap * (rows - 1)) / rows);
  return Array.from({ length: count }, (_, i) => ({ x: 24 + (i % cols) * (width + gap),
    y: 120 + Math.floor(i / cols) * (height + gap), width, height }));
}

export function nextLayout(existing: ChartLayout[], metric = false): ChartLayout {
  const width = metric ? 360 : 600, height = metric ? 180 : 420;
  for (let y = 120; y + height <= 1056; y += 12) {
    for (let x = 24; x + width <= 1896; x += 12) {
      if (existing.every((r) => x + width + 12 <= r.x || r.x + r.width + 12 <= x
        || y + height + 12 <= r.y || r.y + r.height + 12 <= y)) return { x, y, width, height };
    }
  }
  // 空间不足时让新图叠放在上层，保留用户已经排好的位置。
  return { x: 24 + (existing.length % 8) * 24, y: 120 + (existing.length % 8) * 24, width, height };
}

export function transformLayout(start: ChartLayout, dx: number, dy: number, handle: ResizeHandle | 'move', snap = true): ChartLayout {
  const step = snap ? 12 : 1;
  const quantize = (n: number) => Math.round(n / step) * step;
  if (handle === 'move') return constrainLayout({ ...start, x: quantize(start.x + dx), y: quantize(start.y + dy) });
  let left = start.x, right = left + start.width, top = start.y, bottom = top + start.height;
  if (handle.includes('w')) left = clamp(quantize(left + dx), 24, right - MIN_CARD.width);
  if (handle.includes('e')) right = clamp(quantize(right + dx), left + MIN_CARD.width, 1896);
  if (handle.includes('n')) top = clamp(quantize(top + dy), 120, bottom - MIN_CARD.height);
  if (handle.includes('s')) bottom = clamp(quantize(bottom + dy), top + MIN_CARD.height, 1056);
  return { x: left, y: top, width: right - left, height: bottom - top };
}
