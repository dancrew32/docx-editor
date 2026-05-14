import type { WrapTextDirection } from '../../layout-engine/types';

export interface FloatingExclusionRect {
  /** Which side the object is on for simple one-sided wrapping. */
  side: 'left' | 'right';
  /** X position relative to the content area. */
  x: number;
  /** Y position relative to the content area. */
  y: number;
  width: number;
  height: number;
  distTop: number;
  distBottom: number;
  distLeft: number;
  distRight: number;
  wrapText?: WrapTextDirection;
  wrapType?: string;
}

export interface FloatingImageZone {
  leftMargin: number;
  rightMargin: number;
  topY: number;
  bottomY: number;
  segments?: FloatingLineSegmentZone[];
}

export interface FloatingLineSegmentZone {
  leftOffset: number;
  availableWidth: number;
}

export function rectsToFloatingZones(
  rects: FloatingExclusionRect[],
  contentWidth: number
): FloatingImageZone[] {
  return rects.map((rect) => {
    const rectLeft = rect.x - rect.distLeft;
    const rectRight = rect.x + rect.width + rect.distRight;
    const rectTop = rect.y - rect.distTop;
    const rectBottom = rect.y + rect.height + rect.distBottom;

    let leftMargin = 0;
    let rightMargin = 0;
    let segments: FloatingLineSegmentZone[] | undefined;

    const wt = rect.wrapText ?? 'bothSides';

    if (wt === 'right') {
      leftMargin = Math.max(0, rectRight);
    } else if (wt === 'left') {
      rightMargin = Math.max(0, contentWidth - rectLeft);
    } else if (rectLeft > 0 && rectRight < contentWidth) {
      segments = [
        { leftOffset: 0, availableWidth: Math.max(0, rectLeft) },
        {
          leftOffset: Math.max(0, rectRight),
          availableWidth: Math.max(0, contentWidth - rectRight),
        },
      ].filter((segment) => segment.availableWidth > 1);
      leftMargin = 0;
      rightMargin = 0;
    } else if (rect.side === 'left') {
      leftMargin = Math.max(0, rectRight);
    } else {
      rightMargin = Math.max(0, contentWidth - rectLeft);
    }

    return { leftMargin, rightMargin, topY: rectTop, bottomY: rectBottom, segments };
  });
}

export function getFloatingMargins(
  lineY: number,
  lineHeight: number,
  zones: FloatingImageZone[] | undefined,
  paragraphYOffset: number
): { leftMargin: number; rightMargin: number; segments?: FloatingLineSegmentZone[] } {
  if (!zones || zones.length === 0) {
    return { leftMargin: 0, rightMargin: 0 };
  }

  let leftMargin = 0;
  let rightMargin = 0;
  let segments: FloatingLineSegmentZone[] | undefined;

  const absoluteLineTop = paragraphYOffset + lineY;
  const absoluteLineBottom = absoluteLineTop + lineHeight;

  for (const zone of zones) {
    if (absoluteLineBottom <= zone.topY || absoluteLineTop >= zone.bottomY) continue;
    if (zone.segments?.length) {
      segments = segments ? intersectSegments(segments, zone.segments) : zone.segments;
      continue;
    }
    leftMargin = Math.max(leftMargin, zone.leftMargin);
    rightMargin = Math.max(rightMargin, zone.rightMargin);
  }

  return { leftMargin, rightMargin, segments };
}

function intersectSegments(
  a: FloatingLineSegmentZone[],
  b: FloatingLineSegmentZone[]
): FloatingLineSegmentZone[] {
  const result: FloatingLineSegmentZone[] = [];
  for (const left of a) {
    for (const right of b) {
      const start = Math.max(left.leftOffset, right.leftOffset);
      const end = Math.min(
        left.leftOffset + left.availableWidth,
        right.leftOffset + right.availableWidth
      );
      if (end > start) {
        result.push({ leftOffset: start, availableWidth: end - start });
      }
    }
  }
  return result;
}
