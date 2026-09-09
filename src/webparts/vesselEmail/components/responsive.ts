export const MOBILE_MAX = 767;
export const TABLET_MAX = 1024;

export function isMobileWidth(width: number): boolean {
  return width <= MOBILE_MAX;
}

export function isTabletWidth(width: number): boolean {
  return width > MOBILE_MAX && width <= TABLET_MAX;
}

export function isTabletOrBelow(width: number): boolean {
  return width <= TABLET_MAX;
}
