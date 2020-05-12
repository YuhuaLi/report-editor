export function inRange(
  num: number,
  start: number,
  end: number,
  inclusive?: boolean
) {
  return inclusive
    ? num >= Math.min(start, end) && num <= Math.max(start, end)
    : num > Math.min(start, end) && num < Math.max(start, end);
}
