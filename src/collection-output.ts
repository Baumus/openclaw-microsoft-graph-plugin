/** Include JSON string escaping used when a host embeds the model-facing text. */
export function modelTextBytes(value: unknown): number {
  return Buffer.byteLength(JSON.stringify(JSON.stringify(value)), "utf8");
}

/** Fit whole records only; the caller must retain or explicitly report the remainder. */
export function boundedCollectionOutput<T, V>(
  items: T[],
  maximumBytes: number,
  render: (selected: T[], hasRemaining: boolean) => V,
  measure: (value: V) => number = modelTextBytes,
): { value: V; remaining: T[] } {
  if (!Number.isSafeInteger(maximumBytes) || maximumBytes < 1024) throw new Error("invalid_byte_limit");
  for (let count = items.length; count >= (items.length ? 1 : 0); count -= 1) {
    const value = render(items.slice(0, count), count < items.length);
    if (measure(value) <= maximumBytes) return { value, remaining: items.slice(count) };
  }
  throw new Error("invalid_collection_item_too_large");
}
