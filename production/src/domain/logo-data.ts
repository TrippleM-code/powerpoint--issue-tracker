/** Only embedded raster images may cross from document metadata into image sinks. */
export function normalizeLogoDataUrl(value: unknown): string | undefined {
  if (typeof value !== "string" || value.length > 1024 * 1024) return undefined;
  const logo = value.trim();
  return /^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+/]+={0,2}$/i.test(logo)
    ? logo : undefined;
}
