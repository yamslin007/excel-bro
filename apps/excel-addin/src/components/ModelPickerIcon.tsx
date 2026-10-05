/** Shared model-menu glyph; the owning button provides its accessible name. */
export function ModelPickerIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
      <rect x="6" y="6" width="12" height="12" rx="3" />
      <path d="M9 3v3m6-3v3M9 18v3m6-3v3M3 9h3m-3 6h3M18 9h3m-3 6h3" />
      <path d="M10 10h4v4h-4z" />
    </svg>
  );
}
