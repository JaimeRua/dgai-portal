
/** Validates navigation destinations received from configurable SharePoint lists. */
export function getSafePortalUrl(value: string | undefined): string | undefined {
  const raw = (value || '').trim();

  if (
    !raw ||
    raw === '#' ||
    Array.from(raw).some((char) => {
      const code = char.charCodeAt(0);
      return code <= 31 || code === 127;
    })
  ) {
    return undefined;
  }

  // Reject protocol-relative URLs and backslashes to avoid ambiguous origins.
  if (raw.startsWith('//') || raw.includes('\\')) {
    return undefined;
  }

  try {
    const url = new URL(raw, window.location.origin);

    if (url.username || url.password) {
      return undefined;
    }

    if (url.origin === window.location.origin) {
      return url.href;
    }

    return url.protocol === 'https:' ? url.href : undefined;

  } catch {
    return undefined;
  }
}
