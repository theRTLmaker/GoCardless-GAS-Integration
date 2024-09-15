/**
 * Sanitizes a string to prevent potential security issues.
 * This function removes or escapes characters that could be used for injection attacks.
 *
 * @param input The input string to sanitize
 * @returns The sanitized string
 */
export function sanitizeString(input: string): string {
  if (typeof input !== 'string') {
    return '';
  }

  // Remove any HTML tags
  let sanitized = input.replace(/<[^>]*>/g, '');

  // Escape special characters
  sanitized = sanitized
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;')
    .replace(/\//g, '&#x2F;');

  // Remove any non-printable characters
  sanitized = sanitized.replace(/[^\x20-\x7E]/g, '');

  // Trim whitespace from both ends
  sanitized = sanitized.trim();

  return sanitized;
}