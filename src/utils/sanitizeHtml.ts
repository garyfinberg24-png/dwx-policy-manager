import DOMPurify from 'dompurify';

const ALLOWED_TAGS = [
  'p', 'br', 'b', 'i', 'u', 'strong', 'em',
  'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
  'ul', 'ol', 'li',
  'a',
  'table', 'thead', 'tbody', 'tr', 'td', 'th',
  'span', 'div',
  'img',
  'blockquote', 'pre', 'code',
];

const ALLOWED_ATTR = [
  'href', 'target', 'rel', 'src', 'alt', 'width', 'height',
  'class', 'style', 'colspan', 'rowspan',
];

/**
 * Sanitize HTML content using DOMPurify.
 * Strips dangerous tags/attributes while preserving safe formatting.
 */
export function sanitizeHtml(dirty: string): string {
  if (!dirty) return '';
  return DOMPurify.sanitize(dirty, {
    ALLOWED_TAGS,
    ALLOWED_ATTR,
    ALLOW_DATA_ATTR: false,
  });
}

/**
 * Escape HTML special characters for safe embedding in HTML templates.
 * Use this for plain text content in email templates — converts &, <, >, ", ' to entities.
 */
export function escapeHtml(text: string): string {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
