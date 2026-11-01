/**
 * DOCX converter: imports Word documents into the Quire model and exports them back, reusing the
 * original package when available so styles, headers/footers and settings survive a round trip.
 * See docs/SCHEMA.md for the document model contract.
 */

export { importDocx } from './docx/import';
export { exportDocx } from './docx/export';
