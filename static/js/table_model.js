'use strict';

(() => {
// This module is deliberately dependency-free and usable from CommonJS plugin
  // hooks, direct browser scripts (timeslider), and Node's unit-test runner.
  const WIDTH_DECIMAL_PLACES = 4;

  const isFinitePositive = (value) => Number.isFinite(value) && value > 0;

  const asColumnCount = (value) => {
    const count = Number(value);
    return Number.isInteger(count) && count > 0 ? count : 0;
  };

  const roundWidth = (value) => Number(Number(value).toFixed(WIDTH_DECIMAL_PLACES));

  const equalColumnWidths = (columnCount) => {
    const count = asColumnCount(columnCount);
    if (!count) return [];
    const base = roundWidth(100 / count);
    const widths = Array(count).fill(base);
    const difference = roundWidth(100 - widths.reduce((sum, width) => sum + width, 0));
    widths[count - 1] = roundWidth(widths[count - 1] + difference);
    return widths;
  };

  const hasUsableColumnWidths = (widths, columnCount) => {
    const count = asColumnCount(columnCount);
    return count > 0 && Array.isArray(widths) && widths.length === count &&
    widths.every(isFinitePositive);
  };

  // Rendering legacy metadata must not mutate or silently rewrite it. This
  // returns an exact copy of valid historical widths, regardless of precision.
  const columnWidthsForRender = (metadata, columnCount) => {
    const count = asColumnCount(columnCount);
    const widths = metadata && metadata.columnWidths;
    return hasUsableColumnWidths(widths, count) ? widths.slice() : equalColumnWidths(count);
  };

  // Normalize only when metadata is intentionally written. Ratios are retained,
  // values are bounded to four decimals, and the serialized total is exactly 100.
  const normalizeColumnWidthsForWrite = (widths, columnCount) => {
    const count = asColumnCount(columnCount);
    if (!hasUsableColumnWidths(widths, count)) return equalColumnWidths(count);
    const total = widths.reduce((sum, width) => sum + width, 0);
    if (!isFinitePositive(total)) return equalColumnWidths(count);

    const normalized = widths.map((width) => roundWidth((width / total) * 100));
    const difference = roundWidth(100 - normalized.reduce((sum, width) => sum + width, 0));
    const adjustmentIndex = normalized.reduce(
        (largest, width, index, values) => width > values[largest] ? index : largest, 0);
    normalized[adjustmentIndex] = roundWidth(normalized[adjustmentIndex] + difference);
    return normalized;
  };

  const normalizeMetadataForWrite = (metadata, required = {}) => {
    const source = metadata && typeof metadata === 'object' ? metadata : {};
    const result = {...source, ...required};
    const columnCount = asColumnCount(result.cols);
    if (columnCount) {
      result.cols = columnCount;
      result.columnWidths = normalizeColumnWidthsForWrite(result.columnWidths, columnCount);
    }
    return result;
  };

  // DOM collection is not an author action. Preserve all unknown and historical
  // fields byte-for-byte where possible, updating only canonical row identity.
  const mergeCollectedMetadata = (existing, required = {}) => {
    const source = existing && typeof existing === 'object' ? existing : {};
    const result = {...source, ...required};
    const columnCount = asColumnCount(result.cols);
    if (!hasUsableColumnWidths(result.columnWidths, columnCount)) delete result.columnWidths;
    return result;
  };

  const headerRowCount = (metadata) => {
    if (metadata && Object.prototype.hasOwnProperty.call(metadata, 'headerRows')) {
      return Number(metadata.headerRows) === 1 ? 1 : 0;
    }
    // ep_data_tables is a data-table plugin. Legacy tables conventionally use
    // their first row as column labels; authors can explicitly opt out.
    return 1;
  };

  const headerColumnCount = (metadata) => {
    if (!metadata || !Object.prototype.hasOwnProperty.call(metadata, 'headerColumns')) return 0;
    return Number(metadata.headerColumns) === 1 ? 1 : 0;
  };

  const tableCaption = (metadata) => {
    if (!metadata || typeof metadata.caption !== 'string') return '';
    return metadata.caption.replace(/\s+/g, ' ').trim();
  };

  const tablePropertiesForEditing = (metadata, columnCount, rowCount = 0) => ({
    caption: tableCaption(metadata),
    cols: asColumnCount(columnCount),
    columnWidths: columnWidthsForRender(metadata, columnCount),
    headerColumns: headerColumnCount(metadata),
    headerRows: headerRowCount(metadata),
    rows: Math.max(0, Number(rowCount) || 0),
  });

  const metadataWithTableProperties = (metadata, properties, required = {}) => {
    const source = properties && typeof properties === 'object' ? properties : {};
    const caption = typeof source.caption === 'string'
      ? source.caption.replace(/\s+/g, ' ').trim().slice(0, 240) : '';
    return normalizeMetadataForWrite({
      ...(metadata && typeof metadata === 'object' ? metadata : {}),
      caption,
      headerRows: Number(source.headerRows) === 1 ? 1 : 0,
      headerColumns: Number(source.headerColumns) === 1 ? 1 : 0,
      columnWidths: source.columnWidths,
    }, required);
  };

  const api = {
    WIDTH_DECIMAL_PLACES,
    columnWidthsForRender,
    equalColumnWidths,
    hasUsableColumnWidths,
    headerColumnCount,
    headerRowCount,
    mergeCollectedMetadata,
    metadataWithTableProperties,
    normalizeColumnWidthsForWrite,
    normalizeMetadataForWrite,
    tableCaption,
    tablePropertiesForEditing,
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof window !== 'undefined') window.epDataTablesTableModel = api;
})();
