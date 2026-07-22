# ep_data_tables

Accessible, collaborative tables for Etherpad. Tables remain part of Etherpad's line-based document model, so cell content can retain hyperlinks, font styling, inline images, authorship colors, and compatible block styles.

![Table editing in Etherpad](https://i.imgur.com/AOdt59T.png)

## Features

- Grid picker for creating tables
- Add and remove rows or columns
- Keyboard navigation between cells
- Pointer-based column resizing and precise width controls
- Optional caption, column headers, and row headers
- Backward-compatible rendering of existing table metadata
- Read-only table rendering in the Etherpad timeslider
- Preservation of unknown metadata during intentional edits

## Installation

From the Etherpad directory:

```sh
pnpm run plugins i ep_data_tables
```

Restart Etherpad after installation. This release supports Etherpad 3.3.2 and later 3.x releases on Node.js 20 or newer.

## Editing tables

Use the table toolbar button to choose an initial size. The table menu supports row and column insertion or deletion, table deletion, and table properties.

The Table Properties dialog can set:

- an accessible caption;
- whether the first row contains column headers;
- whether the first column contains row headers; and
- exact column widths.

Keyboard users can move through cells with Tab and arrow-key navigation. Width controls provide a keyboard-accessible alternative to drag handles.

## Data model and compatibility

Each table row remains one Etherpad line. A `tbljson` line attribute stores the table ID, row position, column count, widths, and additive accessibility metadata. Cell text remains ordinary attributed Etherpad content.

Historical width values are rendered without opening-time normalization. Intentional metadata changes normalize widths to 100 percent with at most four decimal places. Unknown metadata fields are retained so historical and future content can round-trip safely.

Legacy tables treat the first row as column headers by default. Live-editor cells keep the established `<td>` structure and expose header relationships through ARIA roles and references. The read-only timeslider uses native table-header elements.

## Limitations

- Merged cells are not supported.
- Copying complete table structures between pads is not supported.
- Very large tables can be expensive because Etherpad renders each row through its collaborative line model.

## Export support

This package does not register HTML or document-export hooks. Install a separately reviewed table-export plugin if exported tables are required. Separating export behavior prevents a partial renderer from intercepting Etherpad's general export pipeline.

## Development

```sh
pnpm install --frozen-lockfile
pnpm test
pnpm lint
```

The implementation is derived from the `ep_tables5` family and is distributed under the Apache License 2.0. See `LICENSE` for details.
