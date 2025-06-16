// ep_tables5 – toolbar & context‑menu wiring for the attribute‑based engine
// -----------------------------------------------------------------------------
// This companion file keeps the old UX (insert‑table grid, row/col ops, etc.)
// but calls the **new** Ace helpers (`ace_createTableViaAttributes`,
// `ace_doDatatableOptions`) that are exposed by client_hooks.js.
//
// It relies on jQuery because Etherpad already bundles it.

/* global $ */

const log = (...m) => console.debug('[ep_tables5:initialisation]', ...m);

// NEW: Helper function to get line number (adapted from ep_image_insert)
// Needs to be defined before being used in the mousedown handler
function _getLineNumberOfElement(element) {
    // Implementation similar to ep_image_insert
    let currentElement = element;
    let count = 0;
    while (currentElement = currentElement.previousElementSibling) {
        count++;
    }
    return count;
}

let $tblContextMenu;

exports.postAceInit = (hook, ctx) => {
  // log('postAceInit: START', { hook, ctx });
  // ───────────────────── helpers ─────────────────────
  const $menu      = $('#table-context-menu');
  const $gridWrap  = $('#create-table-container');
  const $gridCells = $('#new-table-size-selector td');
  const $sizeText  = $('#new-table-size');
  const $toolbarBtn = $('#table-menu-button');
  // log('postAceInit: Found UI elements.');

  function position(el, target, dx = 0, dy = 0) {
    // log('position: Calculating position for', el, 'relative to', target);
    const p = target.offset();
    el.css({ left: p.left + dx, top: p.top + dy });
    // log('position: Set position:', { left: p.left + dx, top: p.top + dy });
  }

  // ───────────────────── init once DOM is ready ─────────────────────
  // Using setTimeout to ensure DOM elements are ready
  // log('postAceInit: Setting up UI handlers via setTimeout.');
  setTimeout(() => {
    // log('postAceInit: setTimeout callback - START');
    try {
    // move pop‑ups to <body> so they are not clipped by #editbar overflow
    $menu.add($gridWrap).appendTo('body').hide();
      // log('postAceInit: setTimeout - Moved popups to body.');

    // grid hover: live update highlight & label
    $gridCells.hover(function () {
      const cell = $(this);
      const r = cell.parent().index();
      const c = cell.index();
      $gridCells.removeClass('selected');
      for (let i = 0; i <= r; i++) {
        for (let j = 0; j <= c; j++) {
          $('#new-table-size-selector tr').eq(i).find('td').eq(j).addClass('selected');
        }
      }
        const size = `${c + 1} X ${r + 1}`;
        $sizeText.text(size);
    });

    // main toolbar button toggles context menu
    $toolbarBtn.on('click', (e) => {
        // log('Toolbar Button Click: START');
      e.preventDefault();
      e.stopPropagation();
      position($menu, $toolbarBtn, 0, $toolbarBtn.outerHeight());
      $menu.toggle();
      $gridWrap.hide();
        // log('Toolbar Button Click: END - Toggled menu visibility:', $menu.is(':visible'));
    });

    // "Insert table" hover reveals grid chooser
    $('#tbl_prop_create_table').hover(
      function () {
        position($gridWrap, $(this), $(this).outerWidth(), -12);
        $gridWrap.show();
      },
        () => {
          setTimeout(() => {
            if (!$gridWrap.is(':hover')) {
              $gridWrap.hide();
            }
          }, 100);
        }
    );

    // keep grid visible while hovered
      $gridWrap.hover(
        () => { $gridWrap.show(); },
        () => { $gridWrap.hide(); }
      );

    // selecting a size calls Ace helper then hides menus
    $gridCells.on('click', () => {
        // log('Grid Cell Click: START');
      const [cols, rows] = $sizeText.text().split(' X ').map(n => parseInt(n, 10));
        // log('Grid Cell Click: Parsed size:', { cols, rows });
      ctx.ace.callWithAce((ace) => {
          // log('Grid Cell Click: Calling ace.ace_createTableViaAttributes...');
        ace.ace_createTableViaAttributes(rows, cols);
          // log('Grid Cell Click: ace.ace_createTableViaAttributes call finished.');
      }, 'tblCreate', true);
      $menu.hide();
      $gridWrap.hide();
        // log('Grid Cell Click: END - Hid menus.');
    });

    // other menu actions (row / col insert & delete, etc.)
    $menu.find('.menu-item[data-action]')
      .not('#tbl_prop_create_table')
      .on('click', function (e) {
          const action = $(this).data('action');
          // log('Menu Item Click: START', { action });
        e.preventDefault();
        e.stopPropagation();
        ctx.ace.callWithAce((ace) => {
            // log('Menu Item Click: Calling ace.ace_doDatatableOptions...', { action });
          ace.ace_doDatatableOptions(action);
            // log('Menu Item Click: ace.ace_doDatatableOptions call finished.');
        }, 'tblOptions', true);
        $menu.hide();
          // log('Menu Item Click: END - Hid menu.');
      });

    // manual close button
    $('#tbl_prop_menu_hide').on('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      $menu.hide();
    });

    // global click closes pop‑ups when clicking outside
    $(document).on('click', (e) => {
      if (!$(e.target).closest('#table-context-menu, #table-menu-button, #create-table-container').length) {
        $menu.hide();
        $gridWrap.hide();
      }
    });

    // --- BEGIN NEW: Add mousedown listener for cell selection ---
    // Wrap in callWithAce to ensure we have the editor instance
    ctx.ace.callWithAce((ace) => {
        // log('postAceInit: Inside callWithAce for attaching mousedown listeners.');
        try {
            const $innerIframe = $('iframe[name="ace_outer"]').contents().find('iframe[name="ace_inner"]');
            if ($innerIframe.length === 0) {
                console.error('ep_tables5 postAceInit: ERROR - Could not find inner iframe (ace_inner) for cell selection.');
                return;
            }
            const $inner = $($innerIframe.contents().find('body'));
            const innerDoc = $innerIframe.contents(); // Inner document for event delegation

            if (!$inner || $inner.length === 0) {
                console.error('ep_tables5 postAceInit: ERROR - Could not get body from inner iframe for cell selection.');
                return;
            }

            // Check if ace.editor exists now we are inside callWithAce
            if (!ace.editor) {
                console.error('ep_tables5 postAceInit: ERROR - ace.editor is STILL undefined inside callWithAce. Cannot attach state.');
                return;
            }

            // Initialize the shared state variable on the editor instance
            if (typeof ace.editor.ep_tables5_last_clicked === 'undefined') {
                ace.editor.ep_tables5_last_clicked = null;
                // log('postAceInit: Initialized ace.editor.ep_tables5_last_clicked');
            }

            // log('postAceInit: Attempting to attach mousedown listener to $inner for cell selection...');

            // Mousedown on table TD elements
            $inner.on('mousedown', 'table.dataTable td', function(evt) {
                // log('[ep_tables5 mousedown] RAW MOUSE DOWN detected inside table.dataTable td.');
                
                // Check if the click is on an image or image-related element
                const target = evt.target;
                const $target = $(target);
                const isImageElement = $target.closest('.inline-image, .image-placeholder, .image-inner, .image-resize-handle').length > 0;
                
                if (isImageElement) {
                    // log('[ep_tables5 mousedown] Click detected on image element within table cell. Completely skipping table processing to avoid interference.');
                    // Completely skip all table processing when image is clicked
                    return;
                }
                
                // log('[ep_tables5 mousedown] Click detected on table cell (not image). Processing normally.');
                
                if (evt.button !== 0) return; // Only left clicks

                const tdElement = $(this); 
                const trElement = tdElement.closest('tr');
                const tableElement = trElement.closest('table.dataTable');
                const lineDiv = tdElement.closest('div.ace-line'); 

                if (tdElement.length && trElement.length && tableElement.length && lineDiv.length) {
                    const cellIndex = tdElement.index(); 
                    const tblId = tableElement.attr('data-tblId');
                    const lineNum = _getLineNumberOfElement(lineDiv[0]);

                    if (tblId !== undefined && cellIndex !== -1 && lineNum !== -1) {
                        // log(`[ep_tables5 mousedown] Clicked cell (SUCCESS): Line=${lineNum}, TblId=${tblId}, CellIndex=${cellIndex}`);
                        // Store info on the shared ace editor object
                        ace.editor.ep_tables5_last_clicked = { lineNum, cellIndex, tblId };

                        // Visual feedback (Keep this for now)
                        // --- TEST: Comment out class manipulation ---
                        // tableElement.find('td.selected-table-cell').removeClass('selected-table-cell');
                        // tdElement.addClass('selected-table-cell');
                        // log('[ep_tables5 mousedown] TEST: Skipped adding/removing selected-table-cell class');
                    } else {
                        console.warn('[ep_tables5 mousedown] Could not reliably get cell info (FAIL).', {tblId, cellIndex, lineNum});
                        ace.editor.ep_tables5_last_clicked = null; // Clear shared state
                        // --- TEST: Comment out class manipulation (for consistency on failure) ---
                        // $inner.find('td.selected-table-cell').removeClass('selected-table-cell');
                    }
                } else {
                     // log('[ep_tables5 mousedown] Click was not within a valid TD/TR/TABLE/LINEDIV structure (FAIL).');
                }
            });

            // Add listener to clear selection when clicking outside tables
            $(innerDoc).on('mousedown', function(evt) {
                if (!$(evt.target).closest('table.dataTable').length) {
                    if (ace.editor.ep_tables5_last_clicked) { 
                        // log('[ep_tables5 mousedown] Clicked outside table, clearing cell info.');
                        ace.editor.ep_tables5_last_clicked = null;
                        // --- TEST: Comment out class manipulation ---
                        // $inner.find('td.selected-table-cell').removeClass('selected-table-cell');
                    }
                }
            });

            // log('postAceInit: Mousedown listeners for cell selection attached successfully (inside callWithAce).');

        } catch(e) {
            console.error('[ep_tables5 postAceInit] Error attaching mousedown listener (inside callWithAce):', e);
        }
    }, 'tables5_cell_selection_setup', true); // Unique name for callWithAce task
    // --- END NEW: Add mousedown listener for cell selection ---

    } catch (error) {
      // log('postAceInit: setTimeout callback - ERROR:', error);
      console.error('[ep_tables5:initialisation] Error in setTimeout callback:', error);
    }
    // log('postAceInit: setTimeout callback - END');
  }, 400); // delay so #editbar is in DOM
  // log('postAceInit: END');
};
