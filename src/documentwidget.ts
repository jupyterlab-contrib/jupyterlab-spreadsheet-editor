// Copyright (c) Jupyter Development Team.
// Distributed under the terms of the Modified BSD License.

import { DocumentWidget } from '@jupyterlab/docregistry';
import { SpreadsheetWidget } from './widget';
import { ToolbarButton } from '@jupyterlab/apputils';
import { LabIcon } from '@jupyterlab/ui-components';
import addRowSvg from '../style/icons/mdi-table-row-plus-after.svg';
import removeRowSvg from '../style/icons/mdi-table-row-remove.svg';
import addColumnSvg from '../style/icons/mdi-table-column-plus-after.svg';
import removeColumnSvg from '../style/icons/mdi-table-column-remove.svg';
import numberedSvg from '../style/icons/mdi-format-list-numbered.svg';
import fitColumnWidthSvg from '../style/icons/mdi-table-column-width.svg';
import freezeColumnSvg from '../style/icons/mdi-snowflake.svg';
import unfreezeColumnSvg from '../style/icons/mdi-snowflake-off.svg';
import listTypeSvg from '../style/icons/mdi-format-list-bulleted-type.svg';
import topHeaderSvg from '../style/icons/mdi-page-layout-header.svg';
import { nullTranslator, TranslationBundle } from '@jupyterlab/translation';

export class SpreadsheetEditorDocumentWidget extends DocumentWidget<
  SpreadsheetWidget
> {
  protected _trans: TranslationBundle;

  constructor(options: DocumentWidget.IOptions<SpreadsheetWidget>) {
    super(options);
    const translator = options.translator || nullTranslator;
    this._trans = translator.load('spreadsheet-editor');

    const addRowButton = new ToolbarButton({
      icon: new LabIcon({ name: 'spreadsheet:add-row', svgstr: addRowSvg }),
      onClick: () => {
        this.content.jexcel.insertRow();
        this.content.updateModel();
      },
      tooltip: this._trans.__('Insert a row at the end')
    });
    this.toolbar.addItem('spreadsheet:insert-row', addRowButton);

    const removeRowButton = new ToolbarButton({
      icon: new LabIcon({
        name: 'spreadsheet:remove-row',
        svgstr: removeRowSvg
      }),
      onClick: () => {
        this.content.jexcel.deleteRow(this.content.jexcel.rows.length - 1, 1);
        this.content.updateModel();
      },
      tooltip: this._trans.__('Remove the last row')
    });
    this.toolbar.addItem('spreadsheet:remove-row', removeRowButton);

    const addColumnButton = new ToolbarButton({
      icon: new LabIcon({
        name: 'spreadsheet:add-column',
        svgstr: addColumnSvg
      }),
      onClick: () => {
        this.content.jexcel.insertColumn(null, null);
        this.content.updateModel();
      },
      tooltip: this._trans.__('Insert a column at the end')
    });
    this.toolbar.addItem('spreadsheet:insert-column', addColumnButton);

    const removeColumnButton = new ToolbarButton({
      icon: new LabIcon({
        name: 'spreadsheet:remove-column',
        svgstr: removeColumnSvg
      }),
      onClick: () => {
        this.content.jexcel.deleteColumn(this.content.columnsNumber, null);
        this.content.updateModel();
      },
      tooltip: this._trans.__('Remove the last column')
    });
    this.toolbar.addItem('spreadsheet:remove-column', removeColumnButton);

    let indexVisible = true;
    const showHideIndex = new ToolbarButton({
      icon: new LabIcon({
        name: 'spreadsheet:show-hide-index',
        svgstr: numberedSvg
      }),
      onClick: () => {
        if (indexVisible) {
          this.content.jexcel.hideIndex();
        } else {
          this.content.jexcel.showIndex();
        }
        indexVisible = !indexVisible;
      },
      tooltip: this._trans.__('Show/hide row numbers')
    });
    this.toolbar.addItem('spreadsheet:show-hide-index', showHideIndex);

    const fitColumnWidthButton = new ToolbarButton({
      icon: new LabIcon({
        name: 'spreadsheet:fit-columns',
        svgstr: fitColumnWidthSvg
      }),
      onClick: () => {
        switch (this.content.fitMode) {
          case 'fit-cells':
            this.content.fitMode = 'all-equal-fit';
            break;
          case 'all-equal-fit':
            this.content.fitMode = 'all-equal-default';
            break;
          case 'all-equal-default':
            this.content.fitMode = 'fit-cells';
        }
        this.content.relayout();
      },
      tooltip: this._trans.__('Fit columns width')
    });
    this.toolbar.addItem('spreadsheet:fit-columns', fitColumnWidthButton);

    const freezeColumnButton = new ToolbarButton({
      icon: new LabIcon({
        name: 'spreadsheet:freeze-columns',
        svgstr: freezeColumnSvg
      }),
      onClick: () => {
        this.content.freezeSelectedColumns();
      },
      tooltip: this._trans.__(
        'Freeze the initial columns (up to the selected column)'
      )
    });
    this.toolbar.addItem('spreadsheet:freeze-columns', freezeColumnButton);

    const unfreezeColumnButton = new ToolbarButton({
      icon: new LabIcon({
        name: 'spreadsheet:unfreeze-columns',
        svgstr: unfreezeColumnSvg
      }),
      onClick: () => {
        this.content.unfreezeColumns();
      },
      tooltip: this._trans.__('Unfreeze frozen column')
    });
    this.toolbar.addItem('spreadsheet:unfreeze-columns', unfreezeColumnButton);

    const switchHeadersButton = new ToolbarButton({
      icon: new LabIcon({
        name: 'spreadsheet:switch-headers',
        svgstr: topHeaderSvg
      }),
      onClick: () => {
        this.content.switchHeaders();
      },
      tooltip: this._trans.__('First row as a header')
    });
    this.toolbar.addItem('spreadsheet:switch-headers', switchHeadersButton);

    const columnTypeButton = new ToolbarButton({
      icon: new LabIcon({
        name: 'spreadsheet:column-types',
        svgstr: listTypeSvg
      }),
      onClick: () => {
        this.content.switchTypesBar();
      },
      tooltip: this._trans.__('Show/hide column types bar')
    });
    this.toolbar.addItem(
      'spreadsheet:switch-column-types-bar',
      columnTypeButton
    );
  }
}
