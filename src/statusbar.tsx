import * as React from 'react';

import { VDomModel, VDomRenderer } from '@jupyterlab/apputils';
import {
  ITranslator,
  nullTranslator,
  TranslationBundle
} from '@jupyterlab/translation';
import { TextItem } from '@jupyterlab/statusbar';
import { ISelection, SpreadsheetWidget } from './widget';

/**
 * StatusBar item to display selection span.
 */
export class SelectionStatus extends VDomRenderer<SelectionStatus.Model> {
  constructor(translator?: ITranslator) {
    super(new SelectionStatus.Model());
    this.translator = translator || nullTranslator;
    this._trans = this.translator.load('spreadsheet-editor');
  }

  render() {
    if (!this.model) {
      return null;
    }
    const selection = this.model.selection;
    if (!selection) {
      return <TextItem source={''} />;
    }
    // if only one cell (or zero cells) is selected, do not show anything
    if (selection.rows <= 1 && selection.columns <= 1) {
      return <TextItem source={''} />;
    }
    this.node.title =
      this._trans._n('Selected %1 row', 'Selected %1 rows', selection.rows) +
      this._trans._n(' and %1 column', ' and %1 columns', selection.columns);

    const text =
      this._trans._n('%1 row', '%1 rows', selection.rows) +
      this._trans._n(', %1 column', ', %1 columns', selection.columns);

    return <TextItem source={text} />;
  }

  protected translator: ITranslator;
  private _trans: TranslationBundle;
}

export namespace SelectionStatus {
  export class Model extends VDomModel {
    private _spreadsheetWidget: SpreadsheetWidget | null;

    get selection(): ISelection {
      return this.spreadsheetWidget?.selection;
    }

    set spreadsheetWidget(widget: SpreadsheetWidget) {
      if (this._spreadsheetWidget) {
        this._spreadsheetWidget.selectionChanged.disconnect(
          this._triggerChange,
          this
        );
      }
      this._spreadsheetWidget = widget;
      this._spreadsheetWidget.selectionChanged.connect(
        this._triggerChange,
        this
      );
      this._triggerChange();
    }

    get spreadsheetWidget() {
      return this._spreadsheetWidget;
    }

    private _triggerChange(): void {
      this.stateChanged.emit(void 0);
    }
  }
}
