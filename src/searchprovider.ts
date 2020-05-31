// Copyright (c) Jupyter Development Team.
// Distributed under the terms of the Modified BSD License.

import { ISearchMatch, ISearchProvider } from "@jupyterlab/documentsearch";
import { SpreadsheetEditorDocumentWidget } from "./documentwidget";
import { Widget } from "@lumino/widgets";
import { DocumentWidget } from "@jupyterlab/docregistry";
import { SpreadsheetWidget } from "./widget";
import { ISignal, Signal } from "@lumino/signaling";
import { JExcelElement } from "jexcel";

export class SpreadsheetSearchProvider implements ISearchProvider<SpreadsheetEditorDocumentWidget> {
  /**
   * Report whether or not this provider has the ability to search on the given object
   */
  static canSearchOn(domain: Widget): domain is SpreadsheetEditorDocumentWidget {
    // check to see if the SpreadsheetSearchProvider can search on the
    // first cell, false indicates another editor is present
    return (
      domain instanceof DocumentWidget && domain.content instanceof SpreadsheetWidget
    );
  }

  get changed(): ISignal<this, void> {
    return this._changed;
  }

  get currentMatchIndex() {
    return this._currentMatchIndex;
  };

  readonly isReadOnly: boolean;

  get matches(): ISearchMatch[] {
    return this._matches;
  };

  endQuery(): Promise<void> {
    this._currentMatchIndex = null;
    return Promise.resolve(undefined);
  }

  async endSearch(): Promise<void> {
    //return Promise.resolve(undefined);
  }

  getInitialQuery(searchTarget: SpreadsheetEditorDocumentWidget): any {
    let target = searchTarget.content.jexcel
    let columns = target.getSelectedColumns()
    let rows = target.getSelectedRows(true)
    if (rows.length == 1 && columns.length == 1) {
      let value = target.getValueFromCoords(columns[0], rows[0], false);
      if (value) {
        return value
      }
    }
    return null;
  }

  async highlightNext(): Promise<ISearchMatch | undefined> {
    if (this._currentMatchIndex + 1 < this.matches.length) {
      this._currentMatchIndex += 1;
    } else {
      this._currentMatchIndex = 0;
    }
    let match = this.matches[this.currentMatchIndex]
    if (!match) {
      return;
    }
    this.highlight(match);
    return match;
  }

  async highlightPrevious(): Promise<ISearchMatch | undefined> {
    if (this._currentMatchIndex > 0) {
      this._currentMatchIndex -= 1;
    } else {
      this._currentMatchIndex = this.matches.length - 1;
    }
    let match = this.matches[this.currentMatchIndex]
    if (!match) {
      return;
    }
    this.highlight(match);
    return match;
  }

  highlight(match: ISearchMatch) {
    this._target.updateSelectionFromCoords(match.column, match.line, match.column, match.line, null);
  }

  async replaceAllMatches(newText: string): Promise<boolean> {
    for (let i = 0; i < this.matches.length; i++) {
      this._currentMatchIndex = i
      await this.replaceCurrentMatch(newText);
    }
    return true;
  }

  async replaceCurrentMatch(newText: string): Promise<boolean> {
    let replaceOccurred = false;
    let match = this.matches[this.currentMatchIndex]
    let cell = this._target.getValueFromCoords(match.column, match.line, false);
    let index = -1;
    let newValue = String(cell).replace(this._query, (substring) => {
      index += 1;
      if (index == match.index) {
        replaceOccurred = true;
        return newText;
      }

      return substring
    })
    let subsequentIndex = this.currentMatchIndex + 1;
    while (subsequentIndex < this.matches.length) {
      let subsequent = this.matches[subsequentIndex];
      if (subsequent.column == match.column && subsequent.line == match.line) {
        subsequent.index -= 1;
      } else {
        break;
      }
      subsequentIndex += 1;
    }

    this._target.setValueFromCoords(match.column, match.line, newValue, false);

    await this.highlightNext();
    return replaceOccurred;
  }

  private _onSheetChanged() {
    this.findMatches();
    this._changed.emit(undefined);
  }

  protected findMatches(): ISearchMatch[] {
    let matches: ISearchMatch[] = [];
    let data = this._target.getData();
    let row_number = 0;
    let column_number = -1;
    let index = 0;
    for (let row of data) {
      for (let cell of row) {
        column_number += 1;
        if (!cell) {
          continue;
        }
        let matched = String(cell).match(this._query)
        if (!matched) {
          continue;
        }
        index = 0;
        for (let match of matched) {
          matches.push({
            line: row_number,
            column: column_number,
            index: index,
            fragment: match,
            text: match
          })

          index += 1;
        }
      }
      column_number = -1;
      row_number += 1;
    }
    this._currentMatchIndex = 0;
    this._matches = matches;

    return matches
  }

  async startQuery(query: RegExp, searchTarget: SpreadsheetEditorDocumentWidget): Promise<ISearchMatch[]> {
    if (!SpreadsheetSearchProvider.canSearchOn(searchTarget)) {
      throw new Error('Cannot find Spreadsheet editor instance to search');
    }

    this._sheet = searchTarget.content;
    this._query = query;
    this._target = searchTarget.content.jexcel;
    this._target.resetSelection(true);
    this._target.el.blur();

    this._sheet.changed.connect(() => { this._onSheetChanged() })

    return this.findMatches();
  }

  private _changed = new Signal<this, void>(this);

  private _target: JExcelElement;
  private _sheet: SpreadsheetWidget;
  private _query: RegExp;
  private _matches: ISearchMatch[];
  private _currentMatchIndex: number;
}
