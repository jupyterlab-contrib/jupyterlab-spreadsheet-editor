// Copyright (c) Jupyter Development Team.
// Distributed under the terms of the Modified BSD License.

import { ISearchMatch, ISearchProvider } from "@jupyterlab/documentsearch";
import { SpreadsheetEditorDocumentWidget } from "./documentwidget";
import { Widget } from "@lumino/widgets";
import { DocumentWidget } from "@jupyterlab/docregistry";
import { SpreadsheetWidget } from "./widget";
import { ISignal, Signal } from "@lumino/signaling";
import { JExcelElement } from "jexcel";

interface ICellCoordinates {
  column: number;
  row: number;
}

class CoordinatesSet {

  protected _set: Set<string>;

  constructor() {
    this._set = new Set<string>();
  }

  protected _valueToString(value: ICellCoordinates): string {
    if (typeof value == 'undefined'){
      return value
    }
    return value.column + '|' + value.row;
  }

  protected _stringToValue(value: string): ICellCoordinates {
    if (typeof value == 'undefined'){
      return value
    }
    let parts = value.split('|');
    if (parts.length != 2) {
      console.warn('A problem with stringToValue input detected!')
    }
    return {
      column: parseInt(parts[0], 10),
      row: parseInt(parts[1], 10)
    }
  }

  add(value: ICellCoordinates): this {
    this._set.add(this._valueToString(value));
    return this;
  }

  delete(value: ICellCoordinates): boolean {
    return this._set.delete(this._valueToString(value));
  }

  has(value: ICellCoordinates): boolean {
    return this._set.has(this._valueToString(value));
  }

  values(): IterableIterator<ICellCoordinates> {
    let iterator = this._set.values();
    let stringToValue = this._stringToValue
    return {
      [Symbol.iterator](): IterableIterator<ICellCoordinates> {
        return this;
      },
      return(value: any): IteratorResult<ICellCoordinates> {
        let returned = iterator.return(value)
        return {
          done: returned.done,
          value: stringToValue(returned.value)
        };
      },
      next() {
        let next = iterator.next()
        return {
          done: next.done,
          value: stringToValue(next.value)
        }
      },
      throw() {
        let thrown = iterator.throw()
        return {
          done: thrown.done,
          value: stringToValue(thrown.value)
        }
      }
    }
  }

  clear() {
    return this._set.clear();
  }
}

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
  }

  readonly isReadOnly: boolean;

  get matches(): ISearchMatch[] {
    return this._matches;
  }

  endQuery(): Promise<void> {
    this.backlightOff();
    this._currentMatchIndex = null;
    this._matches = [];
    this._sheet.changed.disconnect(this._onSheetChanged, this)
    return Promise.resolve(undefined);
  }

  private backlightOff() {
    for (let matchCoords of this.backlitMatches.values()) {
      let cell: HTMLElement = this._target.getCellFromCoords(matchCoords.column, matchCoords.row)
      cell.classList.remove('se-backlight');
    }
    this.backlitMatches.clear();
  }

  async endSearch(): Promise<void> {
    //return Promise.resolve(undefined);
    return this.endQuery()
  }

  private getSelectedCellCoordinates(): ICellCoordinates {
    let target = this._target
    let columns = target.getSelectedColumns()
    let rows = target.getSelectedRows(true)
    if (rows.length == 1 && columns.length == 1) {
      return {
        column: columns[0],
        row: rows[0]
      }
    }
  }

  private _initialQueryCoodrs: ICellCoordinates;

  getInitialQuery(searchTarget: SpreadsheetEditorDocumentWidget): any {
    this._target = searchTarget.content.jexcel;
    let coords = this.getSelectedCellCoordinates();
    this._initialQueryCoodrs = coords;
    if (coords) {
      let value = this._target.getValueFromCoords(
        coords.column, coords.row, false
      );
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
    this.backlightMatches();
    this._target.updateSelectionFromCoords(match.column, match.line, match.column, match.line, null);
    let cell = this._target.getCellFromCoords(match.column, match.line)
    cell.scrollIntoView(false);
  }

  async replaceAllMatches(newText: string): Promise<boolean> {
    for (let i = 0; i < this.matches.length; i++) {
      this._currentMatchIndex = i
      await this.replaceCurrentMatch(newText, true);
    }
    this._matches = this.findMatches();
    this.backlightMatches();
    return true;
  }

  async replaceCurrentMatch(newText: string, isReplaceAll=false): Promise<boolean> {
    let replaceOccurred = false;
    let match = this.matches[this.currentMatchIndex]
    let cell = this._target.getValueFromCoords(match.column, match.line, false);
    let index = -1;
    let matchesInCell = 0;

    let newValue = String(cell).replace(this._query, (substring) => {
      index += 1;
      matchesInCell += 1;
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

    if (!isReplaceAll && matchesInCell == 1) {
      let matchCoords = {column: match.column, row: match.line};
      let cell: HTMLElement = this._target.getCellFromCoords(match.column, match.line)
      cell.classList.remove('se-backlight');
      this.backlitMatches.delete(matchCoords)
    }

    if(!isReplaceAll) {
      await this.highlightNext();
    }
    return replaceOccurred;
  }

  private _onSheetChanged() {
    this._matches = this.findMatches();
    this.backlightMatches();
    this._changed.emit(undefined);
  }

  protected backlitMatches: CoordinatesSet;

  /**
   * Highlight n=1000 matches around the current match.
   * The number of highlights is limited to prevent negative impact on the UX in huge notebooks.
   */
  protected backlightMatches(n=1000): void {
    for (
      let i = Math.max(0, this._currentMatchIndex - n / 2);
      i < Math.min(this._currentMatchIndex + n / 2, this.matches.length);
      i++
    ) {
      let match = this.matches[i];
      let matchCoord = {
        column: match.column,
        row: match.line
      }

      if (!this.backlitMatches.has(matchCoord)) {
        let cell: HTMLElement = this._target.getCellFromCoords(match.column, match.line)
        cell.classList.add('se-backlight');
        this.backlitMatches.add(matchCoord)
      }
    }
  }

  protected findMatches(): ISearchMatch[] {
    let currentCellCoordinates = this._initialQueryCoodrs;
    this._initialQueryCoodrs = null;
    let currentMatchIndex = 0;

    let matches: ISearchMatch[] = [];
    let data = this._target.getData();
    let rowNumber = 0;
    let columnNumber = -1;
    let index = 0;
    let totalMatchIndex = 0;
    for (let row of data) {
      for (let cell of row) {
        columnNumber += 1;
        if (!cell) {
          continue;
        }
        let matched = String(cell).match(this._query)
        if (!matched) {
          continue;
        }
        index = 0;
        if (
          currentCellCoordinates != null
          && currentCellCoordinates.row == rowNumber
          && currentCellCoordinates.column == columnNumber
        ) {
          currentMatchIndex = totalMatchIndex;
        }
        for (let match of matched) {
          matches.push({
            line: rowNumber,
            column: columnNumber,
            index: index,
            fragment: match,
            text: match
          })
          index += 1;
          totalMatchIndex += 1;
        }
      }
      columnNumber = -1;
      rowNumber += 1;
    }
    this._currentMatchIndex = currentMatchIndex;
    this._matches = matches;

    if (matches.length) {
      this.highlight(matches[this._currentMatchIndex]);
    }

    return matches
  }

  constructor() {
    this.backlitMatches = new CoordinatesSet();
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

    this._sheet.changed.connect(this._onSheetChanged, this)

    return this.findMatches();
  }

  private _changed = new Signal<this, void>(this);

  private _target: JExcelElement;
  private _sheet: SpreadsheetWidget;
  private _query: RegExp;
  private _matches: ISearchMatch[];
  private _currentMatchIndex: number;
}
