// Copyright (c) Jupyter Development Team.
// Distributed under the terms of the Modified BSD License.

import {
  ISearchMatch,
  ISearchProvider,
  ISearchProviderFactory,
  SearchProvider,
  IFilters
} from '@jupyterlab/documentsearch';
import { SpreadsheetEditorDocumentWidget } from './documentwidget';
import { Widget } from '@lumino/widgets';
import { DocumentWidget } from '@jupyterlab/docregistry';
import { SpreadsheetWidget } from './widget';
import { ISignal, Signal } from '@lumino/signaling';
import { JspreadsheetInstance } from 'jspreadsheet-ce';

export interface ICellCoordinates {
  column: number;
  row: number;
}

interface ICellSearchMatch extends ISearchMatch, ICellCoordinates {
  // no-op
}

class CoordinatesSet {
  protected _set: Set<string>;

  constructor() {
    this._set = new Set<string>();
  }

  protected _valueToString(value: ICellCoordinates): string {
    if (typeof value === 'undefined') {
      return value;
    }
    return value.column + '|' + value.row;
  }

  protected _stringToValue(value: string): ICellCoordinates {
    if (typeof value === 'undefined') {
      return value;
    }
    const parts = value.split('|');
    if (parts.length !== 2) {
      console.warn('A problem with stringToValue input detected!');
    }
    return {
      column: parseInt(parts[0], 10),
      row: parseInt(parts[1], 10)
    };
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
    const iterator = this._set.values();
    const stringToValue = this._stringToValue;
    return {
      [Symbol.iterator](): IterableIterator<ICellCoordinates> {
        return this;
      },
      return(value: any): IteratorResult<ICellCoordinates> {
        const returned = iterator.return!(value);
        return {
          done: returned.done,
          value: stringToValue(returned.value)
        };
      },
      next() {
        const next = iterator.next();
        return {
          done: next.done,
          value: stringToValue(next.value)
        };
      },
      throw() {
        const thrown = iterator.throw!();
        return {
          done: thrown.done,
          value: stringToValue(thrown.value)
        };
      }
    };
  }

  clear() {
    return this._set.clear();
  }
}

export class SpreadsheetSearchProviderFactory
  implements ISearchProviderFactory<Widget>
{
  createNew(
    widget: Widget
    //translator?: ITranslator
  ): ISearchProvider {
    return new SpreadsheetSearchProvider(
      widget as SpreadsheetEditorDocumentWidget
    );
  }

  /**
   * Report whether or not this provider has the ability to search on the given object
   */
  isApplicable(domain: Widget): domain is SpreadsheetEditorDocumentWidget {
    // check to see if the SpreadsheetSearchProvider can search on the
    // first cell, false indicates another editor is present
    return (
      domain instanceof DocumentWidget &&
      domain.content instanceof SpreadsheetWidget
    );
  }
}

export class SpreadsheetSearchProvider extends SearchProvider<SpreadsheetEditorDocumentWidget> {
  constructor(widget: SpreadsheetEditorDocumentWidget) {
    super(widget);
    this._sheet = widget.content;
    this._target = this._sheet.jexcel!;
    this.backlitMatches = new CoordinatesSet();
  }

  private mostRecentSelectedCell: any;

  get changed(): ISignal<this, void> {
    return this._changed;
  }

  get currentMatchIndex(): number | null {
    return this._currentMatchIndex;
  }

  readonly isReadOnly: boolean = false;

  get matches(): ICellSearchMatch[] {
    return this._matches;
  }

  get matchesCount(): number | null {
    return this._matches.length;
  }

  endQuery(): Promise<void> {
    this.backlightOff();
    this._currentMatchIndex = 0;
    this._matches = [];
    this._sheet.changed.disconnect(this._onSheetChanged, this);
    return Promise.resolve(undefined);
  }

  /**
   * Clear currently highlighted match.
   */
  async clearHighlight(): Promise<void> {
    this.backlightOff();
  }

  private backlightOff() {
    for (const matchCoords of this.backlitMatches.values()) {
      const cell: HTMLElement = this._target.getCellFromCoords(
        matchCoords.column,
        matchCoords.row
      );
      cell.classList.remove('se-backlight');
    }
    this.backlitMatches.clear();
  }

  async endSearch(): Promise<void> {
    // restore the selection
    // eslint-disable-next-line eqeqeq
    if (this._target.selectedCell == null && this.mostRecentSelectedCell) {
      this._target.selectedCell = this.mostRecentSelectedCell;
    }
    return this.endQuery();
  }

  private getSelectedCellCoordinates(): ICellCoordinates | null {
    const target = this._target;
    const columns = target.getSelectedColumns();
    const rows = target.getSelectedRows(true) as number[];
    if (rows.length === 1 && columns.length === 1) {
      return {
        column: columns[0],
        row: rows[0]
      };
    }
    return null;
  }

  private _initialQueryCoodrs: ICellCoordinates | null = null;

  getInitialQuery(): string {
    const coords = this.getSelectedCellCoordinates();
    this._initialQueryCoodrs = coords;
    if (coords) {
      const value = this._target.getValueFromCoords(
        coords.column,
        coords.row,
        false
      );
      if (value) {
        return value.toString();
      }
    }
    // Close the editor to avoid overwriting contents of last edited cell
    // as users starts typing into the search box after pressing ctrl + f
    // (but only do that after the initial value was taken)
    this._target.resetSelection(true);
    return '';
  }

  async highlightNext(): Promise<ICellSearchMatch | undefined> {
    if (this._currentMatchIndex + 1 < this.matches.length) {
      this._currentMatchIndex += 1;
    } else {
      this._currentMatchIndex = 0;
    }
    const match = this.matches[this._currentMatchIndex];
    if (!match) {
      return;
    }
    this.highlight(match);
    return match;
  }

  async highlightPrevious(): Promise<ICellSearchMatch | undefined> {
    if (this._currentMatchIndex > 0) {
      this._currentMatchIndex -= 1;
    } else {
      this._currentMatchIndex = this.matches.length - 1;
    }
    const match = this.matches[this._currentMatchIndex];
    if (!match) {
      return;
    }
    this.highlight(match);
    return match;
  }

  highlight(match: ICellSearchMatch) {
    this.backlightMatches();
    // select the matched cell, which leads to a loss of "focus" (or rather jexcel eagerly intercepting events)
    this._target.updateSelectionFromCoords(
      match.column,
      match.row,
      match.column,
      match.row,
      null
    );
    // "regain" focus by erasing selection information (but keeping all the CSS) - this is a workaround (best avoided)
    this.mostRecentSelectedCell = this._target.selectedCell;
    this._target.selectedCell = null;
    this._sheet.scrollCellIntoView({ row: match.row, column: match.column });
  }

  async replaceAllMatches(newText: string): Promise<boolean> {
    for (let i = 0; i < this.matches.length; i++) {
      this._currentMatchIndex = i;
      await this.replaceCurrentMatch(newText, true);
    }
    this._matches = this.findMatches();
    this.backlightMatches();
    return true;
  }

  async replaceCurrentMatch(
    newText: string,
    isReplaceAll = false
  ): Promise<boolean> {
    let replaceOccurred = false;
    const match = this.matches[this._currentMatchIndex];
    const cell = this._target.getValueFromCoords(
      match.column,
      match.row,
      false
    );
    let index = -1;
    let matchesInCell = 0;

    const newValue = String(cell).replace(this._query!, substring => {
      index += 1;
      matchesInCell += 1;
      if (index === match.position) {
        replaceOccurred = true;
        return newText;
      }

      return substring;
    });
    let subsequentIndex = this._currentMatchIndex + 1;
    while (subsequentIndex < this.matches.length) {
      const subsequent = this.matches[subsequentIndex];
      if (subsequent.column === match.column && subsequent.row === match.row) {
        subsequent.position -= 1;
      } else {
        break;
      }
      subsequentIndex += 1;
    }

    this._target.setValueFromCoords(match.column, match.row, newValue, false);

    if (!isReplaceAll && matchesInCell === 1) {
      const matchCoords = { column: match.column, row: match.row };
      const cell: HTMLElement = this._target.getCellFromCoords(
        match.column,
        match.row
      );
      cell.classList.remove('se-backlight');
      this.backlitMatches.delete(matchCoords);
    }

    if (!isReplaceAll) {
      await this.highlightNext();
    }
    return replaceOccurred;
  }

  private _onSheetChanged() {
    // matches may need updating
    this._matches = this.findMatches(false);
    // update backlight
    this.backlightOff();
    this.backlightMatches();
    this._changed.emit(undefined);
  }

  protected backlitMatches: CoordinatesSet;

  /**
   * Highlight n=1000 matches around the current match.
   * The number of highlights is limited to prevent negative impact on the UX in huge notebooks.
   */
  protected backlightMatches(n = 1000): void {
    for (
      let i = Math.max(0, this._currentMatchIndex - n / 2);
      i < Math.min(this._currentMatchIndex + n / 2, this.matches.length);
      i++
    ) {
      const match = this.matches[i];
      const matchCoord = {
        column: match.column,
        row: match.row
      };

      if (!this.backlitMatches.has(matchCoord)) {
        const cell: HTMLElement = this._target.getCellFromCoords(
          match.column,
          match.row
        );
        cell.classList.add('se-backlight');
        this.backlitMatches.add(matchCoord);
      }
    }
  }

  protected findMatches(highlightFirst = true): ICellSearchMatch[] {
    const currentCellCoordinates = this._initialQueryCoodrs;
    this._initialQueryCoodrs = null;
    let currentMatchIndex = 0;
    const query = this._query;
    if (!query) {
      return [];
    }

    const matches: ICellSearchMatch[] = [];
    const data = this._target.getData();
    let rowNumber = 0;
    let columnNumber = -1;
    let index = 0;
    let totalMatchIndex = 0;
    for (const row of data) {
      for (const cell of row) {
        columnNumber += 1;
        if (!cell) {
          continue;
        }
        const matched = String(cell).match(query!);
        if (!matched) {
          continue;
        }
        index = 0;
        if (
          currentCellCoordinates !== null &&
          currentCellCoordinates.row === rowNumber &&
          currentCellCoordinates.column === columnNumber
        ) {
          currentMatchIndex = totalMatchIndex;
        }
        for (const match of matched) {
          matches.push({
            row: rowNumber,
            column: columnNumber,
            position: index,
            text: match
          });
          index += 1;
          totalMatchIndex += 1;
        }
      }
      columnNumber = -1;
      rowNumber += 1;
    }
    this._currentMatchIndex = currentMatchIndex;
    this._matches = matches;

    if (matches.length && highlightFirst) {
      this.highlight(matches[this._currentMatchIndex]);
    }

    return matches;
  }
  async startQuery(query: RegExp, filters: IFilters): Promise<void> {
    const searchTarget = this.widget;
    this._sheet = searchTarget.content;
    this._query = query;

    this._sheet.changed.connect(this._onSheetChanged, this);

    this.findMatches();
  }

  private _changed = new Signal<this, void>(this);

  private _target: JspreadsheetInstance;
  private _sheet: SpreadsheetWidget;
  private _query: RegExp | null = null;
  private _matches: ICellSearchMatch[] = [];
  private _currentMatchIndex: number = 0;
}
