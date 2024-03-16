// Copyright (c) Jupyter Development Team.
// Distributed under the terms of the Modified BSD License.

import { Widget } from '@lumino/widgets';
import { DocumentRegistry } from '@jupyterlab/docregistry';
import { PromiseDelegate, UUID } from '@lumino/coreutils';
import { PathExt } from '@jupyterlab/coreutils';
import Papa from 'papaparse';
import { Message } from '@lumino/messaging';
import jspreadsheet from 'jspreadsheet-ce';
import { Signal } from '@lumino/signaling';
import { ICellCoordinates } from './searchprovider';

type columnTypeId =
  | 'autocomplete'
  | 'calendar'
  | 'checkbox'
  | 'color'
  | 'dropdown'
  | 'hidden'
  | 'html'
  | 'image'
  | 'numeric'
  | 'radio'
  | 'text';

export interface ISelection {
  rows: number;
  columns: number;
}

/**
 * An spreadsheet widget.
 */
export class SpreadsheetWidget extends Widget {
  public jexcel: jspreadsheet.JspreadsheetInstance | null = null;
  protected separator: string;
  protected linebreak: string = '\n';
  public fitMode: 'all-equal-default' | 'all-equal-fit' | 'fit-cells';
  public changed: Signal<this, void>;
  protected hasFrozenColumns: boolean;
  private editor: HTMLDivElement;
  private container: HTMLDivElement;
  private columnTypesBar: HTMLDivElement;
  private selectAllElement: HTMLElement | null = null;

  protected firstRowAsHeader: boolean;
  private header: Array<string> | undefined = undefined;
  private columnTypes: Array<columnTypeId> = [];
  public selectionChanged: Signal<SpreadsheetWidget, ISelection>;

  /**
   * Construct a new Spreadsheet widget.
   */
  constructor(context: DocumentRegistry.CodeContext) {
    super();
    this.id = UUID.uuid4();
    this.title.label = PathExt.basename(context.localPath);
    this.title.closable = true;
    this.context = context;
    this.separator = ''; // Papa auto detect
    this.fitMode = 'all-equal-default';
    if (context.localPath.endsWith('tsv')) {
      this.separator = '\t';
    }
    if (context.localPath.endsWith('csv')) {
      this.separator = ',';
    }
    this.hasFrozenColumns = false;
    this.firstRowAsHeader = false;

    context.ready
      .then(() => {
        this._onContextReady();
      })
      .catch(console.warn);
    this.changed = new Signal<this, void>(this);
    this._selection = {
      rows: 0,
      columns: 0
    };
    this.selectionChanged = new Signal(this);

    const container = document.createElement('div');
    container.className = 'se-area-container';
    this.node.appendChild(container);

    // TODO: move to a separate class/widget
    this.columnTypesBar = document.createElement('div');
    this.columnTypesBar.classList.add('se-column-types');
    this.columnTypesBar.classList.add('se-hidden');
    this.columnTypeSelectors = new Map();
    container.appendChild(this.columnTypesBar);

    this.editor = document.createElement('div');
    container.appendChild(this.editor);
    this.container = container;
  }

  protected parseValue(content: string): jspreadsheet.CellValue[][] {
    const parsed = Papa.parse<string[]>(content, { delimiter: this.separator });
    if (!this.separator) {
      this.separator = parsed.meta.delimiter;
    }
    this.linebreak = parsed.meta.linebreak;
    if (parsed.errors.length) {
      console.warn('Parsing errors encountered', parsed.errors);
    }
    const columnsNumber = this.extractColumnNumber(parsed.data);
    // TODO: read the actual type from a file?
    // TODO only redefine if reading for the first time?
    // TODO add/remove when column added removed?
    if (typeof this.columnTypes === 'undefined') {
      this.columnTypes = [
        ...Array(columnsNumber).map(() => 'text' as columnTypeId)
      ];
    }

    if (this.firstRowAsHeader) {
      this.header = parsed.data.shift();
    } else {
      this.header = undefined;
    }

    return parsed.data;
  }

  extractColumnNumber(data: jspreadsheet.CellValue[][]): number {
    return data.length ? data[0].length : 0;
  }

  columns(columnsNumber: number) {
    const columns: Array<jspreadsheet.Column> = [];

    for (let i = 0; i < columnsNumber; i++) {
      columns.push({
        title: this.header ? this.header[i] : undefined,
        type: this.columnTypes[i]
      });
    }
    return columns;
  }

  private onChange(): void {
    this.context.model.sharedModel.setSource(this.getValue());
    this.changed.emit();
  }

  public get selection(): ISelection {
    return this._selection;
  }

  private _selection: ISelection;

  private _onContextReady(): void {
    if (this.isDisposed) {
      return;
    }
    const contextModel = this.context.model;

    // Set the editor model value.
    const content = contextModel.toString();
    const data = this.parseValue(content);

    const options: jspreadsheet.JSpreadsheetOptions = {
      data: data,
      minDimensions: [1, 1],
      // minSpareCols: 1,
      // minSpareRows: 1,
      csvFileName: this.title.label,
      // @ts-expect-error (boolean missing in typing, but documented in repo)
      contextMenu: false,
      columnDrag: true,
      onchange: () => {
        this.onChange();
      },
      // insert
      oninsertrow: () => {
        this.onChange();
      },
      oninsertcolumn: () => {
        this.onChange();
        this.populateColumnTypesBar();
        this.onResize();
      },
      // move
      onmoverow: () => {
        this.onChange();
      },
      onmovecolumn: () => {
        this.onChange();
      },
      // delete
      ondeleterow: () => {
        this.onChange();
      },
      ondeletecolumn: () => {
        this.onChange();
        this.populateColumnTypesBar();
        this.onResize();
      },
      // resize
      onresizecolumn: () => {
        this.adjustColumnTypesWidth();
      },
      onselection: (
        el: HTMLElement,
        borderLeft: number,
        borderTop: number,
        borderRight: number,
        borderBottom: number,
        origin: any
      ) => {
        this._selection = {
          rows: borderBottom - borderTop + 1,
          columns: borderRight - borderLeft + 1
        };
        this.selectionChanged.emit(this._selection);
        // TODO: support all corners of selection
        this.scrollCellIntoView({ column: borderLeft, row: borderTop });
      },
      columns: this.columns(this.extractColumnNumber(data))
    };

    this.createEditor(options);

    // Wire signal connections.
    contextModel.contentChanged.connect(this._onContentChanged, this);

    this.populateColumnTypesBar();

    // If the sheet is not too big, use the more user-friendly columns width adjustment model
    if (data.length && data[0].length * data.length < 100 * 100) {
      this.fitMode = 'fit-cells';
      this.relayout();
    }

    // Resolve the ready promise.
    this._ready.resolve(undefined);
  }

  reloadEditor(options: jspreadsheet.JSpreadsheetOptions) {
    const config = this.jexcel!.getConfig();
    this.jexcel!.destroy();
    this.createEditor({
      ...config,
      ...options
    });
    this.relayout();
  }

  switchHeaders() {
    const value = this.getValue();
    this.firstRowAsHeader = !this.firstRowAsHeader;
    const data = this.parseValue(value);
    this.reloadEditor({
      data: data,
      columns: this.columns(this.extractColumnNumber(data))
    });
  }

  switchTypesBar() {
    this.columnTypesBar.classList.toggle('se-hidden');
  }

  columnTypeSelectors: Map<number, HTMLSelectElement>;

  protected populateColumnTypesBar() {
    // TODO: interface with name, id, options callback?
    const options = [
      'text',
      'numeric',
      'hidden',
      'dropdown',
      'autocomplete',
      'checkbox',
      'radio',
      'calendar',
      'image',
      'color',
      'html'
    ];
    this.columnTypesBar.innerHTML = '';
    this.columnTypeSelectors.clear();
    for (let columnId = 0; columnId < this.columnsNumber; columnId++) {
      // TODO react widget
      const columnTypeSelector = document.createElement('select');
      for (const option of options) {
        const optionElement = document.createElement('option');
        optionElement.value = option;
        optionElement.text = option;
        columnTypeSelector.appendChild(optionElement);
      }
      columnTypeSelector.onchange = ev => {
        const select = ev.target as HTMLSelectElement;
        this.columnTypes[columnId] = select.options[select.selectedIndex]
          .value as columnTypeId;
        this.reloadEditor({ columns: this.columns(this.columnTypes.length) });
      };
      this.columnTypeSelectors.set(columnId, columnTypeSelector);
      this.columnTypesBar.appendChild(columnTypeSelector);
    }
  }

  protected adjustColumnTypesWidth() {
    if (!this.selectAllElement || this.columnTypeSelectors.size === 0) {
      return;
    }
    this.columnTypesBar.style.marginLeft =
      this.selectAllElement.offsetWidth + 'px';
    const widths = this.jexcel!.getWidth(undefined);
    for (let columnId = 0; columnId < this.columnsNumber; columnId++) {
      this.columnTypeSelectors.get(columnId)!.style.width =
        widths[columnId] + 'px';
    }
  }

  protected createEditor(options: jspreadsheet.JSpreadsheetOptions) {
    this.jexcel = jspreadsheet(this.editor, options);
    this.selectAllElement =
      this.jexcel.headerContainer.querySelector('.jexcel_selectall');
  }

  protected onAfterShow(msg: Message) {
    super.onAfterShow(msg);
    this.relayout();
  }

  get ready(): Promise<void> {
    return this._ready.promise;
  }

  getValue(): string {
    const data = this.jexcel!.getData();
    if (this.firstRowAsHeader) {
      data.unshift(this.jexcel!.getHeaders(true) as string[]);
    }
    return Papa.unparse(data, {
      delimiter: this.separator,
      newline: this.linebreak
    });
  }

  setValue(value: string) {
    const parsed = this.parseValue(value);
    this.jexcel!.setData(parsed);
  }

  private _onContentChanged(): void {
    const oldValue = this.getValue();
    const newValue = this.context.model.toString();

    if (oldValue !== newValue) {
      this.setValue(newValue);
    }
  }

  get wrapper() {
    if (this.hasFrozenColumns) {
      return this.jexcel!.content;
    }
    return this.container;
  }

  onResize() {
    if (!this.jexcel) {
      return;
    }
    if (this.fitMode === 'all-equal-fit') {
      this.relayout();
    }
    if (this.hasFrozenColumns) {
      this.jexcel.content.style.width = this.node.offsetWidth + 'px';
      this.jexcel.content.style.height =
        (this.node.querySelector('.jexcel_content') as HTMLElement)
          .offsetHeight + 'px';
    }
    this.adjustColumnTypesWidth();
  }

  protected onActivateRequest(msg: Message): void {
    // ensure focus
    // TODO
    // this.jexcel.el.focus();
    this.editor.focus();
  }

  dispose(): void {
    if (this.jexcel) {
      this.jexcel.destroy();
    }
    super.dispose();
  }

  updateModel() {
    this.context.model.sharedModel.setSource(this.getValue());
  }

  freezeSelectedColumns() {
    const columns = this.jexcel!.getSelectedColumns();
    this.reloadEditor({
      freezeColumns: Math.max(...columns) + 1,
      tableOverflow: true,
      tableWidth: this.node.offsetWidth + 'px',
      tableHeight: this.node.offsetHeight + 'px'
    });
    this.hasFrozenColumns = true;
  }

  unfreezeColumns() {
    this.reloadEditor({
      freezeColumns: undefined,
      tableOverflow: false,
      tableWidth: undefined,
      tableHeight: undefined
    });
    this.hasFrozenColumns = false;
  }

  get columnsNumber(): number {
    const data = this.jexcel!.getData();
    if (!data.length) {
      return 0;
    }
    return data[0].length;
  }

  getHeaderElements() {
    const headers = [];
    for (const element of this.jexcel!.headerContainer.children) {
      // TODO use data attribute?
      if (element.className !== 'jexcel_selectall') {
        headers.push(element);
      }
    }
    return headers;
  }

  relayout() {
    if (!this.jexcel) {
      return;
    }
    const columns = this.columnsNumber;

    if (!columns) {
      return;
    }

    switch (this.fitMode) {
      case 'all-equal-default': {
        const options = this.jexcel.getConfig();
        for (let i = 0; i < columns; i++) {
          this.jexcel.setWidth(i, options.defaultColWidth, undefined);
        }
        break;
      }
      case 'all-equal-fit': {
        const indexColumn = this.node.querySelector(
          '.jexcel_selectall'
        ) as HTMLElement;
        const availableWidth = this.node.clientWidth - indexColumn.offsetWidth;
        const widthPerColumn = availableWidth / columns;
        for (let i = 0; i < columns; i++) {
          this.jexcel.setWidth(i, widthPerColumn, undefined);
        }
        break;
      }
      case 'fit-cells': {
        const data = this.jexcel.getData();
        const headers = this.getHeaderElements();
        for (let i = 0; i < columns; i++) {
          let maxColumnWidth = Math.max(25, headers[i].scrollWidth);
          for (let j = 0; j < data.length; j++) {
            const cell = this.jexcel.getCellFromCoords(i, j) as HTMLElement;
            maxColumnWidth = Math.max(maxColumnWidth, cell.scrollWidth);
          }
          this.jexcel.setWidth(i, maxColumnWidth, undefined);
        }
        break;
      }
    }
    this.adjustColumnTypesWidth();
  }

  context: DocumentRegistry.CodeContext;
  private _ready = new PromiseDelegate<void>();

  scrollCellIntoView(match: ICellCoordinates) {
    const cell = this.jexcel!.getCellFromCoords(match.column, match.row);
    const cellRect = cell.getBoundingClientRect();
    const wrapperRect = this.wrapper.getBoundingClientRect();
    let alignToTop = false;
    const softMargin = 3;

    if (cellRect.right > wrapperRect.right) {
      this.wrapper.scrollBy(cellRect.right - wrapperRect.right, 0);
    } else if (cellRect.left < wrapperRect.left) {
      this.wrapper.scrollBy(cellRect.left - wrapperRect.left, 0);
    }

    if (cellRect.top - softMargin < wrapperRect.top) {
      alignToTop = true;
    }
    if (alignToTop || cellRect.bottom + softMargin > wrapperRect.bottom) {
      cell.scrollIntoView(alignToTop);
    }
  }
}
