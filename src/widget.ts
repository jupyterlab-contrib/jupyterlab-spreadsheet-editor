// Copyright (c) Jupyter Development Team.
// Distributed under the terms of the Modified BSD License.

import { Widget } from "@lumino/widgets";
import { DocumentRegistry } from "@jupyterlab/docregistry";
import { PromiseDelegate, UUID } from "@lumino/coreutils";
import { PathExt } from "@jupyterlab/coreutils";
import Papa from "papaparse";
import { Message } from "@lumino/messaging";
import jexcel from "jexcel";
import { Signal } from "@lumino/signaling";
import { ISearchMatch } from "@jupyterlab/documentsearch";

type columnTypeId = 'autocomplete'
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

/**
 * An spreadsheet widget.
 */
export class SpreadsheetWidget extends Widget {
  /**
   * Construct a new Spreadsheet widget.
   */
  public jexcel: jexcel.JExcelElement
  protected separator: string;
  protected linebreak: string;
  public fitMode: 'all-equal-default' | 'all-equal-fit' | 'fit-cells';
  public changed: Signal<this, void>
  protected hasFrozenColumns: boolean;
  private editor: HTMLDivElement;
  private container: HTMLDivElement;
  private columnTypesBar: HTMLDivElement;
  private selectAllElement: HTMLElement;

  protected firstRowAsHeader: boolean;
  private header: Array<string>;
  private columnTypes: Array<columnTypeId>;

  constructor(context: DocumentRegistry.CodeContext) {
    super();
    this.id = UUID.uuid4();
    this.title.label = PathExt.basename(context.localPath);
    this.title.closable = true;
    this.context = context;
    this.separator = ""; // Papa auto detect
    this.fitMode = 'all-equal-default';
    if (context.localPath.endsWith('tsv')) {
      this.separator = '\t'
    }
    if (context.localPath.endsWith('csv')) {
      this.separator = ','
    }
    this.hasFrozenColumns = false;
    this.firstRowAsHeader = false;

    context.ready.then(() => {
      this._onContextReady();
    }).catch(console.warn);
    this.changed = new Signal<this, void>(this);
  }

  protected parseValue(content: string): jexcel.CellValue[][] {
    let parsed = Papa.parse(content, {delimiter: this.separator})
    if (!this.separator) {
      this.separator = parsed.meta.delimiter;
    }
    this.linebreak = parsed.meta.linebreak;
    if (parsed.errors.length) {
      console.warn('Parsing errors encountered', parsed.errors)
    }
    let columns_n = this.extractColumnNumber(parsed.data);
    // TODO: read the actual type from a file?
    // TODO only redefine if reading for the first time?
    // TODO add/remove when column added removed?
    if (typeof this.columnTypes === "undefined") {
      this.columnTypes = [...Array(columns_n).map(() => 'text' as columnTypeId)]
    }

    if (this.firstRowAsHeader) {
      this.header = parsed.data.shift();
    } else {
      this.header = null;
    }

    return parsed.data;
  }

  extractColumnNumber(data: jexcel.CellValue[][]): number {
    return data.length ? data[0].length : 0
  }

  columns(columns_n: number) {

    let columns: Array<jexcel.Column> = [];

    for (let i = 0; i < columns_n; i++) {
      columns.push({
        title: this.header ? this.header[i] : null,
        type: this.columnTypes[i]
      })

    }
    return columns
  }

  private _onContextReady(): void {
    if (this.isDisposed) {
      return;
    }
    const contextModel = this.context.model;

    // Set the editor model value.
    let content = contextModel.toString();
    let data = this.parseValue(content)

    let options: jexcel.Options = {
      data: data,
      minDimensions: [1, 1],
      // @ts-ignore
      //minSpareCols: 1,
      // @ts-ignore
      // minSpareRows: 1,
      csvFileName: this.title.label,
      columnDrag: true,
      onchange: () => {
        this.context.model.value.text = this.getValue();
        this.changed.emit();
      },
      oninsertcolumn: () => {
        this.populateColumnTypesBar();
        this.onResize();
      },
      ondeletecolumn: () => {
        this.populateColumnTypesBar();
        this.onResize();
      },
      onresizecolumn: () => {
        this.adjustColumnTypesWidth();
      },
      columns: this.columns(this.extractColumnNumber(data))
    };

    const container = document.createElement('div');
    container.className = 'se-area-container'
    this.node.appendChild(container);

    // TODO: move to a separate class/widget
    this.columnTypesBar = document.createElement('div');
    this.columnTypesBar.classList.add('se-column-types')
    this.columnTypesBar.classList.add('se-hidden')
    this.columnTypeSelectors = new Map();
    container.appendChild(this.columnTypesBar);

    this.editor = document.createElement('div');
    container.appendChild(this.editor)
    this.container = container;

    this.createEditor(options);

    // Wire signal connections.
    contextModel.contentChanged.connect(this._onContentChanged, this);

    this.populateColumnTypesBar();

    // If the sheet is not too big, use the more user-friendly columns width adjustment model
    if (data.length && data[0].length * data.length < 100 * 100) {
      this.fitMode = 'fit-cells';
      this.relayout()
    }

    // Resolve the ready promise.
    this._ready.resolve(undefined);
  }

  reloadEditor(options: jexcel.Options) {
    let config = this.jexcel.getConfig();
    this.jexcel.destroy();
    this.createEditor({
      ...config,
      ...options
    })
    this.relayout();
  }

  switchHeaders() {
    let value = this.getValue();
    this.firstRowAsHeader = !this.firstRowAsHeader;
    let data = this.parseValue(value);
    this.reloadEditor(
      {
        data: data,
        columns: this.columns(this.extractColumnNumber(data))
      }
    )
  }

  switchTypesBar() {
    this.columnTypesBar.classList.toggle('se-hidden')
  }

  columnTypeSelectors: Map<number, HTMLSelectElement>

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
    ]
    this.columnTypesBar.innerHTML = '';
    this.columnTypeSelectors.clear();
    for (let columnId = 0; columnId < this.columns_n; columnId++) {
      // TODO react widget
      let columnTypeSelector = document.createElement("select")
      for (let option of options) {
        let optionElement = document.createElement("option")
        optionElement.value = option;
        optionElement.text = option;
        columnTypeSelector.appendChild(optionElement)
      }
      columnTypeSelector.onchange = (ev) => {
        let select = ev.target as HTMLSelectElement;
        this.columnTypes[columnId] = select.options[select.selectedIndex].value as columnTypeId;
        this.reloadEditor({columns: this.columns(this.columnTypes.length)})

      };
      this.columnTypeSelectors.set(columnId, columnTypeSelector);
      this.columnTypesBar.appendChild(columnTypeSelector);
    }
  }

  protected adjustColumnTypesWidth() {
    if (this.columnTypeSelectors.size == 0) {
      return
    }
    this.columnTypesBar.style.marginLeft = this.selectAllElement.offsetWidth + 'px';
    let widths = this.jexcel.getWidth(null);
    for (let columnId = 0; columnId < this.columns_n; columnId++) {
      this.columnTypeSelectors.get(columnId).style.width = widths[columnId] + 'px';
    }
  }

  protected createEditor(options: jexcel.Options) {
    this.jexcel = jexcel(this.editor, options);
    this.selectAllElement = this.jexcel.headerContainer.querySelector('.jexcel_selectall');
  }

  protected onAfterShow(msg: Message) {
    super.onAfterShow(msg);
    this.relayout()
  }

  get ready(): Promise<void> {
    return this._ready.promise;
  }

  getValue(): string {
    let data = this.jexcel.getData();
    if (this.firstRowAsHeader) {
      data.unshift(this.jexcel.getHeaders(true))
    }
    return Papa.unparse(
      data,
      {
        delimiter: this.separator,
        newline: this.linebreak
      }
    )
  }

  setValue(value: string) {
    let parsed = this.parseValue(value);
    this.jexcel.setData(parsed);
  }

  private _onContentChanged(): void {
    const oldValue = this.getValue();
    const newValue = this.context.model.toString();

    if (oldValue !== newValue) {
      this.setValue(newValue)
    }
  }

  get wrapper() {
    if (this.hasFrozenColumns) {
      return this.jexcel.content;
    }
    return this.container;
  }

  onResize() {
    if (typeof this.jexcel === 'undefined') {
      return
    }
    if (this.fitMode == 'all-equal-fit') {
      this.relayout();
    }
    if (this.hasFrozenColumns) {
      this.jexcel.content.style.width = this.node.offsetWidth + 'px';
      this.jexcel.content.style.height = (
        this.node.querySelector('.jexcel_content') as HTMLElement
      ).offsetHeight + 'px'
    }
    this.adjustColumnTypesWidth();
  }

  protected onActivateRequest(msg: Message): void {
    // ensure focus
    this.jexcel.el.focus()
  }

  dispose(): void {
    if (this.jexcel) {
      this.jexcel.destroy();
    }
    super.dispose();
  }

  updateModel() {
    this.context.model.value.text = this.getValue();
  }

  freezeSelectedColumns() {
    let columns = this.jexcel.getSelectedColumns()
    this.reloadEditor(
      {
        // @ts-ignore
        freezeColumns: Math.max(...columns) + 1,
        tableOverflow: true,
        tableWidth: this.node.offsetWidth + 'px',
        tableHeight: this.node.offsetHeight + 'px'
      }
    )
    this.hasFrozenColumns = true;
  }

  unfreezeColumns() {
    this.reloadEditor(
      {
        // @ts-ignore
        freezeColumns: null,
        tableOverflow: false,
        tableWidth: null,
        tableHeight: null
      }
    )
    this.hasFrozenColumns = false;
  }

  get columns_n(): number {
    let data = this.jexcel.getData();
    if (!data.length) {
      return;
    }
    return data[0].length;
  }

  getHeaderElements() {
    let headers = [];
    for (let element of this.jexcel.headerContainer.children) {
      // TODO use data attribute?
      if (element.className != 'jexcel_selectall') {
        headers.push(element)
      }
    }
    return headers;
  }

  relayout() {
    if (typeof this.jexcel === 'undefined') {
      return
    }
    let columns = this.columns_n;

    if (!columns) {
      return;
    }
    
    switch (this.fitMode) {
      case "all-equal-default": {
        let options = this.jexcel.getConfig();
        for (let i = 0; i < columns; i++) {
          this.jexcel.setWidth(i, options.defaultColWidth, null);
        }
        break;
      }
      case "all-equal-fit": {
        let indexColumn = this.node.querySelector('.jexcel_selectall') as HTMLElement
        let availableWidth = this.node.clientWidth - indexColumn.offsetWidth;
        let widthPerColumn = availableWidth / columns;
        for (let i = 0; i < columns; i++) {
          this.jexcel.setWidth(i, widthPerColumn, null);
        }
        break;
      }
      case "fit-cells": {
        let data = this.jexcel.getData();
        let headers = this.getHeaderElements();
        for (let i = 0; i < columns; i++) {
          let maxColumnWidth = Math.max(25, headers[i].scrollWidth);
          for (let j = 0; j < data.length; j++) {
            let cell = this.jexcel.getCellFromCoords(i, j) as HTMLElement;
            maxColumnWidth = Math.max(maxColumnWidth, cell.scrollWidth);
          }
          this.jexcel.setWidth(i, maxColumnWidth, null);
        }
        break;
      }
    }
    this.adjustColumnTypesWidth();
  }

  context: DocumentRegistry.CodeContext;
  private _ready = new PromiseDelegate<void>();

  scrollCellIntoView(match: ISearchMatch) {
    let cell = this.jexcel.getCellFromCoords(match.column, match.line)
    let cellRect = cell.getBoundingClientRect();
    let wrapperRect = this.wrapper.getBoundingClientRect();
    let alignToTop = false;

    if (cellRect.top < wrapperRect.top) {
      alignToTop = true;
    }
    if (alignToTop || cellRect.bottom > wrapperRect.bottom || cellRect.left < wrapperRect.left || cellRect.right > wrapperRect.right) {
      cell.scrollIntoView(alignToTop);
    }
  }
}
