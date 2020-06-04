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

    context.ready.then(() => {
      this._onContextReady();
    });
    this.changed = new Signal<this, void>(this);
  }

  protected parseValue(content: string): jexcel.CellValue[][] {
    let parsed = Papa.parse(content, {'delimiter': this.separator})
    if (!this.separator) {
      this.separator = parsed.meta.delimiter;
    }
    this.linebreak = parsed.meta.linebreak;
    if (parsed.errors.length) {
      console.warn('Parsing errors encountered', parsed.errors)
    }
    return parsed.data;
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
        this.onResize();
      },
      ondeletecolumn: () => {
        this.onResize();
      }
    };
    this.jexcel = jexcel(this.node as HTMLDivElement, options);

    // Wire signal connections.
    contextModel.contentChanged.connect(this._onContentChanged, this);

    // If the sheet is not too big, use the more user-friendly columns width adjustment model
    if (data.length && data[0].length * data.length < 100 * 100) {
      this.fitMode = 'fit-cells';
      this.relayout()
    }

    // Resolve the ready promise.
    this._ready.resolve(undefined);
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
    return Papa.unparse(
      data,
      {
        delimiter: this.separator,
        newline: this.linebreak
      }
    )
  }

  setValue(value: string) {
    this.jexcel.setData(this.parseValue(value));
  }

  private _onContentChanged(): void {
    const oldValue = this.getValue();
    const newValue = this.context.model.toString();

    if (oldValue !== newValue) {
      this.setValue(newValue)
    }
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
    let config = this.jexcel.getConfig();
    this.jexcel.destroy();
    this.jexcel = jexcel(
      this.node as HTMLDivElement,
      {
        ...config,
        freezeColumns: Math.max(...columns) + 1,
        tableOverflow: true,
        tableWidth: this.node.offsetWidth + 'px',
        tableHeight: this.node.offsetHeight + 'px'
      }
    )
    this.hasFrozenColumns = true;
  }

  unfreezeColumns() {
    let config = this.jexcel.getConfig();
    this.jexcel.destroy();
    this.jexcel = jexcel(
      this.node as HTMLDivElement,
      {
        ...config,
        freezeColumns: null,
        tableOverflow: false,
        tableWidth: null,
        tableHeight: null
      }
    )
    this.hasFrozenColumns = false;
  }

  get columns(): number {
    let data = this.jexcel.getData();
    if (!data.length) {
      return;
    }
    return data[0].length;
  }

  relayout() {
    let columns = this.columns;

    if (!columns) {
      return;
    }
    
    switch (this.fitMode) {
      case "all-equal-default":
        let options = this.jexcel.getConfig();
        for (let i = 0; i < columns; i++) {
          this.jexcel.setWidth(i, options.defaultColWidth, null);
        }
        break;
      case "all-equal-fit":
        let indexColumn = this.node.querySelector('.jexcel_selectall') as HTMLElement
        let availableWidth = this.node.clientWidth - indexColumn.offsetWidth;
        let widthPerColumn = availableWidth / columns;
        for (let i = 0; i < columns; i++) {
          this.jexcel.setWidth(i, widthPerColumn, null);
        }
        break;
      case "fit-cells":
        let data = this.jexcel.getData();
        for (let i = 0; i < columns; i++) {
          let maxColumnWidth = 25;
          let header = this.jexcel.getHeader(i);
          for (let j = 0; j < data.length; j++) {
            let cell = this.jexcel.getCell(header + (j + 1)) as HTMLElement;
            maxColumnWidth = Math.max(maxColumnWidth, cell.scrollWidth);
          }
          this.jexcel.setWidth(i, maxColumnWidth, null);
        }
        break;
    }
  }

  context: DocumentRegistry.CodeContext;
  private _ready = new PromiseDelegate<void>();
}
