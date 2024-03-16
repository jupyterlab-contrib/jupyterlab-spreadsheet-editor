// Copyright (c) Jupyter Development Team.
// Distributed under the terms of the Modified BSD License.

import {
  ILayoutRestorer,
  JupyterFrontEnd,
  JupyterFrontEndPlugin
} from '@jupyterlab/application';
import { IMainMenu } from '@jupyterlab/mainmenu';
import {
  ABCWidgetFactory,
  DocumentRegistry,
  IDocumentWidget
} from '@jupyterlab/docregistry';
import { Widget } from '@lumino/widgets';
import { DocumentWidget } from '@jupyterlab/docregistry';
import { ICommandPalette, WidgetTracker } from '@jupyterlab/apputils';
import {
  redoIcon,
  undoIcon,
  addAboveIcon,
  addBelowIcon,
  addIcon,
  copyIcon,
  pasteIcon
} from '@jupyterlab/ui-components';
import { ISearchProviderRegistry } from '@jupyterlab/documentsearch';
import { CommandRegistry } from '@lumino/commands';

import '../style/index.css';
import { SpreadsheetWidget } from './widget';
import { SpreadsheetEditorDocumentWidget } from './documentwidget';
import { SpreadsheetSearchProviderFactory } from './searchprovider';
import { ILauncher } from '@jupyterlab/launcher';
import { spreadsheetIcon } from '@jupyterlab/ui-components';
import {
  IFileBrowserFactory,
  IDefaultFileBrowser
} from '@jupyterlab/filebrowser';
import { IStatusBar } from '@jupyterlab/statusbar';
import { SelectionStatus } from './statusbar';
import {
  ITranslator,
  nullTranslator,
  TranslationBundle
} from '@jupyterlab/translation';
import { removeColumnIcon, removeRowIcon } from './icons';

const paletteCategory = 'Spreadsheet Editor';

const FACTORY = 'Spreadsheet Editor';

/**
 * A widget factory for editors.
 */
export class SpreadsheetEditorFactory extends ABCWidgetFactory<
  IDocumentWidget<SpreadsheetWidget>,
  DocumentRegistry.ICodeModel
> {
  /**
   * Create a new widget given a context.
   */
  protected createNewWidget(
    context: DocumentRegistry.CodeContext
  ): IDocumentWidget<SpreadsheetWidget> {
    const content = new SpreadsheetWidget(context);
    return new SpreadsheetEditorDocumentWidget({
      content,
      context,
      translator: this.translator
    });
  }
}

/**
 * Add File Editor undo and redo widgets to the Edit menu
 */
export function addUndoRedoToEditMenu(menu: IMainMenu) {
  const isEnabled = (widget: Widget): boolean => {
    return (
      widget instanceof DocumentWidget &&
      widget.content instanceof SpreadsheetWidget
    );
  };
  menu.editMenu.undoers.undo.add({
    id: CommandIDs.undo,
    isEnabled
  });

  menu.editMenu.undoers.redo.add({
    id: CommandIDs.redo,
    isEnabled
  });
}

/**
 * Function to create a new untitled text file, given the current working directory.
 */
function createNew(commands: CommandRegistry, cwd: string, ext = 'tsv') {
  return commands
    .execute('docmanager:new-untitled', {
      path: cwd,
      type: 'file',
      ext
    })
    .then(model => {
      return commands.execute('docmanager:open', {
        path: model.path,
        factory: FACTORY
      });
    });
}

/**
 * The command IDs used by the spreadsheet editor plugin.
 */
export namespace CommandIDs {
  export const createNewCSV = 'spreadsheet-editor:create-new-csv-file';

  export const createNewTSV = 'spreadsheet-editor:create-new-tsv-file';

  export const undo = 'spreadsheet-editor:undo';

  export const redo = 'spreadsheet-editor:redo';

  export const copy = 'spreadsheet-editor:copy';

  export const paste = 'spreadsheet-editor:paste';

  export const insertRowBelow = 'spreadsheet-editor:insert-row-below';

  export const insertRowAbove = 'spreadsheet-editor:insert-row-above';

  export const insertColumnLeft = 'spreadsheet-editor:insert-column-left';

  export const insertColumnRight = 'spreadsheet-editor:insert-column-right';

  export const removeColumn = 'spreadsheet-editor:remove-column';

  export const removeRow = 'spreadsheet-editor:remove-row';
}

/**
 * Add Create New DSV File to the Launcher
 */
export function addCreateNewToLauncher(
  launcher: ILauncher,
  trans: TranslationBundle
) {
  launcher.add({
    command: CommandIDs.createNewCSV,
    category: trans.__('Other'),
    rank: 3
  });
  launcher.add({
    command: CommandIDs.createNewTSV,
    category: trans.__('Other'),
    rank: 3
  });
}

/**
 * Add the New File command
 */
export function addCreateNewCommands(
  commands: CommandRegistry,
  contextMenuHitTest: (
    fn: (node: HTMLElement) => boolean
  ) => HTMLElement | undefined,
  tracker: WidgetTracker<IDocumentWidget<SpreadsheetWidget>>,
  browserFactory: IFileBrowserFactory,
  defaultBrowser: IDefaultFileBrowser,
  trans: TranslationBundle
) {
  const getJexcel = () => {
    return tracker.currentWidget?.content?.jexcel;
  };

  commands.addCommand(CommandIDs.undo, {
    label: trans.__('Undo'),
    icon: undoIcon.bindprops({ stylesheet: 'menuItem' }),
    execute: () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return;
      }
      jexcel.undo();
    },
    isEnabled: () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return false;
      }
      return jexcel.history.length !== 0 && jexcel.historyIndex !== 0;
    }
  });

  commands.addCommand(CommandIDs.redo, {
    label: trans.__('Redo'),
    icon: redoIcon.bindprops({ stylesheet: 'menuItem' }),
    execute: () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return;
      }
      jexcel.redo();
    },
    isEnabled: () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return false;
      }
      return (
        jexcel.history.length !== 0 &&
        jexcel.historyIndex !== jexcel.history.length - 1
      );
    }
  });

  commands.addCommand(CommandIDs.copy, {
    label: trans.__('Copy'),
    icon: copyIcon.bindprops({ stylesheet: 'menuItem' }),
    execute: () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return;
      }
      jexcel.copy(true);
    },
    isEnabled: () => !!getJexcel()
  });

  commands.addCommand(CommandIDs.paste, {
    label: trans.__('Paste'),
    icon: pasteIcon.bindprops({ stylesheet: 'menuItem' }),
    execute: async () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return;
      }
      const selection = jexcel.selectedCell;
      if (!selection) {
        return;
      }
      const text = await navigator.clipboard.readText();
      if (text) {
        jexcel.paste(selection[0] as number, selection[1] as number, text);
      }
    },
    isEnabled: () => !!getJexcel()
  });

  commands.addCommand(CommandIDs.insertRowBelow, {
    label: trans.__('Insert Row Below'),
    icon: addBelowIcon.bindprops({ stylesheet: 'menuItem' }),
    execute: () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return;
      }
      const cell = contextMenuHitTest(node => node.matches('td[data-y]'));
      const rowNumber = cell?.dataset.y
        ? parseInt(cell.dataset.y, 10)
        : undefined;
      jexcel.insertRow(1, rowNumber);
    },
    isEnabled: () => !!getJexcel()
  });

  commands.addCommand(CommandIDs.insertRowAbove, {
    label: trans.__('Insert Row Above'),
    icon: addAboveIcon.bindprops({ stylesheet: 'menuItem' }),
    execute: () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return;
      }
      const cell = contextMenuHitTest(node => node.matches('td[data-y]'));
      const rowNumber = cell?.dataset.y
        ? parseInt(cell.dataset.y, 10)
        : undefined;
      // @ts-expect-error (wrong typing for insertBefore as `number`, should be `boolean`)
      jexcel.insertRow(1, rowNumber, true);
    },
    isEnabled: () => !!getJexcel()
  });

  commands.addCommand(CommandIDs.insertColumnLeft, {
    label: trans.__('Insert Column To The Left'),
    icon: addIcon.bindprops({ stylesheet: 'menuItem' }),
    execute: () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return;
      }
      const cell = contextMenuHitTest(node => node.matches('td[data-x]'));
      const columnNumber = cell?.dataset.x
        ? parseInt(cell.dataset.x, 10)
        : undefined;
      jexcel.insertColumn(1, columnNumber, true);
    },
    isEnabled: () => !!getJexcel()
  });

  commands.addCommand(CommandIDs.insertColumnRight, {
    label: trans.__('Insert Column To The Right'),
    icon: addIcon.bindprops({ stylesheet: 'menuItem' }),
    execute: () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return;
      }
      const cell = contextMenuHitTest(node => node.matches('td[data-x]'));
      const columnNumber = cell?.dataset.x
        ? parseInt(cell.dataset.x, 10)
        : undefined;
      jexcel.insertColumn(1, columnNumber, false);
    },
    isEnabled: () => !!getJexcel()
  });

  commands.addCommand(CommandIDs.removeColumn, {
    label: trans.__('Delete Column'),
    icon: removeColumnIcon.bindprops({ stylesheet: 'menuItem' }),
    execute: () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return;
      }
      const cell = contextMenuHitTest(node => node.matches('td[data-x]'));
      if (!cell) {
        return;
      }
      const columnNumber = parseInt(cell.dataset.x!, 10);
      jexcel.deleteColumn(columnNumber);
    },
    isEnabled: () => !!getJexcel()
  });

  commands.addCommand(CommandIDs.removeRow, {
    label: trans.__('Delete Row'),
    icon: removeRowIcon.bindprops({ stylesheet: 'menuItem' }),
    execute: () => {
      const jexcel = getJexcel();
      if (!jexcel) {
        return;
      }
      const cell = contextMenuHitTest(node => node.matches('td[data-y]'));
      if (!cell) {
        return;
      }
      const rowNumber = parseInt(cell.dataset.y!, 10);
      jexcel.deleteRow(rowNumber);
    },
    isEnabled: () => !!getJexcel()
  });

  commands.addCommand(CommandIDs.createNewCSV, {
    label: args =>
      args['isPalette'] ? trans.__('New CSV File') : trans.__('CSV File'),
    caption: trans.__('Create a new CSV file'),
    icon: args => (args['isPalette'] ? undefined : spreadsheetIcon),
    execute: args => {
      const currentBrowser =
        browserFactory?.tracker.currentWidget ?? defaultBrowser;
      const cwd = args['cwd'] || currentBrowser.model.path;
      return createNew(commands, cwd as string, 'csv');
    }
  });

  commands.addCommand(CommandIDs.createNewTSV, {
    label: args =>
      args['isPalette'] ? trans.__('New TSV File') : trans.__('TSV File'),
    caption: trans.__('Create a new TSV file'),
    icon: args => (args['isPalette'] ? undefined : spreadsheetIcon),
    execute: args => {
      const currentBrowser =
        browserFactory?.tracker.currentWidget ?? defaultBrowser;
      const cwd = args['cwd'] || currentBrowser.model.path;
      return createNew(commands, cwd as string, 'tsv');
    }
  });
}

const PLUGIN_ID = 'spreadsheet-editor:plugin';

/**
 * Initialization data for the spreadsheet-editor extension.
 */
const extension: JupyterFrontEndPlugin<void> = {
  id: PLUGIN_ID,
  autoStart: true,
  requires: [IFileBrowserFactory, IDefaultFileBrowser],
  optional: [
    ICommandPalette,
    ILauncher,
    IMainMenu,
    ILayoutRestorer,
    ISearchProviderRegistry,
    IStatusBar,
    ITranslator
  ],
  activate: (
    app: JupyterFrontEnd,
    browserFactory: IFileBrowserFactory,
    defaultBrowser: IDefaultFileBrowser,
    palette: ICommandPalette | null,
    launcher: ILauncher | null,
    menu: IMainMenu | null,
    restorer: ILayoutRestorer | null,
    searchregistry: ISearchProviderRegistry | null,
    statusBar: IStatusBar | null,
    translator: ITranslator | null
  ) => {
    translator = translator || nullTranslator;
    const trans = translator.load(PLUGIN_ID);

    const factory = new SpreadsheetEditorFactory({
      name: FACTORY,
      fileTypes: ['csv', 'tsv', '*'],
      defaultFor: ['csv', 'tsv'],
      translator: translator
    });

    const tracker = new WidgetTracker<IDocumentWidget<SpreadsheetWidget>>({
      namespace: PLUGIN_ID
    });

    if (restorer) {
      void restorer.restore(tracker, {
        command: 'docmanager:open',
        args: widget => ({ path: widget.context.path, factory: FACTORY }),
        name: widget => widget.context.path
      });
    }

    app.docRegistry.addWidgetFactory(factory);
    const ft = app.docRegistry.getFileType('csv');

    factory.widgetCreated.connect((sender, widget) => {
      // Track the widget.
      void tracker.add(widget);
      // Notify the widget tracker if restore data needs to update.
      widget.context.pathChanged.connect(() => {
        void tracker.save(widget);
      });

      if (ft) {
        widget.title.icon = ft.icon!;
        widget.title.iconClass = ft.iconClass!;
        widget.title.iconLabel = ft.iconLabel!;
      }
    });

    if (searchregistry) {
      searchregistry.add(PLUGIN_ID, new SpreadsheetSearchProviderFactory());
    }

    addCreateNewCommands(
      app.commands,
      app.contextMenuHitTest.bind(app),
      tracker,
      browserFactory,
      defaultBrowser,
      trans
    );

    if (palette) {
      palette.addItem({
        command: CommandIDs.createNewCSV,
        args: { isPalette: true },
        category: paletteCategory
      });
      palette.addItem({
        command: CommandIDs.createNewTSV,
        args: { isPalette: true },
        category: paletteCategory
      });
    }

    if (launcher) {
      addCreateNewToLauncher(launcher, trans);
    }

    if (menu) {
      addUndoRedoToEditMenu(menu);
    }

    if (statusBar) {
      const item = new SelectionStatus(translator);
      // Keep the status item up-to-date with the current spreadsheet editor.
      tracker.currentChanged.connect(() => {
        const current = tracker.currentWidget;
        item.model.spreadsheetWidget = current?.content ?? null;
      });

      statusBar.registerStatusItem(PLUGIN_ID, {
        item,
        align: 'right',
        rank: 4,
        isActive: () =>
          !!app.shell.currentWidget &&
          !!tracker.currentWidget &&
          app.shell.currentWidget === tracker.currentWidget
      });
    }
  }
};

export default extension;
