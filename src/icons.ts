import { LabIcon } from '@jupyterlab/ui-components';

import removeRowSvg from '../style/icons/mdi-table-row-remove.svg';
import addColumnSvg from '../style/icons/mdi-table-column-plus-after.svg';
import removeColumnSvg from '../style/icons/mdi-table-column-remove.svg';
import freezeColumnSvg from '../style/icons/mdi-snowflake.svg';
import unfreezeColumnSvg from '../style/icons/mdi-snowflake-off.svg';
import addRowSvg from '../style/icons/mdi-table-row-plus-after.svg';

export const freezeColumnIcon = new LabIcon({
  name: 'spreadsheet:freeze-columns',
  svgstr: freezeColumnSvg
});

export const unfreezeColumnIcon = new LabIcon({
  name: 'spreadsheet:unfreeze-columns',
  svgstr: unfreezeColumnSvg
});

export const removeColumnIcon = new LabIcon({
  name: 'spreadsheet:remove-column',
  svgstr: removeColumnSvg
});

export const addColumnIcon = new LabIcon({
  name: 'spreadsheet:add-column',
  svgstr: addColumnSvg
});

export const removeRowIcon = new LabIcon({
  name: 'spreadsheet:remove-row',
  svgstr: removeRowSvg
});

export const addRowIcon = new LabIcon({
  name: 'spreadsheet:add-row',
  svgstr: addRowSvg
});
