//Based on: https://github.com/SharePoint/sp-dev-fx-extensions/tree/master/samples/react-command-dialog
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'SampleCommandCommandSetStrings';

import ColorPickerDialog from './ColorPickerDialog';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISampleCommandCommandSetProperties {
  // This is an example; replace with your own property
  disabledCommandIds: string[] | undefined;
}

const LOG_SOURCE: string = 'SampleCommandCommandSet';

export default class SampleCommandCommandSet extends BaseListViewCommandSet<ISampleCommandCommandSetProperties> {

  private _colorCode: string;
  
    @override
    public onInit(): Promise<void> {
      Log.info(LOG_SOURCE, 'Initialized DialogDemoCommandSet');
      return Promise.resolve<void>();
    }
  
    @override
    public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
      if (this.properties.disabledCommandIds) {
        for (const commandId of this.properties.disabledCommandIds) {
          const command: Command | undefined = this.tryGetCommand(commandId);
          if (command && command.visible) {
            Log.info(LOG_SOURCE, `Hiding command ${commandId}`);
            command.visible = false;
          }
        }
      }
    }
  
    @override
    public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
      switch (event.itemId) {
        case 'COMMAND_1':
          const dialog: ColorPickerDialog = new ColorPickerDialog();
          dialog.message = 'Pick a color:';
          // Use 'EEEEEE' as the default color for first usage
          dialog.colorCode = this._colorCode || '#EEEEEE';
          dialog.show().then(() => {
            this._colorCode = dialog.colorCode;
            Dialog.alert(`Picked color: ${dialog.colorCode}`);
          });
          break;
        default:
          throw new Error('Unknown command');
      }
    }
}
