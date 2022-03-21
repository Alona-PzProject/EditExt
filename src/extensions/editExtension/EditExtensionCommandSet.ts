import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { Constants } from './Models/Constants';


/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IEditExtensionCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'EditExtensionCommandSet';

export default class EditExtensionCommandSet extends BaseListViewCommandSet<IEditExtensionCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized EditExtensionCommandSet');
    console.log('10');
    return Promise.resolve();
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    // Get Nevigate_To_New_Form Command
    const compareOneCommand: Command = this.tryGetCommand('Nevigate_To_New_Form');

    // Get Nevigate_To_Edit_Form Command
    const compareSecondCommand: Command = this.tryGetCommand('Nevigate_To_Edit_Form');

    // Get current library name from pageContext
    let LibraryName = this.context.pageContext.list.title;

    if (Constants.ARRAY_OF_ACTIVE_LIBRARY_NAMES.indexOf(LibraryName) !== -1) {

      if (compareOneCommand) {
        compareOneCommand.visible = true;
      }
      if (compareSecondCommand) {
        compareSecondCommand.visible = event.selectedRows.length === 1;
      }
    } else {
      if (compareOneCommand) {
        compareOneCommand.visible = false;
      }
      if (compareSecondCommand) {
        compareSecondCommand.visible = false;
      }
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'Nevigate_To_New_Form':
        let webUri = this.context.pageContext.web.absoluteUrl;
        window.location.href = webUri + Constants.NewFormUrl;
        break;
      case 'Nevigate_To_Edit_Form':
        if(event.selectedRows.length > 0) {
          let ItemID = event.selectedRows[0].getValueByName('ID').toString();
          let webUri = this.context.pageContext.web.absoluteUrl;
          let StringToRedirect = webUri + Constants.EditFormUrl +'?FormID=' + ItemID;
          window.location.href = StringToRedirect;  
        }
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
