import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'CommandSetCommandSetStrings';


import {sp } from "@pnp/sp/presets/all";


/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICommandSetCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'CommandSetCommandSet';

export default class CommandSetCommandSet extends BaseListViewCommandSet<ICommandSetCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CommandSetCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }



  const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
  if (compareTwoCommand) {
    // This command should be hidden unless exactly one row is selected.
    compareTwoCommand.visible = event.selectedRows.length > 1;
  }

  const compareThreeCommand: Command = this.tryGetCommand('COMMAND_3');
  if (compareThreeCommand) {
    // This command should be hidden unless exactly one row is selected.
    compareThreeCommand.visible = event.selectedRows.length > 1;
  }


}



  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        // Dialog.alert(`${this.properties.sampleTextOne}`);
      let title: string= event.selectedRows[0].getValueByName("Title").toString();
      let Status: string = event.selectedRows[0].getValueByName("Status").toString();

      Dialog.alert(`Project Name: ${title} \n Project Status: ${Status}% done`);

        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;

      case 'COMMAND_3':
        Dialog.prompt(`Project Status Remarks`).then((value:string)=>{
          this.UpdateRemarks(event.selectedRows,value)
        });


      default:
        throw new Error('Unknown command');
    }
  }

  private UpdateRemarks(items: any, value: string) {
    let batch = sp.createBatch();

    items.forEach(item => {
      sp.web.lists.getByTitle("ProjectsStatus").items.getById(item.getValueByName('ID')).inBatch(batch).update({ Remarks: value }).then(res => {

      });
    });
    batch.execute().then(res => {
      location.reload();
    });;
  }



}
