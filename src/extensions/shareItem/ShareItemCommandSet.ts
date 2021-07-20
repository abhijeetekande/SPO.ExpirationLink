import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import SharingPanel, {ISharingPanelProps} from '../../components/sharingPanel';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'ShareItemCommandSetStrings';
import { sp } from "@pnp/sp";
import { assign } from '@uifabric/utilities';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IShareItemCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ShareItemCommandSet';

export default class ShareItemCommandSet extends BaseListViewCommandSet<IShareItemCommandSetProperties> {
  private panelPlaceHolder: HTMLDivElement = null;
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ShareItemCommandSet');
   // Setup the PnP JS with SPFx context
   sp.setup({
    spfxContext: this.context
  });

  // Create the container for our React component
  this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));
  return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
            const div = document.createElement('div');
            let selectedItem = event.selectedRows[0];
            const listItemId = selectedItem.getValueByName('ID') as number;
            const title = selectedItem.getValueByName("Title");
            this._showPanel(listItemId, title);
   // break;
     //   Dialog.alert(`${this.properties.sampleTextOne}`);
        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _showPanel(itemId: number, currentTitle: string) {
    this._renderPanelComponent({
      isOpen: true,
      currentTitle,
      itemId,
      listId: this.context.pageContext.list.id.toString(),
      onClose: this._dismissPanel,
      context:this.context,
      siteurl: this.context.pageContext.site.absoluteUrl,
    });
  }

  private _dismissPanel() {
    this._renderPanelComponent({ isOpen: false });
  }

  private _renderPanelComponent(props: any) {
    const element: React.ReactElement<ISharingPanelProps> = React.createElement(SharingPanel, assign({
      onClose: null,
      currentTitle: null,
      itemId: null,
      isOpen: false,
      listId: null
    }, props));
    ReactDom.render(element, this.panelPlaceHolder);
  }
}
