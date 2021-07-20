import * as React from 'react';
import { PanelType } from 'office-ui-fabric-react/lib/Panel';
import { TextField, DefaultButton, PrimaryButton, DialogFooter, Panel, Spinner } from "office-ui-fabric-react";
import { PrincipalType, IOfficeUiFabricPeoplePickerProps, OfficeUiFabricPeoplePicker, TypePicker } from './OfficeUiFabricPeoplePicker';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { sp } from "@pnp/sp";
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';


const today: Date = new Date(Date.now());
//Only allow forecoming month to be selected
const minDate = today;
minDate.setMonth(today.getMonth()+1);
minDate.setDate(1);
var samples;
const DayPickerStrings: IDatePickerStrings = {
  months: [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ],

  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

  goToToday: '',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  isRequiredErrorMessage: 'Field is required.',
};
export interface ISharingPanelState {
    saving: boolean;
    firstDayOfWeek?: DayOfWeek;
    Plannedclosingdateforaction: Date;
    Personsforcorrectiveactions: string;
}

export interface ISharingPanelProps {
    onClose: () => void;
    isOpen: boolean;
    currentTitle: string;
    itemId: number;
    listId: string;
    context:ListViewCommandSetContext;
    siteurl: string;
}

export default class SharingPanel extends React.Component<ISharingPanelProps, ISharingPanelState > {
    private editedTitle: string = null;
    constructor(props: ISharingPanelProps) {
        super(props);
        this.state = {
            saving: false,
            firstDayOfWeek:  DayOfWeek.Monday,
            Plannedclosingdateforaction: null,
            Personsforcorrectiveactions: "",
        };
    }
    private _onSelectDate = (date: Date | null | undefined): void => {
        this.setState({ Plannedclosingdateforaction: date });
        console.log('Date:', date.toDateString());
      }

    
    private _onTitleChanged(title: string) {
        this.editedTitle = title;
    }

    private _getPeoplePickerItems= (items: any[]): void => {
        //console.log(items)
        var person= "";
        if(items.length >0)
        {
          person= items[0].user.Id.toString();
        }
        this.setState({Personsforcorrectiveactions: person});
       }
    
    private _onCancel() {
        this.props.onClose();
    }

    
    private _onSave() {
        this.setState({ saving: true });
        // sp.web.lists.getById(this.props.listId).items.getById(this.props.itemId).update({
        //     'Title': this.editedTitle
        // }).then(() => {
        //     this.setState({ saving: false });
        //     this.props.onClose();
        // });
    }

  public render(): React.ReactElement<ISharingPanelProps> {
    let { isOpen, currentTitle } = this.props;
    const { firstDayOfWeek, Plannedclosingdateforaction } = this.state;
    return (
      
        <Panel isOpen={isOpen}>
                <h2>This is a custom panel with your own content</h2>
                <TextField value={currentTitle} label="Item title" placeholder="Choose the new title" />
                <OfficeUiFabricPeoplePicker
                        spHttpClient= {this.props.context.spHttpClient}
                        siteUrl={this.props.siteurl}
                        typePicker={TypePicker.Normal}
                        principalType={PrincipalType.User}
                        numberOfItems= {10}
                        itemLimit={1}
                        onChange={this._getPeoplePickerItems}
                        >
                    </OfficeUiFabricPeoplePicker>
                    <div className="docs-DatePickerExample">                      
                      <DatePicker
                        isRequired={true}                        
                        firstDayOfWeek={DayOfWeek.Monday}
                        strings={DayPickerStrings}
                        placeholder="Select a date..."
                        minDate={minDate}         
                        showMonthPickerAsOverlay={false}             
                        allowTextInput={true}    
                        value= {Plannedclosingdateforaction}                    
                        onSelectDate= {this._onSelectDate}
                      />
                    </div>   
                <DialogFooter>
                    <DefaultButton text="Cancel" onClick={this._onCancel} />
                    <PrimaryButton text="Save" onClick={this._onSave} />
                </DialogFooter>
            </Panel>
      
    );
  }
}