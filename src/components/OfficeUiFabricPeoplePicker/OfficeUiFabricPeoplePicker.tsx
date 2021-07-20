
import * as React from 'react';
import pnp, { FetchOptions,HttpClient} from 'sp-pnp-js';
//import { HttpClient } from 'sp-pnp-js';
import {
  IOfficeUiFabricPeoplePickerProps,
  IOfficeUiFabricPeoplePickerState,
  IClientPeoplePickerSearchUser,
  SharePointSearchUserPersona,
  IEnsurableSharePointUser,
  IEnsureUser,
  TypePicker
} from '.';
import { NormalPeoplePicker, IPersonaProps, CompactPeoplePicker, IBasePickerSuggestionsProps } from 'office-ui-fabric-react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as lodash from 'lodash';
let httpClient = new HttpClient();

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "Suggestions",
  noResultsFoundText: "No macthing users",
  loadingText: "Loading"
};
export interface HttpClientImpl {
  fetch(url: string, options: FetchOptions): Promise<Response>;
}

export class OfficeUiFabricPeoplePicker extends React.Component<IOfficeUiFabricPeoplePickerProps, IOfficeUiFabricPeoplePickerState> {

  constructor(props: IOfficeUiFabricPeoplePickerProps, context?: any) {
    super(props, context);
    this.state = {
      selectedItems: props.selectedItems
    };
  }

  public render(): React.ReactElement<IOfficeUiFabricPeoplePickerProps> {
    if (this.props.typePicker === TypePicker.Normal) {
      return (
        <NormalPeoplePicker
          onChange={this._onChange.bind(this)}
          onResolveSuggestions={this._onFilterChanged}
          getTextFromItem={(persona: IPersonaProps) => persona.text}
          pickerSuggestionsProps={suggestionProps}
          className={'ms-PeoplePicker'}
          selectedItems={this.state.selectedItems}
          key={'normal'}
          itemLimit={this.props.itemLimit}
        />
      );
    } else {
      return (
        <CompactPeoplePicker
          onChange={this._onChange.bind(this)}
          onResolveSuggestions={this._onFilterChanged}
          getTextFromItem={(persona: IPersonaProps) => persona.text}
          pickerSuggestionsProps={suggestionProps}
          selectedItems={this.state.selectedItems}
          className={'ms-PeoplePicker'}
          key={'normal'}
          itemLimit={this.props.itemLimit}
        />
      );
    }
  }

  private _onChange(items: any[]) {
    this.setState({
      selectedItems: items
    });
    if (this.props.onChange) {
      this.props.onChange(items);
    }
  }

  
  private _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
    if (!filterText || filterText.length < 3) return Promise.resolve([] as IPersonaProps[]);
     return this._searchPeople(filterText);
  }

  /**
   * @function
   * Returns people results after a REST API call
   */
  private _searchPeople(terms: string) {
        
    const userRequestUrl = `${this.props.siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
    const ensureUserUrl = `${this.props.siteUrl}/_api/web/ensureUser`;
    const userQueryParams = {
      'queryParams': {
        'AllowEmailAddresses': true,
        'AllowMultipleEntities': false,
        'AllUrlZones': false,
        'MaximumEntitySuggestions': this.props.numberOfItems,
        'PrincipalSource': 15,
        'PrincipalType': this.props.principalType,
        'QueryString': terms
      }
    };

    return httpClient.post(userRequestUrl,
      { body: JSON.stringify(userQueryParams) })
      .then(response => response.json())
      .then((response: {value: string}) => {
       // const batch =this.props.spHttpClient.beginBatch();
        let userQueryResults: IClientPeoplePickerSearchUser[] = JSON.parse(response.value);

        var batch = pnp.sp.createBatch();
        const batchPromises = userQueryResults.map(p=>
         pnp.sp.web.inBatch(batch).ensureUser(p.Key)
         .then(r => 
          {
             return r.user.get().then((user : IEnsureUser) => ({ ...p, ...user } as IEnsurableSharePointUser));
          })
       );

          return batch.execute().then(() => 
          Promise.all(batchPromises).then(users => users.map(u => SharePointSearchUserPersona(u))));

      });
 }
}
