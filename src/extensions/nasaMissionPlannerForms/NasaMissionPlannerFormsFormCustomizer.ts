import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { IMissionPlan } from '../../models';
import { FormDisplayMode } from '@microsoft/sp-core-library';
import { BaseFormCustomizer } from '@microsoft/sp-listview-extensibility';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import {
  NasaMissionPlannerForms, INasaMissionPlannerFormsProps,
  NasaMissionPlannerViewForm, INasaMissionPlannerViewFormProps
} from './components';

export interface INasaMissionPlannerFormsFormCustomizerProperties {
  columnMissionTitleInternalName?: string;
  columnMissionProjectInternalName?: string;
  columnLaunchVehicleInternalName?: string;
  columnLaunchSpacecraftInternalName?: string;
}

export default class NasaMissionPlannerFormsFormCustomizer
  extends BaseFormCustomizer<INasaMissionPlannerFormsFormCustomizerProperties> {

  private _currentMission: IMissionPlan;
  private _currentMissionEtag: string;

  public async onInit(): Promise<void> {
    // if display / edit, get item
    switch (this.displayMode) {
      case (FormDisplayMode.Display):
      case (FormDisplayMode.Edit):
        const { currentMission, currentMissionEtag } = await this._getItem(this.context.itemId ?? 0);
        this._currentMission = currentMission;
        this._currentMissionEtag = currentMissionEtag;
        break;
    }

    return;
  }

  public render(): void {
    // if display / edit, get item
    switch (this.displayMode) {
      case (FormDisplayMode.Display):
        ReactDOM.render(React.createElement(NasaMissionPlannerViewForm, {
          currentMission: this._currentMission,
          onClose: this._onClose
        } as INasaMissionPlannerViewFormProps), this.domElement);
        break;
      case (FormDisplayMode.Edit):
        // use same control for EDIT & NEW forms
      case (FormDisplayMode.New): // eslint-disable-line no-fallthrough
        ReactDOM.render(React.createElement(NasaMissionPlannerForms, {
          displayMode: this.displayMode,
          currentMission: this._currentMission,
          onSave: this._onSave,
          onClose: this._onClose
        } as INasaMissionPlannerFormsProps), this.domElement);
        break;
    }
  }

  public onDispose(): void {
    // This method should be used to free any resources that were allocated during rendering.
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private _onSave = async (mission: IMissionPlan): Promise<void> => {
    switch (this.displayMode) {
      case (FormDisplayMode.New):
        await this._createItem(mission);
        break;
      case (FormDisplayMode.Edit):
        await this._updateItem(mission, this.context.itemId ?? 0, this._currentMissionEtag);
        break;
    }
    this.formSaved();
  }

  private _onClose = (): void => {
    this.formClosed();
  }

  /**
   * Flag indicating if just the Title field for the list is used, or if unique
   * columns are specified for each of the fields the form uses.
   *
   * @readonly
   * @private
   * @memberof NasaMissionPlannerFormsFormCustomizer
   */
  private get _uniqueColumnsSpecified(): boolean {
    return (
      this.properties.columnMissionTitleInternalName !== undefined
      && this.properties.columnMissionProjectInternalName !== undefined
      && this.properties.columnLaunchVehicleInternalName !== undefined
      && this.properties.columnLaunchSpacecraftInternalName !== undefined
    );
  }

  /**
   * Fetch the specified item from the SharePoint list.
   *
   * @private
   * @param {number} listItemId
   * @return {*}  {Promise<{ currentMission: IMissionPlan, currentMissionEtag: string }>}
   * @memberof NasaMissionPlannerFormsFormCustomizer
   */
  private async _getItem(listItemId: number): Promise<{ currentMission: IMissionPlan, currentMissionEtag: string }> {
    const responseRaw: SPHttpClientResponse = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this.context.list.guid}')/items(${listItemId})`,
      SPHttpClient.configurations.v1,
      { headers: { 'ACCEPT': 'application/json;odata.metadata=none' } }
    );
    const responseJson: { Title: string, [property: string]: string } = await responseRaw.json();

    const resultArray: string[] = responseJson.Title.split('__');
    // get mission either from single field, or seperate fields as
    //  specified in public properties on this extension
    const mission: IMissionPlan = (this._uniqueColumnsSpecified)
      ? {
        title: responseJson[this.properties.columnMissionTitleInternalName ?? ""],
        project: responseJson[this.properties.columnMissionProjectInternalName ?? ""],
        launchVehicle: responseJson[this.properties.columnLaunchVehicleInternalName ?? ""],
        spacecraft: (responseJson[this.properties.columnLaunchSpacecraftInternalName ?? ""].indexOf(',') > -1)
          ? responseJson[this.properties.columnLaunchSpacecraftInternalName ?? ""].split(',')
          : [responseJson[this.properties.columnLaunchSpacecraftInternalName ?? ""]]
      } as IMissionPlan
      : {
        title: resultArray[0],
        project: resultArray[1],
        launchVehicle: resultArray[2],
        spacecraft: (resultArray[3].indexOf(',') >= 0)
          ? resultArray[3].split(',')
          : [resultArray[3]]
      } as IMissionPlan;

    return {
      currentMission: mission,
      currentMissionEtag: responseRaw.headers.get('ETag') ?? ""
    };
  }

  /**
   * Creates the payload object to save to the SharePoint list,
   * depending if specific columns are specified.
   *
   * @private
   * @memberof NasaMissionPlannerFormsFormCustomizer
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _missionSpListItem(mission: IMissionPlan): any {
    return (this._uniqueColumnsSpecified)
      ? {
        [this.properties.columnMissionTitleInternalName ?? ""]: mission.title,
        [this.properties.columnMissionProjectInternalName ?? ""]: mission.project,
        [this.properties.columnLaunchVehicleInternalName ?? ""]: mission.launchVehicle,
        [this.properties.columnLaunchSpacecraftInternalName ?? ""]: mission.spacecraft?.join(',')
      }
      : { Title: `${mission.title}__${mission.project}__${mission.launchVehicle}__${mission.spacecraft?.join(",")}` }
  }

  /**
   * Save the item in the SharePoint list.
   *
   * @private
   * @param {IMissionPlan} mission
   * @return {*}  {Promise<SPHttpClientResponse>}
   * @memberof NasaMissionPlannerFormsFormCustomizer
   */
  private async _createItem(mission: IMissionPlan): Promise<SPHttpClientResponse> {
    return this.context.spHttpClient.post(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this.context.list.guid}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: { 'CONTENT-TYPE': 'application/json;odata.metadata=none' },
        body: JSON.stringify(this._missionSpListItem(mission))
      }
    );
  }

  /**
   * Update the specified item in the SharePoint list.
   *
   * @private
   * @param {IMissionPlan} mission
   * @param {number} listItemId
   * @param {string} listItemETag
   * @return {*}  {Promise<SPHttpClientResponse>}
   * @memberof NasaMissionPlannerFormsFormCustomizer
   */
  private async _updateItem(mission: IMissionPlan, listItemId: number, listItemETag: string): Promise<SPHttpClientResponse> {
    return this.context.spHttpClient.post(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this.context.list.guid}')/items(${listItemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'CONTENT-TYPE': 'application/json;odata.metadata=none',
          'IF-MATCH': listItemETag,
          'X-HTTP-METHOD': 'MERGE'
        },
        body: JSON.stringify(this._missionSpListItem(mission))
      }
    );

  }
}
