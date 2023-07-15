import * as React from 'react';
import {
  useEffect,
  useState
} from 'react';
import { FormDisplayMode } from '@microsoft/sp-core-library';

import styles from './NasaMissionPlannerForms.module.scss';

import { IMissionPlan } from "../../../models";

import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Stack, IStackTokens } from 'office-ui-fabric-react/lib/Stack';
import { htmlElementProperties } from 'office-ui-fabric-react';

export interface INasaMissionPlannerFormsProps {
  displayMode: FormDisplayMode;
  currentMission?: IMissionPlan;
  onSave: (arg0: IMissionPlan) => void;
  onClose: () => void;
}

interface IMissionPlanOptions {
  project: string;
  launchVehicles: string[];
  spacecraft: string[];
};

const nasaMissionModel: IMissionPlanOptions[] = [
  {
    project: 'Mercury',
    launchVehicles: ['Atlas-D', 'Atlas-LV-3B', 'Little Joe', 'Redstone'],
    spacecraft: ['Mercury capsule']
  },
  {
    project: 'Gemini',
    launchVehicles: ['Atlas-Agena', 'Tital II GLV'],
    spacecraft: ['Gemini capsule']
  },
  {
    project: 'Apollo',
    launchVehicles: ['Saturn I', 'Saturn IB', 'Saturn V'],
    spacecraft: ['Command & service module', 'Apollo Lunar Module']
  }
] as IMissionPlanOptions[];

export const NasaMissionPlannerForms: React.FC<INasaMissionPlannerFormsProps> = (props) => {
  const stackTokens: IStackTokens = { childrenGap: 20 };

  // form field options
  const formProjectDropdownOptions: IDropdownOption[] = nasaMissionModel.map((nasaModel: IMissionPlanOptions) => {
    return { key: nasaModel.project, text: `${nasaModel.project}` }
  });
  const [formLaunchVehicleOptions, setFormLaunchVehicleOptions] = useState<IChoiceGroupOption[]>([]);
  const [formSpacecraftOptions, setFormSpacecraftOptions] = useState<string[]>([]);

  const [missionTitle, setMissionTitle] = useState(props.currentMission?.title);
  const [missionProject, setMissionProject] = useState(props.currentMission?.project);
  const [missionVehicle, setMissionVehicle] = useState(props.currentMission?.launchVehicle);
  const [missionSpacecrafts, setMissionSpacecrafts] = useState(props.currentMission?.spacecraft!=undefined && props.currentMission?.spacecraft.length  > 0 ? props.currentMission?.spacecraft : []);

  // init (only run on component mount)
  useEffect(() => {
    // if edit mode, init the currently selection options
    if (props.displayMode === FormDisplayMode.Edit) {
      // get the current mission target project (used to set other options)
      const selectedProjectConfigOptions: IMissionPlanOptions = nasaMissionModel.filter((model: IMissionPlanOptions) => {
        return (model.project === props.currentMission?.project);
      })[0];

      // set the form options
      const launchVehicleOptions: IChoiceGroupOption[] = selectedProjectConfigOptions.launchVehicles?.map((vehicle: string) => {
        return {
          key: vehicle,
          text: vehicle
        }
      });
      setFormLaunchVehicleOptions(launchVehicleOptions);
      setFormSpacecraftOptions(selectedProjectConfigOptions.spacecraft);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* +-+-+-+-+-+-+-+ FORM EVENT HANDLERS +-+-+-+-+-+-+-+ */

  /**
   * Event handler when the selected NASA project changes.
   *
   * @private
   * @memberof NasaMissionPlannerForms
   */
  const onProjectChange = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption, index?: number): void => {
    // get selected project
    const selectedProject: IMissionPlanOptions = nasaMissionModel.filter((model: IMissionPlanOptions) => {
      return (model.project === option.key);
    })[0];

    // update available launch vehicle options for this project config
    setFormLaunchVehicleOptions(selectedProject.launchVehicles.map((vehicle: string) => {
      return { key: vehicle, text: vehicle };
    }));
    setMissionVehicle(undefined);

    // update avail spacecraft options for this project config
    setFormSpacecraftOptions(selectedProject.spacecraft);
    setMissionSpacecrafts([]);

    // selected mission project
    setMissionProject(selectedProject.project);
  }

  /**
   * Event handler when the selected spacecraft changes.
   *
   * @private
   * @memberof NasaMissionPlannerForms
   */
  const onSpacecraftSelectionChange = (event: React.FormEvent<HTMLElement>, isChecked: boolean, vehicle: string): void => {
    // update the collection of spacecrafts selected
    setMissionSpacecrafts((prevSelection: string[]) => {
      const newSelection: string[] = (isChecked)
        // add checked spacecraft
        ? [...prevSelection, vehicle]
        // remove spacecraft
        : prevSelection.filter((spacecraft: string) => { return (spacecraft !== vehicle); });

      return newSelection;
    });
  }

  /**
   * Event handler to handle clicking the SAVE button.
   */
  const onSaveClick = (): void => {
    props.onSave({
      title: missionTitle,
      project: missionProject,
      launchVehicle: missionVehicle,
      spacecraft: missionSpacecrafts
    } as IMissionPlan);
  }

  /* +-+-+-+-+-+-+-+ FORM ELEMENTS +-+-+-+-+-+-+-+ */

  return (
    <section className={`${styles.nasaMissionPlannerForms}`}>
      <Stack tokens={stackTokens}>
        <TextField
          label='Mission plan title'
          required={true}
          value={missionTitle}
          onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => { setMissionTitle(newValue || undefined); }} />
        <Dropdown
          label="NASA project to plan a mission for:"
          placeholder="Select a NASA project"
          required={true}
          options={formProjectDropdownOptions}
          selectedKey={missionProject}
          onChange={onProjectChange} />

        {missionProject && (
          <Stack>
            <ChoiceGroup
              label="Select the mission's launch vehicle:"
              required={true}
              options={formLaunchVehicleOptions}
              selectedKey={missionVehicle}
              onChange={(event: React.FormEvent<HTMLElement>, option: IChoiceGroupOption) => { setMissionVehicle(option.key) }} />

            <Label>Select the spacecraft invovled in this mission:</Label>
            <Stack tokens={{ childrenGap: 10 }}>
              {formSpacecraftOptions.map((vehicle: string) => (
                <Checkbox
                  key={vehicle}
                  label={vehicle}
                  checked={(missionSpacecrafts.indexOf(vehicle) !== -1)}
                  onChange={(event: React.FormEvent<HTMLElement>, checked) => onSpacecraftSelectionChange(event, !checked, vehicle)} />
              ))}
            </Stack>
          </Stack>
        )}

        <Stack horizontal tokens={stackTokens}>
          <PrimaryButton text="Save" onClick={onSaveClick} />
          <DefaultButton text="Cancel" onClick={props.onClose} />
        </Stack>
      </Stack>
    </section>
  );
}
