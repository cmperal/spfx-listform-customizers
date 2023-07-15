import * as React from 'react';
import { IMissionPlan } from "../../../models";

import styles from './NasaMissionPlannerForms.module.scss';

import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Stack, IStackTokens } from 'office-ui-fabric-react/lib/Stack';
import { Text } from 'office-ui-fabric-react/lib/Text';

export interface INasaMissionPlannerViewFormProps {
  currentMission: IMissionPlan;
  onClose: () => void;
}

export const NasaMissionPlannerViewForm: React.FC<INasaMissionPlannerViewFormProps> = (props) => {
    const stackTokens: IStackTokens = { childrenGap: 20 };

    return(
      <section className={`${styles.nasaMissionPlannerForms}`}>
        <Stack tokens={stackTokens}>
          <Label>Mission title:</Label><Text>{props.currentMission.title}</Text>
          <Label>NASA project:</Label><Text>{props.currentMission.project}</Text>
          <Label>Launch vehicle:</Label><Text>{props.currentMission.launchVehicle}</Text>
          <Label>Spacecraft:</Label><Text>{props.currentMission.spacecraft?.toString()}</Text>
          <Stack horizontal tokens={stackTokens}>
            <DefaultButton text="Close" onClick={props.onClose} />
          </Stack>
        </Stack>
      </section>
    );

}
