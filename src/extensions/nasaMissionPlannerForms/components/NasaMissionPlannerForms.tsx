import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

import styles from './NasaMissionPlannerForms.module.scss';

export interface INasaMissionPlannerFormsProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

const LOG_SOURCE: string = 'NasaMissionPlannerForms';

export default class NasaMissionPlannerForms extends React.Component<INasaMissionPlannerFormsProps, {}> {
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: NasaMissionPlannerForms mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: NasaMissionPlannerForms unmounted');
  }

  public render(): React.ReactElement<{}> {
    return <div className={styles.nasaMissionPlannerForms} />;
  }
}
