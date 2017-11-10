import * as React from 'react';
import { Fabric } from 'office-ui-fabric-react';
import { IReactHeaderProps } from './IReactHeaderProps';

import styles from './ReactHeader.module.scss';

export class ReactHeader extends React.Component<IReactHeaderProps, {}> {
  public render(): JSX.Element {
    return (
      <Fabric>
        <div className={styles.app}>
          <div className={`ms-bgColor-themeDark ms-fontColor-white ${styles.header}`}>
            <b>{this.props.description || 'A default header message'}</b>
          </div>
        </div>
      </Fabric>
    );
  }
}