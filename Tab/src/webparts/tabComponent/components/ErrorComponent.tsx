import * as React from 'react';
import styles from './TabComponent.module.scss';

export interface IErrorComponentProps {
    ErrorStr: string;
}

export default class ErrorComponent extends React.Component<IErrorComponentProps, {}> {
    public render(): React.ReactElement<IErrorComponentProps> {
        return (
            <div className={styles.errorComponent}>
                <div className={styles.aligned}>
                    <i className={'ms-Icon ms-Icon--Error ' + styles.icon}></i>
                </div>
                <div className={styles.aligned + ' ' + styles.ml7}>
                    <span>
                        {this.props.ErrorStr}
                    </span>
                </div>
            </div>
        );
    }
}
