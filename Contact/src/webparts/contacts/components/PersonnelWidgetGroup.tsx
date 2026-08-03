import * as React from 'react';

import styles from './Contacts.module.scss';
import PersonnelWidgetPerson from './PersonnelWidgetPerson';
import { PersonGroupModel } from './shared/PersonGroup';

export interface IPersonnelWidgetGroupProps {
    PersonGroup: PersonGroupModel;
}

export default class PersonnelWidgetGroup extends React.Component<IPersonnelWidgetGroupProps, {}> {

    public render(): React.ReactElement<IPersonnelWidgetGroupProps> {
        return (
            <div className='ms-Grid'>
                <div className='ms-Grid-row'>
                    <div className='ms-Grid-col ms-sm12'>
                        <div className={styles.groupHeading}>
                            {this.props.PersonGroup.groupTitle}
                        </div>
                    </div>
                </div>
                <div className='ms-Grid-row'>
                    <div className='ms-Grid-col ms-sm12'>
                        <div className={styles.personnelFrame}>
                            {
                                this.props.PersonGroup.personList.map(d => {
                                    return (
                                        <div className={'ms-Grid-row ' + styles.personnelRow}>
                                            <div className={'ms-Grid-col ms-sm12'} id={d.Id}>
                                                <PersonnelWidgetPerson personObject={d} key={d.Id} />
                                            </div>
                                        </div>
                                    );
                                })
                            }
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}
