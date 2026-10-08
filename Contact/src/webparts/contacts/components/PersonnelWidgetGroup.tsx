import * as React from 'react';

import styles from './Contacts.module.scss';
import PersonnelWidgetPerson from './PersonnelWidgetPerson';
import { PersonGroupModel } from './shared/PersonGroup';

export interface IPersonnelWidgetGroupProps {
    PersonGroup: PersonGroupModel;
    contactsPerRow: number;
}

export default class PersonnelWidgetGroup extends React.Component<IPersonnelWidgetGroupProps, {}> {

    public render(): React.ReactElement<IPersonnelWidgetGroupProps> {
        const contactsPerRow: number = Math.max(1, Math.min(4, Number(this.props.contactsPerRow) || 1));
        const itemClassName: string = styles['contactGridItem' + contactsPerRow];
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
                        <div className={styles.personnelFrame + ' ' + styles.contactGrid}>
                            {
                                this.props.PersonGroup.personList.map(d => {
                                    return (
                                        <div className={styles.personnelRow + ' ' + itemClassName} id={d.Id} key={d.Id}>
                                            <PersonnelWidgetPerson personObject={d} />
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
