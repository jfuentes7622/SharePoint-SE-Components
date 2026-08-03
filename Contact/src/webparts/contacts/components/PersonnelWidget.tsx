import * as React from 'react';
import styles from './Contacts.module.scss';

import { PersonGroupModel } from './shared/PersonGroup';
import PersonnelWidgetGroup from './PersonnelWidgetGroup';

//import DirDivBranch from '../../../../models/DirDivBranch';
import { Guid } from '@microsoft/sp-core-library';

export interface IPersonnelWidgetProps {  
    PersonGroups: Array<PersonGroupModel>;
}

export default class PersonnelWidget extends React.Component<IPersonnelWidgetProps, {}> {
    public render(): React.ReactElement<IPersonnelWidgetProps> {
        return(
            <div className={styles.ContactImages}>
            {
                this.props.PersonGroups.map( d => {
                    const guid = Guid.newGuid().toString();
                    return (
                        <PersonnelWidgetGroup PersonGroup={d} key={guid}/>
                    );
                })
            }
            </div>
        );
    }
}
