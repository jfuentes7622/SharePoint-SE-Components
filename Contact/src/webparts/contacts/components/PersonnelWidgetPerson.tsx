import * as React from 'react';
import styles from './Contacts.module.scss';
import { PersonObjectModel } from './shared/PersonObject';
// import PersonnelWidgetPersonImage from './PersonnelWidgetPersonImage';
import ImageDisplay from './shared/ImageDisplay';
import IFrameDialog from './shared/IFrameDialog';

export interface IPersonnelWidgetPersonProps {
    personObject: PersonObjectModel;
}

export default class PersonnelWidgetPerson extends React.Component<IPersonnelWidgetPersonProps, {}> {
    public render(): React.ReactElement<IPersonnelWidgetPersonProps> {
        return (
            <div className={'ms-Grid ' + styles.p1}>
                <div className='ms-Grid-row'>
                    <div className={'ms-Grid-col ms-sm12 ' + styles.flexRow}>
                        {(
                            this.props.personObject.displayPhoto &&
                            // <PersonnelWidgetPersonImage personObject={this.props.personObject}></PersonnelWidgetPersonImage>
                            <ImageDisplay
                                imageDisplayLink={this.props.personObject.imageLink}
                                imageDisplayWidth={this.props.personObject.imageWidth as number}
                                imageShape={this.props.personObject.imageShape}
                                imageIsDynamicHeight={true}
                            />
                        )}
                        <div className={'ms-Grid ' + styles.ml2 + ' ' + styles.p3 + ' ' + styles.personnelInfoArea}>
                            <div className='ms-Grid-row'>
                                <div className={'ms-Grid-col ms-sm12 ' + styles.personnelTitle}>
                                    {this.props.personObject.Title}
                                </div>
                            </div>
                            <div className='ms-Grid-row'>
                                <div className={'ms-Grid-col ms-sm12 ' + styles.personnelName}>
                                    {this.props.personObject.name}
                                </div>
                            </div>
                            {
                                this.props.personObject.eMail &&
                                <div className='ms-Grid-row'>
                                    <div className={'ms-Grid-col ms-sm12 ' + styles.personnelEmail}>
                                        <a className={styles.personnelEmail} href={'mailto:' + this.props.personObject.eMail}>
                                            {this.props.personObject.eMail}
                                        </a>
                                    </div>
                                </div>
                            }
                            {
                                this.props.personObject.phoneNumber &&
                                <div className='ms-Grid-row'>
                                    <div className={'ms-Grid-col ms-sm12 ' + styles.personnelPhone}>
                                        {this.props.personObject.phoneNumber}
                                    </div>
                                </div>
                            }
                            {
                                this.props.personObject.sVoip &&
                                <div className='ms-Grid-row'>
                                    <div className={'ms-Grid-col ms-sm12 ' + styles.personnelPhone}>
                                        {this.props.personObject.sVoip}
                                    </div>
                                </div>
                            }

                            {
                                this.props.personObject.bioLink &&
                                <div className='ms-Grid-row'>
                                    <div className={'ms-Grid-col ms-sm12 ' + styles.personnelBioLink}>
                                        <span className={styles.personnelBioLink} role="button" onClick={this.bioClicked.bind(this)}>
                                            Read Bio
                                        </span>
                                    </div>
                                </div>
                            }
                        </div>
                    </div>
                </div>
            </div>
        );
    }

    private bioClicked(): void {
        const windowWidth: number = Math.round(document.body.getBoundingClientRect().width * .80);
        const windowHeight: number = Math.round(document.body.getBoundingClientRect().height * .80);
        (new IFrameDialog(this.props.personObject.name + ' Biography', this.props.personObject.bioLink, windowWidth, windowHeight)).show();
    }
}