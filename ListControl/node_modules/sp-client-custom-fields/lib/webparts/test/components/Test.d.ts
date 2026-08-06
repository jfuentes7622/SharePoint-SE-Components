import * as React from 'react';
import { ITestWebPartProps } from '../ITestWebPartProps';
export interface ITestProps extends ITestWebPartProps {
}
export default class Test extends React.Component<ITestProps, {}> {
    render(): JSX.Element;
}
