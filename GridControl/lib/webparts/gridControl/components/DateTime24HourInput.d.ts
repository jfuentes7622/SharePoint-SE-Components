/// <reference types="react" />
import * as React from 'react';
export interface DateTime24HourInputProps {
    value: string;
    timeOnly: boolean;
    minuteIncrement: number;
    ariaLabel: string;
    onChange: (value: string) => void;
}
export declare const DateTime24HourInput: React.StatelessComponent<DateTime24HourInputProps>;
