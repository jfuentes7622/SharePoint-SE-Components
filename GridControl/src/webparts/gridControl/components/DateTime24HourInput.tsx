import * as React from 'react';

export interface DateTime24HourInputProps {
  value: string;
  timeOnly: boolean;
  minuteIncrement: number;
  ariaLabel: string;
  onChange: (value: string) => void;
}

function pad(value: number): string {
  return value < 10 ? '0' + String(value) : String(value);
}

export const DateTime24HourInput: React.StatelessComponent<DateTime24HourInputProps> = (props) => {
  var text = String(props.value || '');
  var separatorIndex = text.indexOf('T');
  var date = props.timeOnly ? '' : text.substring(0, 10);
  var time = props.timeOnly ? text : (separatorIndex >= 0 ? text.substring(separatorIndex + 1) : '');
  var hour = /^\d{2}:\d{2}/.test(time) ? time.substring(0, 2) : '';
  var minute = /^\d{2}:\d{2}/.test(time) ? time.substring(3, 5) : '';
  var increment = [1, 5, 10, 15].indexOf(Number(props.minuteIncrement)) >= 0 ? Number(props.minuteIncrement) : 5;
  var hours: string[] = [];
  var minutes: string[] = [];
  var index: number;
  for (index = 0; index < 24; index += 1) { hours.push(pad(index)); }
  for (index = 0; index < 60; index += increment) { minutes.push(pad(index)); }
  if (minute && minutes.indexOf(minute) < 0) { minutes.push(minute); minutes.sort(); }
  var emit = (nextDate: string, nextHour: string, nextMinute: string): void => {
    var nextTime = (nextHour || '00') + ':' + (nextMinute || '00');
    props.onChange(props.timeOnly ? nextTime : (nextDate ? nextDate + 'T' + nextTime : ''));
  };
  return (
    <div className="gc-24hour-input">
      {!props.timeOnly && <input type="date" value={date} aria-label={props.ariaLabel + ' date'} onChange={(event) => emit(event.currentTarget.value, hour, minute)} />}
      <select value={hour} aria-label={props.ariaLabel + ' hour (24-hour)'} onChange={(event) => emit(date, event.currentTarget.value, minute)}>
        <option value="">HH</option>
        {hours.map((entry) => <option key={entry} value={entry}>{entry}</option>)}
      </select>
      <span aria-hidden="true">:</span>
      <select value={minute} aria-label={props.ariaLabel + ' minute'} onChange={(event) => emit(date, hour, event.currentTarget.value)}>
        <option value="">mm</option>
        {minutes.map((entry) => <option key={entry} value={entry}>{entry}</option>)}
      </select>
    </div>
  );
};
