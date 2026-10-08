import * as React from 'react';

export interface DateTime24HourInputProps {
  value: string;
  displayFormat: 'dateOnly' | 'timeOnly' | 'dateTime';
  minuteIncrement: number;
  disabled?: boolean;
  required?: boolean;
  ariaLabel: string;
  style?: React.CSSProperties;
  onChange: (value: string) => void;
}

function pad(value: number): string {
  return value < 10 ? '0' + String(value) : String(value);
}

function getParts(value: string, displayFormat: string): { date: string; hour: string; minute: string } {
  var text = String(value || '');
  var separatorIndex = text.indexOf('T');
  var date = displayFormat === 'timeOnly' ? '' : text.substring(0, 10);
  var time = displayFormat === 'timeOnly'
    ? text
    : (separatorIndex >= 0 ? text.substring(separatorIndex + 1) : '');
  return {
    date: date,
    hour: /^\d{2}:\d{2}/.test(time) ? time.substring(0, 2) : '',
    minute: /^\d{2}:\d{2}/.test(time) ? time.substring(3, 5) : ''
  };
}

export const DateTime24HourInput: React.StatelessComponent<DateTime24HourInputProps> = (props) => {
  var parts = getParts(props.value, props.displayFormat);
  var increment = [1, 5, 10, 15].indexOf(Number(props.minuteIncrement)) >= 0
    ? Number(props.minuteIncrement) : 5;
  var hours: string[] = [];
  var minutes: string[] = [];
  var index: number;
  for (index = 0; index < 24; index += 1) {
    hours.push(pad(index));
  }
  for (index = 0; index < 60; index += increment) {
    minutes.push(pad(index));
  }
  if (parts.minute && minutes.indexOf(parts.minute) < 0) {
    minutes.push(parts.minute);
    minutes.sort();
  }

  var emit = (date: string, hour: string, minute: string): void => {
    var nextHour = hour || '00';
    var nextMinute = minute || '00';
    if (props.displayFormat === 'timeOnly') {
      props.onChange(nextHour + ':' + nextMinute);
      return;
    }
    props.onChange(date ? date + 'T' + nextHour + ':' + nextMinute : '');
  };
  var dateStyle: React.CSSProperties = Object.assign({}, props.style || {}, {
    flex: '1 1 180px',
    minWidth: '140px'
  });
  var selectStyle: React.CSSProperties = Object.assign({}, props.style || {}, {
    width: 'auto',
    minWidth: '70px'
  });

  if (props.displayFormat === 'dateOnly') {
    return (
      <input
        type="date"
        value={parts.date}
        disabled={props.disabled}
        required={props.required}
        aria-label={props.ariaLabel}
        style={props.style}
        onChange={(event) => props.onChange(event.currentTarget.value)}
      />
    );
  }

  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '8px', width: '100%' }}>
      {props.displayFormat === 'dateTime' && (
        <input
          type="date"
          value={parts.date}
          disabled={props.disabled}
          required={props.required}
          aria-label={props.ariaLabel + ' date'}
          style={dateStyle}
          onChange={(event) => emit(event.currentTarget.value, parts.hour, parts.minute)}
        />
      )}
      <select
        value={parts.hour}
        disabled={props.disabled}
        required={props.required}
        aria-label={props.ariaLabel + ' hour (24-hour)'}
        style={selectStyle}
        onChange={(event) => emit(parts.date, event.currentTarget.value, parts.minute)}
      >
        <option value="">HH</option>
        {hours.map((hour) => <option key={hour} value={hour}>{hour}</option>)}
      </select>
      <span aria-hidden="true">:</span>
      <select
        value={parts.minute}
        disabled={props.disabled}
        required={props.required}
        aria-label={props.ariaLabel + ' minute'}
        style={selectStyle}
        onChange={(event) => emit(parts.date, parts.hour, event.currentTarget.value)}
      >
        <option value="">mm</option>
        {minutes.map((minute) => <option key={minute} value={minute}>{minute}</option>)}
      </select>
    </div>
  );
};
