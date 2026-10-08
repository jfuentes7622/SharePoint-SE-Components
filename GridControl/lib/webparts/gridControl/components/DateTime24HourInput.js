"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
function pad(value) {
    return value < 10 ? '0' + String(value) : String(value);
}
exports.DateTime24HourInput = function (props) {
    var text = String(props.value || '');
    var separatorIndex = text.indexOf('T');
    var date = props.timeOnly ? '' : text.substring(0, 10);
    var time = props.timeOnly ? text : (separatorIndex >= 0 ? text.substring(separatorIndex + 1) : '');
    var hour = /^\d{2}:\d{2}/.test(time) ? time.substring(0, 2) : '';
    var minute = /^\d{2}:\d{2}/.test(time) ? time.substring(3, 5) : '';
    var increment = [1, 5, 10, 15].indexOf(Number(props.minuteIncrement)) >= 0 ? Number(props.minuteIncrement) : 5;
    var hours = [];
    var minutes = [];
    var index;
    for (index = 0; index < 24; index += 1) {
        hours.push(pad(index));
    }
    for (index = 0; index < 60; index += increment) {
        minutes.push(pad(index));
    }
    if (minute && minutes.indexOf(minute) < 0) {
        minutes.push(minute);
        minutes.sort();
    }
    var emit = function (nextDate, nextHour, nextMinute) {
        var nextTime = (nextHour || '00') + ':' + (nextMinute || '00');
        props.onChange(props.timeOnly ? nextTime : (nextDate ? nextDate + 'T' + nextTime : ''));
    };
    return (React.createElement("div", { className: "gc-24hour-input" },
        !props.timeOnly && React.createElement("input", { type: "date", value: date, "aria-label": props.ariaLabel + ' date', onChange: function (event) { return emit(event.currentTarget.value, hour, minute); } }),
        React.createElement("select", { value: hour, "aria-label": props.ariaLabel + ' hour (24-hour)', onChange: function (event) { return emit(date, event.currentTarget.value, minute); } },
            React.createElement("option", { value: "" }, "HH"),
            hours.map(function (entry) { return React.createElement("option", { key: entry, value: entry }, entry); })),
        React.createElement("span", { "aria-hidden": "true" }, ":"),
        React.createElement("select", { value: minute, "aria-label": props.ariaLabel + ' minute', onChange: function (event) { return emit(date, hour, event.currentTarget.value); } },
            React.createElement("option", { value: "" }, "mm"),
            minutes.map(function (entry) { return React.createElement("option", { key: entry, value: entry }, entry); }))));
};

//# sourceMappingURL=DateTime24HourInput.js.map
