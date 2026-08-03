function loadCSS(filename) {
  var link = document.createElement("link");
  link.rel = "stylesheet";
  link.type = "text/css";
  link.href = filename;
  document.head.appendChild(link);
}

// Global variables (or module-scoped variables)
let _spClockList = { value: [] };
let _invalidList = false;
let _spList = null; // Replace with your initial spList value
let _digitalOnly = false; // Replace with your initial digitalOnly value
let _hour12 = false; // Replace with your initial hour12 value
let _displayDay = false; // Replace with your initial displayDay value
let _absoluteUrl = null;
let _spHttpClient = null; // This needs to be the SPHttpClient instance from your SPFx web part
let _settings = {};

// Clock component (moved to separate function)
function renderClock(timeZone, digitalOnly, hour12, displayDay) {
  // Clock rendering logic (from the previous Clock.tsx example)
  let _date = new Date();
  let _formatOptions = {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      second: 'numeric',
      timeZone: timeZone // Use the timezone passed in
  };
  let _dateString;
  let _hoursDegrees;
  let _minutesDegrees;
  let _secondsDegrees;
  let _divStyleHours = {};
  let _divStyleMinutes = {};
  let _divStyleSeconds = {};

  function updateClockValues() {
    _date = new Date();
    _dateString = new Intl.DateTimeFormat('default', _formatOptions).format(_date);
    let hours = parseInt(new Intl.DateTimeFormat('default', { hour: 'numeric', timeZone: _formatOptions.timeZone, hour12: hour12 }).format(_date));
    let minutes = parseInt(new Intl.DateTimeFormat('default', { minute: 'numeric', timeZone: _formatOptions.timeZone }).format(_date));
    let seconds = parseInt(new Intl.DateTimeFormat('default', { second: 'numeric', timeZone: _formatOptions.timeZone }).format(_date));
    _hoursDegrees = hours * 30 + minutes / 2;
    _minutesDegrees = minutes * 6 + seconds / 10;
    _secondsDegrees = seconds * 6;
    _divStyleHours = {
        transform: 'rotateZ(' + _hoursDegrees + 'deg)',
    };
    _divStyleMinutes = {
        transform: 'rotateZ(' + _minutesDegrees + 'deg)',
    };
    _divStyleSeconds = {
        transform: 'rotateZ(' + _secondsDegrees + 'deg)',
    };
  }
  updateClockValues();

  function clockHTML() {
    if (!digitalOnly) {
      return `
      <div class="clockContainer">
        <div class="analogContainer styling-class">
          <div class="analogClock">
            <svg
              class="backgroundNumbers"
              viewBox="0 0 226.6 233.8"
            >
              <path
                class="numbers-class"
                d="M105.5 22.7V6.4h-5.9V4.3c.8 0 1.5-.1 2.2-.2.7-.1 1.4-.3 2-.7.6-.3 1.1-.8 1.5-1.3.4-.6.7-1.3.8-2.1h2.1v22.7h-2.7zM114.1 4.8c.3-1 .8-1.8 1.4-2.5.6-.7 1.4-1.3 2.4-1.7.9-.4 2-.6 3.2-.6 1 0 1.9.1 2.8.4.9.3 1.6.7 2.3 1.2.6.5 1.1 1.2 1.5 2 .4.8.6 1.7.6 2.8 0 1-.2 1.9-.5 2.7s-.7 1.5-1.2 2.1c-.5.6-1.1 1.2-1.8 1.6-.7.5-1.3 1-2 1.4-.7.4-1.4.8-2.1 1.3s-1.3.9-1.9 1.3c-.6.5-1.1 1-1.5 1.5s-.7 1.2-.8 1.9h11.6v2.4h-14.8c.1-1.3.3-2.5.7-3.4.4-.9.8-1.8 1.4-2.5s1.2-1.3 2-1.9c.7-.5 1.5-1 2.3-1.5 1-.6 1.8-1.1 2.5-1.6s1.3-1 1.8-1.5.8-1.1 1.1-1.7.4-1.3.4-2.1c0-.6-.1-1.2-.4-1.7-.2-.5-.5-.9-.9-1.3s-.9-.6-1.4-.8c-.5-.2-1.1-.3-1.7-.3-.8 0-1.5.2-2 .5-.6.3-1 .8-1.4 1.3-.4.5-.6 1.1-.8 1.8s-.2 1.3-.2 2H114c-.3-1-.2-2.1.1-3.1zM166.5 38.2V21.9h-5.9v-2.2c.8 0 1.5-.1 2.2-.2.7-.1 1.4-.3 2-.7.6-.3 1.1-.8 1.5-1.3.4-.6.7-1.3.8-2.1h2.1v22.7h-2.7zM198.9 59.2c.3-1 .8-1.8 1.4-2.5.6-.7 1.4-1.3 2.4-1.7.9-.4 2-.6 3.2-.6 1 0 1.9.1 2.8.4.9.3 1.6.7 2.3 1.2.6.5 1.1 1.2 1.5 2 .4.8.6 1.7.6 2.8 0 1-.2 1.9-.5 2.7s-.7 1.5-1.2 2.1c-.5.6-1.1 1.2-1.8 1.6-.7.5-1.3 1-2 1.4-.7.4-1.4.8-2.1 1.3s-1.3.9-1.9 1.3c-.6.5-1.1 1-1.5 1.5s-.7 1.2-.8 1.9h11.6V77H198c.1-1.3.3-2.5.7-3.4.4-.9.8-1.8 1.4-2.5s1.2-1.3 2-1.9c.7-.5 1.5-1 2.3-1.5 1-.6 1.8-1.1 2.5-1.6s1.3-1 1.8-1.5.8-1.1 1.1-1.7.4-1.3.4-2.1c0-.6-.1-1.2-.4-1.7-.2-.5-.5-.9-.9-1.3s-.9-.6-1.4-.8c-.5-.2-1.1-.3-1.7-.3-.8 0-1.5.2-2 .5-.6.3-1 .8-1.4 1.3-.4.5-.6 1.1-.8 1.8s-.2 1.3-.2 2h-2.7c-.2-1.1-.1-2.1.2-3.1zM217.6 115.1H218.5c.6 0 1.1-.1 1.6-.2s1-.4 1.4-.7c.4-.3.7-.7.9-1.2.2-.5.4-1 .4-1.6 0-1.2-.4-2.1-1.2-2.7-.8-.6-1.7-.9-2.9-.9-.7 0-1.4.1-1.9.4-.5.3-1 .6-1.3 1.1-.4.4-.6 1-.8 1.6-.2.6-.3 1.2-.3 1.9h-2.7c0-1.1.2-2.1.5-3s.8-1.7 1.3-2.3c.6-.6 1.3-1.1 2.2-1.5.9-.4 1.9-.5 3-.5 1 0 1.9.1 2.7.4s1.6.6 2.2 1.1c.6.5 1.1 1.1 1.5 1.9s.5 1.7.5 2.7c0 1-.3 1.9-.9 2.7-.6.8-1.3 1.4-2.2 1.8v.1c1.4.3 2.4 1 3.1 2 .7 1 1 2.2 1 3.6 0 1.1-.2 2.1-.6 3-.4.9-1 1.6-1.7 2.2s-1.5 1-2.5 1.3-2 .4-3 .4c-1.2 0-2.2-.2-3.1-.5-.9-.3-1.7-.8-2.4-1.4-.7-.6-1.2-1.4-1.5-2.3-.4-.9-.5-2-.5-3.1h2.7c0 1.5.5 2.7 1.3 3.6.8.9 2 1.4 3.6 1.4.7 0 1.3-.1 1.9-.3.6-.2 1.1-.5 1.6-.9s.8-.8 1.1-1.4.4-1.2.4-1.8c0-.7-.1-1.3-.4-1.9s-.6-1-1-1.4-.9-.7-1.5-.8-1.2-.3-1.9-.3c-.6 0-1.1 0-1.6.1V115c-.1.1 0 .1.1.1zM214.2 173.8v2.4h-3.1v5.3h-2.6v-5.3h-10v-2.6l10.3-14.8h2.2v15h3.2zm-5.6-11.1l-7.6 11.1h7.6v-11.1zM163.7 199.4l-1.2 6.5.1.1c.5-.6 1.1-1 1.9-1.2.8-.3 1.6-.4 2.3-.4 1 0 2 .2 2.8.5.9.3 1.7.8 2.3 1.5.7.7 1.2 1.5 1.6 2.4s.6 2.1.6 3.4c0 1-.2 1.9-.5 2.8-.3.9-.8 1.7-1.5 2.4s-1.5 1.3-2.5 1.7-2.1.6-3.5.6c-1 0-1.9-.1-2.8-.4s-1.6-.7-2.3-1.2-1.2-1.2-1.6-2c-.4-.8-.6-1.7-.6-2.8h2.7c0 .6.2 1.1.4 1.6s.6.9 1 1.3.9.7 1.5.9c.6.2 1.2.3 1.9.3.6 0 1.3-.1 1.8-.3.6-.2 1.1-.6 1.5-1 .4-.4.8-1 1-1.7.3-.7.4-1.5.4-2.4 0-.7-.1-1.4-.4-2.1s-.6-1.2-1-1.6-1-.8-1.6-1.1-1.3-.4-2.1-.4c-.9 0-1.7.2-2.4.6-.7.4-1.3.9-1.8 1.6l-2.3-.1 2.1-11.8h11.2v2.4h-9zM116.4 214.1c-.7-.6-1.5-.9-2.6-.9-1.2 0-2.1.3-2.8.8s-1.3 1.3-1.6 2.1-.7 1.8-.8 2.8c-.1 1-.2 1.9-.3 2.8l.1.1c.6-1 1.4-1.8 2.4-2.3.9-.5 2-.7 3.3-.7 1.1 0 2.1.2 2.9.6.9.4 1.6.9 2.2 1.6s1 1.4 1.4 2.3c.3.9.5 1.9.5 2.9 0 .8-.1 1.7-.4 2.6-.3.9-.7 1.7-1.3 2.4-.6.7-1.4 1.3-2.3 1.8-1 .5-2.2.7-3.6.7-1.7 0-3-.3-4.1-1s-1.8-1.6-2.4-2.6c-.6-1.1-.9-2.2-1.1-3.5-.2-1.3-.3-2.5-.3-3.7 0-1.6.1-3.1.4-4.5.3-1.5.7-2.8 1.4-3.9.6-1.1 1.5-2 2.6-2.7 1.1-.7 2.4-1 4-1 1.9 0 3.4.5 4.5 1.5s1.7 2.4 1.9 4.3h-2.7c-.2-1.1-.6-1.9-1.3-2.5zm-4.9 7.5c-.6.3-1.1.6-1.5 1.1-.4.5-.7 1-.9 1.6-.2.6-.3 1.3-.3 2s.1 1.4.3 2c.2.6.5 1.2.9 1.6.4.4.9.8 1.5 1.1.6.3 1.3.4 2 .4s1.4-.1 2-.4c.6-.3 1-.6 1.4-1.1.4-.5.7-1 .9-1.6s.3-1.2.3-1.9-.1-1.4-.3-2c-.2-.6-.5-1.2-.8-1.6-.4-.5-.9-.8-1.4-1.1s-1.2-.4-2-.4c-.9-.1-1.5.1-2.1.3zM64.9 203.4c-1 1.6-1.9 3.2-2.7 5-.8 1.8-1.4 3.6-1.9 5.5s-.8 3.7-.9 5.5h-3c.1-1.9.4-3.8.9-5.6.5-1.8 1.1-3.6 1.9-5.2s1.7-3.3 2.7-4.8c1-1.5 2.1-2.9 3.3-4.1H53.5V197h14.7v2.3c-1.2 1.2-2.3 2.5-3.3 4.1zM15.2 162.1c.4-.7.9-1.3 1.5-1.8s1.3-.9 2.1-1.1c.8-.3 1.6-.4 2.5-.4 1.2 0 2.3.2 3.2.5.9.3 1.6.8 2.1 1.3s.9 1.2 1.2 1.9c.3.7.4 1.4.4 2.1 0 1-.3 2-.8 2.8s-1.3 1.5-2.3 1.9c1.4.4 2.4 1.1 3 2.1s1 2.2 1 3.6c0 1.1-.2 2.1-.6 2.9-.4.9-.9 1.6-1.6 2.2s-1.5 1-2.4 1.3-1.9.4-2.9.4c-1.1 0-2.1-.1-3-.4-.9-.3-1.8-.7-2.4-1.3s-1.2-1.3-1.6-2.2c-.4-.9-.6-1.9-.6-3 0-1.3.3-2.5 1-3.5s1.7-1.7 2.9-2.2c-1-.4-1.7-1-2.3-1.9-.6-.9-.9-1.8-.9-2.8-.1-.9.1-1.7.5-2.4zm2.9 16.2c.9.8 2.1 1.2 3.5 1.2.7 0 1.3-.1 1.9-.3.6-.2 1.1-.5 1.5-.9s.7-.9 1-1.4.3-1.1.3-1.8c0-.6-.1-1.2-.4-1.7s-.6-1-1-1.4-.9-.7-1.5-.9c-.6-.2-1.2-.3-1.8-.3-.7 0-1.3.1-1.9.3-.6.2-1.1.5-1.5.9-.4.4-.8.8-1 1.4-.2.5-.4 1.1-.4 1.8-.1 1.2.3 2.3 1.3 3.1zm-.3-12c.2.5.5.8.9 1.1.4.3.8.5 1.3.7.5.1 1 .2 1.6.2 1.1 0 2-.3 2.7-1 .7-.6 1.1-1.5 1.1-2.7s-.4-2-1.1-2.6c-.7-.6-1.6-.9-2.7-.9-.5 0-1 .1-1.5.2s-.9.4-1.3.7c-.4.3-.6.7-.8 1.1-.2.4-.3.9-.3 1.5-.2.7-.1 1.2.1 1.7zM4.5 125.1c.8.6 1.7.9 2.8.9 1.7 0 2.9-.7 3.7-2.2s1.3-3.6 1.4-6.6l-.1-.1c-.5 1-1.2 1.7-2.2 2.3-.9.6-2 .8-3.1.8-1.2 0-2.2-.2-3.1-.6-.9-.4-1.6-.9-2.3-1.6-.6-.7-1.1-1.5-1.4-2.4-.3-.9-.5-2-.5-3.1s.2-2.1.5-3c.4-.9.9-1.7 1.5-2.3.7-.7 1.5-1.2 2.4-1.5.9-.4 1.9-.5 3-.5s2.1.2 3 .5c.9.3 1.8.9 2.5 1.7.7.8 1.3 1.9 1.7 3.3.4 1.4.6 3.2.6 5.3 0 3.9-.6 6.9-1.9 9-1.2 2.1-3.2 3.2-6 3.2-1.9 0-3.5-.5-4.7-1.4s-2-2.4-2.1-4.4h2.7c.4 1.2.9 2.1 1.6 2.7zm7.2-14.2c-.2-.6-.5-1.2-.9-1.6s-.9-.9-1.5-1.1c-.6-.3-1.2-.4-2-.4s-1.5.1-2.1.4-1 .7-1.4 1.2c-.4.5-.6 1.1-.8 1.7-.2.6-.2 1.3-.2 2 0 .6.1 1.2.3 1.8.2.6.5 1.1.9 1.5.4.4.9.8 1.4 1.1s1.1.4 1.8.4 1.3-.1 1.9-.4 1.1-.6 1.5-1.1c.4-.5.7-1 .9-1.6.2-.6.3-1.2.3-1.9.2-.7.1-1.4-.1-2zM13.6 76V59.8H7.8v-2.2c.8 0 1.5-.1 2.2-.2.7-.1 1.4-.3 2-.7.6-.3 1.1-.8 1.5-1.3.4-.6.7-1.3.8-2.1h2.1V76h-2.8zM21.9 62.3c0-.9.1-1.8.3-2.6.2-.9.4-1.7.7-2.4.3-.8.8-1.4 1.3-2 .6-.6 1.3-1 2.1-1.4s1.9-.5 3-.5c1.2 0 2.2.2 3 .5s1.5.8 2.1 1.4c.6.6 1 1.2 1.3 2 .3.8.6 1.6.7 2.4.2.9.3 1.7.3 2.6s.1 1.8.1 2.6 0 1.7-.1 2.6-.1 1.8-.3 2.6c-.2.9-.4 1.7-.7 2.4s-.8 1.4-1.3 2c-.6.6-1.2 1-2.1 1.4s-1.8.5-3 .5-2.2-.2-3-.5-1.5-.8-2.1-1.4c-.6-.6-1-1.2-1.3-2s-.6-1.6-.7-2.4c-.2-.9-.3-1.7-.3-2.6 0-.9-.1-1.8-.1-2.6.1-.8.1-1.7.1-2.6zm2.9 5.4c.1 1.1.2 2 .5 3 .3.9.8 1.7 1.4 2.4s1.5 1 2.7 1c1.2 0 2-.3 2.7-1s1.1-1.4 1.4-2.4c.3-.9.5-1.9.5-3 .1-1.1.1-2 .1-2.9V63c0-.7-.1-1.3-.2-2s-.2-1.3-.4-2c-.2-.6-.4-1.2-.8-1.7s-.8-.9-1.3-1.2-1.2-.4-2-.4-1.4.1-2 .4c-.5.3-1 .7-1.3 1.2-.4.5-.6 1-.8 1.7-.2.6-.3 1.3-.4 2-.1.7-.1 1.3-.2 2v1.8c.1.9.1 1.9.1 2.9z M53.5 38.2V21.9h-5.9v-2.2c.8 0 1.5-.1 2.2-.2.7-.1 1.4-.3 2-.7.6-.3 1.1-.8 1.5-1.3.4-.6.7-1.3.8-2.1h2.1v22.7h-2.7zM69.1 38.2V21.9h-5.9v-2.2c.8 0 1.5-.1 2.2-.2.7-.1 1.4-.3 2-.7.6-.3 1.1-.8 1.5-1.3.4-.6.7-1.3.8-2.1h2.1v22.7h-2.7z"
              />
            </svg>
            <div
              class="indicator hoursIndicator"
              style="transform: rotateZ(${_hoursDegrees}deg);"
            ></div>
            <div
              class="indicator minutesIndicator"
              style="transform: rotateZ(${_minutesDegrees}deg);"
            ></div>
            <div"
              class="indicator secondsIndicator"
              style="transform: rotateZ(${_secondsDegrees}deg);"
            ></div>
            <div class="indicatorCover"></div>
          </div>
        </div>
        <div class="digitalContainer">
        ${_dateString}
        </div>
      </div>
      `;
    } else {
      return `
        <div class="digitalContainer">
          ${_dateString}
        </div>
      `;
    }
  }

  return clockHTML();
}

// Function to fetch data from SharePoint
async function getDataFromSharepoint(listName) {

  if (!_absoluteUrl) {
    console.error("absoluteUrl not initialized.");
    return { value: [] };
  }

  const url = `${_absoluteUrl}/_api/web/lists/GetbyTitle('${listName}')/items?$orderby=SortOrder asc`;

  try {
    const response = await fetch(url, {
      headers: {
        "Accept": "application/json;odata=nometadata"
      }
    });

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const jsonResponse = await response.json();

    if (!jsonResponse.value) {
      return { value: [] };
    }

    const result = jsonResponse.value.map(item => ({
      Title: item.Title,
      Timezone: item.Timezone
    }));

    return { value: result };
  } catch (error) {
    console.error(`Failed to get list: '${listName}' at ${_absoluteUrl}; ${error}`);
    return { value: [] };
  }
}

// Function to check if the clock type is valid
function _isClockType(element) {
  return 'Title' in element && 'Timezone' in element;
}

// Function to update the state with the new clock list
function _updateState(spClockList) {
  if (spClockList && spClockList.value && (spClockList.value.length > 0)) {
      if (_isClockType(spClockList.value[0])) {
          _invalidList = false;
          _spClockList = spClockList;
      } else {
          _invalidList = true;
          _spClockList = { value: [] };
      }
  }
}

async function loadSettings() {
  try {
      if (!_absoluteUrl) {
          console.error("absoluteUrl not initialized.");
          return {};
      }
      // Construct the URL to the settings.json file in the Site Assets library
      const siteRelativeUrl = _absoluteUrl + "/SiteAssets/settings.json";

      const response = await fetch(siteRelativeUrl);

      if (!response.ok) {
          throw new Error(`Failed to load settings.json: ${response.status} ${response.statusText}`);
      }

      const settings = await response.json();
      return settings;
  } catch (error) {
      console.error("Error loading settings.json from Site Assets:", error);
      return {}; // Return an empty object in case of error
  }
}

// Function to update the clock list and DOM
function updateClockList() {
  if (_spList) {
      getDataFromSharepoint(_spList).then((res) => {
          if (res.value.length > 0) {
              _updateState(res);
          } else {
              _invalidList = true;
              _spClockList = { value: [] };
          }
          updateDOM(); // Call updateDOM after getDataFromSharepoint completes
      });
  } else {
      updateDOM();  // Still call updateDOM, even if there's no spList
  }
}

// Function to update the DOM with the clock list
function updateDOM() {
  let worldclock = document.getElementById('worldclock'); 
  
  if (!worldclock) {
      console.error("Element with ID 'worldClockContainer' not found.");
      return;
  }

  let html = '';

  if (_invalidList) {
      html = `
          <div class="invalidList">Invalid World Clock List!</div>
      `;
  } else if (_spClockList.value.length === 0) {
      html = `
          <div class="ms-fontWeight-bold">Please open the property pane and choose a list with locations and timezones.</div>
      `;
  } else {
      html = `
          ${_spClockList.value.map(mapping => `
              <div key="${mapping.Title}" class="clockComponent">
                  <div class="clockTitle">${mapping.Title}</div>
                  ${renderClock(mapping.Timezone, _digitalOnly, _hour12, _displayDay)}
              </div>
          `).join('')}
      `;
  }

  worldclock.innerHTML = html;

  // Update the clock values for each clock instance
  // You might need to use a setInterval for each clock instance if you want them to update independently

  // Select all clock containers
  const clockContainers = worldclock.querySelectorAll('.clockComponent');

  clockContainers.forEach(clockContainer => {
      let _date = new Date();
      let _formatOptions = {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      second: 'numeric',
      timeZone: clockContainer.attributes.key.value // Default Timezone
      };
      let _dateString;
      let _hoursDegrees;
      let _minutesDegrees;
      let _secondsDegrees;
      let _divStyleHours = {};
      let _divStyleMinutes = {};
      let _divStyleSeconds = {};

      function updateClockValues() {
          // Extract timezone from the parent element
          let timezoneElement = clockContainer.querySelector('.clockTitle');
          let timezone = _spClockList.value.find(item => item.Title === timezoneElement.textContent).Timezone;

          if (timezone) {
              _formatOptions.timeZone = timezone;
          }

          _date = new Date();
          _dateString = new Intl.DateTimeFormat('default', _formatOptions).format(_date);
          let hours = parseInt(new Intl.DateTimeFormat('default', { hour: 'numeric', timeZone: _formatOptions.timeZone, hour12: _hour12 }).format(_date));
          let minutes = parseInt(new Intl.DateTimeFormat('default', { minute: 'numeric', timeZone: _formatOptions.timeZone }).format(_date));
          let seconds = parseInt(new Intl.DateTimeFormat('default', { second: 'numeric', timeZone: _formatOptions.timeZone }).format(_date));
          _hoursDegrees = hours * 30 + minutes / 2;
          _minutesDegrees = minutes * 6 + seconds / 10;
          _secondsDegrees = seconds * 6;
          _divStyleHours = {
              transform: 'rotateZ(' + _hoursDegrees + 'deg)',
          };
          _divStyleMinutes = {
              transform: 'rotateZ(' + _minutesDegrees + 'deg)',
          };
          _divStyleSeconds = {
              transform: 'rotateZ(' + _secondsDegrees + 'deg)',
          };
      }

      function updateClockHTML() {
      // Select elements within the clock container
      const hourHand = clockContainer.querySelector('.hoursIndicator');
      const minuteHand = clockContainer.querySelector('.minutesIndicator');
      const secondHand = clockContainer.querySelector('.secondsIndicator');
      const digitalContainer = clockContainer.querySelector('.digitalContainer');

      // Update clock hands and date based on _digitalOnly
      if (!_digitalOnly) {
          if (hourHand) {
              hourHand.style.transform = `rotateZ(${_hoursDegrees}deg)`;
              // hourHand.style.backgroundColor = _hourHandColor;
          }
          if (minuteHand) {
              minuteHand.style.transform = `rotateZ(${_minutesDegrees}deg)`;
              // minuteHand.style.backgroundColor = _minHandColor;
          }
          if (secondHand) {
              secondHand.style.transform = `rotateZ(${_secondsDegrees}deg)`;
              // secondHand.style.backgroundColor = _secondHandColor;
          }
      }

      if (digitalContainer) {
          digitalContainer.textContent = _dateString;
      }
      }

      // Call updateClockValues and updateClockHTML every second
      updateClockValues();
      updateClockHTML(); // Initial render
      setInterval(() => {
          // console.log(_dateString)
          updateClockValues();
          updateClockHTML();
      }, 1000); // 1000 milliseconds = 1 second
  });
}

// Function to initialize the world clock
function initializeWorldClock(settings) {
  _settings = settings; // Store the settings
  _spList = settings.spList;
  _digitalOnly = settings.digitalOnly === 'true'; // Convert to boolean
  _hour12 = settings.hour12 === 'true'; // Convert to boolean
  _displayDay = settings.displayDay === 'true'; // Convert to boolean
  updateClockList(); // Fetch data and render the clock list initially
}

window.addEventListener('load', () => {
  loadCSS("/sites/afcent-a63/siteassets/styles.css");

  fetch('/sites/afcent-a63/siteassets/headerhtml.txt')
       .then(response => response.text())
       .then(data => {
       document.getElementById('suiteBarTop').insertAdjacentHTML("beforebegin",data);

       // IMPORTANT: Set _absoluteUrl to the SharePoint site URL
       _absoluteUrl = _spPageContextInfo.webAbsoluteUrl;

       return loadSettings();
       })
       .then(settings => {
          initializeWorldClock(settings);
       })
       .catch(error => console.error('Error fetching the file:', error));

  fetch('/sites/afcent-a63/siteassets/footerhtml.txt')
       .then(response => response.text())
       .then(data => {
       document.getElementById('s4-workspace').insertAdjacentHTML("beforeend",data);;
       })
       .catch(error => console.error('Error fetching the file:', error));    
});