import { _SPFlight } from '@microsoft/sp-core-library';
var SPLoaderFlights = (function () {
    function SPLoaderFlights() {
    }
    SPLoaderFlights._useNewBootSequence = function () {
        return _SPFlight.isEnabled(182 );
    };
    return SPLoaderFlights;
}());
export default SPLoaderFlights;
