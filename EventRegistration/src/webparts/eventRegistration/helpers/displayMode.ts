export type EventRegistrationDisplayMode = 'combined' | 'catalog' | 'registration' | 'organizer';

export function getDisplayMode(value: string): EventRegistrationDisplayMode {
  return value === 'catalog' || value === 'registration' || value === 'organizer' ? value : 'combined';
}

export function getRegistrationPageUrl(pageUrl: string, currentUrl: string, eventId: number): string {
  var current = new URL(currentUrl);
  var target = new URL(pageUrl, currentUrl);
  if ((target.protocol !== 'https:' && target.protocol !== 'http:') || target.origin !== current.origin) {
    throw new Error('The registration page URL must be an HTTP(S) page on the current SharePoint origin.');
  }
  target.search = target.search.replace(/^\?/, '').split('&').filter(function(part: string): boolean {
    var key = decodeURIComponent(part.split('=')[0]).toLowerCase();
    return key !== 'eventid' && key !== 'sessionid';
  }).join('&');
  target.searchParams.set('eventid', String(eventId));
  return target.href;
}
