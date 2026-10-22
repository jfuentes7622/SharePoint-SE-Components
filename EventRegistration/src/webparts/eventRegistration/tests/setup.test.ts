import { assert } from 'chai';
import EventRegistrationService from '../services/EventRegistrationService';
import EventRegistration from '../components/EventRegistration';
import { listSchemas, listTitle } from '../services/schema';

describe('Event Registration setup detection', function(): void {
  var service: EventRegistrationService;
  var lists: Array<{ Title: string }>;
  var requestedUrl: string;

  beforeEach(function(): void {
    lists = listSchemas.map(function(schema): { Title: string } {
      return { Title: listTitle("Organizer's Events", schema.key) };
    });
    service = new EventRegistrationService('https://example.test', undefined, "Organizer's Events", false);
    (service as any).rest = {
      escape: function(value: string): string { return value.replace(/'/g, "''"); },
      rows: function(data: any): any[] { return data.value; },
      get: function(url: string): Promise<any> {
        requestedUrl = url;
        return Promise.resolve({ data: { value: lists } });
      }
    };
  });

  it('recognizes all required lists and escapes a custom prefix', function(): Promise<void> {
    return service.hasRequiredLists().then(function(configured: boolean): void {
      assert.isTrue(configured);
      assert.include(decodeURIComponent(requestedUrl), "Title eq 'Organizer''s Events Events'");
      assert.include(requestedUrl, '&$top=6');
    });
  });

  it('requires setup when no lists exist', function(): Promise<void> {
    lists = [];
    return service.hasRequiredLists().then(function(configured: boolean): void {
      assert.isFalse(configured);
    });
  });

  it('requires setup when any required list is missing', function(): Promise<void> {
    lists.pop();
    return service.hasRequiredLists().then(function(configured: boolean): void {
      assert.isFalse(configured);
    });
  });

  it('does not turn permission or network failures into missing-list results', function(): Promise<void> {
    var failure = new Error('Access denied');
    (service as any).rest.get = function(): Promise<any> { return Promise.reject(failure); };
    return service.hasRequiredLists().then(function(): void {
      assert.fail('Expected list discovery to fail');
    }, function(error: Error): void {
      assert.strictEqual(error, failure);
    });
  });

  function initializeComponent(configured: boolean, failure?: Error): Promise<any> {
    var component: any = Object.create(EventRegistration.prototype);
    component.props = { organizerGroupId: 0 };
    component.state = { initialized: true, error: 'Previous error' };
    component.setState = function(update: any): void {
      Object.keys(update).forEach(function(key: string): void { component.state[key] = update[key]; });
    };
    component.reloadCalls = 0;
    component.reload = function(): Promise<void> {
      component.reloadCalls += 1;
      return Promise.resolve();
    };
    component.service = {
      getCurrentUser: function(): Promise<any> { return Promise.resolve({ id: 1 }); },
      isOrganizer: function(): Promise<boolean> { return Promise.resolve(true); },
      hasRequiredLists: function(): Promise<boolean> {
        return failure ? Promise.reject(failure) : Promise.resolve(configured);
      }
    };
    component.initialize();
    return new Promise<any>(function(resolve): void {
      setTimeout(function(): void { resolve(component); }, 0);
    });
  }

  it('shows setup without a raw error or data requests when lists are missing', function(): Promise<void> {
    return initializeComponent(false).then(function(component: any): void {
      assert.isFalse(component.state.initialized);
      assert.isFalse(component.state.loading);
      assert.equal(component.state.error, '');
      assert.isTrue(component.state.organizer);
      assert.equal(component.reloadCalls, 0);
    });
  });

  it('loads data when lists exist', function(): Promise<void> {
    return initializeComponent(true).then(function(component: any): void {
      assert.equal(component.reloadCalls, 1);
    });
  });

  it('keeps genuine discovery failures visible', function(): Promise<void> {
    return initializeComponent(false, new Error('Access denied')).then(function(component: any): void {
      assert.isFalse(component.state.loading);
      assert.equal(component.state.error, 'Access denied');
      assert.equal(component.reloadCalls, 0);
    });
  });
});
