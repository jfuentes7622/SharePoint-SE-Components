import { assert } from 'chai';
import EventRegistrationService from '../services/EventRegistrationService';
import { listSchemas, listTitle } from '../services/schema';
import { ISetupResult } from '../models';
import EventRegistration from '../components/EventRegistration';

describe('Event Registration field provisioning', function(): void {
  var service: any;
  var result: ISetupResult;
  var fieldsByList: { [path: string]: { [name: string]: any } };
  var posts: Array<{ url: string; body: any }>;
  var failure: any;
  var createdName: string;
  var persistFields: boolean;
  var reads: string[];
  var updates: Array<{ url: string; body: any }>;
  var ignoreIndex: boolean;
  var ignoreRelationship: boolean;

  function emptyResult(): ISetupResult {
    return { createdLists: [], createdFields: [], validatedLists: [], errors: [] };
  }

  beforeEach(function(): void {
    service = new EventRegistrationService('https://example.test', undefined, "Organizer's Events", false);
    result = emptyResult();
    fieldsByList = {};
    posts = [];
    failure = undefined;
    createdName = '';
    persistFields = true;
    reads = [];
    updates = [];
    ignoreIndex = false;
    ignoreRelationship = false;
    listSchemas.forEach(function(schema, index: number): void {
      var path = "web/lists/getByTitle('" + listTitle("Organizer's Events", schema.key).replace(/'/g, "''") + "')";
      fieldsByList[path] = {};
      service.listIds[schema.key] = '00000000-0000-0000-0000-00000000000' + String(index + 1);
    });
    service.rest = {
      escape: function(value: string): string { return value.replace(/'/g, "''"); },
      rows: function(data: any): any[] { return data.value; },
      item: function(data: any): any { return data.d || data; },
      errorMessage: function(error: any): string { return error.message; },
      get: function(url: string): Promise<any> {
        reads.push(url);
        if (url.indexOf('web/lists?') === 0) {
          return Promise.resolve({ data: { value: listSchemas.map(function(schema): any {
            return { Title: listTitle("Organizer's Events", schema.key) };
          }) } });
        }
        if (url.indexOf('/fields/getByInternalNameOrTitle(') >= 0) {
          var lookupPath = url.split('/fields/')[0];
          var lookupName = /getByInternalNameOrTitle\('([^']+)'\)/.exec(url)[1];
          var lookup = fieldsByList[lookupPath][lookupName];
          assert.isDefined(lookup, 'Lookup metadata may be read only after the field exists');
          return Promise.resolve({ data: JSON.parse(JSON.stringify(lookup)) });
        }
        assert.notInclude(url, 'getByInternalNameOrTitle');
        assert.include(url, '/fields?$select=InternalName,TypeAsString,Required&$filter=');
        if (failure) {
          return Promise.reject(failure);
        }
        var path = url.split('/fields?')[0];
        var matches = decodeURIComponent(url).match(/InternalName eq '[^']+'/g);
        assert.include(url, '&$top=' + String(matches.length));
        var fields = matches.map(function(match: string): any {
          return fieldsByList[path][/InternalName eq '([^']+)'/.exec(match)[1]];
        }).filter(function(field: any): boolean { return !!field; });
        return Promise.resolve({ data: { value: fields } });
      },
      post: function(url: string, body: any, headers: any): Promise<any> {
        if (url.indexOf('/fields/getByInternalNameOrTitle(') >= 0) {
          updates.push({ url: url, body: body });
          assert.equal(headers['X-HTTP-Method'], 'MERGE');
          assert.equal(headers['IF-MATCH'], '*');
          assert.equal(body.__metadata.type, 'SP.FieldLookup');
          var lookupName = /getByInternalNameOrTitle\('([^']+)'\)/.exec(url)[1];
          var lookup = fieldsByList[url.split('/fields/')[0]][lookupName];
          if (body.Indexed === true && !ignoreIndex) {
            lookup.Indexed = true;
          }
          if (body.RelationshipDeleteBehavior !== undefined) {
            if (lookup.Indexed !== true) {
              return Promise.reject(new Error('This lookup field cannot enforce a relationship behavior because it is not indexed.'));
            }
            if (!ignoreRelationship) {
              lookup.RelationshipDeleteBehavior = body.RelationshipDeleteBehavior;
            }
          }
          return Promise.resolve({ data: {} });
        }
        posts.push({ url: url, body: body });
        assert.include(url, '/fields/CreateFieldAsXml');
        assert.equal(body.parameters.__metadata.type, 'SP.XmlSchemaFieldCreationInformation');
        var xml = body.parameters.SchemaXml;
        var name = createdName || (body.parameters.Options & 8
          ? / Name="([^"]+)"/.exec(xml)[1] : / DisplayName="([^"]+)"/.exec(xml)[1]);
        var type = / Type="([^"]+)"/.exec(xml)[1];
        var indexed = xml.indexOf('Indexed="TRUE"') >= 0;
        var relationship = xml.indexOf('RelationshipDeleteBehavior="Restrict"') >= 0 ? 2 : 0;
        if (type === 'Lookup' && relationship === 2 && !indexed) {
          return Promise.reject(new Error('This lookup field cannot enforce a relationship behavior because it is not indexed.'));
        }
        var stored = {
          InternalName: name, TypeAsString: type, Required: xml.indexOf('Required="TRUE"') >= 0,
          Indexed: indexed, RelationshipDeleteBehavior: relationship
        };
        if (persistFields) {
          fieldsByList[url.split('/fields/')[0]][name] = stored;
        }
        assert.notInclude(xml, '{LIST_ID}');
        return Promise.resolve({ data: stored });
      }
    };
  });

  it('creates every missing field across the six existing lists and reuses them on retry', function(): Promise<void> {
    var expectedCount = listSchemas.reduce(function(count: number, schema): number {
      return count + schema.fields.length;
    }, 0);
    return service.provisionFields(0, result).then(function(): Promise<void> {
      assert.deepEqual(result.errors, []);
      assert.lengthOf(result.createdFields, expectedCount);
      assert.lengthOf(posts, expectedCount);
      assert.equal(posts[0].body.parameters.Options, 8);
      assert.include(posts[0].url, "getByTitle('Organizer''s Events Events')");
      var lookups = posts.filter(function(post): boolean {
        return post.body.parameters.SchemaXml.indexOf('Type="Lookup"') >= 0;
      });
      assert.isAbove(lookups.length, 0);
      assert.include(lookups[0].body.parameters.SchemaXml, 'List="{00000000-0000-0000-0000-000000000001}"');
      lookups.forEach(function(post): void {
        assert.include(post.body.parameters.SchemaXml, 'Indexed="TRUE"');
        assert.notInclude(post.body.parameters.SchemaXml, 'RelationshipDeleteBehavior');
      });
      assert.lengthOf(updates, lookups.length);
      updates.forEach(function(update): void { assert.equal(update.body.RelationshipDeleteBehavior, 2); });
      var updateCount = updates.length;
      result = emptyResult();
      return service.provisionFields(0, result).then(function(): void {
        assert.lengthOf(updates, updateCount);
      });
    }).then(function(): void {
      assert.deepEqual(result.errors, []);
      assert.lengthOf(result.createdFields, 0);
      assert.lengthOf(posts, expectedCount);
    });
  });

  it('repairs an existing unindexed lookup before enabling Restrict and is idempotent', function(): Promise<void> {
    var schema = { key: 'sessions', fields: [listSchemas[1].fields[0]] };
    var lookup = {
      InternalName: 'EREvent', TypeAsString: 'Lookup', Required: true, Indexed: false, RelationshipDeleteBehavior: 0
    };
    fieldsByList[Object.keys(fieldsByList)[1]].EREvent = lookup;
    return service.ensureFields(schema, 0, result).then(function(): Promise<void> {
      assert.lengthOf(posts, 0);
      assert.lengthOf(updates, 2);
      assert.isTrue(updates[0].body.Indexed);
      assert.equal(updates[1].body.RelationshipDeleteBehavior, 2);
      assert.isTrue(lookup.Indexed);
      assert.equal(lookup.RelationshipDeleteBehavior, 2);
      return service.ensureFields(schema, 0, result);
    }).then(function(): void {
      assert.lengthOf(updates, 2);
    });
  });

  it('does not enable Restrict when an index update fails to persist', function(): Promise<void> {
    ignoreIndex = true;
    fieldsByList[Object.keys(fieldsByList)[1]].EREvent = {
      InternalName: 'EREvent', TypeAsString: 'Lookup', Required: true, Indexed: false, RelationshipDeleteBehavior: 0
    };
    return service.ensureFields({ key: 'sessions', fields: [listSchemas[1].fields[0]] }, 0, result)
      .then(function(): void {
        assert.fail('Expected index verification failure');
      }, function(error: Error): void {
        assert.include(error.message, 'could not be indexed');
        assert.lengthOf(updates, 1);
        assert.isUndefined(updates[0].body.RelationshipDeleteBehavior);
      });
  });

  it('does not claim success when Restrict fails to persist', function(): Promise<void> {
    ignoreRelationship = true;
    fieldsByList[Object.keys(fieldsByList)[1]].EREvent = {
      InternalName: 'EREvent', TypeAsString: 'Lookup', Required: true, Indexed: true, RelationshipDeleteBehavior: 0
    };
    return service.ensureFields({ key: 'sessions', fields: [listSchemas[1].fields[0]] }, 0, result)
      .then(function(): void {
        assert.fail('Expected relationship verification failure');
      }, function(error: Error): void {
        assert.include(error.message, 'did not retain its index and Restrict relationship');
        assert.lengthOf(result.createdFields, 0);
      });
  });

  it('does not silently replace an existing Cascade relationship', function(): Promise<void> {
    fieldsByList[Object.keys(fieldsByList)[1]].EREvent = {
      InternalName: 'EREvent', TypeAsString: 'Lookup', Required: true, Indexed: true, RelationshipDeleteBehavior: 1
    };
    return service.ensureFields({ key: 'sessions', fields: [listSchemas[1].fields[0]] }, 0, result)
      .then(function(): void {
        assert.fail('Expected incompatible relationship failure');
      }, function(error: Error): void {
        assert.include(error.message, 'incompatible delete behavior');
        assert.lengthOf(updates, 0);
      });
  });

  it('requires lookup indexes and Restrict during startup readiness checks', function(): Promise<void> {
    return service.provisionFields(0, result).then(function(): Promise<boolean> {
      assert.deepEqual(result.errors, []);
      fieldsByList[Object.keys(fieldsByList)[1]].EREvent.Indexed = false;
      return service.hasRequiredSchema();
    }).then(function(ready: boolean): Promise<boolean> {
      assert.isFalse(ready);
      var lookup = fieldsByList[Object.keys(fieldsByList)[1]].EREvent;
      lookup.Indexed = true;
      lookup.RelationshipDeleteBehavior = 0;
      return service.hasRequiredSchema();
    }).then(function(ready: boolean): Promise<void> {
      assert.isFalse(ready);
      return service.provisionFields(0, emptyResult());
    }).then(function(): Promise<boolean> {
      return service.hasRequiredSchema();
    }).then(function(ready: boolean): void {
      assert.isTrue(ready);
    });
  });

  it('reports genuine HTTP 400 errors without treating them as missing columns', function(): Promise<void> {
    failure = { status: 400, message: 'Invalid field collection query' };
    return service.provisionFields(0, result).then(function(): void {
      assert.lengthOf(result.errors, 6);
      assert.include(result.errors[0], 'Invalid field collection query');
      assert.lengthOf(posts, 0);
    });
  });

  it('keeps permission failures visible without attempting field creation', function(): Promise<void> {
    failure = { status: 403, message: 'Access denied' };
    return service.provisionFields(0, result).then(function(): void {
      assert.lengthOf(result.errors, 6);
      assert.include(result.errors[0], 'Access denied');
      assert.lengthOf(posts, 0);
    });
  });

  it('does not convert a missing list into an attempt to create its fields', function(): Promise<void> {
    failure = { status: 404, message: 'List not found' };
    return service.provisionFields(0, result).then(function(): void {
      assert.lengthOf(result.errors, 6);
      assert.include(result.errors[0], 'List not found');
      assert.lengthOf(posts, 0);
    });
  });

  it('preserves incompatible existing fields and reports their schema mismatch', function(): Promise<void> {
    var path = Object.keys(fieldsByList)[0];
    fieldsByList[path].ERDescription = { InternalName: 'ERDescription', TypeAsString: 'Text', Required: false };
    return service.ensureFields(listSchemas[0], 0, result).then(function(): void {
      assert.fail('Expected incompatible field failure');
    }, function(error: Error): void {
      assert.include(error.message, 'expected Note');
      assert.lengthOf(posts, 0);
    });
  });

  it('retains required-field validation without modifying existing columns', function(): Promise<void> {
    var path = Object.keys(fieldsByList)[0];
    var organizer = listSchemas[0].fields[2];
    fieldsByList[path].EROrganizer = { InternalName: 'EROrganizer', TypeAsString: 'User', Required: false };
    return service.ensureFields({ key: 'events', fields: [organizer] }, 0, result).then(function(): void {
      assert.fail('Expected required-field failure');
    }, function(error: Error): void {
      assert.include(error.message, 'not required');
      assert.lengthOf(posts, 0);
    });
  });

  it('repairs missing ER fields without deleting earlier display-named columns', function(): Promise<void> {
    var path = Object.keys(fieldsByList)[0];
    var earlierField = { InternalName: 'Description', TypeAsString: 'Note', Required: false };
    fieldsByList[path].Description = earlierField;
    return service.provisionFields(0, result).then(function(): Promise<boolean> {
      assert.deepEqual(result.errors, []);
      assert.strictEqual(fieldsByList[path].Description, earlierField);
      assert.equal(fieldsByList[path].ERDescription.InternalName, 'ERDescription');
      return service.hasRequiredSchema();
    }).then(function(ready: boolean): void {
      assert.isTrue(ready);
    });
  });

  it('detects an unexpected stored internal name instead of counting the field as created', function(): Promise<void> {
    createdName = 'Description';
    return service.provisionFields(0, result).then(function(): void {
      assert.lengthOf(result.errors, 6);
      assert.include(result.errors[0], 'required internal name');
      assert.lengthOf(result.createdFields, 0);
    });
  });

  it('verifies persisted fields instead of trusting a successful creation response', function(): Promise<void> {
    persistFields = false;
    return service.provisionFields(0, result).then(function(): void {
      assert.lengthOf(result.errors, 6);
      assert.include(result.errors[0], 'required internal name');
      assert.lengthOf(result.createdFields, 0);
    });
  });

  it('checks schema readiness even when all six lists exist', function(): Promise<void> {
    return service.hasRequiredLists().then(function(exists: boolean): Promise<boolean> {
      assert.isTrue(exists);
      return service.hasRequiredSchema();
    }).then(function(ready: boolean): void {
      assert.isFalse(ready);
      assert.lengthOf(reads.filter(function(url: string): boolean { return url.indexOf('/fields?') >= 0; }), 6);
    });
  });

  it('becomes ready only after every required field has been persisted and validated', function(): Promise<void> {
    return service.provisionFields(0, result).then(function(): Promise<boolean> {
      assert.deepEqual(result.errors, []);
      return service.hasRequiredSchema();
    }).then(function(ready: boolean): Promise<boolean> {
      assert.isTrue(ready);
      var path = Object.keys(fieldsByList)[0];
      delete fieldsByList[path].ERDescription;
      return service.hasRequiredSchema();
    }).then(function(ready: boolean): void {
      assert.isFalse(ready);
    });
  });

  it('does not silently accept incompatible field metadata during readiness checks', function(): Promise<void> {
    return service.provisionFields(0, result).then(function(): Promise<boolean> {
      fieldsByList[Object.keys(fieldsByList)[0]].ERDescription.TypeAsString = 'Text';
      return service.hasRequiredSchema();
    }).then(function(): void {
      assert.fail('Expected schema validation failure');
    }, function(error: Error): void {
      assert.include(error.message, 'expected Note');
    });
  });

  it('shows setup without loading event items when the lists exist but their fields do not', function(): Promise<void> {
    var component: any = Object.create(EventRegistration.prototype);
    component.props = { organizerGroupId: 0 };
    component.state = {};
    component.setState = function(update: any): void {
      Object.keys(update).forEach(function(key: string): void { component.state[key] = update[key]; });
    };
    var reloadCalls = 0;
    component.reload = function(): Promise<void> { reloadCalls += 1; return Promise.resolve(); };
    service.getCurrentUser = function(): Promise<any> { return Promise.resolve({ id: 1 }); };
    service.isOrganizer = function(): Promise<boolean> { return Promise.resolve(true); };
    component.service = service;
    component.initialize();
    return new Promise<void>(function(resolve, reject): void {
      setTimeout(function(): void {
        try {
          assert.equal(reloadCalls, 0);
          assert.isFalse(component.state.initialized);
          assert.isFalse(component.state.loading);
          assert.equal(component.state.error, '');
          resolve();
        } catch (error) {
          reject(error);
        }
      }, 0);
    });
  });

  it('reports field creation failures instead of claiming provisioning succeeded', function(): Promise<void> {
    service.rest.post = function(): Promise<any> { return Promise.reject(new Error('CreateFieldAsXml failed')); };
    return service.provisionFields(0, result).then(function(): void {
      assert.lengthOf(result.errors, 6);
      assert.include(result.errors[0], 'CreateFieldAsXml failed');
      assert.lengthOf(result.createdFields, 0);
    });
  });
});
