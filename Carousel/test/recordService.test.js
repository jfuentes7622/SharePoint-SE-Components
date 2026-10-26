'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ts = require('typescript');

const source = fs.readFileSync(path.join(__dirname, '..', 'src', 'services', 'RecordService.ts'), 'utf8');
const compiled = ts.transpileModule(source, {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES5 }
}).outputText;
const context = {
  exports: {},
  console: console,
  require: (name) => {
    assert.strictEqual(name, './SPListService', 'Unexpected service import');
    return { default: function SPListService() {} };
  }
};
vm.runInNewContext(compiled, context, { filename: 'RecordService.js' });
const RecordService = context.exports.default;

function fixture(records, config) {
  const service = new RecordService(Object.assign({
    slideTitleField: 'Title',
    slideDescriptionField: 'Description',
    slideLinkField: 'ClickLink',
    enableDiagnostics: false
  }, config), { aUrl: 'https://example.test' });
  service._spListService.GetSlideData = (library) => {
    assert.strictEqual(library, 'Pictures');
    return Promise.resolve(records);
  };
  return service;
}

function record(id, link) {
  return {
    Id: String(id),
    FileRef: '/sites/demo/Pictures/image ' + id + '.jpg',
    Title: 'Slide ' + id,
    Description: 'Description ' + id,
    ClickLink: link
  };
}

async function run() {
  const missingLink = record(3);
  delete missingLink.ClickLink;
  const records = [
    record(1, null),
    record(2, undefined),
    missingLink,
    record(4, ''),
    record(5, 'https://example.test/text-link'),
    record(6, { Url: 'https://example.test/hyperlink', Description: 'Open destination' })
  ];
  const slides = await fixture(records).GetSlides('Pictures');
  assert.strictEqual(slides.length, 6, 'An empty optional link must not discard any images');
  slides.forEach((slide, index) => {
    assert.strictEqual(slide.slideId, String(index + 1));
    assert.strictEqual(slide.slideTitle, 'Slide ' + (index + 1));
    assert.strictEqual(slide.slideText, 'Description ' + (index + 1));
    assert.strictEqual(slide.slideNumber, index + 1);
    assert.strictEqual(slide.slideImgUrl, 'https://example.test/sites/demo/Pictures/image%20' + (index + 1) + '.jpg');
    assert.strictEqual(slide.slideNavigationUrl, index < 4 ? ''
      : index === 4 ? 'https://example.test/text-link' : 'https://example.test/hyperlink');
    assert.strictEqual(slide.slideLinkText, index === 5 ? 'Open destination' : '');
  });

  const incompleteLinks = [{}, { Url: null, Description: null }, { Url: '/destination' }];
  const incompleteSlides = await fixture(incompleteLinks.map((link, index) => record(index + 1, link))).GetSlides('Pictures');
  incompleteSlides.forEach((slide, index) => {
    assert.strictEqual(slide.slideNavigationUrl, index === 2 ? '/destination' : '');
    assert.strictEqual(slide.slideLinkText, '');
  });

  const custom = record(7, 'ignored');
  custom.Destination = null;
  custom.Heading = null;
  custom.Caption = undefined;
  custom.SlideOrder = 2;
  const customSlides = await fixture([custom], {
    slideLinkField: 'Destination', slideTitleField: 'Heading', slideDescriptionField: 'Caption'
  }).GetSlides('Pictures');
  assert.strictEqual(customSlides[0].slideNavigationUrl, '');
  assert.strictEqual(customSlides[0].slideLinkText, '');
  assert.strictEqual(customSlides[0].slideTitle, '');
  assert.strictEqual(customSlides[0].slideText, '');
  assert.strictEqual(customSlides[0].slideNumber, 2);

  const unmappedSlides = await fixture([record(8, null)], { slideLinkField: '' }).GetSlides('Pictures');
  assert.strictEqual(unmappedSlides[0].slideNavigationUrl, '');
  assert.strictEqual(unmappedSlides[0].slideLinkText, '');

  const service = fixture([]);
  const error = new Error('Library request failed');
  service._spListService.GetSlideData = () => Promise.reject(error);
  try {
    await service.GetSlides('Pictures');
    assert.fail('Library failures must not become an empty successful result');
  } catch (actual) {
    assert.strictEqual(actual, error);
  }

  console.log('Carousel record service tests passed (empty links, six-slide mapping, custom fields, and request failures).');
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
