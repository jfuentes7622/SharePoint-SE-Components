'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const React = require('react');
const ts = require('typescript');

const source = fs.readFileSync(path.join(
  __dirname, '..', 'src', 'webparts', 'carousel', 'components', 'Carousel.tsx'
), 'utf8');
const compiled = ts.transpileModule(source, {
  compilerOptions: {
    module: ts.ModuleKind.CommonJS,
    target: ts.ScriptTarget.ES5,
    jsx: ts.JsxEmit.React
  }
}).outputText;
const styles = {};
['iskmCarousel', 'carouselSurface', 'carouselDiv', 'carouselViewport', 'slideFrame',
  'slideTitlePanel', 'slideTitle', 'slideFooter', 'slideDescription', 'slideLinkRow',
  'imageCircle', 'imageSquare'].forEach((name) => { styles[name] = name; });
const context = {
  exports: {},
  console: console,
  require: (name) => {
    switch (name) {
      case 'react': return React;
      case '@microsoft/sp-lodash-subset': return { assign: Object.assign };
      case './Carousel.module.scss': return { default: styles };
      case 'guid-typescript': return { Guid: { create: () => 'test-image' } };
      case 'react-slick': return { default: 'slider' };
      case 'dompurify': return { sanitize: (value) => value };
      default:
        if (name.indexOf('slick-carousel/') === 0) { return {}; }
        throw new Error('Unexpected import: ' + name);
    }
  }
};
vm.runInNewContext(compiled, context, { filename: 'Carousel.js' });
const Carousel = context.exports.default;

function renderSlide(overrides) {
  const props = Object.assign({
    carouselSlideLibrary: 'Pictures',
    carouselWidth: 500,
    carouselHeight: 300,
    carouselSlideInterval: '5000',
    carouselTransitionInterval: '500',
    enableDiagnostics: false,
    recSvc: { GetSlides: () => new Promise(() => {}) }
  }, overrides);
  const component = new Carousel(props);
  component.state = { imageData: ['image.svg'], dataLoaded: true };
  component._slides = [{
    slideTitle: 'Image title',
    slideText: 'Image description',
    slideNavigationUrl: 'https://example.com',
    slideLinkText: 'Learn more'
  }];
  const root = component.render();
  const carousel = React.Children.toArray(root.props.children)[0];
  const viewport = carousel.props.children;
  const slider = viewport.props.children;
  return { component, slider, slide: slider.props.children[0] };
}

[300, 500].forEach((height) => {
  [false, true].forEach((captions) => {
    const rendered = renderSlide({
      carouselHeight: height,
      showSlideTitle: captions,
      showSlideDescription: captions,
      showSlideLink: captions
    });
    const children = React.Children.toArray(rendered.slide.props.children);
    const area = children[captions ? 1 : 0];
    const image = area.props.children;
    assert.strictEqual(rendered.slide.props.style.height, height + 'px');
    assert.strictEqual(rendered.slider.props.adaptiveHeight, false);
    assert.strictEqual(area.type, 'div');
    assert.strictEqual(area.props.style.flex, '1 1 0%');
    assert.strictEqual(area.props.style.minHeight, 0);
    assert.strictEqual(area.props.style.width, '100%');
    assert.strictEqual(area.props.style.position, 'relative');
    assert.strictEqual(image.props.style.position, 'absolute');
    assert.strictEqual(image.props.style.width, '100%');
    assert.strictEqual(image.props.style.height, '100%');
    assert.strictEqual(image.props.style.objectFit, 'contain');
    assert.strictEqual(image.props.style.marginTop, undefined);
    assert.strictEqual(image.props.style.marginBottom, undefined);
    assert.strictEqual(children.length, captions ? 3 : 1);
    let clicked = -1;
    rendered.component._slideClicked = (index) => { clicked = index; };
    image.props.onClick();
    assert.strictEqual(clicked, 0, 'Image navigation must remain wired');
  });
});

[0, -1, 'invalid'].forEach((height) => {
  const rendered = renderSlide({ carouselHeight: height });
  const area = React.Children.toArray(rendered.slide.props.children)[0];
  assert.strictEqual(rendered.slider.props.adaptiveHeight, true);
  assert.strictEqual(rendered.slide.props.style.height, undefined);
  assert.strictEqual(area.props.style.flex, '0 0 auto');
  assert.strictEqual(area.props.children.props.style.height, 'auto');
  assert.strictEqual(area.props.children.props.style.objectFit, 'contain');
});

[0, 300].forEach((height) => {
  const rendered = renderSlide({ carouselHeight: height, imageIsCircle: true });
  const image = React.Children.toArray(rendered.slide.props.children)[0];
  assert.strictEqual(image.type, 'img', 'Circular images must retain their existing layout');
  assert.strictEqual(image.props.style.objectFit, 'cover');
  assert.strictEqual(image.props.style.width, (height || 500) + 'px');
  assert.strictEqual(image.props.style.height, (height || 500) + 'px');
});

console.log('Carousel image sizing tests passed.');
