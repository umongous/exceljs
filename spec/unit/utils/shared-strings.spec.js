const SharedStrings = verquire('utils/shared-strings');

describe('SharedStrings', () => {
  it('Stores and shares string values', () => {
    const ss = new SharedStrings();

    const iHello = ss.add('Hello');
    const iHelloV2 = ss.add('Hello');
    const iGoodbye = ss.add('Goodbye');

    expect(iHello).to.equal(iHelloV2);
    expect(iGoodbye).to.not.equal(iHelloV2);

    expect(ss.count).to.equal(2);
    expect(ss.totalRefs).to.equal(3);
  });

  it('Does not escape values', () => {
    // that's the job of the xml utils
    const ss = new SharedStrings();

    const iXml = ss.add('<tag>value</tag>');
    const iAmpersand = ss.add('&');

    expect(ss.getString(iXml)).to.equal('<tag>value</tag>');
    expect(ss.getString(iAmpersand)).to.equal('&');
  });

  it('Stores distinct rich-text values separately', () => {
    // Object values used to collide on the '[object Object]' hash key,
    // collapsing every rich-text cell into the first one added.
    const ss = new SharedStrings();

    const title = {richText: [{font: {size: 16}, text: 'Title'}]};
    const arrowUp = {
      richText: [{font: {color: {argb: '00FF00'}}, text: ' up'}],
    };
    const arrowDown = {
      richText: [{font: {color: {argb: 'FF0000'}}, text: ' down'}],
    };

    const iTitle = ss.add(title);
    const iUp = ss.add(arrowUp);
    const iDown = ss.add(arrowDown);
    const iUpAgain = ss.add({
      richText: [{font: {color: {argb: '00FF00'}}, text: ' up'}],
    });

    expect(iTitle).to.not.equal(iUp);
    expect(iUp).to.not.equal(iDown);
    expect(iUpAgain).to.equal(iUp);

    expect(ss.getString(iTitle)).to.equal(title);
    expect(ss.getString(iUp)).to.equal(arrowUp);
    expect(ss.getString(iDown)).to.equal(arrowDown);

    expect(ss.count).to.equal(3);
    expect(ss.totalRefs).to.equal(4);
  });

  it('Does not collide plain strings with serialized rich text', () => {
    const ss = new SharedStrings();

    const rich = {richText: [{text: 'x'}]};
    const lookalike = JSON.stringify(rich);

    const iRich = ss.add(rich);
    const iLookalike = ss.add(lookalike);

    expect(iRich).to.not.equal(iLookalike);
    expect(ss.getString(iRich)).to.equal(rich);
    expect(ss.getString(iLookalike)).to.equal(lookalike);
  });
});
