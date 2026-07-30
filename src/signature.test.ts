import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  hasTextSignature,
  hasHtmlSignature,
  appendTextSignature,
  appendHtmlSignature,
  applyIdentitySignature,
} from './signature.js';

// Real values from a Fastmail identity: the signature block and the "never
// configured" placeholder the server stores for identities with no signature.
const TEXT_SIG = 'Best Regards, \n\nThe Example Team.\n\n\n';
const HTML_SIG = '<div><div><span>Best Regards,&nbsp;</span></div><div><span>The Example Team.</span></div><div><img src="https://example.com/logo.png"><br></div></div><div><br></div>';
const EMPTY_HTML_SIG = '<div><br></div><div><br></div>';

describe('hasTextSignature', () => {
  it('accepts a signature with content', () => {
    assert.equal(hasTextSignature(TEXT_SIG), true);
  });

  it('rejects whitespace-only, the Fastmail empty-signature representation', () => {
    assert.equal(hasTextSignature('\n'), false);
    assert.equal(hasTextSignature('   \n\t '), false);
    assert.equal(hasTextSignature(''), false);
  });

  it('rejects absent values rather than throwing', () => {
    assert.equal(hasTextSignature(undefined), false);
    assert.equal(hasTextSignature(null), false);
  });
});

describe('hasHtmlSignature', () => {
  it('accepts markup with visible text', () => {
    assert.equal(hasHtmlSignature('<div>Best Regards</div>'), true);
  });

  it('rejects markup that renders as nothing', () => {
    assert.equal(hasHtmlSignature(EMPTY_HTML_SIG), false);
    assert.equal(hasHtmlSignature('<div><span></span></div>'), false);
    assert.equal(hasHtmlSignature('<div>&nbsp;&nbsp;</div>'), false);
  });

  it('accepts an image-only signature — a logo is content with no text', () => {
    assert.equal(hasHtmlSignature('<div><img src="https://example.com/logo.png"></div>'), true);
  });

  it('accepts a link-only signature', () => {
    assert.equal(hasHtmlSignature('<div><a href="https://example.com"></a></div>'), true);
  });
});

describe('appendTextSignature', () => {
  it('separates body and signature by exactly one blank line', () => {
    assert.equal(appendTextSignature('Hello.', TEXT_SIG), `Hello.\n\n${TEXT_SIG}`);
  });

  it('collapses existing trailing whitespace instead of compounding it', () => {
    assert.equal(appendTextSignature('Hello.\n\n\n', TEXT_SIG), `Hello.\n\n${TEXT_SIG}`);
  });

  it('is idempotent — re-signing does not stack a second copy', () => {
    const once = appendTextSignature('Hello.', TEXT_SIG);
    assert.equal(appendTextSignature(once, TEXT_SIG), once);
  });

  it('treats a trailing-whitespace variant as already signed', () => {
    const signed = `Hello.\n\n${TEXT_SIG.trimEnd()}`;
    assert.equal(appendTextSignature(signed, TEXT_SIG), signed);
  });

  it('leaves the body untouched when the signature is blank', () => {
    assert.equal(appendTextSignature('Hello.', '\n'), 'Hello.');
    assert.equal(appendTextSignature('Hello.', undefined), 'Hello.');
  });

  it('never synthesises a body from an empty one', () => {
    assert.equal(appendTextSignature('', TEXT_SIG), '');
  });
});

describe('appendHtmlSignature', () => {
  it('appends without inserting a separator', () => {
    assert.equal(appendHtmlSignature('<div>Hi</div>', HTML_SIG), `<div>Hi</div>${HTML_SIG}`);
  });

  it('is idempotent', () => {
    const once = appendHtmlSignature('<div>Hi</div>', HTML_SIG);
    assert.equal(appendHtmlSignature(once, HTML_SIG), once);
  });

  it('skips the empty-signature placeholder', () => {
    assert.equal(appendHtmlSignature('<div>Hi</div>', EMPTY_HTML_SIG), '<div>Hi</div>');
  });

  it('never synthesises a body from an empty one', () => {
    assert.equal(appendHtmlSignature('', HTML_SIG), '');
  });
});

describe('applyIdentitySignature', () => {
  const identity = { textSignature: TEXT_SIG, htmlSignature: HTML_SIG };

  it('signs both parts when both are present', () => {
    const result = applyIdentitySignature({ textBody: 'Hi.', htmlBody: '<div>Hi.</div>' }, identity);
    assert.equal(result.textBody, `Hi.\n\n${TEXT_SIG}`);
    assert.equal(result.htmlBody, `<div>Hi.</div>${HTML_SIG}`);
  });

  it('keeps an absent part absent — a text-only message stays text-only', () => {
    const result = applyIdentitySignature({ textBody: 'Hi.' }, identity);
    assert.equal(result.textBody, `Hi.\n\n${TEXT_SIG}`);
    assert.equal(result.htmlBody, undefined);
  });

  it('keeps an absent text part absent for an html-only message', () => {
    const result = applyIdentitySignature({ htmlBody: '<div>Hi.</div>' }, identity);
    assert.equal(result.textBody, undefined);
    assert.equal(result.htmlBody, `<div>Hi.</div>${HTML_SIG}`);
  });

  it('passes bodies through when the identity has no signature fields', () => {
    const result = applyIdentitySignature({ textBody: 'Hi.', htmlBody: '<div>Hi.</div>' }, {});
    assert.equal(result.textBody, 'Hi.');
    assert.equal(result.htmlBody, '<div>Hi.</div>');
  });

  it('passes bodies through when there is no identity at all', () => {
    const result = applyIdentitySignature({ textBody: 'Hi.' }, undefined);
    assert.equal(result.textBody, 'Hi.');
  });

  it('signs only the configured half when the identity defines one signature', () => {
    const result = applyIdentitySignature(
      { textBody: 'Hi.', htmlBody: '<div>Hi.</div>' },
      { textSignature: TEXT_SIG, htmlSignature: EMPTY_HTML_SIG },
    );
    assert.equal(result.textBody, `Hi.\n\n${TEXT_SIG}`);
    assert.equal(result.htmlBody, '<div>Hi.</div>');
  });
});
