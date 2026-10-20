/**
 * Parses an Actionable Messages Designer template and turns it into the HTML
 * body that Outlook renders as an actionable card.
 */
const CardBuilder = (() => {
  const ADAPTIVE_SCRIPT_TYPE = 'application/adaptivecard+json';
  const MESSAGECARD_SCRIPT_TYPE = 'application/ld+json';
  const MESSAGECARD_CONTEXT = 'https://schema.org/extensions';

  /** Replaces every {{name}} placeholder using the supplied key/value map. */
  function applyVariables(rawJson, variables) {
    if (!variables || Object.keys(variables).length === 0) {
      return rawJson;
    }
    return rawJson.replace(/\{\{\s*([\w.-]+)\s*\}\}/g, (match, name) => {
      if (!Object.prototype.hasOwnProperty.call(variables, name)) {
        return match;
      }
      // The replacement lands inside a JSON string literal, so escape it as one.
      const encoded = JSON.stringify(String(variables[name]));
      return encoded.slice(1, -1);
    });
  }

  /** Parses `name=value` lines into an object. */
  function parseVariables(text) {
    const variables = {};
    (text || '').split(/\r?\n/).forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith('#')) {
        return;
      }
      const index = trimmed.indexOf('=');
      if (index <= 0) {
        return;
      }
      variables[trimmed.slice(0, index).trim()] = trimmed.slice(index + 1).trim();
    });
    return variables;
  }

  function detectCardType(card) {
    if (card && card['@type'] === 'MessageCard') {
      return 'messagecard';
    }
    if (card && card.type === 'AdaptiveCard') {
      return 'adaptive';
    }
    return 'unknown';
  }

  function describe(cardType) {
    if (cardType === 'adaptive') {
      return 'Adaptive Card';
    }
    if (cardType === 'messagecard') {
      return 'MessageCard (legacy)';
    }
    return 'Unrecognized card';
  }

  function collectActionTypes(node, found) {
    if (Array.isArray(node)) {
      node.forEach((item) => collectActionTypes(item, found));
      return found;
    }
    if (node && typeof node === 'object') {
      if (typeof node.type === 'string' && node.type.startsWith('Action.')) {
        found.add(node.type);
      }
      if (typeof node['@type'] === 'string') {
        found.add(node['@type']);
      }
      Object.values(node).forEach((value) => collectActionTypes(value, found));
    }
    return found;
  }

  /**
   * Validates and normalizes a raw template.
   * @returns {{card: object, cardType: string, warnings: string[]}}
   */
  function prepare(rawJson, options) {
    const opts = options || {};
    const substituted = applyVariables(rawJson, opts.variables);

    let card;
    try {
      card = JSON.parse(substituted);
    } catch (error) {
      throw new Error('The template is not valid JSON: ' + error.message);
    }
    if (!card || typeof card !== 'object' || Array.isArray(card)) {
      throw new Error('The template must be a JSON object.');
    }

    const cardType = detectCardType(card);
    if (cardType === 'unknown') {
      throw new Error(
        'Unrecognized template. Expected "type": "AdaptiveCard" or "@type": "MessageCard".'
      );
    }

    const warnings = [];
    const originator = (opts.originator || '').trim();
    if (originator) {
      card.originator = originator;
    }

    if (cardType === 'messagecard' && !card['@context']) {
      card['@context'] = MESSAGECARD_CONTEXT;
    }
    if (cardType === 'adaptive' && !card.version) {
      card.version = '1.0';
      warnings.push('Adaptive Card had no "version"; defaulted to 1.0.');
    }

    const actionTypes = collectActionTypes(card, new Set());
    const hasServiceAction =
      actionTypes.has('Action.Http') ||
      actionTypes.has('Action.ActionCard') ||
      actionTypes.has('HttpPOST') ||
      actionTypes.has('ActionCard');

    if (!card.originator && hasServiceAction) {
      warnings.push(
        'No originator ID set. Outlook only renders cards with service actions when the ' +
          'originator is registered in the Actionable Email Developer Dashboard.'
      );
    }
    if (!actionTypes.size) {
      warnings.push('The card declares no actions; it will render as a static card.');
    }

    return { card, cardType, warnings };
  }

  function escapeHtml(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  /** Serializes the card so that it can never terminate the host <script> tag. */
  function serializeForScriptTag(card) {
    return JSON.stringify(card, null, 2)
      .replace(/</g, '\\u003c')
      .replace(/>/g, '\\u003e')
      .replace(/&/g, '\\u0026')
      .replace(/\u2028/g, '\\u2028')
      .replace(/\u2029/g, '\\u2029');
  }

  function buildHtml(card, cardType, fallbackText) {
    const scriptType = cardType === 'messagecard' ? MESSAGECARD_SCRIPT_TYPE : ADAPTIVE_SCRIPT_TYPE;
    const fallback = escapeHtml(
      fallbackText ||
        'This message contains an actionable card. Open it in Outlook to view and respond.'
    ).replace(/\r?\n/g, '<br />');

    return [
      '<html>',
      '<head>',
      '<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />',
      '<script type="' + scriptType + '">',
      serializeForScriptTag(card),
      '</script>',
      '</head>',
      '<body>',
      '<p>' + fallback + '</p>',
      '</body>',
      '</html>'
    ].join('\n');
  }

  /**
   * One-shot helper: template text in, Graph-ready HTML body out.
   * @returns {{html: string, cardType: string, warnings: string[], card: object}}
   */
  function build(rawJson, options) {
    const prepared = prepare(rawJson, options);
    return {
      ...prepared,
      html: buildHtml(prepared.card, prepared.cardType, (options || {}).fallbackText)
    };
  }

  return { build, prepare, buildHtml, parseVariables, detectCardType, describe };
})();
