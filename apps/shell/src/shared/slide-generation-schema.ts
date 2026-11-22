const color = { type: 'string', pattern: '^#?[0-9A-Fa-f]{6}$' }
const paragraphs = {
  type: 'array',
  minItems: 1,
  items: {
    type: 'object',
    properties: {
      runs: {
        type: 'array',
        minItems: 1,
        items: {
          type: 'object',
          properties: {
            text: { type: 'string', maxLength: 4000 },
            sizePt: { type: 'number', minimum: 6, maximum: 160 },
            bold: { type: 'boolean' },
            italic: { type: 'boolean' },
            color,
            font: { type: 'string' },
          },
          required: ['text'],
        },
      },
      align: { type: 'string', enum: ['left', 'center', 'right', 'justify'] },
    },
    required: ['runs'],
  },
}
/** Matches the structured 1280x720 page format consumed by the official generation pipeline. */
export const SLIDE_PAGE_SCHEMA = {
  type: 'object',
  properties: {
    background: color,
    elements: {
      type: 'array',
      minItems: 1,
      maxItems: 48,
      items: {
        type: 'object',
        properties: {
          type: { type: 'string', enum: ['text', 'shape', 'image'] },
          x: { type: 'number' },
          y: { type: 'number' },
          w: { type: 'number', minimum: 1 },
          h: { type: 'number', minimum: 1 },
          paragraphs,
          shape: { type: 'string' },
          fill: color,
          url: { type: 'string' },
          valign: { type: 'string', enum: ['top', 'middle', 'bottom'] },
        },
        required: ['type', 'x', 'y', 'w', 'h'],
        oneOf: [
          { properties: { type: { const: 'text' } }, required: ['paragraphs'] },
          { properties: { type: { const: 'shape' } }, required: ['shape'] },
          { properties: { type: { const: 'image' } }, required: ['url'] },
        ],
      },
    },
  },
  required: ['elements'],
}
export const SLIDE_CREATE_FIELDS = {
  path: { type: 'string' },
  draftSlides: {
    type: 'array',
    minItems: 1,
    maxItems: 50,
    items: {
      type: 'object',
      properties: {
        layout: {
          type: 'string',
          enum: ['cover', 'explain', 'compare', 'steps', 'practice', 'summary'],
          description:
            'Compare uses exactly two sections, steps uses two to four, and cover uses zero to two.',
        },
        title: { type: 'string' },
        sections: {
          type: 'array',
          maxItems: 4,
          items: {
            type: 'object',
            properties: { heading: { type: 'string' }, text: { type: 'string' } },
          },
        },
      },
      required: ['layout', 'title', 'sections'],
    },
  },
  theme: { type: 'string', enum: ['classroom', 'report'] },
  slides: { type: 'array', minItems: 1, items: SLIDE_PAGE_SCHEMA },
  pages: { type: 'array', minItems: 1, items: SLIDE_PAGE_SCHEMA },
  spec: SLIDE_PAGE_SCHEMA,
  specJson: {
    type: 'string',
    description:
      'JSON encoding of a page spec with elements; every element needs type,x,y,w,h in pixels on a 1280x720 canvas. Text uses paragraphs:[{runs:[{text,color,sizePt}]}].',
  },
}
