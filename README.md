# `@decken/sdk` — Official TypeScript & Node.js SDK + CLI

Zero-dependency TypeScript / Node.js SDK and CLI for **[Decken](https://decken.ai)** — the Native OpenXML (`.pptx`) Slide Generation Kernel that turns any Coding Agent, PowerPoint, WPS Office, or SaaS into a **better Gamma**.

- **Website**: [https://decken.ai](https://decken.ai)
- **Documentation**: [https://docs.decken.ai](https://docs.decken.ai)
- **Python SDK**: [`decken-ai/decken-python`](https://github.com/decken-ai/decken-python)
- **MCP Server**: [`decken-ai/decken-mcp`](https://github.com/decken-ai/decken-mcp)

## Install

```bash
npm install github:decken-ai/decken-node
```

## Unified 3-Scenario Usage (`client.generate` / `POST /v1/decks`)

```ts
import { DeckenClient } from '@decken/sdk';

const decken = new DeckenClient({ apiKey: process.env.DECKEN_API_KEY });

// 1. Generate directly from outline (expressive diagrams + native <a:tbl> tables)
await decken.generate({
  deckTitle: 'Edge Inference Architecture',
  outline: [
    {
      title: 'Zero-Copy KV Routing Pipeline',
      diagram: {
        kind: 'flow',
        nodes: [
          { label: 'Anycast Edge', detail: 'TLS + Auth' },
          { label: 'Prefix Mesh', detail: 'Radix KV tree' },
          { label: 'FP8 Worker', detail: '29ms p99' },
        ],
      },
    },
  ],
  outputPath: './edge-architecture.pptx',
});

// 2. Generate from outline + reference image OR corporate .pptx template
await decken.generate({
  outline: '## FY26 Revenue\n- ARR reached $42.8M (+138% YoY)',
  reference: './corporate-master.pptx', // or './style-reference.png'
  outputPath: './fy26-board.pptx',
});

// 3. Improve an existing rough .pptx OR reconstruct flat screenshots
await decken.generate({
  source: './rough-draft.pptx', // or './screenshot.png'
  prompt: 'Unify visual hierarchy while preserving 100% of numbers verbatim.',
  outputPath: './upgraded.pptx',
});
```

## CLI

```bash
export DECKEN_API_KEY="dk_live_..."

npx github:decken-ai/decken-node generate \
  --outline ./outline.md \
  --reference ./corporate-master.pptx \
  -o ./deck.pptx
```

## License

MIT
