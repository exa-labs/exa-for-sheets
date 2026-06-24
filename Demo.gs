/**
 * Demo.gs — Realistic mock data for Websets Sync.
 *
 * Lets users try out the sheet-sync workflow without an Exa Pro plan.
 * The helpers simulate a live Webset that gradually discovers new items,
 * which is useful for testing the auto-refresh / monitor feature.
 */

/* ---------- Mock Webset metadata ---------- */

var WEBSETS_DEMO_INFO = {
  id: 'demo-webset-001',
  title: 'AI Startups — Series A, US',
  status: 'idle',
  itemCount: 21
};

/* ---------- 8 base items ---------- */

var WEBSETS_DEMO_ITEMS = [
  {
    id: 'item_clkza5aw2f',
    name: 'Anthropic',
    url: 'https://www.anthropic.com',
    title: 'Anthropic — AI Safety Company',
    status: 'verified',
    verifiedAt: '2026-06-20T14:22:00Z',
    enrichments: {
      funding_stage: 'Series D',
      total_funding: '$7.6B',
      founded_year: 2021,
      employee_count: 1500,
      hq_location: 'San Francisco, CA'
    },
    entity: { type: 'Company', domain: 'anthropic.com' }
  },
  {
    id: 'item_dwpiu38spr',
    name: 'Cohere',
    url: 'https://cohere.com',
    title: 'Cohere — Enterprise AI Platform',
    status: 'verified',
    verifiedAt: '2026-06-20T14:25:00Z',
    enrichments: {
      funding_stage: 'Series D',
      total_funding: '$970M',
      founded_year: 2019,
      employee_count: 800,
      hq_location: 'Toronto, ON'
    },
    entity: { type: 'Company', domain: 'cohere.com' }
  },
  {
    id: 'item_sq2hpz2wl3',
    name: 'Mistral AI',
    url: 'https://mistral.ai',
    title: 'Mistral AI — Open & Portable Generative AI',
    status: 'verified',
    verifiedAt: '2026-06-20T14:28:00Z',
    enrichments: {
      funding_stage: 'Series B',
      total_funding: '$1.1B',
      founded_year: 2023,
      employee_count: 600,
      hq_location: 'Paris, France'
    },
    entity: { type: 'Company', domain: 'mistral.ai' }
  },
  {
    id: 'item_w209i7ee0u',
    name: 'Runway',
    url: 'https://runwayml.com',
    title: 'Runway — Applied AI Research Company',
    status: 'verified',
    verifiedAt: '2026-06-20T14:31:00Z',
    enrichments: {
      funding_stage: 'Series D',
      total_funding: '$540M',
      founded_year: 2018,
      employee_count: 500,
      hq_location: 'New York, NY'
    },
    entity: { type: 'Company', domain: 'runwayml.com' }
  },
  {
    id: 'item_4ypx90i9w5',
    name: 'Imbue',
    url: 'https://imbue.com',
    title: 'Imbue — AI Agents That Reason and Code',
    status: 'verified',
    verifiedAt: '2026-06-20T14:34:00Z',
    enrichments: {
      funding_stage: 'Series B',
      total_funding: '$230M',
      founded_year: 2021,
      employee_count: 120,
      hq_location: 'San Francisco, CA'
    },
    entity: { type: 'Company', domain: 'imbue.com' }
  },
  {
    id: 'item_vcyov2fw9d',
    name: 'Adept AI',
    url: 'https://adept.ai',
    title: 'Adept AI — Useful General Intelligence',
    status: 'verified',
    verifiedAt: '2026-06-20T14:37:00Z',
    enrichments: {
      funding_stage: 'Series B',
      total_funding: '$415M',
      founded_year: 2022,
      employee_count: 200,
      hq_location: 'San Francisco, CA'
    },
    entity: { type: 'Company', domain: 'adept.ai' }
  },
  {
    id: 'item_304gtvw72s',
    name: 'Sakana AI',
    url: 'https://sakana.ai',
    title: 'Sakana AI — Nature-Inspired Intelligence',
    status: 'verified',
    verifiedAt: '2026-06-20T14:40:00Z',
    enrichments: {
      funding_stage: 'Series A',
      total_funding: '$155M',
      founded_year: 2023,
      employee_count: 90,
      hq_location: 'Tokyo, Japan'
    },
    entity: { type: 'Company', domain: 'sakana.ai' }
  },
  {
    id: 'item_7z6zzn1xr8',
    name: 'Sierra AI',
    url: 'https://sierra.ai',
    title: 'Sierra AI — Conversational AI for Businesses',
    status: 'verified',
    verifiedAt: '2026-06-20T14:43:00Z',
    enrichments: {
      funding_stage: 'Series B',
      total_funding: '$175M',
      founded_year: 2023,
      employee_count: 150,
      hq_location: 'San Francisco, CA'
    },
    entity: { type: 'Company', domain: 'sierra.ai' }
  },
  {
    id: 'item_3dxav31gx1',
    name: 'Perplexity',
    url: 'https://www.perplexity.ai',
    title: 'Perplexity — AI Search Engine',
    status: 'verified',
    verifiedAt: '2026-06-20T14:46:00Z',
    enrichments: {
      funding_stage: 'Series B',
      total_funding: '$165M',
      founded_year: 2022,
      employee_count: 60,
      hq_location: 'San Francisco, CA'
    },
    entity: { type: 'Company', domain: 'perplexity.ai' }
  },
  {
    id: 'item_36lir407xc',
    name: 'Character.ai',
    url: 'https://character.ai',
    title: 'Character.ai — Customized AI Companions',
    status: 'verified',
    verifiedAt: '2026-06-20T14:49:00Z',
    enrichments: {
      funding_stage: 'Series A',
      total_funding: '$150M',
      founded_year: 2021,
      employee_count: 55,
      hq_location: 'Menlo Park, CA'
    },
    entity: { type: 'Company', domain: 'character.ai' }
  },
  {
    id: 'item_ment8f56xy',
    name: 'Midjourney',
    url: 'https://www.midjourney.com',
    title: 'Midjourney — AI Image Generation',
    status: 'verified',
    verifiedAt: '2026-06-20T14:52:00Z',
    enrichments: {
      funding_stage: 'Bootstrapped',
      total_funding: '$0M',
      founded_year: 2021,
      employee_count: 40,
      hq_location: 'San Francisco, CA'
    },
    entity: { type: 'Company', domain: 'midjourney.com' }
  },
  {
    id: 'item_6k8ifzshk8',
    name: 'Pika',
    url: 'https://pika.art',
    title: 'Pika — AI Video Generation Platform',
    status: 'verified',
    verifiedAt: '2026-06-20T14:55:00Z',
    enrichments: {
      funding_stage: 'Series A',
      total_funding: '$55M',
      founded_year: 2023,
      employee_count: 35,
      hq_location: 'Palo Alto, CA'
    },
    entity: { type: 'Company', domain: 'pika.art' }
  },
  {
    id: 'item_5oaaqb2lo2',
    name: 'Scale AI',
    url: 'https://scale.com',
    title: 'Scale AI — Data Infrastructure for AI',
    status: 'verified',
    verifiedAt: '2026-06-20T14:58:00Z',
    enrichments: {
      funding_stage: 'Series F',
      total_funding: '$1.6B',
      founded_year: 2016,
      employee_count: 1200,
      hq_location: 'San Francisco, CA'
    },
    entity: { type: 'Company', domain: 'scale.com' }
  },
  {
    id: 'item_xqda256m6y',
    name: 'Hugging Face',
    url: 'https://huggingface.co',
    title: 'Hugging Face — The AI Community Building the Future',
    status: 'verified',
    verifiedAt: '2026-06-20T15:01:00Z',
    enrichments: {
      funding_stage: 'Series D',
      total_funding: '$395M',
      founded_year: 2016,
      employee_count: 220,
      hq_location: 'New York, NY'
    },
    entity: { type: 'Company', domain: 'huggingface.co' }
  },
  {
    id: 'item_bm3hgkts24',
    name: 'Inflection AI',
    url: 'https://inflection.ai',
    title: 'Inflection AI — Personal AI for Everyone',
    status: 'verified',
    verifiedAt: '2026-06-20T15:04:00Z',
    enrichments: {
      funding_stage: 'Series C',
      total_funding: '$1.5B',
      founded_year: 2022,
      employee_count: 45,
      hq_location: 'Palo Alto, CA'
    },
    entity: { type: 'Company', domain: 'inflection.ai' }
  },
  {
    id: 'item_ej75ceofve',
    name: 'Databricks',
    url: 'https://www.databricks.com',
    title: 'Databricks — Data and AI Company',
    status: 'verified',
    verifiedAt: '2026-06-20T15:07:00Z',
    enrichments: {
      funding_stage: 'Series I',
      total_funding: '$4.0B',
      founded_year: 2013,
      employee_count: 6000,
      hq_location: 'San Francisco, CA'
    },
    entity: { type: 'Company', domain: 'databricks.com' }
  },
  {
    id: 'item_7o7bqj7c9n',
    name: 'ElevenLabs',
    url: 'https://elevenlabs.io',
    title: 'ElevenLabs — AI Voice Generator',
    status: 'verified',
    verifiedAt: '2026-06-20T15:10:00Z',
    enrichments: {
      funding_stage: 'Series B',
      total_funding: '$101M',
      founded_year: 2022,
      employee_count: 80,
      hq_location: 'London, UK'
    },
    entity: { type: 'Company', domain: 'elevenlabs.io' }
  },
  {
    id: 'item_7cs6bic67z',
    name: 'You.com',
    url: 'https://you.com',
    title: 'You.com — The AI Search Engine',
    status: 'verified',
    verifiedAt: '2026-06-20T15:13:00Z',
    enrichments: {
      funding_stage: 'Series A',
      total_funding: '$45M',
      founded_year: 2020,
      employee_count: 65,
      hq_location: 'Palo Alto, CA'
    },
    entity: { type: 'Company', domain: 'you.com' }
  },
  {
    id: 'item_49ms6wnpfo',
    name: 'Groq',
    url: 'https://groq.com',
    title: 'Groq — LPU Inference Engine',
    status: 'verified',
    verifiedAt: '2026-06-20T15:16:00Z',
    enrichments: {
      funding_stage: 'Series C',
      total_funding: '$362M',
      founded_year: 2016,
      employee_count: 180,
      hq_location: 'Mountain View, CA'
    },
    entity: { type: 'Company', domain: 'groq.com' }
  },
  {
    id: 'item_ftgi9ildox',
    name: 'Together AI',
    url: 'https://www.together.ai',
    title: 'Together AI — Platform for Generative AI',
    status: 'verified',
    verifiedAt: '2026-06-20T15:19:00Z',
    enrichments: {
      funding_stage: 'Series A',
      total_funding: '$122M',
      founded_year: 2022,
      employee_count: 85,
      hq_location: 'San Francisco, CA'
    },
    entity: { type: 'Company', domain: 'together.ai' }
  },
  {
    id: 'item_lntjod8i5g',
    name: 'HeyGen',
    url: 'https://www.heygen.com',
    title: 'HeyGen — AI Video Generation',
    status: 'verified',
    verifiedAt: '2026-06-20T15:22:00Z',
    enrichments: {
      funding_stage: 'Series A',
      total_funding: '$65M',
      founded_year: 2020,
      employee_count: 90,
      hq_location: 'Los Angeles, CA'
    },
    entity: { type: 'Company', domain: 'heygen.com' }
  }
];

/* ---------- 4 extra items (appear gradually during auto-refresh) ---------- */

var WEBSETS_DEMO_EXTRAS = [
  {
    id: 'item_3nbxrer9w5',
    name: 'Harvey AI',
    url: 'https://harvey.ai',
    title: 'Harvey AI — AI for Legal Professionals',
    status: 'verified',
    verifiedAt: '2026-06-21T09:10:00Z',
    enrichments: {
      funding_stage: 'Series C',
      total_funding: '$380M',
      founded_year: 2022,
      employee_count: 350,
      hq_location: 'San Francisco, CA'
    },
    entity: { type: 'Company', domain: 'harvey.ai' }
  },
  {
    id: 'item_o6y0isuzus',
    name: 'Glean',
    url: 'https://glean.com',
    title: 'Glean — AI-Powered Work Assistant',
    status: 'verified',
    verifiedAt: '2026-06-21T11:05:00Z',
    enrichments: {
      funding_stage: 'Series D',
      total_funding: '$800M',
      founded_year: 2019,
      employee_count: 700,
      hq_location: 'Palo Alto, CA'
    },
    entity: { type: 'Company', domain: 'glean.com' }
  },
  {
    id: 'item_c7q3c3145k',
    name: 'Poolside AI',
    url: 'https://poolside.ai',
    title: 'Poolside AI — AI That Writes Software',
    status: 'verified',
    verifiedAt: '2026-06-22T08:30:00Z',
    enrichments: {
      funding_stage: 'Series B',
      total_funding: '$500M',
      founded_year: 2023,
      employee_count: 180,
      hq_location: 'San Francisco, CA'
    },
    entity: { type: 'Company', domain: 'poolside.ai' }
  },
  {
    id: 'item_y5kl3tzv82',
    name: 'Relevance AI',
    url: 'https://relevanceai.com',
    title: 'Relevance AI — Build & Deploy AI Agents',
    status: 'verified',
    verifiedAt: '2026-06-22T15:45:00Z',
    enrichments: {
      funding_stage: 'Series B',
      total_funding: '$50M',
      founded_year: 2020,
      employee_count: 100,
      hq_location: 'Sydney, Australia'
    },
    entity: { type: 'Company', domain: 'relevanceai.com' }
  }
];

/* ---------- State ---------- */

/** Tracks how many extras have been surfaced so far. */
var websets_demoExtraIndex_ = 0;

/* ---------- Public helpers ---------- */

/**
 * Returns demo items, progressively including extras on each call.
 *
 * - First call  → 8 base items
 * - Second call → 8 base + 1 extra  (9 total)
 * - Third call  → 8 base + 2 extras (10 total)
 * - … up to 8 base + 4 extras (12 total)
 *
 * This simulates a Webset Monitor discovering new items over time.
 *
 * @return {Object[]} Array of demo item objects.
 */
function getWebsetsDemoItems() {
  var items = WEBSETS_DEMO_ITEMS.slice();

  if (websets_demoExtraIndex_ < WEBSETS_DEMO_EXTRAS.length) {
    var extras = WEBSETS_DEMO_EXTRAS.slice(0, websets_demoExtraIndex_);
    items = items.concat(extras);
    websets_demoExtraIndex_++;
  } else {
    items = items.concat(WEBSETS_DEMO_EXTRAS);
  }

  return items;
}

/**
 * Returns the mock Webset metadata object.
 *
 * @return {Object} Demo webset info.
 */
function getWebsetsDemoInfo() {
  return WEBSETS_DEMO_INFO;
}
