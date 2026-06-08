const db = require('../db');
const { runReconciliation } = require('./reconciliation');

const REASONS = {
  MISSING_PAYMENT: 'missing_payment',
  AMOUNT_MISMATCH: 'amount_mismatch',
  DUPLICATE_ORDER: 'duplicate_order',
};

const REASON_LABELS = {
  [REASONS.MISSING_PAYMENT]: 'Missing payment',
  [REASONS.AMOUNT_MISMATCH]: 'Amount mismatch',
  [REASONS.DUPLICATE_ORDER]: 'Duplicate order',
};

function getOpenAIConfig() {
  const apiKey = process.env.OPENAI_API_KEY?.trim();
  if (!apiKey) {
    throw new Error('Missing required environment variable: OPENAI_API_KEY');
  }
  return {
    apiKey,
    model: process.env.OPENAI_MODEL?.trim() || 'gpt-4o-mini',
  };
}

function findDuplicateOrderIds() {
  const rows = db
    .prepare(
      `SELECT order_id, COUNT(*) as count
       FROM orders
       GROUP BY order_id
       HAVING count > 1`
    )
    .all();
  return rows.map((r) => r.order_id);
}

function buildIssuePayload(reconciliation) {
  const duplicateOrderIds = findDuplicateOrderIds();
  const issues = [
    ...reconciliation.unmatched,
    ...reconciliation.exception,
  ].map((row) => ({
    orderId: row.orderId,
    status: row.status,
    type: row.type,
    amount: row.amount,
    orderAmount: row.orderAmount,
    paymentAmount: row.paymentAmount,
    systemReason: row.reason,
    isDuplicateOrder: duplicateOrderIds.includes(row.orderId),
  }));

  return { issues, duplicateOrderIds, summary: reconciliation.summary };
}

function fallbackInsight(issue) {
  if (issue.isDuplicateOrder) {
    return {
      orderId: issue.orderId,
      suggestedReason: REASONS.DUPLICATE_ORDER,
      reasonLabel: REASON_LABELS[REASONS.DUPLICATE_ORDER],
      explanation:
        'Multiple orders share this Order ID in the database, which can prevent clean matching.',
      confidence: 'high',
    };
  }
  if (issue.status === 'exception') {
    return {
      orderId: issue.orderId,
      suggestedReason: REASONS.AMOUNT_MISMATCH,
      reasonLabel: REASON_LABELS[REASONS.AMOUNT_MISMATCH],
      explanation: `Order amount (${issue.orderAmount}) does not match payment amount (${issue.paymentAmount}).`,
      confidence: 'high',
    };
  }
  if (issue.type === 'order') {
    return {
      orderId: issue.orderId,
      suggestedReason: REASONS.MISSING_PAYMENT,
      reasonLabel: REASON_LABELS[REASONS.MISSING_PAYMENT],
      explanation:
        'No payment was found with a matching Order ID and amount.',
      confidence: 'high',
    };
  }
  return {
    orderId: issue.orderId,
    suggestedReason: REASONS.MISSING_PAYMENT,
    reasonLabel: REASON_LABELS[REASONS.MISSING_PAYMENT],
    explanation:
      'Payment exists but no matching order was found — possible missing or mislabeled order.',
    confidence: 'medium',
  };
}

function buildFallbackInsights({ issues, duplicateOrderIds, summary }) {
  return {
    overview:
      issues.length === 0
        ? 'All transactions are matched. No issues to analyze.'
        : `Found ${issues.length} issue(s): ${summary.unmatchedCount} unmatched and ${summary.exceptionCount} amount exception(s).${duplicateOrderIds.length ? ` ${duplicateOrderIds.length} Order ID(s) appear more than once.` : ''}`,
    insights: issues.map((issue) => fallbackInsight(issue)),
    generatedBy: 'rules',
  };
}

async function callOpenAI({ issues, duplicateOrderIds, summary }) {
  const { apiKey, model } = getOpenAIConfig();

  const systemPrompt = `You analyze ecommerce reconciliation exceptions.
Suggest exactly one primary reason per transaction from:
- missing_payment (no matching payment or order)
- amount_mismatch (order and payment exist but amounts differ)
- duplicate_order (same order ID appears multiple times)

Respond with valid JSON only:
{
  "overview": "brief summary string",
  "insights": [
    {
      "orderId": "string",
      "suggestedReason": "missing_payment|amount_mismatch|duplicate_order",
      "reasonLabel": "Missing payment|Amount mismatch|Duplicate order",
      "explanation": "1-2 sentence actionable explanation",
      "confidence": "high|medium|low"
    }
  ]
}`;

  const userPrompt = JSON.stringify(
    {
      summary,
      duplicateOrderIds,
      issues,
    },
    null,
    2
  );

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      model,
      temperature: 0.2,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt },
      ],
    }),
  });

  if (!res.ok) {
    const errBody = await res.text();
    throw new Error(`OpenAI API error (${res.status}): ${errBody}`);
  }

  const data = await res.json();
  const content = data.choices?.[0]?.message?.content;
  if (!content) throw new Error('Empty response from OpenAI');

  const parsed = JSON.parse(content);
  return {
    overview: parsed.overview || 'Analysis complete.',
    insights: (parsed.insights || []).map((item) => ({
      orderId: item.orderId,
      suggestedReason: item.suggestedReason,
      reasonLabel:
        item.reasonLabel ||
        REASON_LABELS[item.suggestedReason] ||
        item.suggestedReason,
      explanation: item.explanation,
      confidence: item.confidence || 'medium',
    })),
    generatedBy: 'openai',
    model,
  };
}

const INSIGHTS_LIMIT = 50;

function hasOpenAIKey() {
  return Boolean(process.env.OPENAI_API_KEY?.trim());
}

async function generateInsights(reconciliation) {
  const payload = buildIssuePayload(reconciliation);
  const { issues } = payload;

  if (issues.length === 0) {
    return {
      overview: 'All transactions are matched. No unmatched items to analyze.',
      insights: [],
      generatedBy: 'none',
      issueCount: 0,
    };
  }

  let result;

  if (!hasOpenAIKey()) {
    result = {
      ...buildFallbackInsights(payload),
      generatedBy: 'rules',
      info: 'Set OPENAI_API_KEY in backend/.env to enable AI-generated insights.',
    };
  } else {
    try {
      result = await callOpenAI(payload);
    } catch (err) {
      console.warn('[insights] OpenAI failed, using rule-based fallback:', err.message);
      result = {
        ...buildFallbackInsights(payload),
        warning: `OpenAI unavailable: ${err.message}`,
      };
    }
  }

  const truncated = result.insights.length > INSIGHTS_LIMIT;
  const insights = result.insights.slice(0, INSIGHTS_LIMIT);

  return {
    ...result,
    insights,
    issueCount: issues.length,
    duplicateOrderIds: payload.duplicateOrderIds,
    truncated,
    showing: insights.length,
  };
}

async function getInsights() {
  const reconciliation = runReconciliation();
  const result = await generateInsights(reconciliation);
  return { reconciliation: { summary: reconciliation.summary }, ...result };
}

module.exports = {
  generateInsights,
  getInsights,
  REASONS,
  REASON_LABELS,
};
