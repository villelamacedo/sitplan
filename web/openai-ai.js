import { invokeCloudFunction } from "./cloud-sync.js?v=7";

const MAX_HISTORY_ITEMS = 8;
const MAX_HISTORY_TEXT = 1200;

function cleanText(value, maxLength = 500) {
  return String(value || "")
    .replace(/\b[\w.+-]+@[\w.-]+\.[a-z]{2,}\b/gi, "[e-mail removido]")
    .replace(/\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b/g, "[CPF removido]")
    .replace(/(?:\+?55\s*)?(?:\(?\d{2}\)?\s*)?9?\d{4}[-.\s]?\d{4}\b/g, "[telefone removido]")
    .trim()
    .slice(0, maxLength);
}

function firstName(value) {
  return cleanText(value, 100).split(/\s+/)[0] || "Contato";
}

export function openAIClientContext(client) {
  if (!client) return null;
  return {
    name: firstName(client.name),
    stage: cleanText(client.stage, 40),
    temperature: cleanText(client.temperature, 30),
    priority: cleanText(client.priority, 30),
    profession: cleanText(client.profession, 100),
    incomeRange: cleanText(client.incomeRange, 80),
    maritalStatus: cleanText(client.maritalStatus, 80),
    children: Number(client.children || 0),
    city: cleanText(client.city, 100),
    state: cleanText(client.state, 2),
    source: cleanText(client.source, 100),
    referrer: cleanText(client.referrer, 100),
    lastInteraction: cleanText(client.lastInteraction, 10),
    followUp: cleanText(client.followUp, 10),
    taDate: cleanText(client.taDate, 10),
    oiScheduledDate: cleanText(client.oiScheduledDate, 10),
    oiTime: cleanText(client.oiTime, 5),
    oiDone: Boolean(client.oiDone),
    oiDate: cleanText(client.oiDate, 10),
    pcScheduledDate: cleanText(client.pcScheduledDate, 10),
    pcTime: cleanText(client.pcTime, 5),
    pcDone: Boolean(client.pcDone),
    pcDate: cleanText(client.pcDate, 10),
    outcome: cleanText(client.outcome, 20),
    planStatus: cleanText(client.planStatus, 80),
    planDueDate: cleanText(client.planDueDate, 10),
    c2NextAction: cleanText(client.c2NextAction, 300),
    nextAction: cleanText(client.nextAction, 300),
    notes: cleanText(client.notes, 1200),
    value: Number(client.value || 0),
    recs: Number(client.recs || 0),
    createdAt: cleanText(client.createdAt, 30)
  };
}

export async function askOpenAI({ mode, client, tone = "warm", question = "", history = [], metrics = null }) {
  const safeHistory = history.slice(-MAX_HISTORY_ITEMS).map(item => ({
    role: item.role === "assistant" ? "assistant" : "user",
    content: cleanText(item.content, MAX_HISTORY_TEXT)
  }));

  const data = await invokeCloudFunction("quick-endpoint", {
    mode,
    tone,
    question: cleanText(question, 1000),
    client: openAIClientContext(client),
    history: safeHistory,
    metrics
  });

  if (!data?.text) throw new Error("A IA nao retornou uma resposta.");
  return {
    text: String(data.text).trim(),
    model: data.model || "OpenAI",
    usage: data.usage || null
  };
}

export async function extractOIRecommendations({ fileName, fileData, extractedText = "" }) {
  const data = await invokeCloudFunction("quick-endpoint", {
    mode: "recommendations_pdf",
    fileName: cleanText(fileName, 180),
    fileData,
    extractedText: cleanText(extractedText, 12000)
  });
  return {
    recommendations: Array.isArray(data.recommendations) ? data.recommendations : [],
    model: data.model || "OpenAI",
    usage: data.usage || null
  };
}

export async function extractPostOiDocument({ fileName, fileData, extractedText = "" }) {
  const data = await invokeCloudFunction("quick-endpoint", {
    mode: "post_oi_pdf",
    fileName: cleanText(fileName, 180),
    fileData,
    extractedText: cleanText(extractedText, 12000)
  });
  return {
    diagnosis: data.diagnosis || {},
    recommendations: Array.isArray(data.recommendations) ? data.recommendations : [],
    model: data.model || "OpenAI",
    usage: data.usage || null
  };
}

export async function requestDailyReview(summary) {
  return askOpenAI({
    mode: "daily_review",
    question: JSON.stringify(summary).slice(0, 5000)
  });
}
